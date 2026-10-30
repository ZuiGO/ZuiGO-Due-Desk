"""Dues Desk demo app. Upload an Excel of dealer dues plus the invoice PDFs. The app extracts and validates the
data, matches each invoice to its dealer, then sends WhatsApp / SMS / email reminders (invoice attached to the
email), follows up on a configurable schedule, reads replies, handles promises and instalments, stops on payment.

DEMO MODE (default): every message goes to YOUR phone and email (MY_WHATSAPP, MY_SMS, MY_EMAIL), never to the
dealers in the sheet. Set DEMO_MODE=0 only for real use, after the dealers have agreed to receive messages.

Run:  pip install -r requirements.txt   then   uvicorn app:app --port 8000   and open http://localhost:8000
Credentials go in a .env file next to this script (names are listed at the top of the chat guide).
"""
import asyncio, hashlib, hmac, io, json, os, re, smtplib, time
from contextlib import asynccontextmanager
from collections import Counter
from datetime import date, datetime, timedelta
from email.message import EmailMessage
from pathlib import Path
from urllib.parse import urlencode

import httpx, openpyxl
from fastapi import FastAPI, File, HTTPException, Request, UploadFile
from fastapi.responses import FileResponse, HTMLResponse, Response
from pypdf import PdfReader

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

E = os.environ.get
HERE = Path(__file__).parent
UP = HERE / "uploads"
UP.mkdir(exist_ok=True)
DEMO = E("DEMO_MODE", "1") != "0"
BASE = E("BASE_URL", "http://localhost:8000").rstrip("/")
# Policy. "day" = seconds one policy day lasts in this demo (40 s: a 3-day follow-up gap becomes 2 minutes).
P = dict(pre=3, gap=3, max=4, wa=True, sms=True, em=True, tg=True, inst=True, maxi=4, day=40)
DEALERS, LOG, LAST, META = {}, [], {"id": None}, {}
ACTIVE = ("scheduled", "reminded", "promised", "planned")
TONES = ["Friendly", "Polite follow-up", "Firm", "Final notice"]
LANGUAGES = {"en": "English", "hi": "Hindi", "mr": "Marathi"}
TXT = [
    "Hello {name}, I hope business at {shop} is going well. Our records show {amt} pending for invoice {inv}. Due date noted: {when}. When convenient, could you please review the details here: {link}? If you have already paid, please disregard this message. If you would like any help or want to share a suitable payment date, just reply. We truly appreciate your business.",
    "Hello {name}, a gentle follow-up about invoice {inv}. Our records show {amt} pending. Due date noted: {when}. Whenever convenient, you can review the details here: {link}. If payment is already on its way or you would like to discuss a suitable date, please let us know. We truly appreciate your continued business.",
    "Hello {name}, we wanted to kindly check in about invoice {inv} for {amt}. Due date noted: {when}. When you have a moment, please review it here: {link}. If anything needs clarification or you would like to share a convenient payment date, we are happy to help. Thank you for your understanding.",
    "Hello {name}, we are reaching out once more with a gentle note about invoice {inv} for {amt}. Due date noted: {when}. When convenient, please review the details here: {link}. If you have already arranged payment, need clarification, or would like to discuss timing, please reply whenever it suits you. We value your relationship and appreciate your time.",
]
PROMPT = ('You read a dealer’s WhatsApp reply to a payment reminder. Reply with JSON only: '
          '{"intent": "promise" | "instalment" | "claims_paid" | "dispute" | "opt_out" | "other", '
          '"days": <days until they will pay, or null>, "n": <number of instalments requested, or null>}')
inr = lambda n: f"₹{n:,}"

# ---------------------------------------------------------------- extraction
SYN = {  # field -> accepted header names (order matters: shop before name, due_date before due)
    "cid": ["custid", "customer id", "cust id", "dealer id", "dealer code", "customer code", "client id", "account id", "id"],
    "shop": ["business", "shop", "shop name", "business name", "firm", "company", "company name", "store", "dealer company"],
    "name": ["customer name", "customer", "dealer name", "dealer", "contact person", "contact name", "customer/dealer name", "name", "owner"],
    "loc": ["location", "city", "town", "place", "area"],
    "total": ["total amount", "invoice amount", "bill amount", "grand total", "total"],
    "paid": ["paid amount", "amount paid", "amount received", "paid", "advance", "received"],
    "due_date": ["due date", "duedate", "payment due date", "payment date", "commitment date", "promised date", "settle date"],
    "due": ["amount due", "due amount", "balance due", "pending amount", "outstanding amount", "outstanding", "balance", "due", "pending"],
    "inv": ["invoice no", "invoice number", "invoice id", "invoice", "inv no", "bill no", "bill number"],
    "phone": ["phone", "mobile", "whatsapp", "mobile number", "phone number", "contact number", "telephone"],
    "email": ["email", "e mail", "mail", "email id", "email address"],
}
norm = lambda s: re.sub(r"[^a-z0-9 ]", " ", str(s or "").lower()).strip()


def map_headers(headers):
    m = {}
    used = set()
    normalized = [norm(header) for header in headers]
    for k, synonyms in SYN.items():
        for synonym in synonyms:
            synonym_norm = norm(synonym)
            for i, header_norm in enumerate(normalized):
                if i not in used and header_norm == synonym_norm:
                    m[k] = i
                    used.add(i)
                    break
            if k in m:
                break
    for k, synonyms in SYN.items():
        if k in m:
            continue
        for synonym in synonyms:
            synonym_norm = norm(synonym)
            if len(synonym_norm) <= 3:
                continue
            for i, header_norm in enumerate(normalized):
                if i in used:
                    continue
                if k == "total" and re.search(r"\b(due|paid|pending|outstanding|balance|received)\b", header_norm):
                    continue
                if synonym_norm in header_norm:
                    m[k] = i
                    used.add(i)
                    break
            if k in m:
                break
    if "total" not in m:
        for i, header_norm in enumerate(normalized):
            if i not in used and header_norm == "amount":
                m["total"] = i
                break
    return m


def find_table(workbook):
    best = None
    for ws in workbook.worksheets:
        rows = list(ws.iter_rows(values_only=True))
        for header_index, header in enumerate(rows[:30]):
            mapped = map_headers(header)
            has_identity = "name" in mapped or "shop" in mapped
            has_financial = "total" in mapped or "due" in mapped
            if not has_identity or not (has_financial or "inv" in mapped):
                continue
            score = len(mapped) * 3 + int(has_identity) * 4 + int(has_financial) * 4 + int("inv" in mapped) * 2
            candidate = (score, -header_index, ws, rows, header_index, mapped)
            if best is None or candidate[:2] > best[:2]:
                best = candidate
    if best is None:
        raise ValueError("Could not find a header row. Include a customer/dealer name and an amount due or invoice column in the first 30 rows.")
    _, _, ws, rows, header_index, mapped = best
    return ws, rows, header_index, mapped


def num(v, unit="rupees"):
    if v is None or str(v).strip() == "":
        return None
    try:
        value = float(re.sub(r"[^\d.\-]", "", str(v)))
        return int(round(value * (100_000 if unit == "lakhs" else 1)))
    except ValueError:
        return None


def amount_unit(header):
    label = str(header or "").lower()
    if re.search(r"\b(lakh|lakhs|lac|lacs)\b", label):
        return "lakhs"
    if "₹" in label or re.search(r"\b(inr|rs|rupees?)\b", label):
        return "rupees"
    return E("EXCEL_AMOUNT_UNIT", "rupees").strip().lower()


def pdate(v):
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    for f in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%d %b %Y", "%d-%b-%Y", "%d/%m/%y"):
        try:
            return datetime.strptime(str(v).strip(), f).date()
        except ValueError:
            pass
    return None


def phone(v):
    s = re.sub(r"\D", "", str(int(v) if isinstance(v, float) else v or ""))
    if len(s) == 10:
        return "+91" + s
    return "+" + s if len(s) == 12 and s.startswith("91") else None


def parse_excel(data):
    workbook = openpyxl.load_workbook(io.BytesIO(data), data_only=True, read_only=True)
    ws, rows, header_index, hdr = find_table(workbook)
    header = rows[header_index]
    out = []
    for r in rows[header_index + 1:]:
        if not any(c not in (None, "") for c in r):
            continue
        g = lambda k: r[hdr[k]] if k in hdr and hdr[k] < len(r) else None
        if not any(str(g(k) or "").strip() for k in ("name", "shop", "cid", "inv")):
            continue
        if not g("cid") and not g("inv") and norm(g("name")) in {"total", "grand total", "subtotal", "sub total", "summary"}:
            continue
        total, paid, sheet = (
            num(g(key), amount_unit(header[hdr[key]]) if key in hdr else E("EXCEL_AMOUNT_UNIT", "rupees").strip().lower())
            for key in ("total", "paid", "due")
        )
        dealer_id = len(out) + 1
        d = dict(id=dealer_id, cid=str(g("cid") or f"ROW{dealer_id}").strip(), name=str(g("name") or "").strip(),
                 shop=str(g("shop") or g("name") or "").strip(), loc=str(g("loc") or "").strip(),
                 total=total or 0, paid=paid or 0, sheet_due=sheet, due_date=pdate(g("due_date")),
                 inv=str(g("inv") or "").strip(), phone=phone(g("phone")), email=str(g("email") or "").strip(),
                 pdf=None, flags=[], language="en")
        calc = d["total"] - d["paid"]
        d["due"] = max(sheet, 0) if sheet is not None else (max(calc, 0) if total is not None else 0)
        if total is None:
            d["flags"].append("Total amount missing")
        elif sheet is not None and sheet != calc:
            d["flags"].append(f"Sheet due amount is {inr(sheet)} but total − paid = {inr(calc)}. Using the sheet due amount")
        if not d["due_date"]:
            d["flags"].append("Due date missing or unreadable")
        if not d["phone"] and not d["email"]:
            d["flags"].append("No phone or email")
        elif not d["phone"]:
            d["flags"].append("Phone missing or invalid")
        if not d["shop"] and not d["name"]:
            d["flags"].append("Customer/dealer name missing")
        d.update(bal=d["due"], cur=d["due"], att=0, next_at=None, plan=None, k=0,
                 status="skipped" if d["due"] <= 0 else "ready")
        out.append(d)
    mapped = [[str(header[i]), k] for k, i in hdr.items()]
    return out, mapped, {"sheet": ws.title, "header_row": header_index + 1, "dealers": len(out)}


def match_invoices(pdfs):
    for name, data in pdfs:
        try:
            text = PdfReader(io.BytesIO(data)).pages[0].extract_text() or ""
        except Exception:
            text = ""
        hay = (name + " " + text).lower()
        for d in DEALERS.values():
            if d["pdf"]:
                continue
            if (d["inv"] and d["inv"].lower() in hay) or d["cid"].lower() in hay:
                path = UP / f"{d['id']}_{Path(name).name}"
                path.write_bytes(data)
                d["pdf"] = str(path)
                if not d["inv"]:
                    m = re.search(r"INV[-\s]?\d+", text + " " + name, re.I)
                    d["inv"] = m.group().upper() if m else ""
                break
    for d in DEALERS.values():
        if not d["pdf"]:
            d["flags"].append("No invoice PDF matched")
        d["inv"] = d["inv"] or f"INV-{d['cid']}"


def load(rows, mapped, pdfs):
    DEALERS.clear()
    LOG.clear()
    DEALERS.update({d["id"]: d for d in rows})
    META["mapped"] = mapped
    match_invoices(pdfs)


# ---------------------------------------------------------------- sample data
SAMPLE = [  # name, cid, shop, location, total, paid, due as typed in the sheet, days to due date, invoice, mobile, email
    ("Rajesh Sharma", "C1001", "Sharma Traders", "Nashik", 250000, 150000, 100000, -3, "INV-2041", "5550000101", "sharma@example.com"),
    ("Suresh Patil", "C1002", "Patil Agencies", "Pune", 180000, 80000, 100000, -8, "INV-2038", "5550000102", "patil@example.com"),
    ("Anil Gupta", "C1003", "Gupta Hardware", "Nagpur", 95000, 50000, 45000, -2, "INV-2044", "5550000103", "gupta@example.com"),
    ("Imran Mehta", "C1004", "Mehta Distributors", "Mumbai", 320000, 120000, 200000, -13, "INV-2029", "5550000104", "mehta@example.com"),
    ("Vikas Kulkarni", "C1005", "Kulkarni Enterprises", "Kolhapur", 120000, 45000, 80000, -4, "INV-2040", "5550000105", "kulkarni@example.com"),
    ("Prakash Joshi", "C1006", "Joshi Motors", "Aurangabad", 150000, 90000, 60000, 3, "INV-2047", "5550000106", "joshi@example.com"),
]


def invoice_pdf(inv, shop, name, total, paid, due):
    from reportlab.pdfgen import canvas
    b = io.BytesIO()
    c = canvas.Canvas(b)
    y = 800
    for line in [f"TAX INVOICE {inv}", "Acme Distributors", f"Billed to: {shop} ({name})", f"Total amount: Rs.{total:,}",
                 f"Paid at purchase: Rs.{paid:,}", f"Balance due: Rs.{total - paid:,}", f"Due date: {due:%d %b %Y}"]:
        c.drawString(72, y, line)
        y -= 24
    c.save()
    return b.getvalue()


def make_sample():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Customer Name", "CustID", "Business/ShopName", "Location", "Total Amount (₹)", "Paid Amount (₹)", "Due Amount (₹)",
               "Due Date", "Invoice No", "Mobile", "Email"])
    pdfs = []
    for n, c, s, l, t, p, sd, dd, inv, ph, em in SAMPLE:
        due = date.today() + timedelta(days=dd)
        ws.append([n, c, s, l, t, p, sd, due, inv, ph, em])
        pdfs.append((f"{inv}.pdf", invoice_pdf(inv, s, n, t, p, due)))
    b = io.BytesIO()
    wb.save(b)
    return b.getvalue(), pdfs


# ---------------------------------------------------------------- sending
def mail(to, subject, body, pdf):
    m = EmailMessage()
    m["From"], m["To"], m["Subject"] = E("SMTP_USER"), to, subject
    m.set_content(body)
    if pdf and Path(pdf).exists():
        m.add_attachment(Path(pdf).read_bytes(), maintype="application", subtype="pdf", filename=Path(pdf).name.split("_", 1)[-1])
    port = toint(E("SMTP_PORT"), 465)
    use_ssl = E("SMTP_USE_SSL", "1" if port == 465 else "0") == "1"
    smtp = smtplib.SMTP_SSL if use_ssl else smtplib.SMTP
    with smtp(E("SMTP_HOST", "smtp.gmail.com"), port, timeout=20) as s:
        if not use_ssl:
            s.starttls()
        s.login(E("SMTP_USER"), E("SMTP_PASS"))
        s.send_message(m)


async def send_all(d, m, attach=True):
    """Send on every enabled and configured channel. In demo mode the recipient is always you."""
    def to(k, mine):
        return E(mine) if DEMO else d[k]
    out = []
    async with httpx.AsyncClient(timeout=20) as c:
        if P["wa"] and to("phone", "MY_WHATSAPP"):
            try:
                provider = E("WHATSAPP_PROVIDER", "meta").strip().lower()
                if provider == "aisensy":
                    if d.get("language", "en") != "en":
                        result = "FAILED AiSensy sends its approved campaign template, not this translated message; use manual WhatsApp or a free-text provider"
                    else:
                        result = await aisensy_send_campaign(c, to("phone", "MY_WHATSAPP"), d, m)
                elif provider == "meta":
                    result = await meta_send_text(c, to("phone", "MY_WHATSAPP"), m["wa"])
                else:
                    result = f"FAILED unsupported WHATSAPP_PROVIDER: {provider}"
                out.append(["WhatsApp", result])
            except Exception as e:
                out.append(["WhatsApp", f"FAILED {e}"])
        if m.get("sms") and P["sms"] and E("SMS_GATEWAY_URL") and to("phone", "MY_SMS"):
            try:
                sms_text = m["sms"]
                if len(sms_text) > 300:
                    out.append(["SMS", "FAILED generated reminder exceeds the 300-character gateway limit"])
                else:
                    r = await c.post(E("SMS_GATEWAY_URL").rstrip("/") + "/message",
                                     auth=(E("SMS_GATEWAY_USER", ""), E("SMS_GATEWAY_PASS", "")),
                                     json={"textMessage": {"text": sms_text}, "phoneNumbers": [to("phone", "MY_SMS")]})
                    out.append(["SMS", "sent" if r.status_code < 300 else f"FAILED HTTP {r.status_code}: {r.text[:200]}"])
            except Exception as e:
                out.append(["SMS", f"FAILED {e}"])
        if P["tg"] and E("TELEGRAM_BOT_TOKEN") and E("TELEGRAM_CHAT_ID"):
            try:
                result = await telegram_send(c, E("TELEGRAM_CHAT_ID"), m["wa"])
                out.append(["Telegram", result])
            except Exception as e:
                out.append(["Telegram", f"FAILED connection error ({type(e).__name__})"])
    if P["em"] and E("SMTP_USER") and E("SMTP_PASS") and to("email", "MY_EMAIL"):
        try:
            await asyncio.to_thread(mail, to("email", "MY_EMAIL"), m["subject"], m["body"], d["pdf"] if attach else None)
            out.append(["Email", "sent" + (" with invoice attached" if attach and d["pdf"] else "")])
        except Exception as e:
            out.append(["Email", f"FAILED {e}"])
    return out or [["Preview only", "no channel is configured yet, so nothing was sent"]]


async def telegram_send(client, chat_id, text):
    token = E("TELEGRAM_BOT_TOKEN")
    if not token or not chat_id:
        return "FAILED Telegram requires TELEGRAM_BOT_TOKEN and TELEGRAM_CHAT_ID"
    r = await client.post(
        f"https://api.telegram.org/bot{token}/sendMessage",
        json={"chat_id": chat_id, "text": text},
    )
    if r.status_code >= 300:
        return f"FAILED HTTP {r.status_code}: {r.text[:200]}"
    try:
        response = r.json()
    except ValueError:
        return "FAILED Telegram returned an invalid response"
    return "sent" if response.get("ok") else f"FAILED Telegram API: {json.dumps(response)[:200]}"


async def aisensy_send_campaign(client, recipient, dealer, message):
    api_key = E("AISENSY_API_KEY")
    campaign = E("AISENSY_CAMPAIGN_NAME")
    phone = re.sub(r"\D", "", str(recipient or ""))
    if not api_key or not campaign or not phone:
        return "FAILED AiSensy requires AISENSY_API_KEY, AISENSY_CAMPAIGN_NAME, and a valid recipient phone"
    r = await client.post(
        "https://backend.aisensy.com/campaign/t1/api/v2",
        json={
            "apiKey": api_key,
            "campaignName": campaign,
            "destination": phone,
            "userName": dealer.get("name") or dealer["shop"],
            "templateParams": list(message["wa_template_params"]),
            "source": "Dues Desk",
        },
    )
    if r.status_code < 300:
        return "sent"
    return f"FAILED HTTP {r.status_code}: {r.text[:200]}"


async def meta_send_text(client, recipient, body):
    phone = re.sub(r"\D", "", str(recipient or ""))
    if not phone or not E("META_WA_PHONE_NUMBER_ID") or not E("META_WA_ACCESS_TOKEN") or not E("META_GRAPH_API_VERSION"):
        return "FAILED Meta WhatsApp Cloud API is not fully configured"
    r = await client.post(
        f"https://graph.facebook.com/{E('META_GRAPH_API_VERSION')}/{E('META_WA_PHONE_NUMBER_ID')}/messages",
        headers={"Authorization": f"Bearer {E('META_WA_ACCESS_TOKEN')}"},
        json={
            "messaging_product": "whatsapp",
            "recipient_type": "individual",
            "to": phone,
            "type": "text",
            "text": {"preview_url": False, "body": body},
        },
    )
    if r.status_code < 300:
        return "sent"
    return f"FAILED HTTP {r.status_code}: {r.text[:200]}"


def compose(d, att):
    a = min(att - 1, 3)
    inv = d["inv"] + (f" (instalment {d['k'] + 1} of {len(d['plan'])})" if d["plan"] else "")
    when = reminder_date(d, "en")
    link = f"{BASE}/pay/{d['id']}"
    wa = TXT[a].format(name=d["name"] or d["shop"], shop=d["shop"], amt=inr(d["cur"]), inv=inv, when=when, link=link)
    return dict(
        wa=wa,
        wa_template_params=[
            d["name"] or d["shop"],
            inv,
            f"is still pending (due {when})",
            inr(d["cur"]),
            link,
        ],
        wa_variables={
            "1": d["name"] or d["shop"],
            "2": inv,
            "3": f"is still pending (due {when})",
            "4": inr(d["cur"]),
            "5": link,
        },
        sms=f"Dear {d['shop']}, Rs.{d['cur']:,} is {'overdue' if att > 1 else 'due'} for {d['inv']}. Pay: {link}",
        subject=f"{['Friendly reminder', 'Follow-up', 'Payment overdue', 'Final reminder'][a]}: invoice {d['inv']}",
        body=wa + ("\n\nThe invoice is attached." if d["pdf"] else "") + "\n\nRegards,\nAccounts team, Acme Distributors")


async def compose_message(d, att):
    language = d.get("language", "en")
    if language not in LANGUAGES:
        raise ValueError(f"Unsupported reminder language: {language}")
    api_key = E("GROQ_API_KEY")
    if not api_key:
        raise RuntimeError("GROQ_API_KEY is not configured")

    a = min(att - 1, len(TXT) - 1)
    sms_template = "Hello {shop}, invoice {inv} shows {amt} pending. Due date noted: {when}. When convenient, please review: {link}. Thank you for your business."
    prompt = (
        f"Write BOTH reminder templates entirely in {LANGUAGES[language]} using natural, polished language "
        f"appropriate for a business customer in India. For Marathi, use natural Marathi grammar and vocabulary "
        "(not Hindi); for Hindi, use natural Hindi grammar and vocabulary. Use the language's normal script. "
        "Keep the voice warm, humble, appreciative, and "
        "gentle, as a trusted small-business partner writing personally. Show appreciation for the customer's "
        "continued business and acknowledge they may be busy. Make the payment request clear but low-pressure. "
        "Never sound stern, impatient, legalistic, sarcastic, or guilt-inducing. Do not use words equivalent to "
        "'final notice', 'clear immediately', 'failure', 'overdue warning', or threats of escalation. "
        "Acknowledge that payment may already have been made, and invite the customer to reply if they need "
        "clarification or want to share a convenient payment date. Keep WhatsApp around 45-75 words and SMS "
        "concise. Translate every phrase and label, including 'Due date noted:', into the selected language. "
        "Do not leave English words in Hindi or Marathi messages; only brace-delimited placeholders may remain "
        "in English. Translate/localize the template text while preserving every brace-delimited placeholder "
        "exactly, including its spelling and capitalization. Do not translate, alter, omit, or invent any "
        "placeholder. Text after 'Due date noted:' is a value, not a sentence; preserve that simple label/value "
        "structure naturally in the target language. Do not add or change amounts, dates, invoice IDs, or payment instructions. "
        "Return only json: one valid JSON object with exactly two string fields named whatsapp and sms. "
        'Use this shape: {"whatsapp":"your message with placeholders","sms":"short message with placeholders"}. '
        f"WhatsApp template: {TXT[a]}\nSMS template: {sms_template}"
    )
    try:
        async with httpx.AsyncClient(timeout=35) as client:
            response = await client.post(
                "https://api.groq.com/openai/v1/chat/completions",
                headers={"Authorization": f"Bearer {api_key}"},
                json={
                    "model": E("GROQ_MODEL", "openai/gpt-oss-120b"),
                    "temperature": 0.2,
                    "max_completion_tokens": 1200,
                    "reasoning_effort": "low",
                    "response_format": {"type": "json_object"},
                    "messages": [
                        {"role": "system", "content": "You write kind, clear, culturally natural customer payment reminders."},
                        {"role": "user", "content": prompt},
                    ],
                },
            )
    except httpx.HTTPError as error:
        raise RuntimeError(f"Groq request failed ({type(error).__name__})") from error
    if response.status_code >= 300:
        raise RuntimeError(f"Groq returned HTTP {response.status_code}: {response.text[:250]}")
    try:
        content = response.json()["choices"][0]["message"]["content"]
        if not isinstance(content, str):
            raise TypeError("response content is not text")
        match = re.search(r"\{.*\}", content, re.S)
        translated = json.loads(match.group() if match else content)
        if not isinstance(translated, dict):
            raise TypeError("response JSON is not an object")
    except (ValueError, KeyError, IndexError, TypeError) as error:
        raise RuntimeError("Groq returned an invalid message response") from error

    templates = {"whatsapp": TXT[a], "sms": sms_template}
    for key, source in templates.items():
        translated_template = translated.get(key)
        if not isinstance(translated_template, str):
            raise RuntimeError(f"Groq response is missing its {key} template")
        if Counter(re.findall(r"\{[a-z_]+\}", translated_template)) != Counter(re.findall(r"\{[a-z_]+\}", source)):
            raise RuntimeError(f"Groq changed placeholders in the {key} template; message was not sent")
        devanagari = len(re.findall(r"[\u0900-\u097f]", translated_template))
        latin = len(re.findall(r"[A-Za-z]", translated_template))
        if language == "en" and devanagari:
            raise RuntimeError(f"Groq did not return the reminder in {LANGUAGES[language]}; message was not sent")
        if language in ("hi", "mr") and devanagari < 8:
            raise RuntimeError(f"Groq did not return the reminder in {LANGUAGES[language]} script; message was not sent")
        if language in ("hi", "mr") and re.search(r"[A-Za-z]", re.sub(r"\{[a-z_]+\}", "", translated_template)):
            raise RuntimeError(f"Groq left English wording in the {LANGUAGES[language]} reminder; message was not sent")
        if language == "en" and latin < 20:
            raise RuntimeError(f"Groq returned too little English reminder text; message was not sent")

    message = compose(d, att)
    inv = d["inv"] + (f" (instalment {d['k'] + 1} of {len(d['plan'])})" if d["plan"] else "")
    when = reminder_date(d, language)
    values = {
        "name": d["name"] or d["shop"],
        "shop": d["shop"],
        "amt": inr(d["cur"]),
        "inv": inv,
        "when": when,
        "link": f"{BASE}/pay/{d['id']}",
    }
    message["wa"] = translated["whatsapp"].format(**values)
    message["sms"] = translated["sms"].format(**values)
    if d["pdf"]:
        attachment_note = {
            "en": "\n\nInvoice attached.",
            "hi": "\n\nइनवॉइस संलग्न है।",
            "mr": "\n\nचलन संलग्न केले आहे.",
        }
        message["body"] = message["wa"] + attachment_note[language]
    else:
        message["body"] = message["wa"]
    return message


def reminder_date(d, language):
    if d.get("plan"):
        return {"en": "as agreed", "hi": "हमारी सहमति के अनुसार", "mr": "आपल्या सहमतीनुसार"}[language]
    if d.get("due_date"):
        return d["due_date"].strftime("%d %b" if language == "en" else "%d-%m-%Y")
    return {
        "en": "not recorded",
        "hi": "दर्ज नहीं",
        "mr": "नोंद नाही",
    }[language]


def log(d, kind, text, **x):
    LOG.append(dict(t=time.strftime("%H:%M:%S"), id=d["id"], shop=d["shop"], kind=kind, text=text, **x))
    del LOG[:-300]


# ---------------------------------------------------------------- the agent
async def remind(d):
    try:
        m = await compose_message(d, d["att"] + 1)
    except (RuntimeError, ValueError) as error:
        d.update(status="delivery_failed", next_at=None)
        log(d, "delivery_failed", f"Reminder text could not be generated; no attempt was counted: {error}",
            ch=[["Groq", f"FAILED {error}"]])
        return
    res = await send_all(d, m)
    if not any(result.startswith("sent") for _, result in res):
        d.update(status="delivery_failed", next_at=None)
        log(d, "delivery_failed", f"Reminder {d['att'] + 1} was not delivered; no attempt was counted", ch=res, msg=m["wa"])
        return
    d["att"] += 1
    LAST["id"] = d["id"]
    d["status"] = "reminded"
    log(d, "reminder", f"Reminder {d['att']} of {P['max']} ({TONES[min(d['att'] - 1, 3)].lower()}) for {inr(d['cur'])}", ch=res, msg=m["wa"])
    if d["att"] >= P["max"]:
        d.update(status="escalated", next_at=None)
        log(d, "escalate", f"Escalated to the account manager: {inr(d['bal'])} still unpaid after {P['max']} reminders")
    else:
        d["next_at"] = time.time() + P["gap"] * P["day"]


async def loop():
    while True:
        await asyncio.sleep(2)
        for d in list(DEALERS.values()):
            if d["status"] in ACTIVE and d["next_at"] and time.time() >= d["next_at"]:
                d["next_at"] = None
                await remind(d)


def toint(v, default):
    try:
        return int(v)
    except (TypeError, ValueError):
        return default


def classify(t):
    if E("ANTHROPIC_API_KEY"):
        try:
            import anthropic
            r = anthropic.Anthropic().messages.create(model=E("CLAUDE_MODEL", "claude-sonnet-5-5"), max_tokens=100,
                                                      system=PROMPT, messages=[{"role": "user", "content": t}])
            return json.loads(re.search(r"\{.*\}", r.content[0].text, re.S).group())
        except Exception as e:
            print(f"Claude unavailable ({e}), using keyword rules")
    l = t.lower()
    n, d = re.search(r"(\d+)\s*(instal|emi|part)", l), re.search(r"(\d+)\s*day", l)
    if re.search(r"instal|emi\b", l):
        return {"intent": "instalment", "n": int(n.group(1)) if n else 3}
    if re.search(r"stop|unsubscribe", l):
        return {"intent": "opt_out"}
    if re.search(r"paid|transferred|done", l):
        return {"intent": "claims_paid"}
    if re.search(r"wrong|short|dispute|mistake|not received", l):
        return {"intent": "dispute"}
    if d or re.search(r"tomorrow|week|friday|monday", l):
        return {"intent": "promise", "days": int(d.group(1)) if d else (1 if "tomorrow" in l else 3)}
    return {"intent": "other"}


@asynccontextmanager
async def life(app):
    task = asyncio.create_task(loop())
    yield
    task.cancel()


app = FastAPI(lifespan=life)


def snap():
    now = time.time()
    ds = []
    for d in DEALERS.values():
        ds.append(dict(id=d["id"], cid=d["cid"], name=d["name"], shop=d["shop"], loc=d["loc"], due=d["due"], bal=d["bal"], cur=d["cur"],
                       due_date=d["due_date"].isoformat() if d["due_date"] else None, inv=d["inv"], has_pdf=bool(d["pdf"]),
                       flags=d["flags"], status=d["status"], att=d["att"], inst=f"{d['k'] + 1}/{len(d['plan'])}" if d["plan"] else "",
                       next_in=max(0, int(d["next_at"] - now)) if d["next_at"] else None,
                       language=d.get("language", "en")))
    return dict(dealers=ds, log=LOG[-80:][::-1], policy=P, demo=DEMO, mapped=META.get("mapped", []),
                extraction=META.get("extraction"),
                channels=dict(wa=bool(
                    E("MY_WHATSAPP")
                    and (
                        E("WHATSAPP_PROVIDER", "meta").strip().lower() == "aisensy"
                        and E("AISENSY_API_KEY")
                        and E("AISENSY_CAMPAIGN_NAME")
                        or E("WHATSAPP_PROVIDER", "meta").strip().lower() == "meta"
                        and E("META_WA_PHONE_NUMBER_ID")
                        and E("META_WA_ACCESS_TOKEN")
                        and E("META_GRAPH_API_VERSION")
                    )
                ), sms=bool(E("SMS_GATEWAY_URL") and E("MY_SMS")),
                              em=bool(E("SMTP_USER") and E("SMTP_PASS") and E("MY_EMAIL")),
                              tg=bool(E("TELEGRAM_BOT_TOKEN") and E("TELEGRAM_CHAT_ID"))))


@app.get("/")
def index():
    return FileResponse(HERE / "index.html")


@app.get("/sample.xlsx")
def sample_file():
    return Response(make_sample()[0], media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    headers={"Content-Disposition": "attachment; filename=sample_dealer_dues.xlsx"})


@app.get("/api/state")
def state():
    return snap()


@app.post("/api/upload")
async def upload(excel: UploadFile = File(...), invoices: list[UploadFile] = File(default=[])):
    try:
        rows, mapped, extraction = parse_excel(await excel.read())
    except (ValueError, KeyError, IndexError) as error:
        raise HTTPException(status_code=400, detail=f"Excel extraction failed: {error}") from error
    load(rows, mapped, [(f.filename, await f.read()) for f in invoices])
    META["extraction"] = extraction
    return snap()


@app.post("/api/sample")
def sample():
    xl, pdfs = make_sample()
    rows, mapped, extraction = parse_excel(xl)
    load(rows, mapped, pdfs)
    META["extraction"] = extraction
    return snap()


@app.post("/api/policy")
async def policy(request: Request):
    for k, v in (await request.json()).items():
        if k in P:
            P[k] = bool(v) if isinstance(P[k], bool) else max(toint(v, P[k]), 1)
    return P


@app.post("/api/language/{did}")
async def set_language(did: int, request: Request):
    d = DEALERS.get(did)
    if not d:
        raise HTTPException(status_code=404, detail="Dealer not found")
    language = (await request.json()).get("language")
    if language not in LANGUAGES:
        raise HTTPException(status_code=400, detail="Choose English, Hindi, or Marathi")
    d["language"] = language
    return {"id": did, "language": language}


@app.post("/api/manual-message/{did}")
async def manual_message(did: int, request: Request):
    d = DEALERS.get(did)
    if not d:
        raise HTTPException(status_code=404, detail="Dealer not found")
    payload = await request.json()
    language = payload.get("language", d.get("language", "en"))
    if language not in LANGUAGES:
        raise HTTPException(status_code=400, detail="Choose English, Hindi, or Marathi")
    d["language"] = language
    try:
        message = await compose_message(d, d["att"] + 1)
    except (RuntimeError, ValueError) as error:
        raise HTTPException(status_code=502, detail=f"Groq message generation failed: {error}") from error
    recipient = E("MY_WHATSAPP") if DEMO else d["phone"]
    digits = re.sub(r"\D", "", str(recipient or ""))
    text = message["wa"]
    return {
        "message": text,
        "whatsapp_url": f"https://wa.me/{digits}?{urlencode({'text': text})}" if 10 <= len(digits) <= 15 else None,
    }


@app.post("/api/start")
async def start(request: Request):
    payload = await request.json()
    n = 0
    languages = payload.get("languages", {})
    for i in payload.get("ids", []):
        d = DEALERS.get(i)
        if d and d["status"] == "ready":
            language = languages.get(str(i), languages.get(i, d.get("language", "en")))
            if language not in LANGUAGES:
                raise HTTPException(status_code=400, detail=f"Unsupported reminder language for dealer {i}")
            d["language"] = language
            days = (d["due_date"] - date.today()).days if d["due_date"] else 0
            wait = max(0, days - P["pre"]) * P["day"] + n * 12  # stagger so your phone isn't flooded
            d.update(status="scheduled", next_at=time.time() + wait)
            log(d, "scheduled", "Agent scheduled reminder 1 " + (f"in {wait} s (demo time)" if wait else "right away"))
            n += 1
    return snap()


@app.post("/api/send/{did}")
async def send_now(did: int, request: Request):
    d = DEALERS.get(did)
    if d and d["status"] not in ("paid", "skipped"):
        payload = await request.json()
        language = payload.get("language", d.get("language", "en"))
        if language not in LANGUAGES:
            raise HTTPException(status_code=400, detail="Choose English, Hindi, or Marathi")
        d["language"] = language
        d["next_at"] = None
        await remind(d)
    return snap()


@app.post("/api/reset")
def reset():
    DEALERS.clear()
    LOG.clear()
    return snap()


@app.get("/whatsapp")
async def verify_whatsapp(request: Request):
    query = request.query_params
    if query.get("hub.mode") == "subscribe" and hmac.compare_digest(
        query.get("hub.verify_token", ""), E("META_WA_VERIFY_TOKEN", "")
    ) and E("META_WA_VERIFY_TOKEN"):
        return Response(query.get("hub.challenge", ""), media_type="text/plain")
    return Response("Webhook verification failed", status_code=403)


async def process_whatsapp_message(body, frm):
    # Demo mode: all messages went to one phone, so a reply belongs to the dealer messaged last.
    d = DEALERS.get(LAST["id"]) if DEMO else next((x for x in DEALERS.values() if x["phone"] and re.sub(r"\D", "", x["phone"]) == frm), None)
    if not d or d["status"] in ("ready", "skipped"):
        return "Thanks! There is no active reminder for this number."
    c = await asyncio.to_thread(classify, body)
    i, now, day = c.get("intent", "other"), time.time(), P["day"]
    if i == "instalment" and P["inst"]:
        n = min(max(toint(c.get("n"), 3), 2), P["maxi"])
        per = d["bal"] // n
        amts = [per] * (n - 1) + [d["bal"] - per * (n - 1)]
        d["plan"] = [dict(amt=a, day=1 + 2 * j, at=now + (1 + 2 * j) * day, paid=False) for j, a in enumerate(amts)]
        d.update(k=0, cur=amts[0], att=0, status="planned", next_at=d["plan"][0]["at"])
        r = f"Plan confirmed: {n} instalments of about {inr(per)}, due on {', '.join('day %d' % p['day'] for p in d['plan'])}. I’ll remind you before each one."
    elif i == "instalment":
        d.update(status="approval", next_at=None)
        r = "Thanks! Instalments need approval from our sales rep, who will get back to you."
    elif i == "promise":
        days = toint(c.get("days"), 2)
        d.update(status="promised", next_at=now + (days + 1) * day)
        r = f"Thank you! I’ve noted that you’ll pay within {days} day(s). I’ll check back after that."
    elif i == "claims_paid":
        d.update(status="verifying", next_at=None)
        r = "Thanks! Please send the UTR number or a screenshot. Our accounts team will verify and confirm."
    elif i == "dispute":
        d.update(status="disputed", next_at=None)
        r = "Sorry about that. Our sales rep will call you shortly to sort it out. Reminders are paused."
    elif i == "opt_out":
        d.update(status="opted_out", next_at=None)
        r = "Understood. No more reminders from us."
    else:
        r = "Thanks, I’ve passed your message to our team."
    return r


@app.post("/whatsapp")
async def inbound(request: Request):
    raw_body = await request.body()
    secret = E("META_WA_APP_SECRET")
    signature = request.headers.get("X-Hub-Signature-256", "")
    expected = "sha256=" + hmac.new(secret.encode(), raw_body, hashlib.sha256).hexdigest() if secret else ""
    if not secret or not hmac.compare_digest(signature, expected):
        return Response("Invalid webhook signature or missing app secret", status_code=403)
    try:
        payload = json.loads(raw_body)
    except json.JSONDecodeError:
        return Response("Invalid JSON", status_code=400)
    for entry in payload.get("entry", []):
        for change in entry.get("changes", []):
            for message in change.get("value", {}).get("messages", []):
                if message.get("type") != "text":
                    continue
                frm = re.sub(r"\D", "", str(message.get("from", "")))
                body = str(message.get("text", {}).get("body", ""))
                if not frm or not body:
                    continue
                reply = await process_whatsapp_message(body, frm)
                result = "FAILED Meta WhatsApp Cloud API is not fully configured"
                if E("META_WA_PHONE_NUMBER_ID") and E("META_WA_ACCESS_TOKEN") and E("META_GRAPH_API_VERSION"):
                    async with httpx.AsyncClient(timeout=20) as client:
                        result = await meta_send_text(client, frm, reply)
                dealer = DEALERS.get(LAST["id"]) if DEMO else next((x for x in DEALERS.values() if x["phone"] and re.sub(r"\D", "", x["phone"]) == frm), None)
                if dealer:
                    log(dealer, "reply", f"Replied “{body}”", ch=[["WhatsApp reply", result]], msg=reply)
    return {"status": "ok"}


@app.get("/pay/{did}", response_class=HTMLResponse)
async def pay(did: int):
    d = DEALERS.get(did)
    if not d or d["status"] in ("paid", "ready", "skipped"):
        return "<h2>Nothing is due right now</h2>"
    amt = d["cur"]
    d["bal"] -= amt
    if d["plan"]:
        d["plan"][d["k"]]["paid"] = True
    if d["plan"] and d["k"] + 1 < len(d["plan"]):
        d["k"] += 1
        p = d["plan"][d["k"]]
        d.update(cur=p["amt"], att=0, next_at=p["at"], status="planned")
        msg = f"Thank you! {inr(amt)} received. Balance {inr(d['bal'])}. Next instalment: {inr(p['amt'])} on day {p['day']}."
    else:
        d.update(status="paid", next_at=None, bal=0)
        msg = f"Thank you! {inr(amt)} received. Invoice {d['inv']} is fully paid."
    log(d, "payment", f"Payment of {inr(amt)} received through the pay link", msg=msg)
    await send_all(d, dict(
        wa=msg,
        wa_template_params=[
            d["name"] or d["shop"],
            d["inv"],
            "payment received; " + (f"balance {inr(d['bal'])}" if d["bal"] else "invoice fully paid"),
            inr(amt),
            f"{BASE}/pay/{d['id']}",
        ],
        wa_variables={
            "1": d["name"] or d["shop"],
            "2": d["inv"],
            "3": "payment received; " + (f"balance {inr(d['bal'])}" if d["bal"] else "invoice fully paid"),
            "4": inr(amt),
            "5": f"{BASE}/pay/{d['id']}",
        },
        sms=None,
        subject=f"Payment received: {d['inv']}",
        body=msg,
    ), attach=False)
    return f"<h2 style='font-family:sans-serif'>Demo payment of {inr(amt)} received. You can close this tab.</h2>"
