import asyncio
import ipaddress
import os
import re
import socket
import smtplib
from datetime import datetime, timedelta, timezone
from email.message import EmailMessage
from collections.abc import Callable
from typing import Any
from urllib.parse import urlparse

import httpx
from dotenv import load_dotenv
from sqlalchemy import select, update
from sqlalchemy.orm import Session, joinedload

load_dotenv(os.path.join(os.path.dirname(__file__), ".env"))

from database import SessionLocal
from models import (
    ChannelSettings,
    Contact,
    Customer,
    Installment,
    Invoice,
    MessageLog,
    MessageTemplate,
    ReminderEvent,
    Tenant,
    utcnow,
)
from security import decrypt_secrets


HTTP_TIMEOUT = 20
MAX_BATCH_SIZE = 50
CLAIM_TIMEOUT = timedelta(minutes=5)


def _failure(session: Session, event_id: int, code: str) -> None:
    event = session.get(ReminderEvent, event_id)
    if event is not None and event.status == "processing":
        event.status = "failed"
        event.failure_code = code[:80]
        event.claimed_at = None
        session.commit()


def _recipient(session: Session, tenant_id: int, customer_id: int, channel: str) -> str | None:
    contacts = session.scalars(
        select(Contact)
        .where(Contact.tenant_id == tenant_id, Contact.customer_id == customer_id)
        .order_by(Contact.is_primary.desc(), Contact.id)
    ).all()
    for contact in contacts:
        if channel == "sms" and contact.sms_opted_in and contact.phone:
            return contact.phone
        if channel == "whatsapp" and contact.whatsapp_opted_in and contact.phone:
            return contact.phone
        if channel == "email" and contact.email_opted_in and contact.email:
            return contact.email
        if channel == "telegram" and contact.telegram_chat_id:
            return contact.telegram_chat_id
    return None


def _render_template(template: str, values: dict[str, str]) -> str:
    try:
        return template.format_map(values)
    except (KeyError, ValueError) as error:
        raise ValueError("template_render_failed") from error


def _smtp_send(config: dict[str, Any], secrets: dict[str, str], recipient: str, subject: str, body: str) -> None:
    host = str(config.get("host") or "")
    port = int(config.get("port") or 465)
    sender = str(config.get("sender_email") or secrets.get("username") or "")
    username, password = secrets.get("username", ""), secrets.get("password", "")
    if not host or not sender or not username or not password or not 1 <= port <= 65535:
        raise ValueError("email_not_configured")
    _validate_public_host(host, "SMTP_ALLOWED_HOSTS")
    message = EmailMessage()
    message["From"], message["To"], message["Subject"] = sender, recipient, subject
    message.set_content(body)
    use_ssl = bool(config.get("use_ssl", port == 465))
    if use_ssl:
        with smtplib.SMTP_SSL(host, port, timeout=HTTP_TIMEOUT) as server:
            server.login(username, password)
            server.send_message(message)
    else:
        with smtplib.SMTP(host, port, timeout=HTTP_TIMEOUT) as server:
            server.starttls()
            server.login(username, password)
            server.send_message(message)


async def _http_send(
    channel: str,
    provider: str | None,
    config: dict[str, Any],
    secrets: dict[str, str],
    recipient: str,
    text: str,
    template_params: list[str] | None = None,
) -> None:
    async with httpx.AsyncClient(timeout=HTTP_TIMEOUT) as client:
        if channel == "sms":
            endpoint = str(config.get("gateway_url") or "").strip()
            parsed = urlparse(endpoint)
            if parsed.scheme != "https" or not parsed.hostname or parsed.username or parsed.password:
                raise ValueError("sms_gateway_url_invalid")
            await asyncio.to_thread(_validate_public_host, parsed.hostname, "SMS_GATEWAY_ALLOWED_HOSTS")
            credentials = secrets.get("api_key")
            if not credentials and not (secrets.get("username") and secrets.get("password")):
                raise ValueError("sms_credentials_missing")
            auth = None if credentials else (secrets.get("username", ""), secrets.get("password", ""))
            headers = {"Authorization": f"Bearer {credentials}"} if credentials else {}
            response = await client.post(
                endpoint.rstrip("/") + "/message",
                auth=auth,
                headers=headers,
                json={"textMessage": {"text": text}, "phoneNumbers": [re.sub(r"\D", "", recipient)]},
            )
        elif channel == "whatsapp" and provider == "meta":
            version = str(config.get("graph_api_version") or "").strip()
            phone_id = str(config.get("phone_number_id") or "").strip()
            template_name = str(config.get("template_name") or "").strip()
            template_language = str(config.get("template_language") or "").strip()
            access_token = secrets.get("access_token")
            if (
                not re.fullmatch(r"v\d+\.\d+", version)
                or not phone_id.isdigit()
                or not re.fullmatch(r"[a-z0-9_]+", template_name)
                or not re.fullmatch(r"[a-z]{2}(?:_[A-Z]{2})?", template_language)
                or not access_token
                or not template_params
            ):
                raise ValueError("whatsapp_meta_not_configured")
            response = await client.post(
                f"https://graph.facebook.com/{version}/{phone_id}/messages",
                headers={"Authorization": f"Bearer {access_token}"},
                json={
                    "messaging_product": "whatsapp",
                    "recipient_type": "individual",
                    "to": re.sub(r"\D", "", recipient),
                    "type": "template",
                    "template": {
                        "name": template_name,
                        "language": {"code": template_language},
                        "components": [
                            {
                                "type": "body",
                                "parameters": [{"type": "text", "text": value} for value in template_params],
                            }
                        ],
                    },
                },
            )
        elif channel == "whatsapp" and provider == "aisensy":
            api_key = secrets.get("api_key")
            campaign = str(config.get("campaign_name") or "")
            if not api_key or not campaign or not template_params:
                raise ValueError("whatsapp_aisensy_not_configured")
            response = await client.post(
                "https://backend.aisensy.com/campaign/t1/api/v2",
                json={
                    "apiKey": api_key,
                    "campaignName": campaign,
                    "destination": re.sub(r"\D", "", recipient),
                    "userName": template_params[0],
                    "templateParams": template_params,
                    "source": "Dues Desk",
                },
            )
        elif channel == "telegram":
            token = secrets.get("bot_token")
            if not token:
                raise ValueError("telegram_not_configured")
            response = await client.post(
                f"https://api.telegram.org/bot{token}/sendMessage",
                json={"chat_id": recipient, "text": text},
            )
            try:
                result = response.json()
            except ValueError as error:
                raise ValueError("telegram_provider_invalid_response") from error
            if not isinstance(result, dict) or not result.get("ok"):
                raise ValueError("telegram_provider_rejected_message")
        else:
            raise ValueError("unsupported_channel_provider")
        if response.status_code >= 300:
            raise ValueError(f"provider_http_{response.status_code}")


def _claim_due(db: Session, now: datetime) -> list[int]:
    rows = db.scalars(
        select(ReminderEvent)
        .where(
            ReminderEvent.status == "scheduled",
            ReminderEvent.scheduled_at <= now,
        )
        .order_by(ReminderEvent.scheduled_at, ReminderEvent.id)
        .limit(MAX_BATCH_SIZE)
        .with_for_update(skip_locked=True)
    ).all()
    ids = []
    for event in rows:
        event.status = "processing"
        event.claimed_at = now
        event.attempts += 1
        ids.append(event.id)
    db.commit()
    return ids


def _validate_public_host(host: str, allowlist_name: str) -> None:
    allowed_hosts = {
        value.strip().lower()
        for value in os.getenv(allowlist_name, "").split(",")
        if value.strip()
    }
    if os.getenv("APP_ENV") == "production" and host.lower() not in allowed_hosts:
        raise ValueError("outbound_host_not_allowlisted")
    try:
        addresses = socket.getaddrinfo(host, None, type=socket.SOCK_STREAM)
    except socket.gaierror as error:
        raise ValueError("outbound_host_unresolvable") from error
    if not addresses:
        raise ValueError("outbound_host_unresolvable")
    try:
        if any(not ipaddress.ip_address(address[4][0]).is_global for address in addresses):
            raise ValueError("outbound_host_resolves_to_non_public_address")
    except ValueError as error:
        if str(error).startswith("outbound_"):
            raise
        raise ValueError("outbound_host_invalid") from error


def _load_delivery(db: Session, event_id: int) -> dict[str, Any] | None:
    event = db.scalar(
        select(ReminderEvent)
        .options(
            joinedload(ReminderEvent.invoice)
            .joinedload(Invoice.customer)
            .selectinload(Customer.contacts),
            joinedload(ReminderEvent.invoice).selectinload(Invoice.installments),
        )
        .where(ReminderEvent.id == event_id)
    )
    if event is None or event.status != "processing":
        return None
    invoice = event.invoice
    customer = invoice.customer
    tenant = db.get(Tenant, event.tenant_id)
    setting = db.scalar(
        select(ChannelSettings).where(
            ChannelSettings.tenant_id == event.tenant_id,
            ChannelSettings.channel == event.channel,
        )
    )
    if tenant is None or setting is None or not setting.enabled or not tenant.live_sending_enabled:
        raise ValueError("live_sending_disabled")
    if invoice.status != "open" or invoice.due_amount is None or invoice.due_amount <= 0:
        raise ValueError("invoice_not_sendable")
    installment = None
    if event.installment_id is not None:
        installment = db.scalar(
            select(Installment).where(
                Installment.id == event.installment_id,
                Installment.tenant_id == event.tenant_id,
            )
        )
        if installment is None or installment.status != "planned":
            raise ValueError("installment_not_sendable")
    recipient = _recipient(db, event.tenant_id, customer.id, event.channel)
    if recipient is None:
        raise ValueError("recipient_missing_or_consent_not_recorded")
    language = customer.language if customer.language in {"en", "hi", "mr"} else "en"
    stage = "installment" if installment else "reminder"
    template = None
    template_scopes = [(language, stage)]
    if installment:
        template_scopes.append((language, "reminder"))
    if language != "en":
        template_scopes.append(("en", stage))
        if installment:
            template_scopes.append(("en", "reminder"))
    for template_language, template_stage in template_scopes:
        template = db.scalar(
            select(MessageTemplate).where(
                MessageTemplate.tenant_id == event.tenant_id,
                MessageTemplate.channel == event.channel,
                MessageTemplate.language == template_language,
                MessageTemplate.stage == template_stage,
            )
        )
        if template is not None:
            break
    if template is None:
        raise ValueError("message_template_missing")
    amount = installment.amount if installment else invoice.due_amount
    values = {
        "name": customer.name,
        "shop": customer.company or customer.name,
        "invoice": invoice.invoice_number,
        "amount": f"{amount:.2f}",
        "due_date": invoice.due_date.isoformat() if invoice.due_date else "not recorded",
        "installment_date": installment.due_date.isoformat() if installment else "",
        "payment_link": "",
    }
    template_params = [
        values["name"],
        values["invoice"],
        f"payment reminder; due {values['due_date']}",
        values["amount"],
        values["payment_link"],
    ]
    rendered_message = _render_template(template.body, values)
    history_message = rendered_message
    if event.channel == "whatsapp":
        provider_template = (
            (setting.config or {}).get("template_name")
            or (setting.config or {}).get("campaign_name")
            or "provider-approved template"
        )
        history_message = f"{provider_template}; parameters: {' | '.join(template_params)}"
    return {
        "channel": event.channel,
        "provider": setting.provider,
        "config": setting.config or {},
        "secrets": decrypt_secrets(setting.encrypted_secrets),
        "recipient": recipient,
        "subject": _render_template(template.subject or "Payment reminder", values),
        "message": rendered_message,
        "history_message": history_message,
        "template_params": template_params,
    }


async def _deliver(item: dict[str, Any]) -> None:
    if item["channel"] == "email":
        await asyncio.to_thread(
            _smtp_send,
            item["config"],
            item["secrets"],
            item["recipient"],
            item["subject"],
            item["message"],
        )
    else:
        await _http_send(
            item["channel"],
            item["provider"],
            item["config"],
            item["secrets"],
            item["recipient"],
            item["message"],
            item.get("template_params"),
        )


async def dispatch_once(session_factory: Callable[..., Session] = SessionLocal) -> int:
    with session_factory() as db:
        stale_before = datetime.now(timezone.utc) - CLAIM_TIMEOUT
        db.execute(
            update(ReminderEvent)
            .where(
                ReminderEvent.status == "processing",
                ReminderEvent.claimed_at < stale_before,
            )
            .values(status="needs_review", failure_code="worker_interrupted", claimed_at=None)
        )
        db.commit()
        event_ids = _claim_due(db, datetime.now(timezone.utc))

    for event_id in event_ids:
        payload = None
        try:
            with session_factory() as db:
                payload = _load_delivery(db, event_id)
            if payload is None:
                _failure_with_code(session_factory, event_id, "event_unavailable")
                _record_reminder_message(session_factory, event_id, None, "event_unavailable")
                continue
            await _deliver(payload)
        except (ValueError, RuntimeError, httpx.HTTPError, smtplib.SMTPException, OSError) as error:
            code = str(error) if isinstance(error, ValueError) else f"provider_{type(error).__name__.lower()}"
            _failure_with_code(session_factory, event_id, code)
            _record_reminder_message(session_factory, event_id, payload, code)
        else:
            with session_factory() as db:
                event = db.get(ReminderEvent, event_id)
                if event is not None and event.status == "processing":
                    event.status = "sent"
                    event.delivered_at = utcnow()
                    event.claimed_at = None
                    event.failure_code = None
                    _add_reminder_message(db, event, payload, "sent")
                    db.flush()
                    _flag_unresponsive_customer(db, event.tenant_id, event.invoice.customer_id)
                    db.commit()
    return len(event_ids)


def _failure_with_code(session_factory: Callable[..., Session], event_id: int, code: str) -> None:
    with session_factory() as db:
        _failure(db, event_id, code)


def _add_reminder_message(
    db: Session,
    event: ReminderEvent,
    payload: dict[str, Any] | None,
    status: str,
    failure_code: str | None = None,
) -> None:
    db.add(
        MessageLog(
            tenant_id=event.tenant_id,
            customer_id=event.invoice.customer_id,
            reminder_event_id=event.id,
            channel=event.channel,
            direction="reminder",
            status=status,
            recipient=payload.get("recipient") if payload else None,
            subject=payload.get("subject") if payload else None,
            body=payload.get("history_message", payload.get("message")) if payload else None,
            failure_code=failure_code[:200] if failure_code else None,
            sent_at=utcnow() if status == "sent" else None,
        )
    )


def _record_reminder_message(
    session_factory: Callable[..., Session],
    event_id: int,
    payload: dict[str, Any] | None,
    failure_code: str,
) -> None:
    with session_factory() as db:
        event = db.get(ReminderEvent, event_id)
        if event is not None:
            _add_reminder_message(db, event, payload, "failed", failure_code)
            db.commit()


def _flag_unresponsive_customer(db: Session, tenant_id: int, customer_id: int) -> None:
    customer = db.scalar(
        select(Customer).where(Customer.id == customer_id, Customer.tenant_id == tenant_id)
    )
    if customer is None or customer.risk_flagged:
        return
    today = utcnow().date()
    overdue_invoices = db.scalars(
        select(Invoice).where(
            Invoice.tenant_id == tenant_id,
            Invoice.customer_id == customer_id,
            Invoice.status == "open",
            Invoice.due_amount > 0,
            Invoice.due_date < today,
        )
    ).all()
    if not overdue_invoices:
        return
    invoice_ids = [invoice.id for invoice in overdue_invoices]
    deliveries = db.execute(
        select(ReminderEvent.delivered_at, Invoice.due_date)
        .join(Invoice, Invoice.id == ReminderEvent.invoice_id)
        .where(
            ReminderEvent.tenant_id == tenant_id,
            ReminderEvent.invoice_id.in_(invoice_ids),
            ReminderEvent.status == "sent",
            ReminderEvent.delivered_at.is_not(None),
        )
    ).all()
    overdue_delivery_days = {
        delivered_at.date()
        for delivered_at, due_date in deliveries
        if delivered_at is not None
        and due_date is not None
        and delivered_at.date() > due_date
        and (customer.risk_flagged_at is None or delivered_at > customer.risk_flagged_at)
    }
    if len(overdue_delivery_days) >= 3:
        customer.risk_flagged = True
        customer.risk_flag_reason = (
            "Automatically flagged after 3 delivered payment reminders on separate days "
            "while an invoice remained overdue."
        )
        customer.risk_flagged_at = utcnow()


async def run_worker() -> None:
    interval = max(1, int(os.getenv("REMINDER_WORKER_INTERVAL_SECONDS", "10")))
    while True:
        await dispatch_once()
        await asyncio.sleep(interval)


if __name__ == "__main__":
    asyncio.run(run_worker())
