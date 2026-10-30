from datetime import date, datetime, timedelta, timezone
from io import BytesIO

from sqlalchemy import select

from saas_app import app
from models import ChannelSettings, Contact, Customer, Installment, Invoice, MessageTemplate, ReminderEvent, Tenant, User


def new_account(client, name, email):
    csrf = client.get("/auth/csrf").json()["csrf_token"]
    response = client.post(
        "/auth/signup",
        headers={"X-CSRF-Token": csrf},
        json={"tenant_name": name, "email": email, "password": "a-secure-passphrase-42"},
    )
    return response, client.get("/auth/csrf").json()["csrf_token"]


def test_signup_requires_csrf_and_creates_tenant_admin(client):
    assert client.get("/auth/status").json() == {"authenticated": False}
    response = client.post(
        "/auth/signup",
        json={"tenant_name": "Company", "email": "owner@example.com", "password": "a-secure-passphrase-42"},
    )
    assert response.status_code == 403

    response, csrf = new_account(client, "Company One", "owner@example.com")
    assert response.status_code == 200
    assert response.json()["user"]["role"] == "admin"
    status = client.get("/auth/status").json()
    assert status["authenticated"] is True
    assert status["tenant"]["name"] == "Company One"
    assert client.get("/auth/me").json()["tenant"]["name"] == "Company One"
    assert client.post("/auth/logout", headers={"X-CSRF-Token": csrf}).json() == {"ok": True}
    assert client.get("/api/customers").status_code == 401


def test_sample_workbook_shows_repeatable_installment_columns(client):
    import openpyxl

    response = client.get("/sample.xlsx")
    assert response.status_code == 200
    workbook = openpyxl.load_workbook(BytesIO(response.content), data_only=True)
    sheet = workbook.active
    headers = [cell.value for cell in sheet[1]]
    assert "Installment Date" in headers
    invoice_column = headers.index("Invoice Number") + 1
    installment_column = headers.index("Installment Date") + 1
    assert sheet.cell(2, invoice_column).value == sheet.cell(3, invoice_column).value
    assert sheet.cell(2, installment_column).value != sheet.cell(3, installment_column).value


def test_customers_and_invoices_are_isolated_by_tenant(client):
    response, csrf = new_account(client, "First Tenant", "first@example.com")
    assert response.status_code == 200
    created = client.post(
        "/api/customers",
        headers={"X-CSRF-Token": csrf},
        json={
            "name": "Asha",
            "contacts": [{"phone": "+919999999999", "is_primary": True}],
            "invoices": [{"invoice_number": "INV-1", "total_amount": "1000", "due_date": (date.today() + timedelta(days=8)).isoformat()}],
        },
    )
    assert created.status_code == 200, created.text
    assert created.json()["invoices"][0]["due_amount"] == "1000.00"

    client.post("/auth/logout", headers={"X-CSRF-Token": csrf})
    response, second_csrf = new_account(client, "Second Tenant", "second@example.com")
    assert response.status_code == 200
    assert client.get("/api/customers").json() == []
    assert client.patch(
        f"/api/customers/{created.json()['id']}/schedule",
        headers={"X-CSRF-Token": second_csrf},
        json={"reminder_frequency_days": 7, "reminder_dates": []},
    ).status_code == 404


def test_channel_secrets_are_encrypted_and_never_returned(client, signed_in):
    client, csrf = signed_in
    secret = "sample-provider-secret-do-not-return"
    response = client.put(
        "/api/settings/channels/telegram",
        headers={"X-CSRF-Token": csrf},
        json={"enabled": True, "config": {"chat_id": "12345"}, "secrets": {"bot_token": secret}},
    )
    assert response.status_code == 200
    assert secret not in response.text

    settings = client.get("/api/settings/channels").json()
    telegram = next(item for item in settings if item["channel"] == "telegram")
    assert telegram["enabled"] is True
    assert telegram["configured_secrets"] == {"bot_token": True}
    assert secret not in str(settings)

    with app.state.test_sessions() as db:
        record = db.scalar(select(ChannelSettings).where(ChannelSettings.channel == "telegram"))
        assert record is not None
        assert b"sample-provider-secret-do-not-return" not in record.encrypted_secrets


def test_templates_validate_placeholders_and_are_tenant_scoped(client, signed_in):
    client, csrf = signed_in
    bad = client.put(
        "/api/templates",
        headers={"X-CSRF-Token": csrf},
        json={"channel": "sms", "language": "en", "body": "Hello {customer_secret}"},
    )
    assert bad.status_code == 422

    saved = client.put(
        "/api/templates",
        headers={"X-CSRF-Token": csrf},
        json={"channel": "sms", "language": "hi", "body": "नमस्ते {name}, बिल {invoice}"},
    )
    assert saved.status_code == 200
    assert client.get("/api/templates").json()[0]["language"] == "hi"


def test_incomplete_excel_import_reports_rows_and_schedules_installment_reminders(client, signed_in):
    client, csrf = signed_in
    setting = client.put(
        "/api/settings/channels/sms",
        headers={"X-CSRF-Token": csrf},
        json={"enabled": True, "config": {"gateway_url": "https://sms.example.invalid/send"}, "secrets": {}},
    )
    assert setting.status_code == 200

    due = date.today() + timedelta(days=30)
    commitment = date.today() + timedelta(days=10)
    import openpyxl

    book = openpyxl.Workbook()
    sheet = book.active
    sheet.append(["Customer Name", "Customer ID", "Company", "Phone", "Invoice No", "Total Amount", "Due Date", "Installment Amount", "Installment Date", "Reminder Frequency Days", "WhatsApp Consent"])
    sheet.append(["Mira", "M-1", "Mira Stores", "+919800000001", "INV-42", 1200, due, 400, commitment, 6, "yes"])
    second_commitment = commitment + timedelta(days=10)
    sheet.append(["Mira", "M-1", "Mira Stores", "+919800000001", "INV-42", 1200, due, 400, second_commitment, 6, "yes"])
    sheet.append([None, None, None, None, "INV-SKIP", 100, due, None, None, None])
    file_data = BytesIO()
    book.save(file_data)
    file_data.seek(0)

    response = client.post(
        "/api/import",
        headers={"X-CSRF-Token": csrf},
        files={"workbook": ("dues.xlsx", file_data.getvalue(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
    )
    assert response.status_code == 200, response.text
    result = response.json()
    assert result["customers_created"] == 1
    assert result["invoices_created"] == 1
    assert result["rows_skipped"] == 1
    assert any("customer name" in item["message"].lower() for item in result["warnings"])

    customer = client.get("/api/customers").json()[0]
    assert customer["reminder_frequency_days"] == 6
    assert len(customer["invoices"][0]["installments"]) == 2
    assert customer["invoices"][0]["installments"][0]["reminder_days_before"] == [5, 3]
    assert customer["contacts"][0]["whatsapp_opted_in"] is True
    reminders = client.get("/api/reminders").json()
    scheduled_dates = {item["scheduled_at"][:10] for item in reminders}
    assert (commitment - timedelta(days=5)).isoformat() in scheduled_dates
    assert (commitment - timedelta(days=3)).isoformat() in scheduled_dates
    assert (second_commitment - timedelta(days=5)).isoformat() in scheduled_dates
    assert (second_commitment - timedelta(days=3)).isoformat() in scheduled_dates
    event_count = len(reminders)
    duplicate = client.post(
        "/api/import",
        headers={"X-CSRF-Token": csrf},
        files={"workbook": ("dues.xlsx", file_data.getvalue(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
    )
    assert duplicate.status_code == 200, duplicate.text
    assert duplicate.json()["invoices_created"] == 0
    assert len(client.get("/api/reminders").json()) == event_count


def test_channel_test_uses_only_explicit_test_recipient(client, signed_in, monkeypatch):
    import asyncio

    import worker

    client, csrf = signed_in
    saved = client.put(
        "/api/settings/channels/telegram",
        headers={"X-CSRF-Token": csrf},
        json={
            "enabled": True,
            "config": {"chat_id": "workspace-chat", "test_recipient": "test-chat"},
            "secrets": {"bot_token": "private-token"},
        },
    )
    assert saved.status_code == 200
    calls = []

    async def record_send(*args, **kwargs):
        calls.append((args, kwargs))

    monkeypatch.setattr(worker, "_http_send", record_send)
    response = client.post("/api/settings/channels/telegram/test", headers={"X-CSRF-Token": csrf})
    assert response.status_code == 200, response.text
    assert response.json() == {"status": "sent", "channel": "telegram", "recipient": "test-chat"}
    assert calls[0][0][4] == "test-chat"
    assert "private-token" not in response.text


def test_foreign_tenant_rows_are_not_visible(client, signed_in):
    client, csrf = signed_in
    with app.state.test_sessions() as db:
        original_tenant_id = db.scalar(select(Tenant.id).where(Tenant.name == "Northwind Traders"))
    assert original_tenant_id is not None

    client.post("/auth/logout", headers={"X-CSRF-Token": csrf})
    response, second_csrf = new_account(client, "Second company", "other@example.com")
    assert response.status_code == 200
    assert client.get("/api/customers").json() == []
    with app.state.test_sessions() as db:
        customer = Customer(tenant_id=original_tenant_id, name="Tenant one's record")
        db.add(customer)
        db.commit()
        customer_id = customer.id
    assert client.patch(
        f"/api/customers/{customer_id}/schedule",
        headers={"X-CSRF-Token": second_csrf},
        json={"reminder_frequency_days": 4, "reminder_dates": []},
    ).status_code == 404


def test_worker_requires_live_mode_and_contact_consent(client, signed_in, monkeypatch):
    import asyncio

    import worker

    client, csrf = signed_in
    tenant_id = client.get("/auth/me").json()["tenant"]["id"]
    with app.state.test_sessions() as db:
        db.add(ChannelSettings(tenant_id=tenant_id, channel="sms", enabled=True, config={"gateway_url": "https://sms.example.invalid"}))
        customer = Customer(tenant_id=tenant_id, name="Consent check", language="en")
        db.add(customer)
        db.flush()
        db.add(Contact(tenant_id=tenant_id, customer_id=customer.id, phone="+919999999999", sms_opted_in=False))
        invoice = Invoice(
            tenant_id=tenant_id,
            customer_id=customer.id,
            invoice_number="CONSENT-1",
            due_amount="250.00",
            total_amount="250.00",
            paid_amount="0",
            due_date=date.today(),
            status="open",
        )
        db.add(invoice)
        db.flush()
        customer_id, invoice_id = customer.id, invoice.id
        event = ReminderEvent(
            tenant_id=tenant_id,
            invoice_id=invoice.id,
            channel="sms",
            scheduled_at=datetime.now(timezone.utc) - timedelta(minutes=1),
            idempotency_key="consent-check-event",
        )
        db.add(event)
        db.commit()
        event_id = event.id

    calls = []

    async def record_delivery(item):
        calls.append(item)

    monkeypatch.setattr(worker, "_deliver", record_delivery)
    asyncio.run(worker.dispatch_once(app.state.test_sessions))
    with app.state.test_sessions() as db:
        saved = db.get(ReminderEvent, event_id)
        assert saved.status == "failed"
        assert saved.failure_code == "live_sending_disabled"
        tenant = db.get(Tenant, tenant_id)
        tenant.live_sending_enabled = True
        contact = db.scalar(select(Contact).where(Contact.customer_id == customer_id))
        contact.sms_opted_in = True
        db.add(
            MessageTemplate(
                tenant_id=tenant_id,
                channel="sms",
                language="en",
                stage="reminder",
                body="Hello {name}, {invoice} has {amount} due on {due_date}.",
            )
        )
        next_event = ReminderEvent(
            tenant_id=tenant_id,
            invoice_id=invoice_id,
            channel="sms",
            scheduled_at=datetime.now(timezone.utc) - timedelta(minutes=1),
            idempotency_key="consent-check-event-2",
        )
        db.add(next_event)
        db.commit()
        next_event_id = next_event.id

    asyncio.run(worker.dispatch_once(app.state.test_sessions))
    assert len(calls) == 1
    assert calls[0]["recipient"] == "+919999999999"
    assert "Consent check" in calls[0]["message"]
    with app.state.test_sessions() as db:
        assert db.get(ReminderEvent, next_event_id).status == "sent"
    assert asyncio.run(worker.dispatch_once(app.state.test_sessions)) == 0
