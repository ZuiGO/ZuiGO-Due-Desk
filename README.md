# Dues Desk

Dues Desk is a receivables-reminder application. The hosted MVP entry point is `saas_app:app`; it supports tenant accounts, customer/contact/invoice records, messaging settings, editable templates, and reminder scheduling.

## Local development

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
Copy-Item .env.example .env
alembic upgrade head
python -m uvicorn saas_app:app --reload
```

Open http://127.0.0.1:8000 and create a company workspace. For automatic reminder delivery, run `python -m worker` in a second terminal from the project directory; the web app stores and schedules reminders, while the separate worker dispatches them. Customer sends remain disabled until live messaging is explicitly enabled and a channel is configured. Local development uses SQLite. Keep `.env`, `.credentials.key`, `.session.key`, and the local database private.

Run tests with `python -m pytest`.

The default SQLite file is resolved beside `database.py`, not relative to the shell's current directory. If `DATABASE_URL` is set, that configured database is used instead.

## Database tables

Apply migrations with `alembic upgrade head` before starting the app. A new database is created from the migration history; the ORM does not silently create or alter production tables at web-process startup.

- `tenants` and `users`: workspace ownership, account login, and workspace roles.
- `customers` and `contacts`: deduplicated customer identity, destinations, and recorded channel consent. Customer rows also retain the manual/automatic business-risk flag.
- `invoices` and `installments`: outstanding balances and any promised payment commitments.
- `reminder_policies` and `reminder_events`: workspace/customer reminder rules and the idempotent delivery queue.
- `channel_settings`: per-workspace channel configuration and encrypted provider credentials.
- `message_templates`: per-workspace, channel, language, and reminder-stage message content.
- `message_logs`: tenant-scoped history of manual and automated send attempts, including delivery status and rendered content.
- `alembic_version`: the migration revision installed in the selected database.

Excel imports match customers by customer ID, then normalized name/company, and match invoices by invoice number within the workspace. Re-imports update existing contact/invoice data and do not duplicate customers, invoices, or installment commitments. Repeated header lines and summary rows such as “Total” are ignored.

## Hosted deployment

Use PostgreSQL, HTTPS behind a trusted reverse proxy, and deployment-managed random values for `SESSION_SECRET_KEY` and `CREDENTIALS_ENCRYPTION_KEY`. Generate the latter with `python -c "from cryptography.fernet import Fernet; print(Fernet.generate_key().decode())"`. Set `APP_ENV=production`, `ALLOWED_HOSTS`, `SMS_GATEWAY_ALLOWED_HOSTS`, `SMTP_ALLOWED_HOSTS`, and `DATABASE_URL` to a PostgreSQL connection string using the psycopg driver. The SMS/SMTP allowlists must contain only approved public provider hosts. Configure network egress controls as well; do not rely on hostname checks alone to prevent DNS-rebinding/SSRF. Do not rely on local key files for hosted secrets.

Apply schema migrations before starting the web process:

```powershell
alembic upgrade head
python -m uvicorn saas_app:app --host 0.0.0.0 --port 8000
```

Production startup refuses a non-PostgreSQL database, and sessions require HTTPS. Back up PostgreSQL and the credential-encryption key together; losing the key makes stored provider credentials unrecoverable.

## Import columns

The `.xlsx` importer recognizes customer name/ID, company, location, phone, email, communication-consent flags, Telegram chat ID, invoice number, total/paid/due amounts, due date, reminder frequency (days), custom reminder dates, preferred language, and one installment amount/date pair per row. Repeat the invoice number on additional rows for multiple commitments. Missing data is reported per row; rows without a customer identity are skipped, and incomplete invoice data is marked for review. Invoice reminders default to 10, 7, and 5 days before the due date, then use the configured follow-up frequency and maximum. Users can override lead days, frequency, and custom dates per customer. Installment reminders default to 5 and 3 days before each commitment date. Download the sample workbook from the workspace.

## Current MVP boundaries

- Signup creates one tenant administrator. Email verification, invitations, user management, subscription billing, and SSO are not included yet.
- Tenant-scoped messaging credentials are encrypted at rest and write-only in the API. Channel settings can send a fixed test message to the explicitly configured test recipient. From a customer's contact row, workspace users can manually send SMS, email, and Telegram text; SMS, email, and WhatsApp require recorded contact consent. Manual sends require the channel to be enabled and the tenant's live-send gate to be on.
- Reminder occurrences are persisted in an idempotent queue. Run a separate worker with `python -m worker`; it only sends when the tenant's live-send gate and channel are enabled, a channel template exists, and the contact has recorded consent. Configure SMS/SMTP host allowlists and production egress restrictions.
- A customer is automatically flagged for review after three delivered payment reminders on separate days while a positive invoice balance is overdue. Workspace users can also flag or clear this status manually. A flag is deliberately retained after payment until a user reviews it.
- WhatsApp business-initiated reminders and manual WhatsApp sends use the configured approved Meta template or AiSensy campaign with its exact provider-side name/locale. Manual WhatsApp sends use the first open invoice with a positive balance; free-form text is not sent because provider-approved template rules apply. App templates do not override provider approval requirements. Payment links and payment reconciliation are not implemented.
- The original `app.py` is the earlier single-process demo. Do not expose it as the hosted multi-tenant service; run `saas_app:app`.
- Invoices are modeled in the database, but private PDF storage/serving and migration of legacy demo records remain future work.

Keep uploaded invoices and real provider credentials out of Git. Configure each tenant's channels in that tenant's authenticated workspace.
