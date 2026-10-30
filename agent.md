# Dues Desk Engineering Handoff

## Current objective

Audit the existing project against the product request and evolve it into a hosted, multi-tenant receivables/reminder application. The owner confirmed:

- The product is intended to be hosted and sold to multiple business customers.
- MVP onboarding: email/password signup creates a tenant administrator.
- Email verification, team invitations, and subscription billing are later work.
- Messaging credentials and channel selection must be managed in the application.
- Customers may have multiple contacts and invoices.
- Imports must tolerate incomplete Excel data and make missing/invalid data visible.
- Installment commitments need dates/amounts and advance reminders.
- Reminder frequency/dates should be configurable per customer.
- Reminder templates should be editable and support multiple languages and channels.

## Repository baseline and findings

- Repository: `ZuiGO/ZuiGO-Due-Desk`, branch `main`; the workspace is clean at the initial import commit when this work began.
- Existing application: one `app.py` FastAPI module and a single-page `index.html`; requirements are in `requirements.txt`.
- Current state is held in process-global dictionaries/lists. Upload, settings, policies, messages, replies, and payments are lost on restart and not isolated between businesses.
- Credentials are read from `.env`; the UI has channel on/off policy checkboxes but no credential/configuration screen.
- Excel parsing guesses columns and tolerates some missing fields, but creates one in-memory dealer row and does not model customers, contacts, invoices, or committed installment rows.
- Existing installment handling only creates an automatically spaced in-memory plan after an inbound reply. It does not import commitment dates or schedule configurable advance reminders.
- Existing multilingual copy depends on a Groq API key and generated Hindi/Marathi text. There is no user-managed template library.
- There are no test files or database models/migrations in the repository.
- Existing WhatsApp, SMS, email, and Telegram integrations should be preserved where practical, but only selected/configured channels should send.
- `.env` and invoice uploads are private local data. Do not read, print, commit, or copy `.env` values. Keep both `.env` and uploaded invoices excluded from Git.

## Architecture decision

Use PostgreSQL for the hosted multi-tenant deployment. SQLite is useful for single-user local prototypes/tests but is not the selected production database for tenant isolation, concurrent web/scheduler workers, and future growth. Use a normalized relational schema rather than forcing everything into only three tables:

1. `tenants`
2. `users` (tenant membership and role)
3. `customers` (tenant-scoped)
4. `contacts` (many contacts per customer)
5. `invoices` (many invoices per customer)
6. `installments` (dated amount commitments per invoice)
7. `reminder_policies` (defaults and per-customer overrides)
8. `channel_settings` (tenant channel enablement and encrypted provider credentials)
9. `message_templates` (tenant/platform/language/template-stage scoped content)
10. `reminder_events` (audit/delivery/idempotency history)

All tenant-owned records and queries must be tenant-scoped; unique constraints and foreign keys should include tenant ownership where appropriate. Encrypt provider secrets at rest using a deployment-managed encryption key (KMS/secret manager or an explicit environment key); never return stored secrets from APIs. Do not make a hosted deployment claim until authentication, tenant isolation, migrations, secret handling, and authorization tests pass.

## Plan and work tracking

1. Audit application behavior, integration boundaries, and expected import fields.
2. Implement account authentication and a PostgreSQL-backed tenant-scoped data model/migrations.
3. Implement safe channel settings, multiple channel selection, configurable credentials, and credential status/test UX.
4. Implement incomplete Excel import validation and normalized customer/contact/invoice/installment persistence.
5. Implement installment-aware, configurable reminders, customer-specific schedules, and editable multilingual per-channel templates.
6. Add focused unit/API/integration tests (including cross-tenant access isolation) and document setup/deployment.
7. Run migrations/tests/build checks; summarize what is done and explicitly list deferred production work.

## Known safety and scope constraints

- Never enable live sends by default. Keep demo/test delivery behavior safe and require an explicit live-mode action/configuration.
- Avoid placing payment links or customer/invoice details in logs unless necessary; keep an auditable delivery history without leaking provider credentials.
- Do not expose customer data, invoice files, or credentials across tenant boundaries.
- SaaS MVP excludes billing, email verification, invitations, and enterprise SSO unless explicitly requested later.
- Confirm production configuration requirements (PostgreSQL URL, session signing key, HTTPS, encryption key, and SMTP/identity settings) in deployment documentation.

## Progress log

- 2026-10-30: Confirmed the repository is the previously imported `main` branch and current local state is clean.
- 2026-10-30: Completed initial read-only feature assessment. Confirmed no database, persistence, authentication, tenant isolation, settings UI, per-customer schedules, committed installment import, or editable template library.
- 2026-10-30: User selected hosted multi-tenant SaaS MVP with signup-created tenant admin; verification, invites, and billing are deferred.
- 2026-10-30: Recommended PostgreSQL for hosted production and a normalized schema with supporting settings, installment, policy, template, and event tables in addition to customer/contact/invoice.
- 2026-10-30: Implemented the hosted SaaS MVP in `saas_app.py`, SQLAlchemy models, migrations, encrypted channel credentials, tenant-scoped customer/invoice APIs, Excel import, per-customer reminder scheduling, multilingual templates, and a separate delivery worker. Live delivery remains disabled by default.
- 2026-10-30: Added `/auth/status` for the dashboard's initial session check; it reports anonymous sessions without returning an authorization error and clears stale sessions. Added regression assertions for anonymous and authenticated status.
- 2026-10-30: Validation passed: `python -m pytest -q` (9 tests) and `python -m py_compile saas_app.py worker.py models.py`. A browser reload confirmed sign-in/signup views render and the company-name field only appears during signup.
- 2026-10-30: Fixed the signup/sign-in browser error where `FormData` received a non-form value after an `await`. The submit handler now captures its form element before awaiting CSRF retrieval. Regression tests pass (9 tests), and a browser smoke test successfully created a workspace and reached the dashboard.
- 2026-10-30: Remaining before production: deploy/test against PostgreSQL and HTTPS, complete email verification and rate limiting, production object storage, team invites/billing, and provider-specific live-send verification. Do not expose the legacy `app.py` demo as the hosted entry point; run `saas_app:app` and the separate `python -m worker`.
