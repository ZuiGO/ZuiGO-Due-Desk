# Dues Desk

Dues Desk is a FastAPI demo app for matching dealer dues with invoices and managing payment reminders.

## Run locally

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
uvicorn app:app --port 8000
```

Open http://localhost:8000. The app creates an `uploads` directory for uploaded files. Add any local credentials to a `.env` file next to `app.py`; do not commit that file.

Demo mode is enabled by default. Keep it enabled unless you are authorized and ready to send real messages.

## Repository contents

Local environment files, virtual environments, Python caches, and uploaded invoices are excluded from Git.
