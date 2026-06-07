# Krishna Engineering — Attendance Tool

A web application that accepts the monthly ONtime punch PDF and leave balance XLSX,
processes them, and generates a fully-formatted Excel attendance register.

## Quick Start

### 1. Install dependencies

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

### 2. Configure login

Copy `.env.example` to `.env` and fill in:

```bash
cp .env.example .env
```

Required for Google OAuth:

- `APP_BASE_URL` — local example: `http://127.0.0.1:8000`
- `SESSION_SECRET` — a long random secret
- `GOOGLE_CLIENT_ID`
- `GOOGLE_CLIENT_SECRET`
- `ALLOWED_GOOGLE_DOMAINS` — optional comma-separated Workspace domains
- `ADMIN_EMAILS` — optional comma-separated admin emails
- `MAKER_EMAILS` — optional comma-separated maker emails

In Google Cloud Console, create an OAuth Web Client and add this authorized redirect URI:

```text
http://127.0.0.1:8000/auth/callback
```

For production, replace the host with your deployed URL, for example:

```text
https://your-app.example.com/auth/callback
```

### 3. Run the server

```bash
python run.py
```

Then open **http://localhost:8000** in your browser.

### Options

```bash
python run.py --port 9000           # different port
python run.py --host 0.0.0.0        # expose on local network (other machines)
python run.py --reload              # auto-reload on code changes (dev mode)
```

---

## Access Control

The app now requires Google login.

- **Maker** users can upload/process/import registers, edit attendance, download generated files, and submit a register for admin approval.
- **Admin** users can do everything Makers can do, plus review pending approvals and freeze registers.
- The first user who signs in becomes an admin automatically.
- Emails listed in `ADMIN_EMAILS` are always admins.
- Emails listed in `MAKER_EMAILS` are always makers.
- All other new users are Makers by default.

Current built-in role defaults:

| Email | Role |
|-------|------|
| saianirudh@krishna-engineering.com | Admin |
| accounts@krishna-engineering.com | Maker |

---

## Hosting

This is a FastAPI backend app, so it can be stored on GitHub but cannot run on GitHub Pages. Use GitHub as the source repository and deploy it to a Python web host such as Render, Railway, Fly.io, Heroku-style platforms, Azure App Service, or a VM.

The included `Procfile` starts the app with:

```bash
uvicorn backend.main:app --host 0.0.0.0 --port ${PORT:-8000}
```

Set the environment variables from `.env.example` in the hosting provider dashboard. For production, use an HTTPS `APP_BASE_URL` and add the matching `/auth/callback` URL in Google Cloud Console.

---

## Project Structure

```
attendance-tool/
├── run.py                  ← start the server
├── requirements.txt
├── backend/
│   ├── main.py             ← FastAPI routes
│   ├── parser.py           ← PDF + XLSX data extraction
│   └── processor.py        ← Excel workbook builder
├── frontend/
│   ├── index.html          ← Single-page UI
│   └── static/
│       ├── css/style.css
│       └── js/app.js
├── uploads/                ← temp files (auto-cleaned)
└── outputs/                ← generated registers
```

---

## How It Works

1. **Upload** the ONtime Secureye punch PDF and the leave/workforce XLSX
2. The backend **parses** employee records, daily punch times, and leave balances
3. A new Excel workbook is **built** with one sheet per employee:
   - Full-time employees get Short/Long/Net Hours, Late Mark, Half Day, and a full Leave Summary
   - Labour workers get Hours Worked, Overtime, and a simplified Monthly Summary
4. The workbook is **returned** as a download

---

## Business Rules Encoded

| Rule | Corporate Office | Factory |
|------|-----------------|---------|
| Week Off | Sunday | Tuesday |
| Working Days | Mon–Sat | Mon, Wed–Sun |
| Short Hour (In) | After 09:45 | After 09:15 |
| Short Hour (Out) | Before 18:15 | Before 17:45 |
| Long Hour (In) | Before 09:15 | Before 08:30 |
| Long Hour (Out) | After 19:15 | After 18:45 |
| Late Mark | 10:30–14:00 | 10:00–13:30 |
| Half Day | In ≥ 14:00 or Out ≤ 14:00 | In ≥ 13:30 or Out ≤ 13:30 |

Leave logic: CL → SL → EL → LWP waterfall, with EL accrual from net hours.
