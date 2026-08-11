# Timber Yard System

## First-time setup

1. **Google Sheet** — create one sheet (any name, e.g. "Timber Yard Data").
   Copy its ID from the URL (the part between `/d/` and `/edit`).

2. **Code.gs** — open script.google.com, create a new **standalone** project (not bound to a sheet),
   paste in `Code.gs`. Fill in:
   - `SHEET_ID`
   - `BUSINESS_NAME`, `BUSINESS_LOCATION`, `BUSINESS_PHONE`
   - `AT_API_KEY` (from africastalking.com, free sandbox to start) for SMS on payment

3. **Deploy the script**: Deploy > New deployment > Web app > Execute as Me > Anyone. Copy the URL.

4. **Create the first boss login** — open your Google Sheet, find the "Users" tab
   (created automatically the first time the script runs), and add one row by hand:
   `ID | Username | Password | Role | Name` → e.g. `u1 | boss | yourpassword | boss | Mathenge`
   After that, the boss can create employee logins from the dashboard itself.

5. Paste the Apps Script Web App URL into `APPS_SCRIPT_URL` in **index.html**, **employees.html**, and **boss.html**.

6. Deploy to Vercel (see below).

## Deploying to Vercel via Git

```bash
git init
git add .
git commit -m "Initial commit"
git branch -M main
git remote add origin https://github.com/YOUR_USERNAME/timber-yard-app.git
git push -u origin main
```

Then on vercel.com: **New Project > Import** your GitHub repo > Deploy. No build settings needed — it's static HTML.

Any time you change `APPS_SCRIPT_URL` or edit a page, just `git add . && git commit -m "update" && git push` — Vercel redeploys automatically.
