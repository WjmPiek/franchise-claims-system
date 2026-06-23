# Martins Direct Franchise Claims System - Render Deployment

## 1. Push to GitHub

Create a new private GitHub repository, then run these commands from this project folder:

```powershell
git init
git add .
git commit -m "Production Render deployment"
git branch -M main
git remote add origin https://github.com/YOUR-GITHUB-USER/franchise-claims-system.git
git push -u origin main
```

Do not commit `.env`. This package excludes real secrets.

## 2. Create Render PostgreSQL

In Render, create a new PostgreSQL database. For this system, use a paid PostgreSQL plan for production-size data. Free PostgreSQL is not suitable for the 1.9M+ PolicyData rows.

Copy the **Internal Database URL** from Render.

## 3. Create Render Web Service

Create a new Web Service from the GitHub repository.

Build command:

```bash
pip install -r requirements.txt
```

Start command:

```bash
gunicorn app:app --workers 1 --threads 4 --timeout 900 --access-logfile - --error-logfile -
```

Health check path:

```text
/healthz
```

## 4. Render environment variables

Set these in Render:

```text
APP_ENV=production
FLASK_DEBUG=0
SECRET_KEY=<generate a long random value>
DATABASE_URL=<Render Internal Database URL>
SESSION_COOKIE_SECURE=1
SESSION_COOKIE_SAMESITE=Lax
SESSION_TIMEOUT_MINUTES=30
GOOGLE_MAPS_API_KEY=<your Google Maps key>
APP_BASE_URL=https://your-render-service.onrender.com
CRON_SECRET=<generate a long random value>
MAINTENANCE_MODE=0
ENABLE_POLICY_AGE_SCHEDULER=0
```

Optional email variables:

```text
SMTP_HOST=smtp.office365.com
SMTP_PORT=587
SMTP_USERNAME=<mailbox>
SMTP_PASSWORD=<mailbox password or app password>
SMTP_FROM=<mailbox>
SMTP_FROM_NAME=Martins Direct Analytics
SMTP_USE_TLS=1
```

## 5. Move local PostgreSQL data to Render PostgreSQL

From your local machine, export the local database:

```powershell
pg_dump -Fc -d franchise_claims -f franchise_claims.dump
```

Restore to Render using the External Database URL from Render:

```powershell
pg_restore --clean --if-exists --no-owner --dbname "RENDER_EXTERNAL_DATABASE_URL" franchise_claims.dump
```

After restore, open the live app and go to:

```text
/admin/policydata_current_members/rebuild/start
```

Then confirm:

```sql
SELECT COUNT(*) FROM policydata_detail_raw;
SELECT COUNT(*) FROM app_policydata_current_members;
SELECT COUNT(*) FROM app_policy_relation_summary;
```

## 6. Cron jobs on Render

Create Render Cron Jobs for:

Daily age check:

```text
https://YOUR-LIVE-URL/cron/daily_policy_age_check?token=YOUR_CRON_SECRET
```

Daily backup:

```text
https://YOUR-LIVE-URL/cron/daily_backup?token=YOUR_CRON_SECRET
```

Do not enable the in-process APScheduler on Render. Keep `ENABLE_POLICY_AGE_SCHEDULER=0` and use Render Cron instead.

## 7. Google Maps

Enable these APIs in Google Cloud:

- Maps JavaScript API
- Geocoding API

Then set `GOOGLE_MAPS_API_KEY` in Render.

## 8. Post-deploy checks

Open:

```text
/healthz
/admin/deployment_check
/dashboard
```

Then verify map cache and Policy Relation Counts.
