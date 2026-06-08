# Franchise Claims Analytics System

Updated build with Tab 5 Payable To import.

Claims workbook import now reads:
- Tab 3 `Claims per Branch(sum)`: B8:B56 franchise names, C8:K56 paid claims per month
- Tab 4 `Claims per Branch(count)`: B8:B56 franchise names, C8:K56 claim counts per month
- Tab 5 `Payable To`: B4:F52
  - B = Franchise name
  - C = Claim paid to franchise
  - D = Claim paid to client
  - E = Repudiated / Pending
  - F = Grand Total Claims

These values are normalised and allocated to the matching franchise.

Update: Claims workbook tabs 3 and 4 now import all monthly columns from C through the month before Grand Total. This supports rolling claims files such as Apr 2025 through Apr 2026 and future months without changing code.

## Build update - Advertising Fund Commission removed
- Removed Advertising Fund Commission KPI card.
- Removed Advertising Fund % inputs from Commission Analysis and Scenario Comparison.
- Removed Advertising Fund Commission columns from Commission Analysis, Commission Comparison, Scenario Comparison, Board Report, PDF/Excel exports, and dashboard totals.
- Total Commission now uses BrightRock + Inkulu + MFF percentage commissions, plus separate R1 Policy Fee and 2.1% Underwriter Fee where shown.

## Build update - import progress and PolicyData validation
- Import popup now uses the browser upload progress for the upload stage instead of jumping to 95%.
- After upload, the popup switches to an animated server-processing state with elapsed time while Excel files are read and MEM rows are allocated.
- The popup reaches 100% only when the server has completed import and is reloading the dashboard.
- PolicyData_YYYYMMDD_to_YYYYMMDD files are supported as monthly transaction imports.
- Required PolicyData columns confirmed: Franchise, Relation, AUL Risk, Retail, MPIA.

## PostgreSQL / DBeaver mode

This build supports PostgreSQL persistence for large monthly PolicyData imports.

### 1. Create database in DBeaver
Create a PostgreSQL database, for example:

```sql
CREATE DATABASE franchise_claims;
```

### 2. Set connection string
Copy `.env.example` to `.env` and update the password:

```text
DATABASE_URL=postgresql://postgres:your_password@localhost:5432/franchise_claims
```

### 3. Install requirements

```bash
pip install -r requirements.txt
```

### 4. Start the app

```bash
python app.py
```

On first import the app creates these PostgreSQL tables automatically:

- `policy_monthly_raw`
- `claims_monthly_raw`
- `import_history`
- `franchise_mapping_pg`

Monthly PolicyData imports are saved to PostgreSQL. If the same month is imported again, that month is replaced first so totals are not duplicated. Dashboard pages then reload from PostgreSQL instead of relying only on Excel/memory.

### 5. Useful DBeaver checks

```sql
SELECT import_month, COUNT(*) franchises, SUM(retail_premium) retail, SUM(policy_qty) policies
FROM policy_monthly_raw
GROUP BY import_month
ORDER BY import_month;
```

```sql
SELECT claim_month, SUM(claims_amount) claims, SUM(claim_count) claim_count
FROM claims_monthly_raw
GROUP BY claim_month
ORDER BY claim_month;
```

Claims Tab 1 Allocation Update
------------------------------
Claims import now uses Tab 1 transaction rows as the source of truth:
- Column E: franchise name
- Column Y: Date Of Claim Paid
- Column AC: Amount of Claim Paid Before Net Off

The importer allocates each paid claim to the matching system franchise and paid month, then claim ratios are calculated as Total Claims / Risk Premium for month, 6-month, or 12-month views.

## Martins Direct import calculation update

This build applies the requested monthly import logic:

- PolicyData imports use Column A Franchise, Column I Relation, Column M Risk Premium, Column N Retail Premium, and Column R MPIA.
- Only MEM rows are used for the R1 / ADV fee risk calculation.
- R1/ADV calculation per row:
  - Single premium = Column M / Column R
  - R1 fee = R1 per paid policy month
  - ADV fund = (single premium - R1) x 2.1%, multiplied by Column R
  - New Risk Premium = (single premium - R1 - ADV per month) x Column R
- Claims imports use Column E Franchise, Column Y paid month, and Column Z claim amount.
- Monthly claim ratio is calculated as monthly claims divided by monthly risk premium per franchise. Six- and twelve-month views use weighted totals for reporting only; the monthly rows keep their own monthly claim ratio.
- Default commission rates are BrightRock 10%, Mkhulu/Inkulu 2.5%, and MFF 2.5%, calculated from Retail Premium.


## PostgreSQL / DBeaver setup

1. Create a PostgreSQL database, for example `franchise_claims`.
2. Open DBeaver, create a PostgreSQL connection, and connect to that database.
3. Run `postgres_schema.sql` in DBeaver to create the tables.
4. Set the app connection string in `.env`:

```env
DATABASE_URL=postgresql+psycopg2://postgres:your_password@localhost:5432/franchise_claims
```

5. Restart the Flask app and import the PolicyData workbook.

PolicyData storage:
- `policydata_detail_raw` keeps row-level imported PolicyData, allocated by Column A franchise name, with the original row in `raw_data` JSONB.
- `policy_monthly_raw` keeps monthly totals per franchise used by the dashboard: Retail Premium from Column N and calculated Risk Premium from Column M after R1 and 2.1% ADV deductions.
- Re-importing the same file/month replaces the row-level data for that file/month and replaces the monthly totals for that month.

Useful DBeaver checks:

```sql
-- Monthly totals by franchise
SELECT franchise_name, import_month,
       SUM(retail_premium) AS retail_premium,
       SUM(risk_premium) AS risk_premium
FROM policy_monthly_raw
GROUP BY franchise_name, import_month
ORDER BY franchise_name, import_month;

-- Row-level PolicyData audit for one month
SELECT franchise_name, import_month, COUNT(*) AS rows_loaded,
       SUM(CASE WHEN is_mem THEN retail_premium ELSE 0 END) AS mem_retail,
       SUM(CASE WHEN is_mem THEN new_risk_premium ELSE 0 END) AS mem_new_risk
FROM policydata_detail_raw
GROUP BY franchise_name, import_month
ORDER BY franchise_name, import_month;
```

## Test database fields in DBeaver

1. Connect to your PostgreSQL database in DBeaver.
2. Open `database_field_audit.sql` and run it.
3. If it returns no rows, all required PolicyData fields exist.
4. If it returns rows, those fields are missing. Run `postgres_schema.sql` in DBeaver, or restart the app with `DATABASE_URL` configured. The app also runs `postgres_schema.sql` automatically on startup.

The schema uses:

```sql
CREATE TABLE IF NOT EXISTS ...
ALTER TABLE ... ADD COLUMN IF NOT EXISTS ...
```

That means it creates missing tables and missing fields without deleting existing imported data.

## Database self-healing schema check

This version includes a PostgreSQL self-healing schema check. On startup the app now verifies required tables, columns and indexes, then creates any missing fields without deleting existing data.

Admin pages:

- `/admin/database_health` shows table row counts and missing fields.
- `/admin/repair_database` runs the repair/check manually and returns to the health page.

The schema repair is idempotent and safe to run multiple times.


## Hotfix: dashboard period and claim ratio period labels
- Dashboard KPI period now uses the latest PolicyData month, not a claims-only month. This prevents Total Retail Premium/Risk Premium from showing R0 when the claims workbook contains a later month than PolicyData.
- Claim ratio is recalculated from Claims / Risk Premium for monthly, six-month, and yearly views.
- Six-month analysis periods now display the real month range, e.g. `Apr 2025 - Sep 2025`, instead of generic `Period 1`, `Period 2`, `Period 3`.
- Advanced Analytics and Claim Ratio Analysis use the same recalculated period/month values.

## Production user rollout notes

This build includes production user management:

- The first registered user becomes the approved admin.
- All later registrations are created as Pending / Disabled.
- Admin approves users from `/admin/users` by changing them to Approved / Active.
- Admin can assign roles: Admin, User, or View only.
- Admin can reset passwords and disable accounts.
- Admin can review login / registration / user-management activity at `/admin/audit_log`.

For local testing, use PostgreSQL and set `DATABASE_URL` in `.env`. User accounts require PostgreSQL.

Recommended Render environment variables:

```env
DATABASE_URL=postgresql://...
SECRET_KEY=replace-with-a-long-random-production-secret
FLASK_DEBUG=0
```

Recommended Render start command:

```bash
gunicorn app:app
```

Before giving access to other users:

1. Deploy the app and connect the production PostgreSQL database.
2. Open the live URL and register your admin account first.
3. Ask other users to register.
4. Approve their accounts under Admin > User Management.
5. Review `/admin/audit_log` after testing.

## Phase 2: Franchise Security + Login Tracking

This build adds production-facing user visibility controls and login monitoring.

### Added

- `app_users.last_login` timestamp.
- `app_users.failed_login_count` counter.
- `app_users.last_failed_login` timestamp.
- Successful login updates `last_login` and resets failed login count.
- Failed login attempts against an existing user increment `failed_login_count`.
- Admin User Management now displays Last Login and Failed Logins.
- Main dashboard, workspaces, Excel exports, payover exports, and board reports are scoped through the logged-in user's assigned franchises.

### Testing

1. Login as Super Admin and assign a normal user one or more franchises.
2. Login as that normal user.
3. Confirm dashboard/workspaces/export/report data only shows assigned franchises.
4. Enter an incorrect password once and confirm Failed Logins increases in Admin Users.
5. Login successfully and confirm Last Login updates and Failed Logins resets to zero.

## Live Phase 3 - Backup Management

This build adds an Admin Backups page:

- `/admin/backups` shows backup history.
- `Create Backup Now` creates a ZIP backup in the local `backups/` folder.
- Backup ZIP files contain CSV exports of all managed database tables.
- Backup history is stored in `app_backup_history` and is auto-created by the self-healing schema.
- Old `/admin/backup_database` links still work and create/download a fresh backup immediately.

Recommended test:

1. Run `python app.py`.
2. Login as Super Admin.
3. Open `/admin/backups`.
4. Click `Create Backup Now`.
5. Download the generated ZIP and confirm it contains CSV files.

Before live launch, download and keep at least one backup outside the server.


## Live Phase 5 - Production Launch and Scheduled Backups

This build adds the final go-live tooling for Render production deployment.

### Added

- `/admin/launch_center` for the production go-live checklist.
- `/healthz` public health check for Render.
- `/cron/daily_backup` protected scheduled backup endpoint.
- `/admin/cron_log` for scheduled job history.
- `CRON_SECRET` protection for external cron calls.
- `MAINTENANCE_MODE` environment switch.
- Render `healthCheckPath: /healthz`.

### Render Cron Job

Create a Render Cron Job or external scheduler to call:

```text
https://your-live-domain/cron/daily_backup
```

Add this request header:

```text
X-Cron-Secret: your CRON_SECRET value
```

Recommended schedule:

```text
0 2 * * *
```

### Final go-live order

1. Set all environment variables on Render.
2. Deploy the web service.
3. Open `/healthz` and confirm it returns OK.
4. Login as Super Admin.
5. Open `/admin/launch_center`.
6. Create a manual backup.
7. Test forgot password email.
8. Test a normal franchise user.
9. Create the daily backup cron job.
10. Invite pilot users.


## Phase 6 - Claims Workflow Engine

Added a protected `/claims` workflow for live operational claim tracking:

- Create claim cases per franchise
- Track claimant, policy number, date, amount, status and priority
- Add claim notes and status changes
- Franchise users only see assigned franchises
- Admin and Super Admin see all claims
- Viewer users can view claims but cannot submit updates
- Claim creation and updates are written to the audit log
- Database self-healing creates `app_claim_cases` and `app_claim_notes`

Use:

- `/claims` for the claim case list
- `/claims/new` to create a new claim
- `/claims/<id>` to review and update a claim


## Phase 10 Claims Rules and Document Validation
Admin users can configure document requirements at `/admin/claims_rules`. Uploads are checked against the active admin checklist. Client heatmap now loads cached coordinates in visible batches with a progress bar.
