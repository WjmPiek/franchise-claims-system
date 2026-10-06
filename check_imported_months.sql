-- Read-only diagnostics. Run against the same Render database used by DATABASE_URL.
BEGIN TRANSACTION READ ONLY;

SELECT current_database() AS connected_database, current_schema() AS connected_schema;

-- Actual monthly policy rows: these determine the dashboard's latest month.
SELECT import_month, COUNT(*) AS franchise_rows,
       SUM(retail_premium) AS retail, SUM(risk_premium) AS risk,
       SUM(original_risk_premium) AS original_risk,
       STRING_AGG(DISTINCT source_file, ', ') AS source_files
FROM policy_monthly_raw
GROUP BY import_month ORDER BY import_month;

-- Actual paid-claims months: newer claims alone do not advance the premium dashboard.
SELECT claim_month, COUNT(*) AS franchise_rows, SUM(claims_amount) AS claims,
       SUM(claim_count) AS claim_count,
       STRING_AGG(DISTINCT source_file, ', ') AS source_files
FROM claims_monthly_raw
GROUP BY claim_month ORDER BY claim_month;

-- Import receipts: inspect filenames versus the allocated months for May-September.
SELECT created_at, import_type, source_file, imported_months, row_count, status
FROM import_history ORDER BY created_at DESC LIMIT 100;

-- Detail rows can exist even if the separate monthly-summary save failed.
SELECT import_month, source_file, COUNT(*) AS detail_rows,
       COUNT(*) FILTER (WHERE is_mem) AS mem_rows,
       SUM(retail_premium) FILTER (WHERE is_mem) AS mem_retail
FROM policydata_detail_raw
GROUP BY import_month, source_file ORDER BY import_month, source_file;

COMMIT;

