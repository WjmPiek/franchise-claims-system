-- Run this in DBeaver to test if all required PolicyData database fields exist.
-- It returns only missing fields. If the result is empty, the database has all fields.
WITH required_fields(table_name, column_name, data_type) AS (
    VALUES
    ('policy_monthly_raw','franchise_name','text'),
    ('policy_monthly_raw','import_month','date'),
    ('policy_monthly_raw','retail_premium','numeric'),
    ('policy_monthly_raw','risk_premium','numeric'),
    ('policy_monthly_raw','original_risk_premium','numeric'),
    ('policy_monthly_raw','r1_policy_fee','numeric'),
    ('policy_monthly_raw','underwriter_2_1_fee','numeric'),
    ('policy_monthly_raw','risk_after_r1','numeric'),
    ('policy_monthly_raw','single_monthly_premium_total','numeric'),
    ('policy_monthly_raw','source_file','text'),
    ('policydata_detail_raw','source_file','text'),
    ('policydata_detail_raw','import_month','date'),
    ('policydata_detail_raw','row_number','integer'),
    ('policydata_detail_raw','franchise_name','text'),
    ('policydata_detail_raw','relation','text'),
    ('policydata_detail_raw','is_mem','boolean'),
    ('policydata_detail_raw','retail_premium','numeric'),
    ('policydata_detail_raw','original_risk_premium','numeric'),
    ('policydata_detail_raw','mpia','numeric'),
    ('policydata_detail_raw','single_premium','numeric'),
    ('policydata_detail_raw','r1_policy_fee','numeric'),
    ('policydata_detail_raw','adv_fund_2_1_fee','numeric'),
    ('policydata_detail_raw','risk_after_r1','numeric'),
    ('policydata_detail_raw','new_risk_premium','numeric'),
    ('policydata_detail_raw','raw_data','jsonb')
)
SELECT r.table_name, r.column_name, r.data_type AS expected_type
FROM required_fields r
LEFT JOIN information_schema.columns c
       ON c.table_schema = 'public'
      AND c.table_name = r.table_name
      AND c.column_name = r.column_name
WHERE c.column_name IS NULL
ORDER BY r.table_name, r.column_name;

-- To create missing fields automatically, run postgres_schema.sql.
-- The schema file uses CREATE TABLE IF NOT EXISTS and ALTER TABLE ADD COLUMN IF NOT EXISTS.
