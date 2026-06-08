CREATE TABLE IF NOT EXISTS policy_monthly_raw (
    id BIGSERIAL PRIMARY KEY,
    franchise_name TEXT NOT NULL,
    import_month DATE NOT NULL,
    retail_premium NUMERIC(18,2) DEFAULT 0,
    risk_premium NUMERIC(18,2) DEFAULT 0,
    claims NUMERIC(18,2) DEFAULT 0,
    claim_count NUMERIC(18,2) DEFAULT 0,
    claim_paid_franchise NUMERIC(18,2) DEFAULT 0,
    claim_paid_client NUMERIC(18,2) DEFAULT 0,
    repudiated_pending NUMERIC(18,2) DEFAULT 0,
    grand_total_claims NUMERIC(18,2) DEFAULT 0,
    policy_qty NUMERIC(18,2) DEFAULT 0,
    original_risk_premium NUMERIC(18,2) DEFAULT 0,
    r1_policy_fee NUMERIC(18,2) DEFAULT 0,
    underwriter_2_1_fee NUMERIC(18,2) DEFAULT 0,
    risk_after_r1 NUMERIC(18,2) DEFAULT 0,
    single_monthly_premium_total NUMERIC(18,2) DEFAULT 0,
    current_scenario TEXT DEFAULT '100% Claim Ratio',
    source_file TEXT,
    created_at TIMESTAMP DEFAULT NOW(),
    UNIQUE(franchise_name, import_month)
);

CREATE INDEX IF NOT EXISTS idx_policy_monthly_raw_month ON policy_monthly_raw(import_month);
CREATE INDEX IF NOT EXISTS idx_policy_monthly_raw_franchise ON policy_monthly_raw(franchise_name);

CREATE TABLE IF NOT EXISTS claims_monthly_raw (
    id BIGSERIAL PRIMARY KEY,
    claim_key TEXT,
    claims_franchise_name TEXT NOT NULL,
    claim_month DATE NOT NULL,
    claims_amount NUMERIC(18,2) DEFAULT 0,
    claim_count NUMERIC(18,2) DEFAULT 0,
    claim_paid_franchise NUMERIC(18,2) DEFAULT 0,
    claim_paid_client NUMERIC(18,2) DEFAULT 0,
    repudiated_pending NUMERIC(18,2) DEFAULT 0,
    grand_total_claims NUMERIC(18,2) DEFAULT 0,
    source_file TEXT,
    created_at TIMESTAMP DEFAULT NOW(),
    UNIQUE(claim_key, claims_franchise_name, claim_month)
);

CREATE INDEX IF NOT EXISTS idx_claims_monthly_raw_month ON claims_monthly_raw(claim_month);
CREATE INDEX IF NOT EXISTS idx_claims_monthly_raw_key ON claims_monthly_raw(claim_key);


CREATE TABLE IF NOT EXISTS policydata_detail_raw (
    id BIGSERIAL PRIMARY KEY,
    source_file TEXT NOT NULL,
    import_month DATE NOT NULL,
    row_number INTEGER NOT NULL,
    franchise_name TEXT NOT NULL,
    relation TEXT,
    is_mem BOOLEAN DEFAULT FALSE,
    retail_premium NUMERIC(18,2) DEFAULT 0,
    original_risk_premium NUMERIC(18,2) DEFAULT 0,
    mpia NUMERIC(18,2) DEFAULT 0,
    single_premium NUMERIC(18,6) DEFAULT 0,
    r1_policy_fee NUMERIC(18,2) DEFAULT 0,
    adv_fund_2_1_fee NUMERIC(18,2) DEFAULT 0,
    risk_after_r1 NUMERIC(18,2) DEFAULT 0,
    new_risk_premium NUMERIC(18,2) DEFAULT 0,
    raw_data JSONB,
    created_at TIMESTAMP DEFAULT NOW(),
    UNIQUE(source_file, import_month, row_number)
);

CREATE INDEX IF NOT EXISTS idx_policydata_detail_raw_month ON policydata_detail_raw(import_month);
CREATE INDEX IF NOT EXISTS idx_policydata_detail_raw_franchise ON policydata_detail_raw(franchise_name);
CREATE INDEX IF NOT EXISTS idx_policydata_detail_raw_relation ON policydata_detail_raw(relation);
CREATE INDEX IF NOT EXISTS idx_policydata_detail_raw_raw_data ON policydata_detail_raw USING GIN(raw_data);

CREATE TABLE IF NOT EXISTS import_history (
    id BIGSERIAL PRIMARY KEY,
    import_type TEXT NOT NULL,
    source_file TEXT,
    imported_months TEXT[],
    row_count INTEGER DEFAULT 0,
    status TEXT DEFAULT 'success',
    created_at TIMESTAMP DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS franchise_mapping_pg (
    id BIGSERIAL PRIMARY KEY,
    source_name TEXT UNIQUE NOT NULL,
    mapped_name TEXT NOT NULL,
    approved BOOLEAN DEFAULT TRUE,
    created_at TIMESTAMP DEFAULT NOW(),
    updated_at TIMESTAMP DEFAULT NOW()
);

-- Safety migration: if an older database already has the tables but is missing fields,
-- PostgreSQL will create the missing columns without deleting existing data.
ALTER TABLE policy_monthly_raw ADD COLUMN IF NOT EXISTS retail_premium NUMERIC(18,2) DEFAULT 0;
ALTER TABLE policy_monthly_raw ADD COLUMN IF NOT EXISTS risk_premium NUMERIC(18,2) DEFAULT 0;
ALTER TABLE policy_monthly_raw ADD COLUMN IF NOT EXISTS original_risk_premium NUMERIC(18,2) DEFAULT 0;
ALTER TABLE policy_monthly_raw ADD COLUMN IF NOT EXISTS r1_policy_fee NUMERIC(18,2) DEFAULT 0;
ALTER TABLE policy_monthly_raw ADD COLUMN IF NOT EXISTS underwriter_2_1_fee NUMERIC(18,2) DEFAULT 0;
ALTER TABLE policy_monthly_raw ADD COLUMN IF NOT EXISTS risk_after_r1 NUMERIC(18,2) DEFAULT 0;
ALTER TABLE policy_monthly_raw ADD COLUMN IF NOT EXISTS single_monthly_premium_total NUMERIC(18,2) DEFAULT 0;
ALTER TABLE policy_monthly_raw ADD COLUMN IF NOT EXISTS source_file TEXT;
ALTER TABLE policy_monthly_raw ADD COLUMN IF NOT EXISTS created_at TIMESTAMP DEFAULT NOW();

ALTER TABLE policydata_detail_raw ADD COLUMN IF NOT EXISTS source_file TEXT;
ALTER TABLE policydata_detail_raw ADD COLUMN IF NOT EXISTS import_month DATE;
ALTER TABLE policydata_detail_raw ADD COLUMN IF NOT EXISTS row_number INTEGER;
ALTER TABLE policydata_detail_raw ADD COLUMN IF NOT EXISTS franchise_name TEXT;
ALTER TABLE policydata_detail_raw ADD COLUMN IF NOT EXISTS relation TEXT;
ALTER TABLE policydata_detail_raw ADD COLUMN IF NOT EXISTS is_mem BOOLEAN DEFAULT FALSE;
ALTER TABLE policydata_detail_raw ADD COLUMN IF NOT EXISTS retail_premium NUMERIC(18,2) DEFAULT 0;
ALTER TABLE policydata_detail_raw ADD COLUMN IF NOT EXISTS original_risk_premium NUMERIC(18,2) DEFAULT 0;
ALTER TABLE policydata_detail_raw ADD COLUMN IF NOT EXISTS mpia NUMERIC(18,2) DEFAULT 0;
ALTER TABLE policydata_detail_raw ADD COLUMN IF NOT EXISTS single_premium NUMERIC(18,6) DEFAULT 0;
ALTER TABLE policydata_detail_raw ADD COLUMN IF NOT EXISTS r1_policy_fee NUMERIC(18,2) DEFAULT 0;
ALTER TABLE policydata_detail_raw ADD COLUMN IF NOT EXISTS adv_fund_2_1_fee NUMERIC(18,2) DEFAULT 0;
ALTER TABLE policydata_detail_raw ADD COLUMN IF NOT EXISTS risk_after_r1 NUMERIC(18,2) DEFAULT 0;
ALTER TABLE policydata_detail_raw ADD COLUMN IF NOT EXISTS new_risk_premium NUMERIC(18,2) DEFAULT 0;
ALTER TABLE policydata_detail_raw ADD COLUMN IF NOT EXISTS raw_data JSONB;
ALTER TABLE policydata_detail_raw ADD COLUMN IF NOT EXISTS created_at TIMESTAMP DEFAULT NOW();

-- Application users for production login / role management
CREATE TABLE IF NOT EXISTS app_users (
    id BIGSERIAL PRIMARY KEY,
    name TEXT,
    email TEXT,
    password_hash TEXT,
    role TEXT DEFAULT 'user',
    is_active BOOLEAN DEFAULT true,
    created_at TIMESTAMP DEFAULT NOW(),
    updated_at TIMESTAMP DEFAULT NOW()
);
CREATE UNIQUE INDEX IF NOT EXISTS uq_app_users_email ON app_users (LOWER(TRIM(email)));

-- Audit log for production user activity
CREATE TABLE IF NOT EXISTS app_audit_log (
    id BIGSERIAL PRIMARY KEY,
    user_id INTEGER,
    user_email TEXT,
    action TEXT,
    details TEXT,
    ip_address TEXT,
    created_at TIMESTAMP DEFAULT NOW()
);
CREATE INDEX IF NOT EXISTS idx_app_audit_log_created ON app_audit_log(created_at DESC);
CREATE INDEX IF NOT EXISTS idx_app_audit_log_user ON app_audit_log(user_email);
