import os
import pandas as pd

try:
    import psycopg2
    from psycopg2.extras import execute_values
except Exception:
    psycopg2 = None
    execute_values = None

DATABASE_URL = os.environ.get('DATABASE_URL', '').strip()
SCHEMA_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'postgres_schema.sql')


def available():
    return bool(DATABASE_URL) and psycopg2 is not None


def connect():
    if not available():
        raise RuntimeError('PostgreSQL is not configured. Set DATABASE_URL in .env or environment.')
    return psycopg2.connect(DATABASE_URL)


def init_schema():
    if not available() or not os.path.exists(SCHEMA_FILE):
        return False
    with connect() as conn:
        with conn.cursor() as cur:
            cur.execute(open(SCHEMA_FILE, 'r', encoding='utf-8').read())
        conn.commit()
    return True


def _num(v):
    try:
        if v is None or pd.isna(v):
            return 0.0
        return float(v)
    except Exception:
        return 0.0


def save_policy_months(df, source_files=None):
    if not available() or df is None or df.empty:
        return False
    init_schema()
    work = df.copy()
    work['month'] = pd.to_datetime(work['month'], errors='coerce').dt.to_period('M').dt.to_timestamp()
    work = work.dropna(subset=['month'])
    months = sorted(set(work['month'].dt.date))
    for col in ['franchise','retail_premium','risk_premium','claims','claim_count','claim_paid_franchise','claim_paid_client','repudiated_pending','grand_total_claims','policy_qty','original_risk_premium','r1_policy_fee_imported','underwriter_2_1_fee','risk_after_r1','single_monthly_premium_total','current_scenario']:
        if col not in work.columns:
            work[col] = '100% Claim Ratio' if col == 'current_scenario' else 0
    rows = []
    source = ', '.join(source_files or [])[:500]
    for _, r in work.iterrows():
        rows.append((
            str(r.get('franchise','')).strip(), r['month'].date(),
            _num(r.get('retail_premium')), _num(r.get('risk_premium')),
            _num(r.get('claims')), _num(r.get('claim_count')),
            _num(r.get('claim_paid_franchise')), _num(r.get('claim_paid_client')),
            _num(r.get('repudiated_pending')), _num(r.get('grand_total_claims')),
            _num(r.get('policy_qty')), _num(r.get('original_risk_premium')),
            _num(r.get('r1_policy_fee_imported')), _num(r.get('underwriter_2_1_fee')),
            _num(r.get('risk_after_r1')), _num(r.get('single_monthly_premium_total')),
            str(r.get('current_scenario') or '100% Claim Ratio'), source,
        ))
    if not rows:
        return False
    with connect() as conn:
        with conn.cursor() as cur:
            if months:
                cur.execute('DELETE FROM policy_monthly_raw WHERE import_month = ANY(%s)', (months,))
            execute_values(cur, """
                INSERT INTO policy_monthly_raw (
                    franchise_name, import_month, retail_premium, risk_premium, claims,
                    claim_count, claim_paid_franchise, claim_paid_client, repudiated_pending,
                    grand_total_claims, policy_qty, original_risk_premium, r1_policy_fee,
                    underwriter_2_1_fee, risk_after_r1, single_monthly_premium_total,
                    current_scenario, source_file
                ) VALUES %s
            """, rows, page_size=1000)
            cur.execute('INSERT INTO import_history(import_type, source_file, imported_months, row_count, status) VALUES(%s,%s,%s,%s,%s)', ('policy', source, [str(m) for m in months], len(rows), 'success'))
        conn.commit()
    return True


def load_policy_months():
    if not available():
        return pd.DataFrame()
    init_schema()
    sql = """
        SELECT franchise_name AS franchise, import_month AS month, retail_premium,
               risk_premium, claims, claim_count, claim_paid_franchise,
               claim_paid_client, repudiated_pending, grand_total_claims, policy_qty,
               original_risk_premium, r1_policy_fee AS r1_policy_fee_imported,
               underwriter_2_1_fee, risk_after_r1, single_monthly_premium_total,
               current_scenario
        FROM policy_monthly_raw
        ORDER BY franchise_name, import_month
    """
    with connect() as conn:
        df = pd.read_sql(sql, conn)
    if not df.empty:
        df['month'] = pd.to_datetime(df['month'])
    return df


def save_claims(df, source_files=None):
    if not available() or df is None or df.empty:
        return False
    init_schema()
    work = df.copy()
    if 'claim_key' not in work.columns:
        work['claim_key'] = ''
    work['month'] = pd.to_datetime(work['month'], errors='coerce').dt.to_period('M').dt.to_timestamp()
    work = work.dropna(subset=['month'])
    for col in ['claim_key','franchise','claims','claim_count','claim_paid_franchise','claim_paid_client','repudiated_pending','grand_total_claims']:
        if col not in work.columns:
            work[col] = '' if col in {'claim_key','franchise'} else 0
    months = sorted(set(work['month'].dt.date))
    rows = []
    source = ', '.join(source_files or [])[:500]
    for _, r in work.iterrows():
        rows.append((
            str(r.get('claim_key','')).strip(), str(r.get('franchise','')).strip(), r['month'].date(),
            _num(r.get('claims')), _num(r.get('claim_count')),
            _num(r.get('claim_paid_franchise')), _num(r.get('claim_paid_client')),
            _num(r.get('repudiated_pending')), _num(r.get('grand_total_claims')), source,
        ))
    if not rows:
        return False
    with connect() as conn:
        with conn.cursor() as cur:
            if months:
                cur.execute('DELETE FROM claims_monthly_raw WHERE claim_month = ANY(%s)', (months,))
            execute_values(cur, """
                INSERT INTO claims_monthly_raw (
                    claim_key, claims_franchise_name, claim_month, claims_amount,
                    claim_count, claim_paid_franchise, claim_paid_client,
                    repudiated_pending, grand_total_claims, source_file
                ) VALUES %s
            """, rows, page_size=1000)
            cur.execute('INSERT INTO import_history(import_type, source_file, imported_months, row_count, status) VALUES(%s,%s,%s,%s,%s)', ('claims', source, [str(m) for m in months], len(rows), 'success'))
        conn.commit()
    return True


def load_claims():
    if not available():
        return pd.DataFrame()
    init_schema()
    sql = """
        SELECT claim_key, claims_franchise_name AS franchise, claim_month AS month,
               claims_amount AS claims, claim_count, claim_paid_franchise,
               claim_paid_client, repudiated_pending, grand_total_claims
        FROM claims_monthly_raw
        ORDER BY claims_franchise_name, claim_month
    """
    with connect() as conn:
        df = pd.read_sql(sql, conn)
    if not df.empty:
        df['month'] = pd.to_datetime(df['month'])
    return df
