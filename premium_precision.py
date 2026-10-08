"""Preserve source premium precision; round only displayed totals."""
from sqlalchemy import text

PRECISION_COLUMNS = (
    ('policydata_detail_raw', 'original_risk_premium'),
    ('policydata_detail_raw', 'retail_premium'),
    ('policy_monthly_raw', 'original_risk_premium'),
    ('policy_monthly_raw', 'retail_premium'),
)

def ensure_premium_precision(connection):
    """Widen only scaled source-premium columns, within the import transaction.

    Existing values are preserved. Reimport the original workbook to recover
    decimals already discarded by an older import.
    """
    rows = connection.execute(text("""
        SELECT table_name, column_name, numeric_scale
        FROM information_schema.columns
        WHERE table_schema = current_schema()
          AND table_name IN ('policydata_detail_raw', 'policy_monthly_raw')
          AND column_name IN ('original_risk_premium', 'retail_premium')
          AND data_type = 'numeric'
    """)).mappings()
    scaled = {(r['table_name'], r['column_name']) for r in rows if r['numeric_scale'] is not None}
    for table, column in PRECISION_COLUMNS:
        if (table, column) in scaled:
            # Identifiers come exclusively from the fixed allowlist above.
            connection.execute(text(f'ALTER TABLE "{table}" ALTER COLUMN "{column}" TYPE NUMERIC'))
