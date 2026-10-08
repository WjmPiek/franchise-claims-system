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

def prepare_premium_precision(engine):
    """Commit schema locks before the long-running Excel/COPY transaction."""
    from import_progress import report
    if engine is None:
        raise RuntimeError('PostgreSQL is unavailable; policy details were not saved')
    report('Checking premium storage precision')
    from sqlalchemy.exc import DBAPIError
    try:
        with engine.begin() as connection:
            connection.execute(text("SET LOCAL lock_timeout = '5s'"))
            connection.execute(text("SET LOCAL statement_timeout = '60s'"))
            ensure_premium_precision(connection)
    except DBAPIError as exc:
        code = getattr(exc.orig, 'pgcode', None)
        if code in {'55P03', '57014'}:
            raise RuntimeError(
                'The database precision update is busy or timed out. No rows from this '
                'import have been replaced. Wait for the current import to finish '
                'before retrying.') from exc
        raise
    report('Premium storage precision ready')
