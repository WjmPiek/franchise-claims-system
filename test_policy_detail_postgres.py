"""Exercise the actual detail importer against a temporary PostgreSQL table.

Run from the repository: python test_policy_detail_postgres.py.
Uses the configured DATABASE_URL, with no production row writes. The temporary
table copies production column types and uses the same conditional unique index.
"""
import pandas as pd
from contextlib import nullcontext
from sqlalchemy import text
import app

engine = app.get_db_engine()
assert engine is not None, 'DATABASE_URL must point to PostgreSQL'
engine.hide_parameters = True
original_get_engine = app.get_db_engine

class TransactionEngine:
    def __init__(self, connection):
        self.connection = connection
    def begin(self):
        return nullcontext(self.connection)

with engine.begin() as conn:
    conn.execute(text('CREATE TEMP TABLE policydata_detail_raw AS SELECT * FROM public.policydata_detail_raw WITH NO DATA'))
    conn.execute(text('CREATE UNIQUE INDEX test_detail_source_key ON pg_temp.policydata_detail_raw(source_row_key) WHERE source_row_key IS NOT NULL'))
    conn.execute(text('SET LOCAL search_path = pg_temp, public'))
    row = {'source_file':'test.xlsx','import_month':'2026-05-01','row_number':2,
           'source_row_key':'test.xlsx|2026-05-01|2','franchise':'TEST FRANCHISE',
           'relation':'MEM','is_mem':True,'retail_premium':10,'original_risk_premium':8,
           'mpia':1,'raw_data':{}}
    try:
        app.get_db_engine = lambda: TransactionEngine(conn)
        assert app.save_policy_detail_to_postgres(pd.DataFrame([row]),source_file='test.xlsx')
        # A different source argument avoids the delete-by-source branch and tests
        # the actual ON CONFLICT update of an existing key.
        row['retail_premium'] = 20
        assert app.save_policy_detail_to_postgres(pd.DataFrame([row]),source_file='test-reimport.xlsx')
    finally:
        app.get_db_engine = original_get_engine
    actual = conn.execute(text('SELECT COUNT(*),SUM(retail_premium) FROM pg_temp.policydata_detail_raw')).one()
    assert actual[0] == 1 and actual[1] == 20
print('PASS: detail insert and conflict update work with PostgreSQL conditional unique index; production rows unchanged.')

