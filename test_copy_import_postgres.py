"""Verify COPY against a temporary shadow table; production rows are untouched."""
import os
import json
import time
from datetime import date
from sqlalchemy import create_engine, text
from import_bulk import copy_detail_rows, DETAIL_COLUMNS

engine = create_engine(os.environ['DATABASE_URL'].replace('postgres://', 'postgresql://', 1), hide_parameters=True)
try:
    with engine.connect() as conn:
        transaction = conn.begin()
        try:
            conn.execute(text('CREATE TEMP TABLE policydata_detail_raw AS SELECT * FROM public.policydata_detail_raw WITH NO DATA'))
            conn.execute(text('CREATE UNIQUE INDEX test_copy_unique_key ON pg_temp.policydata_detail_raw(source_row_key) WHERE source_row_key IS NOT NULL'))
            conn.execute(text('SET LOCAL search_path = pg_temp, public'))
            row = dict.fromkeys(DETAIL_COLUMNS, 0)
            row.update(source_file='Test, "quoted".xlsx', import_month=date(2000,1,1),
                       row_number=2, source_row_key='test|2000-01-01|2', client_address_o='line1\nline2',
                       id_number_f=None, franchise_name="O'Brien Müller", relation='', is_mem=False,
                       retail_premium=162198.00, original_risk_premium=-319.69,
                       raw_data=json.dumps({'name':'a\nb\\c"d'}))
            copy_detail_rows(conn, [row])
            actual = conn.execute(text('SELECT source_file, client_address_o, id_number_f, franchise_name, relation, is_mem, retail_premium, original_risk_premium, raw_data FROM pg_temp.policydata_detail_raw')).one()
            assert actual[0:6] == (row['source_file'],row['client_address_o'],None,row['franchise_name'],'',False)
            assert float(actual[6]) == row['retail_premium'] and float(actual[7]) == row['original_risk_premium']
            assert actual[8] == json.loads(row['raw_data'])
            conn.execute(text('DELETE FROM pg_temp.policydata_detail_raw'))
            rows = [dict(row, source_row_key=f'synthetic|{i}', row_number=i + 2) for i in range(10000)]
            started = time.monotonic()
            for offset in range(0,len(rows),2000): copy_detail_rows(conn,rows[offset:offset+2000])
            assert conn.execute(text('SELECT COUNT(*) FROM pg_temp.policydata_detail_raw')).scalar_one() == 10000
            print(f'PASS: COPY preserves types, NULL/blank, Unicode, quotes, newlines and JSON. Synthetic 10000-row shadow-table write: {time.monotonic()-started:.2f}s. Production rows unchanged; this is not a production speed forecast.')
        finally:
            transaction.rollback()
finally:
    engine.dispose()
