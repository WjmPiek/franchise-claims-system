"""Fast, parameterized PostgreSQL writes using the caller's transaction."""
import re
import json
from contextlib import contextmanager
from sqlalchemy import text
from psycopg2.extras import execute_values
from import_progress import report


@contextmanager
def policy_detail_batches(engine, source_file, month):
    if engine is None:
        raise RuntimeError('PostgreSQL is unavailable; policy details were not saved')
    written = 0
    with engine.begin() as connection:
        report('Replacing this month in the database')
        connection.execute(text('DELETE FROM policydata_detail_raw WHERE import_month = :month'), {'month': month})
        def write_batch(batch):
            nonlocal written
            converted = []
            for item in batch:
                row = dict(item)
                row['franchise_name'] = row.pop('franchise')
                row['source_file'] = source_file
                row['import_month'] = month
                row['raw_data'] = json.dumps(row['raw_data'], default=str)
                converted.append(row)
            insert_detail_rows(connection, DETAIL_INSERT_SQL, converted)
            written += len(converted)
            report('Excel rows written (awaiting commit)', written)
        yield write_batch
        if written == 0:
            raise RuntimeError('No applicable policy rows found; existing month was preserved')
        report('Committing policy rows', written, written)
    report('Policy rows committed', written, written)


def insert_detail_rows(connection, statement, rows):
    sql = str(statement)
    values = re.search(r'VALUES\s*\((.*?)\)\s*(ON CONFLICT)', sql, re.S)
    if values is None:
        raise ValueError('Expected detail INSERT with conflict handling')
    keys = re.findall(r':(\w+)', values.group(1))
    template = re.sub(r':\w+', '%s', values.group(1))
    bulk_sql = sql[:values.start()] + 'VALUES %s ' + sql[values.start(2):]
    # psycopg2 adapts every value; filenames, names and JSON never become SQL.
    # Keep the delete and every insert in the same SQLAlchemy transaction.
    with connection.connection.cursor() as cursor:
        for start in range(0, len(rows), 2000):
            batch = [tuple(row[key] for key in keys) for row in rows[start:start + 2000]]
            execute_values(cursor, bulk_sql, batch, template='(' + template + ')', page_size=2000)
            report('Writing policy rows (awaiting commit)', min(start + len(batch), len(rows)), len(rows))


DETAIL_INSERT_SQL = '\n            INSERT INTO policydata_detail_raw (\n                source_file, import_month, row_number, source_row_key, client_address_o, id_number_f, franchise_name, relation, is_mem,\n                retail_premium, original_risk_premium, mpia, single_premium,\n                r1_policy_fee, adv_fund_2_1_fee, risk_after_r1, new_risk_premium, raw_data\n            ) VALUES (\n                :source_file, :import_month, :row_number, :source_row_key, :client_address_o, :id_number_f, :franchise_name, :relation, :is_mem,\n                :retail_premium, :original_risk_premium, :mpia, :single_premium,\n                :r1_policy_fee, :adv_fund_2_1_fee, :risk_after_r1, :new_risk_premium, CAST(:raw_data AS jsonb)\n            )\n            ON CONFLICT (source_row_key) WHERE source_row_key IS NOT NULL DO UPDATE SET\n                source_file = EXCLUDED.source_file,\n                import_month = EXCLUDED.import_month,\n                row_number = EXCLUDED.row_number,\n                client_address_o = EXCLUDED.client_address_o,\n                id_number_f = EXCLUDED.id_number_f,\n                franchise_name = EXCLUDED.franchise_name,\n                relation = EXCLUDED.relation,\n                is_mem = EXCLUDED.is_mem,\n                retail_premium = EXCLUDED.retail_premium,\n                original_risk_premium = EXCLUDED.original_risk_premium,\n                mpia = EXCLUDED.mpia,\n                single_premium = EXCLUDED.single_premium,\n                r1_policy_fee = EXCLUDED.r1_policy_fee,\n                adv_fund_2_1_fee = EXCLUDED.adv_fund_2_1_fee,\n                risk_after_r1 = EXCLUDED.risk_after_r1,\n                new_risk_premium = EXCLUDED.new_risk_premium,\n                raw_data = EXCLUDED.raw_data\n        '
