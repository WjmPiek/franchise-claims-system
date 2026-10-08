"""PostgreSQL integration: cent-scale columns must not round source rows."""
import os
import unittest
from decimal import Decimal
from datetime import date
from sqlalchemy import create_engine, text, event
from premium_precision import ensure_premium_precision, prepare_premium_precision
from import_bulk import policy_detail_batches

@unittest.skipUnless(os.environ.get('PRECISION_TEST_DATABASE_URL'), 'requires isolated PostgreSQL')
class PremiumPrecisionDatabaseTest(unittest.TestCase):
    def setUp(self):
        self.engine = create_engine(os.environ['PRECISION_TEST_DATABASE_URL'])
        with self.engine.begin() as c:
            c.execute(text('DROP TABLE IF EXISTS policydata_detail_raw, policy_monthly_raw'))
            c.execute(text("""
                CREATE TABLE policydata_detail_raw (
                    source_file TEXT, import_month DATE, row_number INTEGER,
                    source_row_key TEXT, client_address_o TEXT, id_number_f TEXT,
                    franchise_name TEXT, relation TEXT, is_mem BOOLEAN,
                    retail_premium NUMERIC(18,2), original_risk_premium NUMERIC(18,2),
                    mpia NUMERIC, single_premium NUMERIC, r1_policy_fee NUMERIC,
                    adv_fund_2_1_fee NUMERIC, risk_after_r1 NUMERIC,
                    new_risk_premium NUMERIC, raw_data JSONB
                )
            """))
            c.execute(text('CREATE TABLE policy_monthly_raw (retail_premium NUMERIC(18,2), original_risk_premium NUMERIC(18,2))'))
            c.execute(text("INSERT INTO policy_monthly_raw VALUES (150.12,100.12)"))

    def tearDown(self):
        self.engine.dispose()

    def test_reports_can_read_while_excel_transaction_is_open(self):
        with policy_detail_batches(self.engine, 'August.xlsx', date(2024,8,1)) as write:
            # The import transaction remains open here while Excel is parsed.
            # A separate dashboard connection must not wait for ALTER TABLE locks.
            with self.engine.begin() as reader:
                reader.execute(text("SET LOCAL statement_timeout = '1s'"))
                self.assertEqual(reader.execute(text('SELECT COUNT(*) FROM policydata_detail_raw')).scalar_one(), 0)
                self.assertEqual(reader.execute(text('SELECT COUNT(*) FROM policy_monthly_raw')).scalar_one(), 1)
            write([dict(source_file='August.xlsx', import_month=date(2024,8,1),
                row_number=2, source_row_key='row', client_address_o='', id_number_f='',
                franchise='ERMELO', relation='MEM', is_mem=True, retail_premium=Decimal('150.004'),
                original_risk_premium=Decimal('100.004'), mpia=1, single_premium=100,
                r1_policy_fee=1, adv_fund_2_1_fee=Decimal('2.08'), risk_after_r1=99,
                new_risk_premium=Decimal('96.92'), raw_data={'Risk':100.004})])

    def test_busy_migration_fails_before_replacing_existing_rows(self):
        import time
        with self.engine.begin() as locker:
            locker.execute(text('SELECT COUNT(*) FROM policydata_detail_raw'))
            started = time.monotonic()
            with self.assertRaisesRegex(RuntimeError, 'No rows.*have been replaced'):
                prepare_premium_precision(self.engine)
            self.assertLess(time.monotonic() - started, 10)
        with self.engine.begin() as check:
            self.assertEqual(check.execute(text('SELECT original_risk_premium FROM policy_monthly_raw')).scalar_one(), Decimal('100.12'))

    def test_widening_copy_totals_and_idempotent_reimport(self):
        batch = []
        for i in range(3):
            batch.append(dict(source_file='August.xlsx', import_month=date(2024,8,1),
                row_number=i+2, source_row_key=str(i), client_address_o='', id_number_f='',
                franchise='ERMELO', relation='MEM', is_mem=True, retail_premium=Decimal('150.004'),
                original_risk_premium=Decimal('100.004'), mpia=1, single_premium=100,
                r1_policy_fee=1, adv_fund_2_1_fee=Decimal('2.08'), risk_after_r1=99,
                new_risk_premium=Decimal('96.92'), raw_data={'Risk':100.004}))
        for unused in range(2):
            with policy_detail_batches(self.engine, 'August.xlsx', date(2024,8,1)) as write:
                write(batch)
        with self.engine.begin() as c:
            risk, retail, count = c.execute(text('SELECT SUM(original_risk_premium), SUM(retail_premium), COUNT(*) FROM policydata_detail_raw')).one()
            self.assertEqual((risk,retail,count), (Decimal('300.012'),Decimal('450.012'),3))
            self.assertEqual(c.execute(text('SELECT original_risk_premium FROM policy_monthly_raw')).scalar_one(), Decimal('100.12'))
            c.execute(text('INSERT INTO policy_monthly_raw VALUES (150.004,100.004)'))
            statements = []
            def capture(conn,cursor,statement,parameters,context,executemany):
                statements.append(statement)
            event.listen(self.engine, 'before_cursor_execute', capture)
            try:
                ensure_premium_precision(c)
            finally:
                event.remove(self.engine, 'before_cursor_execute', capture)
            self.assertFalse(any('ALTER TABLE' in s for s in statements))
            self.assertEqual(c.execute(text('SELECT MAX(retail_premium) FROM policy_monthly_raw')).scalar_one(), Decimal('150.12'))
            self.assertEqual(c.execute(text('SELECT MIN(original_risk_premium) FROM policy_monthly_raw')).scalar_one(), Decimal('100.004'))

if __name__ == '__main__':
    unittest.main()
