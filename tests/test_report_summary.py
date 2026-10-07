import unittest
from decimal import Decimal
import pandas as pd
from report_summary import reconciliation


class ReconciliationTest(unittest.TestCase):
    def setUp(self):
        self.monthly = pd.DataFrame([
            dict(Franchise=f, Month=pd.Timestamp(m), **{'Original Risk Premium': 100,
                 'Retail Premium': 150, 'Risk Premium': 96.92})
            for f, m in [('A', '2026-09-01'), ('B', '2026-09-01'), ('C', '2026-09-01'), ('A', '2026-08-01')]])
        records = []
        for f, m in [('A', '2026-09-01'), ('B', '2026-09-01'), ('C', '2026-09-01'), ('A', '2026-08-01')]:
            for risk, retail, r1, adv, net in [(120, 180, 1, 2.08, 116.92), (-20, -30, 0, 0, -20)]:
                records.append(dict(franchise=f, import_month=pd.Timestamp(m), source_file=f'Policy_{m}.xlsx',
                                    original_risk_premium=risk, retail_premium=retail,
                                    r1_policy_fee=r1, adv_fund_2_1_fee=adv, new_risk_premium=net))
        self.detail = pd.DataFrame(records)

    def test_individual_scope_and_negative_rows(self):
        result = reconciliation(self.monthly.iloc[:1], memory=self.detail)
        self.assertEqual(result['count'], 2)
        self.assertEqual(result['totals'][:2], [Decimal('100'), Decimal('150')])
        self.assertEqual(result['totals'][-3:], [0, 0, 0])
        self.assertIn('match', result['status'])
        self.assertEqual({s[1] for s in result['sources']}, {'A'})

    def test_group_exclusions_and_month_range(self):
        result = reconciliation(self.monthly, memory=self.detail,
            config={'excluded': ['C'], 'use_groups': True, 'groups': {'A': 'Group', 'B': 'Group'}})
        self.assertEqual(result['count'], 6)
        self.assertEqual(len(result['rows']), 2)
        self.assertEqual(result['rows'][1][:3], ['2026-09', 'Group', 4])
        self.assertEqual(result['totals'][0], 300)
        self.assertNotIn('C', {s[1] for s in result['sources']})

    def test_missing_detail_and_actual_difference(self):
        missing = reconciliation(self.monthly.iloc[:1], memory=pd.DataFrame())
        self.assertIn('unavailable', missing['status'])
        self.assertEqual(missing['rows'][0][8], None)
        changed = self.monthly.iloc[:1].copy()
        changed['Retail Premium'] = 151
        result = reconciliation(changed, memory=self.detail)
        self.assertEqual(result['rows'][0][9], 1)
        self.assertIn('review', result['status'])

    def test_database_query_and_exact_scope_pairs(self):
        # Broad indexed query is narrowed to the exact franchise/month pairs afterwards.
        records = self.detail.rename(columns={'franchise': 'franchise_name'}).to_dict('records')
        for row in records:
            row['rows'] = 1
        class Result:
            def mappings(self): return records
        class Connection:
            def __enter__(self): return self
            def __exit__(self, *args): pass
            def execute(self, query, params):
                self.params = params
                return Result()
        class Engine:
            def connect(self): return Connection()
        selected = self.monthly.iloc[[1, 3]]
        result = reconciliation(selected, engine=Engine(), sql_text=lambda s: s)
        self.assertEqual(result['count'], 4)
        self.assertEqual(result['totals'][0], 200)
        self.assertEqual({(s[1], s[0]) for s in result['sources']}, {('B', '2026-09'), ('A', '2026-08')})

if __name__ == '__main__':
    unittest.main()
