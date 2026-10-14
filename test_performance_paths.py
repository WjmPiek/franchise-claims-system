import unittest,json
from unittest.mock import patch
from flask import g
from decimal import Decimal
import test_claim_insights as fixture
from claim_insights import populations,analyse
from performance_index import backfill

class PerformancePathsTest(fixture.ClaimInsightsTest):
    def test_published_product_statistics_need_no_policy_record_scan(self):
        self.publish([self.source()])
        with patch.object(self.store,'records',side_effect=AssertionError('Policy records were scanned')):
            result=analyse(self.store,self.dates)
        self.assertEqual(result['tables']['risk'][0]['premium'],Decimal('194.82'))
        self.assertEqual(result['events'][0]['cover'],'R40,000.00')
    def test_benefit_backfill_is_idempotent_and_preserves_source_and_amounts(self):
        with self.store.direct.connection() as conn:
            before=conn.execute('SELECT record_json FROM central_records').fetchone()[0]
            conn.execute('DELETE FROM central_benefit_versions');conn.execute('DELETE FROM central_product_benefits')
        result=backfill(self.store);self.assertEqual(result['indexed_months'],1)
        self.assertEqual(backfill(self.store)['indexed_months'],0)
        with self.store.direct.connection() as conn:self.assertEqual(before,conn.execute('SELECT record_json FROM central_records').fetchone()[0])
    def test_source_cache_filters_each_scope_and_returns_independent_records(self):
        self.publish([self.source()]);first=populations(self.store,self.dates)[0];first[0]['cover']='Modified caller copy'
        self.assertNotIn('cover',populations(self.store,self.dates)[0][0])
        self.assertFalse(populations(self.store,self.dates,{'OTHER BRANCH'})[0])
        self.publish([self.source('C2',**{'Net amount paid via bank':'230'})],batch='new')
        self.assertEqual(populations(self.store,self.dates)[1][0]['bank'],230)

if __name__=='__main__':unittest.main()
