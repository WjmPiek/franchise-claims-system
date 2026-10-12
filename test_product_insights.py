import unittest
from test_direct_analytics import DirectAnalyticsTest
from test_direct_data import record

class ProductInsightsTest(DirectAnalyticsTest):
    def test_counts_products_without_counting_payment_rows_as_new_policies(self):
        self.save([record(product='Family',cover='40000',mpia='2'),record(product='Family',cover='40000',mpia='1'),record(product='Cremation',cover='20000'),record(relation='SPS',product='Family',cover='40000'),record(relation='MEM',product='GROCERY DISCOUNT COUPONS FOR MARTINS CLIENTS',risk='8.5')])
        before=self.store.overview(1)['categories']
        result=self.store.product_insights(1)
        self.assertEqual(sum(row['policies'] for row in result['rows']),2)
        self.assertEqual(result['top'][0]['value'],2)
        self.assertEqual(next(r for r in result['rows'] if r['product']=='Family')['paid_months'],3)
        self.assertEqual(self.store.overview(1)['categories'],before)
        self.assertEqual(self.store.product_insights(2)['rows'],[])
        self.assertEqual(len(self.store.product_insights(1,product='Family')['rows']),1)
        self.save([record(product='Replacement')],batch='new')
        self.assertEqual(self.store.product_insights(1)['products'],['Replacement'])
    def test_date_range_counts_unique_policies_and_includes_all_payments(self):
        self.save([record(product='Family',risk='100')],month='2026-08',batch='aug')
        self.save([record(product='Family',risk='120')],month='2026-09',batch='sep')
        result=self.store.product_insights(1,start='2026-08',end='2026-09')
        self.assertEqual(result['top'][0]['value'],1)
        self.assertEqual(result['rows'][0]['paid_months'],2)
        self.assertEqual(result['rows'][0]['risk'],220)
        overview,rows,count,page=self.store.detail(1,'all','source',start='2026-09',end='2026-09')
        self.assertEqual(count,1);self.assertEqual(rows[0]['month'],'2026-09')
        self.assertEqual(overview['categories']['source']['original_risk'],120)
        self.assertEqual(self.store.product_insights(2,start='2026-08',end='2026-09')['rows'],[])

    def test_top_ten_and_single_month_scope(self):
        self.save([record(franchise=f'Branch {i}',policy=f'P{j}',product='Family',cover='20000') for i in range(12) for j in range(i+1)])
        self.save([record(franchise='Old Branch',product='Old')],month='2026-08',batch='old')
        result=self.store.product_insights(1)
        self.assertEqual(result['month'],'2026-09')
        self.assertEqual(len(result['top']),10)
        self.assertEqual(result['top'][0]['label'],'Branch 11')
        self.assertNotIn('Old',result['products'])
        self.assertEqual(self.store.product_insights(1,'2026-08')['products'],['Old'])

if __name__=='__main__':unittest.main()
