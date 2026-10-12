import csv
import io
import sqlite3
import tempfile
import unittest
from pathlib import Path
from decimal import Decimal
from flask import Flask,g,render_template
from direct_data import stage_source,prepare_outputs
from direct_analytics import DirectStore,data_month,register_direct_analytics
from test_direct_data import record


class DirectAnalyticsTest(unittest.TestCase):
    def setUp(self):
        self.temp=tempfile.TemporaryDirectory();self.root=Path(self.temp.name)
        self.store=DirectStore(self.root/'analytics.db')
    def tearDown(self):self.temp.cleanup()

    def stage(self, rows):
        path=self.root/'source.csv'
        with path.open('w',newline='',encoding='utf-8') as f:csv.writer(f).writerows(rows)
        conn=sqlite3.connect(':memory:')
        stage_source(path,conn);prepare_outputs(conn,self.root/'reports','source.csv')
        return conn

    def save(self, rows, owner=1,month='2026-09',batch='batch-one'):
        conn=self.stage(rows)
        try:self.store.save(conn,batch,owner,month,'source.csv')
        finally:conn.close()

    def test_month_inference_and_mismatch(self):
        self.assertEqual(data_month('PolicyData_20251101_to_20251130.csv'),'2025-11')
        self.assertEqual(data_month('file.csv','2026-09'),'2026-09')
        for file,chosen in [('file.csv',''),('PolicyData_20251101_to_20251130.csv','2026-09'),('PolicyData_20260901_to_20261001.csv','')]:
            with self.assertRaises(ValueError):data_month(file,chosen)

    def test_exact_totals_month_replacement_and_owner_scope(self):
        self.save([record(),record(relation='MEMBEN',product='GROCERY DISCOUNT COUPONS FOR MARTINS CLIENTS',risk='17',retail='20',mpia='2')])
        overview=self.store.overview(1)
        self.assertEqual(overview['categories']['source']['original_risk'],Decimal('117.005'))
        self.assertEqual(overview['categories']['vouchers']['supplier'],Decimal('13'))
        self.assertEqual(overview['categories']['insurers']['rows'],1)
        self.assertFalse(self.store.overview(2)['available'])
        self.save([record(risk='200',retail='300')],batch='replacement')
        self.assertEqual(self.store.overview(1)['categories']['source']['rows'],1)
        self.assertEqual(self.store.overview(1)['categories']['source']['original_risk'],200)
        self.assertIsNone(self.store.row(1,'batch-one',1))
        self.save([record(risk='50',retail='80')],month='2026-08',batch='august')
        self.assertEqual(self.store.overview(1,'all')['categories']['source']['original_risk'],250)
        fresh=DirectStore(self.root/'analytics.db')
        self.assertEqual(fresh.overview(1)['categories']['source']['rows'],1)

    def test_pagination_plan_group_search_and_record_links(self):
        self.save([record(policy=f'POL{i:03d}') for i in range(61)])
        overview,rows,count,page=self.store.detail(1,None,'plan2',page=2)
        self.assertEqual((len(rows),count,page),(11,61,2))
        self.assertEqual(self.store.detail(1,None,'insurers',query='POL060')[2],1)
        self.assertEqual(self.store.detail(1,None,'vouchers')[2],0)
        self.assertEqual(self.store.detail(1,None,'source',destination='Alberton')[2],0)
        self.assertIsNone(self.store.row(2,'batch-one',1))
        app=Flask(__name__,template_folder=str(Path(__file__).parent/'templates'))
        who={'id':1,'role':'admin'}
        @app.before_request
        def user():g.user=who
        app.add_url_rule('/dashboard','dashboard',lambda:render_template('direct_dashboard.html'))
        app.add_url_rule('/direct-data','direct_data',lambda:'prepare')
        register_direct_analytics(app,self.store)
        client=app.test_client()
        response=client.get('/dashboard')
        self.assertIn(b'Direct Data Analytics',response.data)
        self.assertIn(b'category=plan2',response.data)
        self.assertIn(b'61',response.data)
        response=client.get('/direct-data/analytics?category=plan2&page=2')
        self.assertEqual(response.status_code,200)
        self.assertIn(b'Page 2',response.data)
        self.assertIn(b'POL060',response.data)
        response=client.get('/direct-data/analytics/record/batch-one/1')
        self.assertEqual(response.status_code,200)
        self.assertIn(b'0012312345678',response.data)
        who['id']=2;self.assertEqual(client.get('/direct-data/analytics/record/batch-one/1').status_code,404)
        who['role']='viewer';self.assertEqual(client.get('/direct-data/analytics').status_code,403)

if __name__=='__main__':unittest.main()
