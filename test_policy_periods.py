import io
import unittest
import pandas as pd
from flask import Flask
from policy_reports import register_policy_reports,MONEY

class PolicyPeriodTest(unittest.TestCase):
    def setUp(self):
        app=Flask(__name__,template_folder='templates')
        app.add_url_rule('/dashboard','dashboard',lambda:'dashboard')
        data=pd.DataFrame([dict(import_month='2026-'+month+'-01',franchise='Faithnet',relation='MEM',source_file='source.xlsx',row_number=1,raw_data={'Policy':'0001'},**{key:amount for key in MONEY}) for month,amount in [('08',100),('09',120),('10',140)]])
        register_policy_reports(app,lambda:None,None,lambda:data,lambda frame:frame)
        self.client=app.test_client()
    def test_range_keeps_exact_totals_and_row_months(self):
        args='?date_from=2026-08-17&date_to=2026-09-19'
        response=self.client.get('/policy_reports'+args)
        self.assertEqual(response.status_code,200)
        self.assertIn(b'220.00',response.data)
        detail=self.client.get('/policy_reports'+args+'&download=detail')
        self.assertIn(b'2026-08',detail.data);self.assertIn(b'2026-09',detail.data);self.assertNotIn(b'2026-10',detail.data)
        for kind in ['xlsx','pdf']:
            response=self.client.get('/policy_reports'+args+'&download='+kind)
            self.assertEqual(response.status_code,200);response.close()
    def test_invalid_dates_rejected(self):
        for args in ['date_from=bad&date_to=2026-09-01','date_from=2026-10-01&date_to=2026-09-01','date_from=2026-09-01']:
            self.assertEqual(self.client.get('/policy_reports?'+args).status_code,400)

if __name__=='__main__':unittest.main()
