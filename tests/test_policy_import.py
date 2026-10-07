import ast
import os
import re
import unittest
from pathlib import Path
from datetime import datetime
from decimal import Decimal, ROUND_HALF_UP
import pandas as pd
from flask import Flask
from policy_reports import register_policy_reports


class ImportAndReportsTest(unittest.TestCase):
    def setUp(self):
        header = ['Franchise'] + [f'C{i}' for i in range(1,18)]
        for i,name in [(8,'Relation'),(12,'Risk'),(13,'Retail'),(17,'MPIA')]: header[i]=name
        data=[]
        for relation,risk,retail,mpia in [('MEM',100,150,2),('EXT',20,30,1),('',5,8,1),('EXT',-3,-4,1),('MEM',0,0,1)]:
            row=[None]*18
            row[0],row[8],row[12],row[13],row[17]='ERMELO',relation,risk,retail,mpia
            data.append(row)
        class Sheet:
            def iter_rows(self, max_row=None, **kwargs):return iter([header] if max_row==1 else data)
        class Book:
            sheetnames=['Sheet1']
            def __getitem__(self,name):return Sheet()
            def close(self):pass
        names={'clean_money','month_from_filename','read_policydata_streaming','_normalise_header','display_source_filename','parse_policy_transaction_sheet','_underwriter_allocation_view'}
        source=Path(__file__).resolve().parents[1].joinpath('app.py').read_text(encoding='utf-8')
        tree=ast.Module(body=[n for n in ast.parse(source).body if isinstance(n,ast.FunctionDef) and n.name in names],type_ignores=[])
        self.env=dict(pd=pd,os=os,re=re,datetime=datetime,Decimal=Decimal,ROUND_HALF_UP=ROUND_HALF_UP,load_workbook=lambda *a,**k:Book(),_clean_map_value=lambda v:v,_clean_id_number=lambda v:v)
        exec(compile(tree,'import-functions','exec'),self.env)
        self.total=self.env['read_policydata_streaming']('PolicyData_20260901_to_20260930.xlsx').iloc[0]
        self.detail=self.env['LAST_POLICY_DETAIL_DF']
        app=Flask(__name__,template_folder=str(Path(__file__).resolve().parents[1]/'templates'))
        app.add_url_rule('/dashboard',endpoint='dashboard',view_func=lambda:'dashboard')
        register_policy_reports(app,lambda:None,None,lambda:self.detail)
        self.client=app.test_client()
        self.rows=data
        self.headers=header

    def test_all_relations_blanks_and_negative_adjustments(self):
        self.assertAlmostEqual(self.total.original_risk_premium,122)
        self.assertAlmostEqual(self.total.retail_premium,184)
        self.assertAlmostEqual(self.total.risk_premium,117.94)
        self.assertEqual(self.total.policy_qty,3)
        self.assertEqual(len(self.detail),5)
        self.assertFalse(self.detail[~self.detail.is_mem].adv_fund_2_1_fee.any())
        self.assertAlmostEqual(self.detail.new_risk_premium.sum(),self.total.risk_premium)

    def test_blank_relation_export_and_controls(self):
        response=self.client.get('/policy_reports?relation=(Blanks)&download=detail')
        self.assertEqual(response.status_code,200)
        self.assertEqual(len(response.data.decode('utf-8-sig').splitlines()),2)
        response=self.client.get('/policy_reports?external_risk=122&external_retail=184')
        self.assertEqual(response.status_code,200)
        self.assertEqual(response.data.count(b': Balanced'),2)
        self.assertEqual(self.client.get('/policy_reports?external_risk=NaN').status_code,400)
        response=self.client.get('/policy_reports?relation=DEP&external_risk=0')
        self.assertIn(b'comparison unavailable',response.data)

    def test_months_reimport_and_fallback(self):
        for month in range(1,13):
            total=self.env['read_policydata_streaming'](f'PolicyData_2026{month:02d}01_to_2026{month:02d}28.xlsx').iloc[0]
            self.assertEqual(total.month,pd.Timestamp(2026,month,1))
            self.assertAlmostEqual(total.retail_premium,184)
            self.assertAlmostEqual(total.risk_premium,117.94)
        fallback=self.env['parse_policy_transaction_sheet'](pd.DataFrame(self.rows,columns=self.headers),'PolicyData_20260901_to_20260930.xlsx').iloc[0]
        self.assertAlmostEqual(fallback.retail_premium,184)
        self.assertAlmostEqual(fallback.risk_premium,117.94)

    def test_payover_equals_calculated_net_and_scope_is_enforced(self):
        frame=pd.DataFrame([{'Franchise':'ERMELO','Risk Premium':117.94,'Original Risk Premium':122,'Retail Premium':184,'R1 Policy Fee':3,'ADV Fee 2.1%':2.06}])
        report=self.env['_underwriter_allocation_view'](frame)
        self.assertAlmostEqual(report.iloc[0]['Underwriter Premium Payable'],117.94)
        self.assertAlmostEqual(report.iloc[0]['Retail Less Payover'],66.06)
        app=Flask('scoped',template_folder=str(Path(__file__).resolve().parents[1]/'templates'))
        app.add_url_rule('/dashboard',endpoint='dashboard',view_func=lambda:'dashboard')
        register_policy_reports(app,lambda:None,None,lambda:self.detail,lambda frame:frame.iloc[0:0])
        client=app.test_client()
        self.assertEqual(client.get('/policy_reports?franchise=ERMELO&download=detail').status_code,400)
        self.assertNotIn(b'ERMELO',client.get('/policy_reports?download=detail').data)

if __name__=='__main__':unittest.main()
