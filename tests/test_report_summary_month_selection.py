import ast
import io
import os
import re
import uuid
import tempfile
import unittest
from pathlib import Path
from datetime import datetime
import pandas as pd
from flask import Flask, request, render_template, url_for, send_file
from reportlab.lib.pagesizes import landscape, A4
from reportlab.lib.units import cm
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Paragraph, Table, PageBreak
from pypdf import PdfReader

ROOT = Path(__file__).resolve().parents[1]

class ReportMonthSelectionTest(unittest.TestCase):
    def setUp(self):
        self.app = Flask('month-tests', template_folder=str(ROOT / 'templates'))
        self.app.add_url_rule('/dashboard', endpoint='dashboard', view_func=lambda:'')
        self.app.add_url_rule('/export', endpoint='export', view_func=lambda:'')
        self.app.add_url_rule('/board_report', endpoint='board_report', view_func=lambda:'')
        data = []
        for month, risk in [('2024-08',80.004),('2024-09',100.004),('2024-10',200.004)]:
            for franchise in ['ERMELO','SOWETO']:
                data.append({'Franchise':franchise,'Month':pd.Timestamp(month+'-01'),
                    'Original Risk Premium':risk,'Risk Premium':risk-5,
                    'Retail Premium':risk+50,'Policy Qty':1,'Claims':0,
                    'Total Commission':15,'MFF Book Value 2.5%':2.5,
                    'Franchise Book Value 2.5%':2.5,'Total Book Value':5})
        self.monthly = pd.DataFrame(data)
        names = {'_select_report_months','_friendly_export_df','_financial_report_frames','_financial_download'}
        tree = ast.parse(ROOT.joinpath('app.py').read_text(encoding='utf-8'))
        nodes = [n for n in tree.body if (isinstance(n,ast.FunctionDef) and n.name in names)
                 or (isinstance(n,ast.Assign) and any(isinstance(t,ast.Name) and t.id == 'FINANCIAL_REPORT_COLUMNS' for t in n.targets))]
        def configured(frame, config):
            periods = frame.copy()
            periods['Period'] = periods.Month.dt.strftime('%b %Y')
            return frame.copy(), periods, {'total_retail':frame['Retail Premium'].sum()}
        self.env = dict(pd=pd,re=re,uuid=uuid,os=os,datetime=datetime,request=request,
            render_template=render_template,url_for=url_for,send_file=send_file,
            load_franchise_config=lambda:{},apply_franchise_config=configured,
            LAST_RESULT={'monthly':self.monthly},apply_user_franchise_scope=lambda x:x,
            _import_report_summary=lambda *a:{}, add_excel_summary=lambda *a:None,
            _add_excel_cover=lambda *a:None,_format_excel_sheet=lambda *a:None,
            landscape=landscape,A4=A4,cm=cm,SimpleDocTemplate=SimpleDocTemplate,
            getSampleStyleSheet=getSampleStyleSheet,Paragraph=Paragraph,PageBreak=PageBreak,
            _page_footer=lambda *a:None,add_pdf_summary=lambda *a:None,
            _add_pdf_cover=lambda story,styles,title,label:story.append(Paragraph(title+' '+label,styles['Normal'])),
            _rows_for_pdf=lambda frame,cols,**kw:[cols]+frame.astype(str).values.tolist(),
            _pdf_table=lambda rows,*a:Table(rows))
        exec(compile(ast.Module(body=nodes,type_ignores=[]),'report-functions','exec'),self.env)

    def test_explicit_and_analysis_month_never_choose_latest(self):
        for args in [{'report_period':'month','report_month':'2024-09'},
                     {'report_month':'2024-09'},
                     {'premium_analysis_month':'2024-09'}]:
            work,label = self.env['_select_report_months'](self.monthly,args,'month')
            self.assertEqual(label,'Sep 2024')
            self.assertEqual(set(work.Month.dt.strftime('%Y-%m')),{'2024-09'})
        work,label = self.env['_select_report_months'](self.monthly,
            {'report_period':'month','report_month':'2024-08','premium_analysis_month':'2024-09'})
        self.assertEqual(label,'Aug 2024')
        for args in [{'report_month':'2024-13'}, {'report_month':'2023-09'},
                     {'report_period':'range','report_start':'2024-10','report_end':'2024-09'}]:
            with self.assertRaises(ValueError):
                self.env['_select_report_months'](self.monthly,args)
        work,label = self.env['_select_report_months'](self.monthly,
            {'report_period':'range','report_start':'2024-08','report_end':'2024-09'})
        self.assertEqual(set(work.Month.dt.strftime('%Y-%m')),{'2024-08','2024-09'})

    def test_september_excel_and_pdf_both_report_types(self):
        with tempfile.TemporaryDirectory() as directory:
            self.env['EXPORT_DIR'] = directory
            for report_type in ['commissions','book_value']:
                for output in ['xlsx','pdf']:
                    with self.app.test_request_context('/export?report_type='+report_type+'&report_period=month&report_month=2024-09'):
                        response = self.env['_financial_download'](output)
                        response.direct_passthrough = False
                        body = response.get_data()
                        self.assertIn('Sep_2024',response.headers['Content-Disposition'])
                        if output == 'xlsx':
                            frame = pd.read_excel(io.BytesIO(body),sheet_name='Monthly Detail')
                            self.assertEqual(set(pd.to_datetime(frame.Month).dt.strftime('%Y-%m')),{'2024-09'})
                            self.assertAlmostEqual(frame['Retail Premium'].sum(),300.008)
                            if report_type == 'commissions':
                                self.assertAlmostEqual(frame['Risk Premium'].sum(),200.008)
                                self.assertAlmostEqual(frame['Payover Less Comm.'].sum(),190.008)
                                self.assertNotIn('Original Risk Premium',frame.columns)
                                self.assertFalse(frame.columns.duplicated().any())
                        else:
                            text = ' '.join(p.extract_text() for p in PdfReader(io.BytesIO(body)).pages)
                            self.assertIn('Sep 2024',text)
                            self.assertNotIn('Oct 2024',text)
                            self.assertNotIn('200.004',text)
            # Repeated requests have independent files even in the same second.
            paths = list(Path(directory).iterdir())
            self.assertEqual(len(paths),4)

    def test_gross_display_mapping_preserves_source_and_derived_amounts(self):
        friendly = self.env['_friendly_export_df'](self.monthly)
        self.assertEqual(friendly['Risk Premium'].tolist(),self.monthly['Original Risk Premium'].tolist())
        self.assertEqual(friendly['Payover Less Comm.'].tolist(),self.monthly['Risk Premium'].tolist())
        self.assertIn('Original Risk Premium',self.monthly.columns)

    def test_download_form_initializes_to_selected_analysis_month(self):
        source = ROOT.joinpath('templates/dashboard.html').read_text(encoding='utf-8')
        start = source.index('<select name="report_month" required>')
        fragment = source[start:source.index('</select>',start)+9]
        with self.app.test_request_context('/dashboard?premium_analysis_month=2024-09'):
            html = self.app.jinja_env.from_string(fragment).render(
                report_selected_month='2024-09',
                report_month_options=[{'value':'2024-09','label':'Sep 2024'},{'value':'2024-10','label':'Oct 2024'}])
        self.assertRegex(html,r'value="2024-09" selected')
        self.assertNotRegex(html,r'value="2024-10" selected')

if __name__ == '__main__':
    unittest.main()
