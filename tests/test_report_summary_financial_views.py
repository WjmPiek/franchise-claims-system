import ast
import os
import re
import uuid
import tempfile
import unittest
from pathlib import Path
from datetime import datetime
from unittest.mock import Mock
import pandas as pd
from flask import Flask, request, render_template, url_for
from report_summary import HEADERS

ROOT = Path(__file__).resolve().parents[1]

class FinancialPreviewTest(unittest.TestCase):
    def setUp(self):
        self.app = Flask(__name__, template_folder=str(ROOT / 'templates'))
        self.app.add_url_rule('/dashboard', endpoint='dashboard', view_func=lambda: '')
        self.app.add_url_rule('/export', endpoint='export', view_func=lambda: '')
        self.app.add_url_rule('/board_report', endpoint='board_report', view_func=lambda: '')
        source = ROOT.joinpath('app.py').read_text(encoding='utf-8')
        node = next(n for n in ast.parse(source).body if isinstance(n, ast.FunctionDef) and n.name == '_financial_download')
        self.env = dict(uuid=uuid, re=re, request=request, render_template=render_template, url_for=url_for,
            LAST_RESULT={'monthly': pd.DataFrame({'Month': ['2024-08-01']})},
            apply_user_franchise_scope=lambda x: x,
            _select_report_months=lambda x, args: (x, 'Aug 2024 - Oct 2024'),
            _financial_report_frames=lambda *args: [pd.DataFrame()] * 3,
            _import_report_summary=lambda *args: {}, load_franchise_config=lambda: {})
        exec(compile(ast.Module(body=[node], type_ignores=[]), 'app.py', 'exec'), self.env)

    def test_both_views_embed_matching_pdf_and_preserve_scope(self):
        for report_type, title in [('book_value', 'Book Value'), ('commissions', 'Commissions')]:
            with self.app.test_request_context('/export?report_type='+report_type+'&report_format=view&report_period=range&report_start=2024-08&report_end=2024-10&franchise=ERMELO'):
                html = self.env['_financial_download']('view')
                self.assertIn('type="application/pdf"', html)
                self.assertIn(title+' PDF', html)
                self.assertIn('preview=1', html)
                self.assertIn('report_start=2024-08', html)
                self.assertIn('report_end=2024-10', html)
                self.assertIn('franchise=ERMELO', html)
                self.assertNotIn('report_format=view', html)
                self.assertNotIn('report-table', html)

    def test_view_does_not_generate_report_twice(self):
        frames = Mock()
        audit = Mock()
        self.env['_financial_report_frames'] = frames
        self.env['_import_report_summary'] = audit
        with self.app.test_request_context('/export?report_type=commissions&report_format=view&report_period=month&report_month=2024-09'):
            html = self.env['_financial_download']('view')
        self.assertIn('application/pdf', html)
        frames.assert_not_called()
        audit.assert_not_called()

    def test_pdf_inline_only_for_preview(self):
        from reportlab.lib.pagesizes import landscape, A4
        from reportlab.lib.units import cm
        from reportlab.lib.styles import getSampleStyleSheet
        from reportlab.platypus import SimpleDocTemplate, Paragraph, PageBreak
        with tempfile.TemporaryDirectory() as folder:
            self.env.update(os=os, datetime=datetime, EXPORT_DIR=folder, landscape=landscape,
                A4=A4, cm=cm, SimpleDocTemplate=SimpleDocTemplate,
                getSampleStyleSheet=getSampleStyleSheet, Paragraph=Paragraph, PageBreak=PageBreak,
                _add_pdf_cover=lambda *a: None, add_pdf_summary=lambda *a: None,
                _pdf_table=lambda *a: Paragraph('Report detail', getSampleStyleSheet()['Normal']),
                _rows_for_pdf=lambda *a, **k: [], _page_footer=lambda *a: None, pd=pd)
            for preview, attachment in [('', True), ('&preview=1', False)]:
                send = Mock(return_value='pdf response')
                self.env['send_file'] = send
                with self.app.test_request_context('/board_report?report_type=commissions'+preview):
                    self.env['_financial_download']('pdf')
                self.assertEqual(send.call_args.kwargs['as_attachment'], attachment)

    def test_summary_headings(self):
        self.assertIn('Risk Premium', HEADERS)
        self.assertIn('Retail Premium', HEADERS)
        self.assertIn('Payover Less Comm.', HEADERS)
        self.assertNotIn('Original risk', HEADERS)
        self.assertNotIn('Risk (Payover)', HEADERS)

if __name__ == '__main__':
    unittest.main()
