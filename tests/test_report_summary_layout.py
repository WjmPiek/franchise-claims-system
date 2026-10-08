import io
import unittest
from decimal import Decimal
from reportlab.lib.pagesizes import landscape, A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Table
from pypdf import PdfReader
from report_summary import add_pdf_summary, add_excel_summary
import pandas as pd
from openpyxl import load_workbook

class ReportLayoutTest(unittest.TestCase):
    def result(self):
        return dict(status='Imported detail totals match this report.', count=390773,
            sources=[(month,'FRANCHISE SOURCE MUST NOT APPEAR', 'PolicyData_'+month+'.xlsx') for month in ('2024-08','2024-09','2024-10')],
            rows=[[month, 'Branch', 130000, Decimal('4309576.03'),Decimal('6946877.00'),Decimal('100'),Decimal('100'),Decimal('4309376.03'),Decimal('0'),Decimal('0'),Decimal('0')] for month in ('2024-08','2024-09','2024-10')],
            totals=[Decimal('13015639.50'),Decimal('20568892'),Decimal('102980'),Decimal('260639.24'),Decimal('12652020.26'),Decimal('0'),Decimal('0'),Decimal('0')],missing=[])

    def test_pdf_sources_removed_and_cells_fit(self):
        out=io.BytesIO()
        doc=SimpleDocTemplate(out,pagesize=landscape(A4),leftMargin=14,rightMargin=14,topMargin=16,bottomMargin=24)
        story=[]
        add_pdf_summary(story,getSampleStyleSheet(),self.result(),doc.width)
        for table in [t for t in story if isinstance(t,Table)]:
            width,height=table.wrap(doc.width,doc.height)
            self.assertLessEqual(width,doc.width+.01)
            for row in table._cellvalues:
                for i,value in enumerate(row):
                    if isinstance(value,str):
                        from reportlab.pdfbase.pdfmetrics import stringWidth
                        self.assertLessEqual(stringWidth(value,'Helvetica',9)+12,table._colWidths[i]+1)
        doc.build(story)
        reader=PdfReader(out)
        text=' '.join(page.extract_text() for page in reader.pages)
        self.assertNotIn('Source franchises:',text)
        self.assertNotIn('FRANCHISE SOURCE MUST NOT APPEAR',text)
        self.assertIn('PolicyData_2024-10.xlsx',text)
        self.assertEqual(len(reader.pages),1)

    def test_excel_sources_do_not_list_franchises(self):
        out=io.BytesIO()
        with pd.ExcelWriter(out,engine='xlsxwriter') as writer:
            add_excel_summary(writer,self.result())
        wb=load_workbook(io.BytesIO(out.getvalue()))
        cells=[cell.value for row in wb['Import Summary'] for cell in row]
        self.assertNotIn('Source franchise',cells)
        self.assertNotIn('FRANCHISE SOURCE MUST NOT APPEAR',cells)

if __name__=='__main__':
    unittest.main()
