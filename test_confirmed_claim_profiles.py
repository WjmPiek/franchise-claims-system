import csv,json,sqlite3,tempfile,unittest
from pathlib import Path
from contextlib import closing
from openpyxl import Workbook
from decimal import Decimal
from central_claims import read_transactions,component
from central_identity import member_snapshot
import test_central_compliance as base
from test_direct_data import record
from central_reports import COMBINE

class ConfirmedProfiles(unittest.TestCase):
    setUp=base.CentralComplianceTest.setUp
    tearDown=base.CentralComplianceTest.tearDown
    import_rows=base.CentralComplianceTest.import_rows
    def test_bank_rule_keeps_netoffs_without_double_counting_claims(self):
        path=self.root/'bank.xlsx';book=Workbook();sheet=book.active
        sheet.append(['Final Amount','Claim Status','Net amount paid via bank','Date Of Claim Paid','Participating Group','Amount net off against claim'])
        sheet.append([100,'Approved and paid',70,'2026-09-10','ERMELO',30])
        sheet.append([50,'Approved and net off',0,'2026-09-11','ERMELO',50])
        sheet.append([20000,'Repudiated by Approver',0,None,'ERMELO',350])
        book.save(path);controls={};rows=list(read_transactions(path,'Net amount paid via bank',controls=controls))
        self.assertEqual([Decimal(r['amount']) for r in rows],[Decimal(70),Decimal(0)])
        self.assertEqual(sum(Decimal(r['amount'])!=0 for r in rows),1)
        self.assertEqual(sum(component(r['source'],'Amount net off against claim') for r in rows),80)
        self.assertEqual(controls['netoff_only_rows'],1);self.assertEqual(controls['undated_netoff_amount'],'350')
        self.assertEqual(controls['source_rows'],3)
    def test_paid_components_and_rejected_rows(self):
        path=self.root/'claims.xlsx';book=Workbook();s=book.active
        s.append(['Participating Group','Date Of Claim Paid','Final Amount','Amount net off against claim','Net amount paid via bank','Claim Status'])
        s.append(['ERMELO','2026-09-10',100,30,70,'Paid'])
        s.append(['ERMELO',None,200,0,0,'Claim - Repudiated by Approver'])
        book.save(path);controls={};rows=list(read_transactions(path,'Final Amount',controls=controls))
        self.assertEqual(len(rows),1);self.assertEqual(controls['unpaid_rows'],1)
        self.assertEqual(component(rows[0]['source'],'Amount net off against claim'),Decimal(30))
        s.append(['ERMELO',None,300,0,300,'Claim - Awaiting outstanding information']);book.save(path)
        with self.assertRaisesRegex(ValueError,'payment components'):list(read_transactions(path,'Final Amount'))
    def test_latest_snapshot_uses_system_ref_and_preserves_all_rows(self):
        a=record();a[16]='ref1';b=list(a);b[16]='ref2'
        self.import_rows([a,a,b],'september','2026-09')
        self.import_rows([a],'august','2026-08')
        snap=member_snapshot(self.store,'2026-08','2026-09')
        self.assertEqual(snap['count'],2);self.assertEqual(snap['month'],'2026-09')
        self.assertEqual(len(list(self.store.records('september'))),3)
        self.assertEqual(member_snapshot(self.store,'2026-08','2026-09',set())['count'],0)
    def test_aul_risk_conflict_uses_original_risk(self):
        from direct_data import HEADERS,stage_source,prepare_outputs
        path=self.root/'extra.csv'
        with path.open('w',newline='',encoding='utf-8') as f:csv.writer(f).writerows([HEADERS+['AUL Risk'],record(risk='100')+['200']])
        with closing(sqlite3.connect(self.root/'extra.sqlite')) as stage:
            controls=stage_source(path,stage)
            prepare_outputs(stage,self.root/'unused',path.name,'mpia','mpia','policy',generate_files=False)
            self.direct.save(stage,'extra',1,'2026-09',path.name)
        row=next(self.store.records('extra'))
        self.assertEqual(row['original_risk'],'100');self.assertEqual(row['source_extra']['AULRISK'],'200')
        self.assertTrue(any('Risk / AUL Risk conflict' in issue for issue in row['issues']))
    def test_confirmed_combinations(self):
        self.assertEqual(COMBINE['FAITHNETSA'],COMBINE['MARTINSFUNERALSBROKERS'])
        self.assertEqual(COMBINE['WESTONARIA'],COMBINE['MARTINSFUNERALSRANDFONTEIN'])
    def test_partial_month_scope_preserves_other_publications(self):
        from central_claims import publish_claims
        def stage_rows(values):
            stage=sqlite3.connect(':memory:');stage.execute('CREATE TABLE claims(n INTEGER PRIMARY KEY,month TEXT,franchise TEXT,product TEXT,amount TEXT,source TEXT,issues TEXT)')
            for n,(month,amount) in enumerate(values,1):stage.execute('INSERT INTO claims VALUES (?,?,?,?,?,?,?)',(n,month,'ERMELO','Plan',str(amount),json.dumps({'Net amount paid via bank':str(amount)}),'[]'))
            return stage
        with closing(stage_rows([('2026-08',100),('2026-09',200)])) as stage:publish_claims(self.store,stage,'first.xlsx',1,'Net amount paid via bank',['2026-08','2026-09'])
        with closing(stage_rows([('2026-09',250),('2026-10',50)])) as stage:
            result=publish_claims(self.store,stage,'second.xlsx',1,'Net amount paid via bank',['2026-09'])
        self.assertEqual(result['unpublished_source_months'],['2026-10'])
        rows=list(self.store.claims());self.assertEqual({r['month']:r['amount'] for r in rows},{'2026-08':'100','2026-09':'250'})

if __name__=='__main__':unittest.main()
