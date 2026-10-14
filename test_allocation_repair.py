import json,sqlite3,unittest,os
os.environ['DATABASE_URL']=''
os.environ['APP_ENV']='local'
from contextlib import closing
from unittest.mock import patch
from allocation_repair import repair,UNRESOLVED,AUL_EXAMPLE_NOTE
import test_central_compliance as base
from test_direct_data import record
from central_claims import publish_claims

class AllocationRepairTest(unittest.TestCase):
    setUp=base.CentralComplianceTest.setUp
    tearDown=base.CentralComplianceTest.tearDown
    import_rows=base.CentralComplianceTest.import_rows
    def test_obsolete_example_warning_removed_but_real_note_preserved(self):
        self.import_rows([record(franchise='ANOTHER AUL BRANCH')])
        with self.direct.connection() as conn:
            r=json.loads(conn.execute('SELECT record_json FROM central_records').fetchone()[0]);r['issues']=[AUL_EXAMPLE_NOTE,'Missing voucher contact']
            conn.execute('UPDATE central_records SET record_json=?',(json.dumps(r),))
            s=json.loads(conn.execute('SELECT summary FROM central_monthly').fetchone()[0]);s['review_rows']=1
            conn.execute('UPDATE central_monthly SET summary=?',(json.dumps(s),));conn.execute('UPDATE central_batches SET review_rows=1')
            conn.execute('CREATE TABLE central_import_verification(batch_id TEXT PRIMARY KEY,kind TEXT,checked_at TEXT,result TEXT)')
            conn.execute('INSERT INTO central_import_verification VALUES (?,?,?,?)',('batch1','policy','today',json.dumps(dict(matched=True,controls=[dict(difference='0')],review=1))))
        result=repair(self.store,1)
        self.assertEqual(result['obsolete_notes_cleared'],1)
        after=list(self.store.records('batch1'))[0];self.assertEqual(after['issues'],['Missing voucher contact']);self.assertEqual(after['calculation'],r['calculation'])
        with self.direct.connection() as conn:
            checked=json.loads(conn.execute('SELECT result FROM central_import_verification').fetchone()[0]);self.assertTrue(checked['matched']);self.assertEqual(checked['review_reasons'],{'Missing voucher contact':1})
        self.assertEqual(repair(self.store,1)['obsolete_notes_cleared'],0)
    def test_plan_repair_is_logged_preserves_amounts_and_is_idempotent(self):
        self.import_rows([record(product='Older unnamed family plan',cover='7000')])
        with self.direct.connection() as conn:
            r=json.loads(conn.execute('SELECT record_json FROM central_records').fetchone()[0]);r['plan']='Unclassified';r['issues']=['[BLOCK] AUL product or related policy plan is unclassified; record retained in AUL Unclassified.']
            conn.execute("UPDATE central_records SET plan='Unclassified',record_json=?",(json.dumps(r),))
            s=json.loads(conn.execute('SELECT summary FROM central_monthly').fetchone()[0]);s['aul_historic_rows']=s['aul_historic_payover']=0;s['review_rows']=1
            conn.execute('UPDATE central_monthly SET summary=?',(json.dumps(s),));conn.execute('UPDATE central_batches SET review_rows=1');conn.execute("UPDATE central_product_monthly SET plan='Unclassified'")
        before=r['calculation'];result=repair(self.store,1)
        self.assertEqual(result['policy_classifications'],1)
        after=list(self.store.records('batch1'))[0]
        self.assertEqual(after['plan'],'Historic');self.assertEqual(after['calculation'],before)
        self.assertEqual(repair(self.store,1)['policy_classifications'],0)
        with self.direct.connection() as conn:self.assertEqual(conn.execute('SELECT COUNT(*) FROM central_allocation_audit').fetchone()[0],1)
    def test_existing_alias_and_own_branch_keep_bank_amounts(self):
        self.import_rows([record(franchise='MARTINS FUNERALS ERMELO')])
        with closing(sqlite3.connect(':memory:')) as stage:
            stage.execute('CREATE TABLE claims(n INTEGER PRIMARY KEY,month TEXT,franchise TEXT,product TEXT,amount TEXT,source TEXT,issues TEXT)')
            for n,name in enumerate(['Martins Funeral Franchising - Ermelo','Lavender Properties'],1):stage.execute('INSERT INTO claims VALUES (?,?,?,?,?,?,?)',(n,'2026-09',name,'Example','70',json.dumps({'Net amount paid via bank':'70'}),json.dumps([UNRESOLVED])))
            publish_claims(self.store,stage,'claims.xlsx',1,'Net amount paid via bank',['2026-09'])
        result=repair(self.store,1)
        self.assertEqual(result['claim_allocations'],2)
        values=list(self.store.claims('2026-09','2026-09'))
        self.assertEqual({r['franchise'] for r in values},{'MARTINS FUNERALS ERMELO','Lavender Properties'})
        self.assertEqual(sum(float(r['amount']) for r in values),140)
        self.assertTrue(all(not json.loads(r['issues']) for r in values))

if __name__=='__main__':unittest.main()
