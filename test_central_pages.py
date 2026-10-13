import csv,json,sqlite3,tempfile,unittest
from pathlib import Path
from contextlib import closing
from unittest.mock import patch
from flask import g
import test_dashboard_sections as baseline
from test_direct_data import record
from direct_data import stage_source,prepare_outputs
from central_claims import publish_claims
from central_ledger import publish as publish_ledger

class CentralPages(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        baseline.DashboardSectionsTest.setUpClass();cls.app=baseline.claims.app;cls.client=cls.app.test_client()
        cls.temp=tempfile.TemporaryDirectory();cls.root=Path(cls.temp.name)
        cls.store=cls.app.extensions['central_store'];cls.saved=(cls.store.direct.path,cls.store.direct.ready,cls.store.ready)
        cls.store.direct.path=cls.root/'central.sqlite';cls.store.direct.ready=False;cls.store.ready=False
        cls.queue=cls.app.extensions['central_jobs'];cls.queue_saved=(cls.queue.root,cls.queue.ready,cls.queue.kick)
        cls.queue.root=cls.root/'jobs';cls.queue.root.mkdir();cls.queue.ready=False;cls.queue.kick=lambda:None
        path=cls.root/'input.csv'
        with path.open('w',newline='',encoding='utf-8') as f:
            csv.writer(f).writerows([record(product='Gold Family Plan 31 to 55 Years R 40 000',risk='200',retail='300'),record(relation='DEP',risk='0',retail='20'),record(relation='MEMBEN',product='GROCERY DISCOUNT COUPONS FOR MARTINS CLIENTS',risk='17',retail='20',mpia='2')])
        with closing(sqlite3.connect(':memory:')) as stage:
            stage_source(path,stage);prepare_outputs(stage,cls.root/'unused',path.name,'mpia','mpia','policy',generate_files=False);cls.store.direct.save(stage,'pages',999999,'2026-09',path.name,central_only=True)
        with closing(sqlite3.connect(':memory:')) as stage:
            stage.execute('CREATE TABLE claims(n INTEGER PRIMARY KEY,month TEXT,franchise TEXT,product TEXT,amount TEXT,source TEXT,issues TEXT)')
            stage.execute('INSERT INTO claims VALUES (1,?,?,?,?,?,?)',('2026-09','MARTINS FUNERALS ERMELO','Gold Family Plan 31 to 55 Years R 40 000','100',json.dumps({'Amount net off against claim':'30','Net amount paid via bank':'70'}),'[]'))
            publish_claims(cls.store,stage,'claims.xlsx',999999,'Final Amount',['2026-09'])
        publish_ledger(cls.store,'balance',999999,'2026-09','control.xlsx',[dict(franchise='MARTINS FUNERALS ERMELO',opening='20',payment='10',system_control='217',account='TEST',paid_on='')])
    @classmethod
    def tearDownClass(cls):
        cls.store.direct.path,cls.store.direct.ready,cls.store.ready=cls.saved
        cls.queue.root,cls.queue.ready,cls.queue.kick=cls.queue_saved
        baseline.DashboardSectionsTest.tearDownClass();cls.temp.cleanup()
    def test_all_central_pages_and_separate_claim_components(self):
        args='?date_from=2026-09-01&date_to=2026-09-30'
        for section in ('risk','commissions','policies','top_franchise','claims','downloads','graphs','admin','imports'):
            response=self.client.get('/overview/'+section+args)
            self.assertEqual(response.status_code,200,(section,response.data[:150]))
        response=self.client.get('/calculations'+args+'&category=claims')
        self.assertEqual(response.status_code,200);self.assertIn(b'Bank payments 70',response.data);self.assertIn(b'Net-offs 30',response.data)
        response=self.client.get('/calculations'+args+'&category=commissions')
        self.assertIn(b'2 matching rows',response.data);self.assertIn(b'DEP',response.data)
    def test_scope_protects_details_and_catalogue(self):
        original=self.app.before_request_funcs
        self.app.before_request_funcs={None:[lambda:setattr(g,'user',dict(id=999998,role='viewer'))]}
        try:
            with patch.object(baseline.claims,'get_user_franchise_access',return_value=['OTHER FRANCHISE']):
                response=self.client.get('/calculations?date_from=2026-09-01&date_to=2026-09-30')
                self.assertEqual(response.status_code,200);self.assertNotIn(b'MARTINS FUNERALS ERMELO',response.data)
                self.assertEqual(self.client.get('/reports/central/pages/0').status_code,403)
            self.assertEqual(self.client.post('/admin/claims/confirm-zero',data={'month':'2026-08','confirm':'yes'}).status_code,403)
        finally:self.app.before_request_funcs=original
    def test_zero_confirmation_cannot_erase_paid_claims(self):
        self.assertEqual(self.client.post('/admin/claims/confirm-zero',data={'month':'2026-09','confirm':'yes'}).status_code,409)
        self.assertEqual(self.client.post('/admin/claims/confirm-zero',data={'month':'2026-08','confirm':'yes'}).status_code,302)
        self.assertIn('2026-08',self.store.verified_claim_months())
    def test_legacy_upload_paths_are_disabled(self):
        import io
        self.assertEqual(self.client.post('/direct-data',data={'direct_file':(io.BytesIO(b'test'),'test.csv')}).status_code,410)
        self.assertEqual(self.client.post('/dashboard',data={'excel_file':(io.BytesIO(b'test'),'test.xlsx')}).status_code,410)
    def test_claims_and_balances_invalidate_dependent_exports(self):
        response=self.client.get('/reports/central/pages/0?prepare=json')
        self.assertEqual(response.status_code,202)
        identifier=response.json['job']
        with self.store.direct.connection() as conn:
            args=json.loads(self.store.direct.execute(conn,'SELECT args FROM central_jobs WHERE id=:id',{'id':identifier}).fetchone()[0])
        self.assertEqual(args['ledger_version'],'balance')
        with self.app.test_request_context('/'):
            g.user={'id':999999,'role':'admin'}
            self.assertTrue(self.app.extensions['central_export_access'](args))
            publish_ledger(self.store,'balance-new',999999,'2026-09','control2.xlsx',[dict(franchise='MARTINS FUNERALS ERMELO',opening='50',payment='10',system_control='217',account='TEST',paid_on='')])
            self.assertFalse(self.app.extensions['central_export_access'](args))

if __name__=='__main__':unittest.main()
