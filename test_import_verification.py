import csv,io,json,sqlite3,unittest
from contextlib import closing
from flask import Flask,g
from werkzeug.datastructures import FileStorage
from openpyxl import Workbook
from central_jobs import JobQueue
from central_claims import ANNUAL_HEADERS,read_transactions,publish_claims
from import_audit import ImportAudit
from import_verification import policy_check,claims_check,verify_targets,register,inventory
from direct_data import stage_source,prepare_outputs
from test_direct_data import record
import test_central_compliance as base

class ImportVerificationTest(unittest.TestCase):
    def setUp(self):
        base.CentralComplianceTest.setUp(self)
        self.app=Flask(__name__);self.app.secret_key='test'
        self.queue=JobQueue(self.app,self.store,self.root,ImportAudit(self.root/'audit.sqlite'));self.queue.kick=lambda:None
    tearDown=base.CentralComplianceTest.tearDown
    def policy(self):
        stream=io.StringIO();csv.writer(stream).writerows([record(risk='123.456',retail='200.005'),record(relation='DEP',risk='0',retail='0')])
        batch=self.queue.create(1,'direct',{'filename':'PolicyData_DTSTAMP_RECEIPTSFOR_20260901_to_20260930.csv'},FileStorage(stream=io.BytesIO(stream.getvalue().encode()),filename='source.csv'))
        with closing(sqlite3.connect(':memory:')) as stage:
            source=self.root/'input.csv';source.write_text(stream.getvalue(),encoding='utf-8');stage_source(source,stage)
            prepare_outputs(stage,self.root/'unused',source.name,'mpia','mpia','policy',generate_files=False)
            self.direct.save(stage,batch,1,'2026-09',source.name,central_only=True)
        return batch
    def test_original_file_records_and_fractional_amounts_match(self):
        batch=self.policy();result=policy_check(self.store,self.queue,batch)
        self.assertTrue(result['matched']);self.assertEqual(result['controls'][1]['source'],'123.456')
        self.assertEqual(result['controls'][2]['source'],'200.005')
    def test_changed_saved_amount_is_detected_without_republishing(self):
        batch=self.policy();before=self.store.publications()
        with self.direct.connection() as conn:
            r=json.loads(conn.execute('SELECT record_json FROM central_records WHERE batch_id=? AND n=1',(batch,)).fetchone()[0]);r['retail']='999'
            conn.execute('UPDATE central_records SET record_json=? WHERE batch_id=? AND n=1',(json.dumps(r),batch))
        self.assertFalse(policy_check(self.store,self.queue,batch)['matched'])
        self.assertEqual(self.store.publications(),before)
    def test_unavailable_source_does_not_stop_next_check(self):
        batch=self.policy();before=self.store.publications()
        result=verify_targets(self.store,self.queue,[dict(kind='policy',month='2026-08',batch='unavailable'),dict(kind='policy',month='2026-09',batch=batch)],'check')
        self.assertEqual((result['checked'],result['matched'],result['needs_review']),(2,1,1))
        self.assertEqual(self.store.publications(),before)
    def test_claims_check_preserves_unpaid_cases_and_paid_months(self):
        book=Workbook();s=book.active;s.append(ANNUAL_HEADERS)
        for date,bank,netoff,status in [('2026-08-15',70,30,'Approved'),('2026-09-15',80,20,'Approved'),('',0,350,'Repudiated')]:
            values={'Claim Number':str(s.max_row),'Participating Group':'ERMELO','Claim Reported Date':'2026-08-01','Date Of Claim Paid':date,'Net amount paid via bank':bank,'Amount net off against claim':netoff,'Final Amount':100,'Claim Status':status,'Product':'Example','Age Category':'40-49','PayableTo':'Individual'}
            s.append([values.get(h,'') for h in ANNUAL_HEADERS])
        buf=io.BytesIO();book.save(buf);payload=buf.getvalue()
        batch=self.queue.create(1,'claims',{'filename':'Martins Funeral All Claims.xlsx'},FileStorage(stream=io.BytesIO(payload),filename='claims.xlsx'))
        path=self.root/'claims.xlsx';path.write_bytes(payload)
        with closing(sqlite3.connect(':memory:')) as stage:
            stage.execute('CREATE TABLE claims(n INTEGER PRIMARY KEY,month TEXT,franchise TEXT,product TEXT,amount TEXT,source TEXT,issues TEXT)');stage.execute('CREATE TABLE claim_events(n INTEGER PRIMARY KEY,source TEXT)')
            for r in read_transactions(path,'Net amount paid via bank',event_sink=lambda n,source:stage.execute('INSERT INTO claim_events VALUES (?,?)',(n,json.dumps(source,default=str)))):
                stage.execute('INSERT INTO claims VALUES (?,?,?,?,?,?,?)',(r['n'],r['month'],r['franchise'],r['product'],r['amount'],json.dumps(r['source'],default=str),'[]'))
            publish_claims(self.store,stage,path.name,1,'Net amount paid via bank',['2026-08','2026-09'],batch_id=batch)
        result=claims_check(self.store,self.queue,batch)
        self.assertTrue(result['matched']);self.assertEqual(result['controls'][0]['source'],'3');self.assertEqual(result['review'],1)
    def test_verification_route_is_admin_only(self):
        self.app.add_url_rule('/overview/<section>','overview_page',lambda section:'Overview')
        self.app.add_url_rule('/calculations','central_calculation_detail',lambda:'Records')
        self.app.add_url_rule('/imports/log','import_log',lambda:'Log')
        self.app.jinja_env.filters['money']=str
        register(self.app,self.store,self.queue)
        self.role='viewer'
        @self.app.before_request
        def user():g.user={'id':1,'role':self.role}
        client=self.app.test_client()
        self.assertEqual(client.get('/imports/verification').status_code,403)
        self.role='admin';batch=self.policy()
        response=client.get('/imports/verification')
        self.assertEqual(response.status_code,200);self.assertIn(b'Not checked',response.data)
        response=client.post('/imports/verification/run',data={'date_from':'2026-09-01','date_to':'2026-09-30'})
        self.assertEqual(response.status_code,200)
        self.queue.run()
        self.assertTrue(self.queue.status(response.json['job'],1)['result']['matched'])
        response=client.get('/imports/verification')
        self.assertIn('Verified — totals match'.encode(),response.data)

if __name__=='__main__':unittest.main()
