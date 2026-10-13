import json,sqlite3,tempfile,csv,unittest
from pathlib import Path
from contextlib import closing
from decimal import Decimal
from openpyxl import Workbook
from test_direct_data import record
from direct_data import stage_source,prepare_outputs
from direct_analytics import DirectStore
from central_store import CentralStore
from central_claims import read_transactions,publish_claims,VERIFIED_AMOUNT_HEADER
from claim_insights import analyse,classify,SPECS,register
from flask import Flask,g

PRODUCT='Gold Family Plan 31 to 55 Years R 40 000'
class ClaimInsightsTest(unittest.TestCase):
    def setUp(self):
        self.temp=tempfile.TemporaryDirectory();self.root=Path(self.temp.name)
        self.store=CentralStore(DirectStore(self.root/'store.sqlite'));self.store.direct.central=self.store
        source=self.root/'direct.csv'
        with source.open('w',newline='',encoding='utf-8') as f:csv.writer(f).writerow(record(franchise='ERMELO',product=PRODUCT,cover='40000',risk='200',retail='300'))
        with closing(sqlite3.connect(':memory:')) as stage:
            stage_source(source,stage);prepare_outputs(stage,self.root/'unused',source.name,'mpia','mpia','policy',generate_files=False)
            self.store.direct.save(stage,'direct',1,'2026-09',source.name,central_only=True)
        self.dates=dict(date_from='2026-09-01',date_to='2026-09-30')
    def tearDown(self):self.temp.cleanup()
    def source(self,claim='C1',**extra):
        return {'Claim Number':claim,'Administrator':'Direct','Participating Group':'ERMELO','Product':PRODUCT,
          'Claim Reported Date':'2026-09-01','Date Of Claim Paid':'2026-09-10','Net amount paid via bank':'150',
          'Amount net off against claim':'20','Final Amount':'170','Relationship of Life Assured/Member':'Main Life',
          'Policyholder Gender':'Female','Incident Cause':'Natural','Claim Status':'Claim - Approved by Approver and paid',
          'Age Category':'66 - 75 Years','PayableTo':'Entity','Number of days between Incident Date/Entry Date':'365.25',**extra}
    def publish(self,rows,batch='claims'):
        with closing(sqlite3.connect(':memory:')) as stage:
            stage.execute('CREATE TABLE claims(n INTEGER PRIMARY KEY,month TEXT,franchise TEXT,product TEXT,amount TEXT,source TEXT,issues TEXT)')
            stage.execute('CREATE TABLE claim_events(n INTEGER PRIMARY KEY,source TEXT)')
            for n,s in enumerate(rows,2):
                raw=json.dumps(s);stage.execute('INSERT INTO claim_events VALUES (?,?)',(n,raw))
                if s.get('Date Of Claim Paid') and (Decimal(s['Net amount paid via bank']) or Decimal(s['Amount net off against claim'])):
                    stage.execute('INSERT INTO claims VALUES (?,?,?,?,?,?,?)',(n,s['Date Of Claim Paid'][:7],'ERMELO',s['Product'],s['Net amount paid via bank'],raw,'[]'))
            publish_claims(self.store,stage,'claims.xlsx',1,VERIFIED_AMOUNT_HEADER,['2026-09'],batch_id=batch)
    def test_unpaid_repudiations_are_analysed_without_becoming_paid_claims(self):
        self.publish([self.source(),self.source('C2',**{'Date Of Claim Paid':'','Net amount paid via bank':'0','Amount net off against claim':'0','Claim Status':'Repudiated','Number of days between Incident Date/Entry Date':'-5','Relationship of Life Assured/Member':'Child - Stillborn','PayableTo':'','Policyholder Gender':'Other','Incident Cause':'Unnatural'})])
        a=analyse(self.store,self.dates);self.assertEqual(a['reported_count'],2);self.assertEqual(a['paid_count'],1)
        self.assertEqual(a['bank'],150);self.assertEqual(a['netoff'],20)
        self.assertEqual({r['label']:r['count'] for r in a['tables']['status']},{'Approved':1,'Repudiated':1})
        self.assertEqual(a['tables']['duration'][0]['days_count'],1);self.assertEqual(a['tables']['duration'][0]['years'],1)
        self.assertEqual(a['tables']['payee'][0]['label'],'Franchise holder')
        self.assertIn('R40,000.00',a['tables']['cover'][0]['label'])
        self.assertEqual(a['tables']['risk'][0]['action'],'Must not be sold')
        self.assertEqual(a['tables']['risk_franchise'][0]['count'],1)
    def test_exact_dates_and_partial_month_do_not_use_full_month_premium_as_a_rate(self):
        self.publish([self.source()]);a=analyse(self.store,dict(date_from='2026-09-11',date_to='2026-09-30'))
        self.assertEqual(a['paid_count'],0);self.assertEqual(a['reported_count'],0)
        self.assertIsNone(a['tables']['risk'][0]['ratio'])
        self.assertIsNone(next(c['value'] for c in a['cards'] if c['key']=='risk'))
    def test_newest_source_snapshot_replaces_status_without_double_counting(self):
        self.publish([self.source()],'first');self.publish([self.source(**{'Net amount paid via bank':'100'})],'second')
        a=analyse(self.store,self.dates);self.assertEqual(a['reported_count'],1);self.assertEqual(a['bank'],100)
        with self.store.direct.connection() as c:self.assertEqual(c.execute('SELECT COUNT(*) FROM central_claim_events').fetchone()[0],2)
    def test_netoff_only_rows_do_not_increase_paid_count(self):
        self.publish([self.source(),self.source('C2',**{'Net amount paid via bank':'0','Amount net off against claim':'30','PayableTo':'Individual'})])
        a=analyse(self.store,self.dates);self.assertEqual(a['paid_count'],1);self.assertEqual(a['netoff_count'],2);self.assertEqual(a['netoff'],50)
    def test_scope_applies_to_every_cohort_and_ratio(self):
        self.publish([self.source()]);a=analyse(self.store,self.dates,{'OTHER'})
        self.assertEqual(a['reported_count'],0);self.assertEqual(a['paid_count'],0);self.assertEqual(a['tables']['risk'],[])
    def test_all_detail_pages_and_recipient_restrictions(self):
        self.publish([self.source()]);app=Flask(__name__);app.secret_key='test';app.template_folder=str(Path(__file__).parent/'templates')
        app.add_url_rule('/overview/<section>','overview_page',lambda section:'overview')
        @app.before_request
        def user():g.user={'id':1,'role':'admin'}
        register(app,self.store,lambda:None);client=app.test_client()
        for key in SPECS:
            response=client.get('/claims/analytics/'+key+'?date_from=2026-09-01&date_to=2026-09-30')
            self.assertEqual(response.status_code,200,(key,response.data[:200]))
        app.before_request_funcs={None:[lambda:setattr(g,'user',{'id':2,'role':'voucher_supplier'})]}
        self.assertEqual(client.get('/claims/analytics/status').status_code,403)
    def test_reader_captures_all_status_rows_but_preserves_bank_financial_reader(self):
        path=self.root/'source.xlsx';book=Workbook();sheet=book.active;one=self.source();two=self.source('C2',**{'Date Of Claim Paid':'','Net amount paid via bank':'0','Amount net off against claim':'0','Claim Status':'Repudiated'})
        sheet.append(list(one));sheet.append(list(one.values()));sheet.append(list(two.values()));book.save(path)
        events=[];paid=list(read_transactions(path,VERIFIED_AMOUNT_HEADER,event_sink=lambda n,s:events.append(s)))
        self.assertEqual(len(events),2);self.assertEqual(len(paid),1)
    def test_archived_source_backfill_never_replaces_published_bank_amounts(self):
        from flask import Flask
        from central_jobs import JobQueue
        from import_audit import ImportAudit
        from unittest.mock import patch
        from io import BytesIO
        from claim_insights import backfill
        path=self.root/'archived.xlsx';book=Workbook();sheet=book.active;one=self.source();two=self.source('C2',**{'Date Of Claim Paid':'','Net amount paid via bank':'0','Amount net off against claim':'0','Claim Status':'Repudiated'})
        sheet.append(list(one));sheet.append(list(one.values()));sheet.append(list(two.values()));book.save(path)
        app=Flask(__name__);queue=JobQueue(app,self.store,self.root,ImportAudit(self.root/'audit.sqlite'))
        class Upload:stream=BytesIO(path.read_bytes())
        with patch.object(queue,'kick'):identifier=queue.create(1,'claims',dict(filename='archived.xlsx',amount_header=VERIFIED_AMOUNT_HEADER,sheet='',replace_months=['2026-09']),Upload())
        self.publish([one],identifier)
        with self.store.direct.connection() as c:
            c.execute('DELETE FROM central_claim_events');c.execute('DELETE FROM central_claim_event_versions')
        before=list(self.store.claims())
        result=backfill(self.store,queue,identifier)
        self.assertEqual(result['source']['rows'],2);self.assertFalse(result['financial_data_changed'])
        self.assertEqual(list(self.store.claims()),before)
        self.assertEqual(analyse(self.store,self.dates)['reported_count'],2)
    def test_70_percent_threshold_is_inclusive_and_zero_premium_unavailable(self):
        net=Decimal(self.store.summaries()[0]['summary']['payover'])
        self.publish([self.source(**{'Net amount paid via bank':str(net*Decimal('.70'))})]);a=analyse(self.store,self.dates)
        self.assertEqual(a['tables']['risk'][0]['ratio'],70);self.assertEqual(a['tables']['risk'][0]['colour'],'red')
        self.assertEqual(classify(self.source(**{'Number of days between Incident Date/Entry Date':'bad'}))['days'],None)

if __name__=='__main__':unittest.main()
