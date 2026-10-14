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
            publish_claims(self.store,stage,'claims.xlsx',1,VERIFIED_AMOUNT_HEADER,[r[0] for r in stage.execute('SELECT DISTINCT month FROM claims')],batch_id=batch)
    def complete_year(self):
        from dateutil.relativedelta import relativedelta
        from datetime import date
        months=[(date(2025,10,1)+relativedelta(months=i)).strftime('%Y-%m') for i in range(11)]
        with closing(sqlite3.connect(':memory:')) as stage:
            stage_source(self.root/'direct.csv',stage);prepare_outputs(stage,self.root/'unused','direct.csv','mpia','mpia','policy',generate_files=False)
            for month in months:self.store.direct.save(stage,'direct-'+month,1,month,'direct.csv',central_only=True)
        with self.store.direct.connection() as conn:
            batch=conn.execute("SELECT batch_id FROM central_claim_publications WHERE month='2026-09'").fetchone()[0]
            for month in months:conn.execute('INSERT INTO central_claim_publications VALUES (?,?) ON CONFLICT(month) DO NOTHING',(month,batch))
    def test_unpaid_repudiations_are_analysed_without_becoming_paid_claims(self):
        self.publish([self.source(),self.source('C2',**{'Date Of Claim Paid':'','Net amount paid via bank':'0','Amount net off against claim':'0','Claim Status':'Repudiated','Number of days between Incident Date/Entry Date':'-5','Relationship of Life Assured/Member':'Child - Stillborn','PayableTo':'','Policyholder Gender':'Other','Incident Cause':'Unnatural'})])
        a=analyse(self.store,self.dates);self.assertEqual(a['reported_count'],2);self.assertEqual(a['paid_count'],1)
        self.assertEqual(a['bank'],150);self.assertEqual(a['netoff'],20)
        self.assertEqual({r['label']:r['count'] for r in a['tables']['status']},{'Approved':1,'Repudiated':1})
        self.assertEqual(a['tables']['duration'][0]['days_count'],1);self.assertEqual(a['tables']['duration'][0]['years'],1)
        self.assertEqual(a['tables']['payee'][0]['label'],'Franchise holder')
        self.assertIn('R40,000.00',a['tables']['cover'][0]['label'])
        self.assertEqual(a['tables']['risk'][0]['action'],'Needs complete claims and monthly premiums')
        self.assertEqual(a['tables']['risk_franchise'],[])
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
        self.publish([self.source('SEPT'),self.source('AUG',**{'Date Of Claim Paid':'2026-08-15','Product':PRODUCT.lower()})],batch='annual-detail')
        self.complete_year()
        response=client.get('/claims/analytics/risk?date_from=2026-09-01&date_to=2026-09-30&group='+PRODUCT)
        self.assertIn(b'2025-10-01 to 2026-09-30',response.data)
        self.assertIn(b'2 matching rows',response.data);self.assertIn(b'2026-08-15',response.data)
        response=client.get('/claims/analytics/branch_month?date_from=2026-09-01&date_to=2026-09-30')
        self.assertIn(b'2 matching rows',response.data);self.assertIn(b'2025-10-01 to 2026-09-30',response.data)
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
        self.publish([self.source(**{'Net amount paid via bank':str(net*12*Decimal('.70'))})]);self.complete_year();a=analyse(self.store,self.dates)
        self.assertEqual(a['tables']['risk'][0]['ratio'],70);self.assertEqual(a['tables']['risk'][0]['colour'],'red')
        self.assertEqual(classify(self.source(**{'Number of days between Incident Date/Entry Date':'bad'}))['days'],None)
    def test_sales_flags_use_matching_twelve_month_claims_and_premiums(self):
        self.publish([self.source('OLD',**{'Date Of Claim Paid':'2025-09-30'}),self.source('AUG',**{'Date Of Claim Paid':'2026-08-15','Net amount paid via bank':'200'}),self.source('SEP')]);self.complete_year()
        a=analyse(self.store,self.dates);r=a['tables']['risk'][0]
        self.assertEqual(a['paid_count'],1)
        self.assertEqual(r['count'],2);self.assertEqual(r['bank'],350)
        self.assertEqual(r['premium'],Decimal('194.82')*12)
        self.assertEqual(r['ratio'],Decimal(350)/(Decimal('194.82')*12)*100)
        self.assertEqual(a['risk_dates'],dict(date_from='2025-10-01',date_to='2026-09-30'))
        partial=analyse(self.store,dict(date_from='2026-09-20',date_to='2026-09-25'))
        self.assertEqual(partial['paid_count'],0);self.assertEqual(partial['tables']['risk'][0]['ratio'],r['ratio'])
    def test_sales_ranking_excludes_products_without_paid_claims(self):
        self.publish([self.source()])
        with self.store.direct.connection() as conn:
            row=conn.execute('SELECT * FROM central_product_monthly').fetchone()
            values=list(row);values[2]='Never claimed';conn.execute('INSERT INTO central_product_monthly VALUES (?,?,?,?,?,?)',values)
        self.assertEqual([r['label'] for r in analyse(self.store,self.dates)['tables']['risk']],[PRODUCT])
    def test_sales_combine_branches_without_changing_other_selected_cards(self):
        self.publish([self.source('ONE'),self.source('TWO')]);self.complete_year()
        with self.store.direct.connection() as conn:
            conn.execute("UPDATE central_claims SET franchise='OTHER' WHERE n=3")
            conn.execute("UPDATE central_claim_events SET franchise='OTHER' WHERE n=3")
            for row in conn.execute('SELECT * FROM central_product_monthly').fetchall():
                values=list(row);values[1]='OTHER';conn.execute('INSERT INTO central_product_monthly VALUES (?,?,?,?,?,?)',values)
        combined=analyse(self.store,self.dates,{'ERMELO'},sales_allowed=None)
        r=combined['tables']['risk'][0]
        self.assertEqual(combined['paid_count'],1);self.assertEqual(r['count'],2)
        self.assertEqual(r['franchises'],2);self.assertEqual(r['bank'],300)
        self.assertEqual(r['premium'],Decimal('194.82')*24)
        self.assertEqual(r['remaining'],r['premium']-300)
        restricted=analyse(self.store,self.dates,{'ERMELO'},sales_allowed={'ERMELO'})
        self.assertEqual(restricted['tables']['risk'][0]['count'],1)
        self.assertEqual(restricted['tables']['risk'][0]['premium'],Decimal('194.82')*12)
    def test_first_seven_in_each_sales_band_are_ranked_by_ratio(self):
        amounts=list(range(99,91,-1))+list(range(69,61,-1))+list(range(49,41,-1))
        self.publish([self.source(str(i),**{'Product':'PRODUCT '+str(i),'Net amount paid via bank':str(amount)}) for i,amount in enumerate(amounts)])
        self.complete_year()
        with self.store.direct.connection() as conn:
            row=list(conn.execute("SELECT * FROM central_product_monthly WHERE batch_id='direct'").fetchone())
            summary=json.loads(row[5]);summary['payover']='100';row[5]=json.dumps(summary)
            for i in range(len(amounts)):
                row[2]='PRODUCT '+str(i);conn.execute('INSERT INTO central_product_monthly VALUES (?,?,?,?,?,?)',row)
        a=analyse(self.store,self.dates)
        self.assertEqual([g['count'] for g in a['risk_groups']],[8,8,8])
        self.assertEqual([[r['ratio'] for r in g['rows']] for g in a['risk_groups']],[list(range(99,92,-1)),list(range(69,62,-1)),list(range(49,42,-1))])
        self.assertEqual(a['risk_groups'][0]['rows'][0]['remaining'],1)
    def test_branch_totals_use_twelve_actual_paid_months_and_selected_scope(self):
        self.publish([self.source('OLD',**{'Date Of Claim Paid':'2025-09-30'}),self.source('AUG',**{'Date Of Claim Paid':'2026-08-15','Net amount paid via bank':'200','Amount net off against claim':'40'}),self.source('SEP')])
        a=analyse(self.store,self.dates)
        self.assertEqual(a['paid_count'],1)
        self.assertEqual(a['branch_totals'],dict(claims=350,claim_net_off=60,claim_count=2))
        self.assertEqual(a['branch_summaries'][0]['claims'],350)
        self.assertEqual(sum(r['count'] for r in a['tables']['branch_month']),2)
        self.assertFalse(a['branch_complete'])
        restricted=analyse(self.store,self.dates,{'OTHER'})
        self.assertEqual(restricted['branch_totals']['claims'],0);self.assertEqual(restricted['branch_summaries'],[])
    def test_product_dates_do_not_change_annual_branch_totals(self):
        from claim_insights import sales_selection
        from werkzeug.exceptions import BadRequest
        self.publish([self.source('AUG',**{'Date Of Claim Paid':'2026-08-15','Net amount paid via bank':'200'}),self.source('SEP')]);self.complete_year()
        custom=sales_selection(dict(product_from='2026-08-01',product_to='2026-08-31'),self.dates)
        a=analyse(self.store,self.dates,sales_dates=custom)
        self.assertEqual(a['tables']['risk'][0]['bank'],200)
        self.assertEqual(a['risk_dates'],custom);self.assertEqual(a['branch_totals']['claims'],350)
        self.assertEqual(a['branch_dates'],dict(date_from='2025-10-01',date_to='2026-09-30'))
        with self.assertRaises(BadRequest):sales_selection(dict(product_from='2026-09-01',product_to='2026-08-01'),self.dates)
        partial=analyse(self.store,self.dates,sales_dates=dict(date_from='2026-08-10',date_to='2026-08-20'))
        self.assertEqual(partial['tables']['risk'][0]['bank'],200);self.assertIsNone(partial['tables']['risk'][0]['ratio'])
    def test_most_claimed_ranks_count_in_green_and_blue_instead_of_ratio(self):
        cases=[]
        for product,amounts in [('GREEN_MANY',[20,20,20]),('GREEN_HIGH',[69]),('BLUE_MANY',[10,10,10,10]),('BLUE_HIGH',[49]),('RED',[99])]:
            for i,amount in enumerate(amounts):cases.append(self.source(product+str(i),**{'Product':product,'Net amount paid via bank':str(amount)}))
        self.publish(cases);self.complete_year()
        with self.store.direct.connection() as conn:
            row=list(conn.execute("SELECT * FROM central_product_monthly WHERE batch_id='direct'").fetchone())
            summary=json.loads(row[5]);summary['payover']='100';row[5]=json.dumps(summary)
            for product in ['GREEN_MANY','GREEN_HIGH','BLUE_MANY','BLUE_HIGH','RED']:
                row[2]=product;conn.execute('INSERT INTO central_product_monthly VALUES (?,?,?,?,?,?)',row)
        a=analyse(self.store,self.dates)
        self.assertEqual([g['rows'][0]['label'] for g in a['most_groups']],['GREEN_MANY','BLUE_MANY'])
        self.assertEqual([g['rows'][0]['count'] for g in a['most_groups']],[3,4])
        self.assertEqual(a['most_groups'][0]['rows'][0]['ratio'],60)
    def test_branch_dates_and_rankings_are_independent_of_products(self):
        from claim_insights import branch_selection,rank_branches
        from werkzeug.exceptions import BadRequest
        self.publish([self.source('AUG',**{'Date Of Claim Paid':'2026-08-15','Net amount paid via bank':'200'}),self.source('SEP')]);self.complete_year()
        custom=branch_selection(dict(branch_from='2026-08-01',branch_to='2026-08-31',branch_scenario='count'),self.dates)
        a=analyse(self.store,self.dates,branch_dates=custom)
        self.assertEqual(a['branch_totals']['claims'],200);self.assertEqual(a['branch_totals']['claim_count'],1)
        self.assertEqual(a['tables']['risk'][0]['bank'],350)
        rows=[dict(franchise='A',count=3,claims=100),dict(franchise='B',count=1,claims=500)]
        self.assertEqual(rank_branches(rows,'count')[0]['franchise'],'A');self.assertEqual(rank_branches(rows,'amount')[0]['franchise'],'B')
        with self.assertRaises(BadRequest):branch_selection(dict(branch_from='2026-09-30',branch_to='2026-09-01'),self.dates)
    def test_franchise_products_group_paid_claims_and_keep_scope_and_dates(self):
        self.publish([self.source('SEP'),self.source('AUG',**{'Date Of Claim Paid':'2026-08-15','Product':PRODUCT.lower()}),self.source('OTHER',**{'Product':'Other Product','Net amount paid via bank':'200'}),self.source('OLD',**{'Date Of Claim Paid':'2025-08-15'})])
        app=Flask(__name__);app.secret_key='test';app.template_folder=str(Path(__file__).parent/'templates')
        app.add_url_rule('/overview/<section>','overview_page',lambda section:'overview')
        app.before_request(lambda:setattr(g,'user',{'id':1,'role':'admin'}))
        register(app,self.store,lambda:{'ERMELO'});client=app.test_client()
        query=dict(franchise='ERMELO',date_from='2026-09-01',date_to='2026-09-30',branch_from='2026-08-01',branch_to='2026-09-30')
        response=client.get('/claims/branch-products',query_string=query)
        self.assertEqual(response.status_code,200);self.assertIn(b'3 matching rows',response.data)
        self.assertIn(b'>2</td><td>R300.00</td>',response.data)
        self.assertLess(response.data.index(PRODUCT.encode()),response.data.index(b'>Other Product</a>'))
        query['product']=PRODUCT;response=client.get('/claims/branch-products',query_string=query)
        self.assertIn(b'2 matching rows',response.data)
        query['franchise']='OTHER';self.assertEqual(client.get('/claims/branch-products',query_string=query).status_code,403)

if __name__=='__main__':unittest.main()
