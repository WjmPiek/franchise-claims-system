import csv,json,sqlite3,tempfile,unittest
from decimal import Decimal
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
            stage.execute('INSERT INTO claims VALUES (1,?,?,?,?,?,?)',('2026-09','MARTINS FUNERALS ERMELO','Gold Family Plan 31 to 55 Years R 40 000','100',json.dumps({'Amount net off against claim':'30','Net amount paid via bank':'70','Date Of Claim Paid':'2026-09-10','Product':'Gold Family Plan 31 to 55 Years R 40 000'}),'[]'))
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
        response=self.client.get('/calculations'+args+'&category=claims&view=records')
        self.assertEqual(response.status_code,200);self.assertIn(b'Bank payments 70',response.data);self.assertIn(b'Net-offs 30',response.data)
        response=self.client.get('/calculations'+args+'&category=commissions&view=records')
        self.assertIn(b'2 matching rows',response.data);self.assertIn(b'DEP',response.data)
    def test_calculation_cards_default_to_combined_franchises_for_every_category(self):
        from calculation_groups import policy_groups,money_expression
        dates='?date_from=2026-09-01&date_to=2026-09-30'
        for category in ('source','review','insurers','aul','plan2','historic','unclassified','commissions','active','faithnet','fees','vouchers','claims'):
            page=self.client.get('/calculations'+dates+'&category='+category)
            self.assertEqual(page.status_code,200,category)
            self.assertIn(b'One combined row per franchise',page.data,category)
            self.assertNotIn(b'Each row is a separate source record',page.data)
        page=self.client.get('/calculations'+dates+'&category=insurers')
        self.assertIn(b'2 matching source rows combined into 1 franchise rows',page.data)
        self.assertIn(b'R320.00',page.data)
        page=self.client.get('/calculations'+dates+'&category=vouchers')
        self.assertIn(b'R13.00',page.data);self.assertIn(b'R4.00',page.data)
        page=self.client.get('/calculations'+dates+'&category=claims')
        self.assertIn(b'R70.00',page.data);self.assertIn(b'R30.00',page.data)
        page=self.client.get('/calculations'+dates+'&q=missing-product')
        self.assertIn(b'0 matching source rows combined into 0 franchise rows',page.data)
        self.assertEqual(self.client.get('/calculations'+dates+'&view=bad').status_code,400)
        expression=money_expression(True,'paid_commissions')
        self.assertIn('substring(r.record_json',expression);self.assertNotIn('::json',expression)
        base=" FROM central_records r JOIN central_publications p ON p.batch_id=r.batch_id WHERE p.month='2026-09'"
        rows=policy_groups(self.store,base,{})
        self.assertEqual(len(rows),1);self.assertEqual(str(rows[0]['retail']),'340')
        from calculation_groups import summary_groups,FIELDS
        for category in ('source','insurers','commissions'):
            fast=summary_groups(self.store,'2026-09','2026-09',None,category)
            selected_base=base+(" AND r.destination<>'Vouchers'" if category!='source' else '')
            exact=policy_groups(self.store,selected_base,{})
            for field in FIELDS:self.assertEqual(Decimal(str(fast[0][field])),Decimal(str(exact[0][field])),(category,field))
        old=self.client.get('/franchise/MARTINS%20FUNERALS%20ERMELO'+dates)
        self.assertEqual(old.status_code,302);self.assertIn('/franchise-summary?',old.location)

    def test_top_product_links_use_typed_product_filter_and_keep_it_for_source_view(self):
        product='Gold Family Plan 31 to 55 Years R 40 000'
        params=dict(date_from='2026-09-01',date_to='2026-09-30',category='active',product=product)
        page=self.client.get('/calculations',query_string=params)
        self.assertEqual(page.status_code,200);self.assertIn(b'1 matching source rows combined into 1 franchise rows',page.data)
        self.assertIn(b'R300.00',page.data);self.assertIn(b'name="product"',page.data)
        page=self.client.get('/calculations',query_string=dict(params,view='records'))
        self.assertEqual(page.status_code,200);self.assertIn(b'1 matching rows',page.data)
        rankings=self.client.get('/overview/top_franchise?date_from=2026-09-01&date_to=2026-09-30')
        self.assertIn(b'product=Gold+Family',rankings.data)
        self.assertNotIn(b'q=Gold+Family',rankings.data)

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
    def test_graphs_default_to_latest_twelve_months_without_heavy_page_work(self):
        with self.client.session_transaction() as session:session['period_dates']={'date_from':'2026-09-01','date_to':'2026-09-30'};session.pop('graph_periods',None)
        with patch.object(baseline.claims,'build_executive_dashboard_context',side_effect=AssertionError('Unrelated dashboard work')),patch('central_identity.member_snapshot',side_effect=AssertionError('Member index scan')),patch('central_identity.products_by_retail',side_effect=AssertionError('Product scan')):
            page=self.client.get('/overview/graphs?date_from=2026-09-01&date_to=2026-09-30')
        self.assertEqual(page.status_code,200,page.data[:200]);self.assertEqual(page.data.count(b'data-graph-card='),17)
        self.assertIn(b'value="2025-10-01"',page.data);self.assertIn(b'value="2026-09-30"',page.data)
        with self.client.session_transaction() as session:self.assertEqual(session['period_dates']['date_from'],'2026-09-01')
    def test_update_one_graph_preserves_other_periods_and_skips_claim_rows(self):
        with self.client.session_transaction() as session:session.pop('graph_periods',None)
        with patch.object(self.store,'claims',side_effect=AssertionError('Risk graph must not load claims')):
            page=self.client.get('/graphs/original_risk?date_from=2026-01-01&date_to=2026-01-31')
        self.assertEqual(page.status_code,200);self.assertEqual(page.data.count(b'data-graph-card='),1);self.assertIn(b'2026-01-01',page.data)
        with self.client.session_transaction() as session:self.assertEqual(set(session['graph_periods']),{'original_risk'})
        page=self.client.get('/overview/graphs');self.assertIn(b'value="2025-10-01"',page.data);self.assertIn(b'value="2026-01-01"',page.data)
        self.assertEqual(self.client.get('/graphs/original_risk?date_from=2026-10-01&date_to=2026-01-01').status_code,400)
        self.assertEqual(self.client.get('/graphs/not-a-graph?date_from=2026-01-01&date_to=2026-01-31').status_code,404)
    def test_non_policy_tabs_do_not_load_notices_or_relations(self):
        with patch.object(baseline.claims,'get_policy_relation_counts',side_effect=AssertionError('Relation scan')),patch.object(baseline.claims,'get_policy_age_notifications',side_effect=AssertionError('Age scan')),patch.object(baseline.claims,'get_policy_age_notification_age_bands',side_effect=AssertionError('Age bands scan')):
            for section in ('risk','commissions','claims','top_franchise','imports','downloads','admin'):
                page=self.client.get('/overview/'+section+'?date_from=2026-09-01&date_to=2026-09-30');self.assertEqual(page.status_code,200,section)
        with patch('central_routes.catalog',side_effect=AssertionError('Unrelated report catalogue')):
            for section in ('risk','commissions','claims','top_franchise','imports','admin'):
                self.assertEqual(self.client.get('/overview/'+section).status_code,200,section)
    def test_published_sidebar_pages_skip_legacy_reanalysis(self):
        with patch.object(baseline.claims,'build_executive_dashboard_context',side_effect=AssertionError('Legacy analysis rebuilt')),patch.object(baseline.claims,'apply_franchise_config',side_effect=AssertionError('Legacy portfolio rebuilt')):
            for section in ('risk','commissions','policies','top_franchise','claims','imports','admin','downloads'):
                self.assertEqual(self.client.get('/overview/'+section+'?date_from=2026-09-01&date_to=2026-09-30').status_code,200,section)
    def test_claims_single_product_block_and_annual_branch_detail_links(self):
        page=self.client.get('/overview/claims?date_from=2026-09-01&date_to=2026-09-30')
        self.assertEqual(page.status_code,200)
        self.assertEqual(page.data.count(b'<th>Product</th>'),1)
        self.assertNotIn(b'<h3><span class="claim-flag',page.data)
        self.assertIn(b'The first seven products in each category',page.data)
        self.assertIn(b'Period: 2025-10-01 to 2026-09-30',page.data)
        self.assertIn(b'date_from=2025-10-01&amp;date_to=2026-09-30',page.data)
        page=self.client.get('/overview/claims?date_from=2026-09-01&date_to=2026-09-30&product_from=2026-09-01&product_to=2026-09-30&product_scenario=count')
        self.assertEqual(page.status_code,200);self.assertIn(b'Most Paid Claims (Green &amp; Blue)',page.data)
        self.assertIn(b'name="product_from" value="2026-09-01"',page.data)
        self.assertIn(b'Period: 2025-10-01 to 2026-09-30',page.data)
        self.assertEqual(self.client.get('/overview/claims?product_from=2026-09-30&product_to=2026-09-01').status_code,400)
    def test_table_drag_assets_load_once_on_standalone_and_dashboard_pages(self):
        for path in ('/overview/claims','/overview/imports','/calculations?date_from=2026-09-01&date_to=2026-09-30'):
            response=self.client.get(path)
            self.assertEqual(response.status_code,200,path)
            self.assertEqual(response.data.count(b'table-pan.js'),1,path)
    def test_combined_franchise_summary_reconciles_products_and_keeps_one_total_row(self):
        from franchise_summary import grouped_products
        dates=dict(date_from='2026-09-01',date_to='2026-09-30')
        with patch.object(self.store,'records',side_effect=AssertionError('Grouped summary must not scan source policies')):
            rows=grouped_products(self.store,'MARTINS FUNERALS ERMELO',dates)
            page=self.client.get('/franchise-summary',query_string=dict(franchise='MARTINS FUNERALS ERMELO',**dates))
        self.assertEqual(page.status_code,200)
        self.assertEqual(sum(r['rows'] for r in rows),3)
        self.assertEqual(sum(r['retail'] for r in rows),320)
        self.assertEqual(sum(r['bank'] for r in rows),70)
        self.assertEqual(sum(r['netoff'] for r in rows),30)
        self.assertEqual(len(rows),3)
        self.assertIn(b'One combined row for this franchise',page.data)
        self.assertIn(b'All policy source records',page.data)
        self.assertEqual(page.data.count(b'table-pan.js'),1)
    def test_franchise_links_and_balanced_top_cards(self):
        for section in ('risk','commissions','policies','top_franchise'):
            page=self.client.get('/overview/'+section+'?date_from=2026-09-01&date_to=2026-09-30')
            self.assertEqual(page.status_code,200)
            self.assertIn(b'/franchise-summary?',page.data,section)
        page=self.client.get('/overview/top_franchise?date_from=2026-09-01&date_to=2026-09-30')
        self.assertIn(b'class="franchise-ranking-grid"',page.data)
        original=self.app.before_request_funcs
        self.app.before_request_funcs={None:[lambda:setattr(g,'user',dict(id=999998,role='viewer'))]}
        try:
            with patch.object(baseline.claims,'get_user_franchise_access',return_value=['OTHER FRANCHISE']):
                self.assertEqual(self.client.get('/franchise-summary?franchise=MARTINS+FUNERALS+ERMELO').status_code,403)
        finally:self.app.before_request_funcs=original
    def test_financial_cache_copies_and_rechecks_publication_versions(self):
        from central_routes import financial_result
        from published_cache import FINANCIAL
        FINANCIAL.clear()
        def read():
            with self.app.test_request_context('/overview/risk'):
                g.user={'id':999999,'role':'admin'}
                return financial_result(self.app,baseline.claims.LAST_RESULT)
        with patch.object(self.store,'frame',wraps=self.store.frame) as frame:
            first=read();original=first['monthly'].iloc[0]['Retail Premium']
            first['monthly'].iloc[0,first['monthly'].columns.get_loc('Retail Premium')]=-1
            self.assertEqual(read()['monthly'].iloc[0]['Retail Premium'],original)
            self.assertEqual(frame.call_count,1)
            with self.store.direct.connection() as conn:
                before=conn.execute("SELECT published_at FROM central_publications WHERE month='2026-09'").fetchone()[0]
                conn.execute("UPDATE central_publications SET published_at=? WHERE month='2026-09'",(before+'cache-test',))
            try:read();self.assertEqual(frame.call_count,2)
            finally:
                with self.store.direct.connection() as conn:conn.execute("UPDATE central_publications SET published_at=? WHERE month='2026-09'",(before,))
                FINANCIAL.clear()
    def test_claim_profile_updates_read_claims_only_and_keep_independent_dates(self):
        with patch.object(self.store,'records',side_effect=AssertionError('Profile chart must not scan policies')):
            for key in ('claim_relationship_products','claim_product_age','claim_gender','claim_cause'):
                response=self.client.get('/graphs/'+key+'?date_from=2026-08-01&date_to=2026-09-30')
                self.assertEqual(response.status_code,200,key);self.assertEqual(response.data.count(b'data-graph-card='),1)
                self.assertIn(key.encode(),response.data);self.assertIn(b'Claim Reported Date',response.data)
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
        self.assertIsNotNone(args['claims_version'])
        with self.app.test_request_context('/'):
            g.user={'id':999999,'role':'admin'}
            self.assertTrue(self.app.extensions['central_export_access'](args))
            publish_ledger(self.store,'balance-new',999999,'2026-09','control2.xlsx',[dict(franchise='MARTINS FUNERALS ERMELO',opening='50',payment='10',system_control='217',account='TEST',paid_on='')])
            self.assertFalse(self.app.extensions['central_export_access'](args))

if __name__=='__main__':unittest.main()
