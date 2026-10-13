"""Explicit integration check; every table is TEMP and shadows production tables."""
import os,csv,json,sqlite3,tempfile
from pathlib import Path
from contextlib import closing
from decimal import Decimal
from sqlalchemy import create_engine
from sqlalchemy.pool import StaticPool
from direct_analytics import DirectStore
from direct_data import stage_source,prepare_outputs
from central_store import CentralStore
from central_claims import publish_claims
from central_identity import member_snapshot
from central_reports import catalog_all as catalog,generate
from central_ledger import publish as publish_ledger
from test_direct_data import record

class TemporaryStore(DirectStore):
    def execute(self,conn,sql,params=None):
        if sql.startswith('CREATE TABLE IF NOT EXISTS '):sql=sql.replace('CREATE TABLE IF NOT EXISTS ','CREATE TEMP TABLE IF NOT EXISTS ',1)
        return super().execute(conn,sql,params)

def run():
    url=os.environ.get('DATABASE_URL','')
    if not url:print('Skipped: DATABASE_URL not configured.');return
    if url.startswith('postgres://'):url='postgresql://'+url[len('postgres://'):]
    engine=create_engine(url,poolclass=StaticPool)
    try:
        with tempfile.TemporaryDirectory() as directory:
            root=Path(directory);direct=TemporaryStore(root/'unused',lambda:engine);store=CentralStore(direct);direct.central=store
            with direct.connection() as conn:direct.execute(conn,'CREATE TABLE IF NOT EXISTS central_jobs (id TEXT PRIMARY KEY,created_at TEXT NOT NULL)')
            path=root/'source.csv'
            with path.open('w',newline='',encoding='utf-8') as f:csv.writer(f).writerows([record(product='Gold Family Plan 31 to 55 Years R 40 000',risk='200',retail='300')]*2)
            for identifier in ('original','replacement'):
                with closing(sqlite3.connect(':memory:')) as stage:
                    stage_source(path,stage);prepare_outputs(stage,root/'unused-export',path.name,'mpia','mpia','policy',generate_files=False)
                    direct.save(stage,identifier,900000001,'2026-09',path.name,central_only=True)
            assert store.publications()[0]['id']=='replacement'
            assert member_snapshot(store,'2026-09','2026-09')['count']==1
            assert Decimal(store.summaries()[0]['summary']['payover'])==Decimal('389.64')
            with direct.connection() as conn:
                assert direct.execute(conn,"SELECT COUNT(*) FROM central_records WHERE batch_id='original'").fetchone()[0]==2
                assert direct.execute(conn,"SELECT COUNT(*) FROM direct_data_rows").fetchone()[0]==0
            with closing(sqlite3.connect(':memory:')) as stage:
                stage.execute('CREATE TABLE claims(n INTEGER PRIMARY KEY,month TEXT,franchise TEXT,product TEXT,amount TEXT,source TEXT,issues TEXT)')
                source={'Claim Number':'TEMP-C1','Participating Group':'MARTINS FUNERALS ERMELO','Product':'Gold Family Plan 31 to 55 Years R 40 000','Claim Reported Date':'2026-09-01','Claim Status':'Approved','Amount net off against claim':'30','Net amount paid via bank':'70','Final Amount':'100','Date Of Claim Paid':'2026-09-10','Number of days between Incident Date/Entry Date':'365.25'}
                stage.execute('INSERT INTO claims VALUES (1,?,?,?,?,?,?)',('2026-09','MARTINS FUNERALS ERMELO','Gold Family Plan 31 to 55 Years R 40 000','70',json.dumps(source),'[]'))
                stage.execute('CREATE TABLE claim_events(n INTEGER PRIMARY KEY,source TEXT)')
                stage.execute('INSERT INTO claim_events VALUES (?,?)',(1,json.dumps(source)))
                stage.execute('INSERT INTO claim_events VALUES (?,?)',(2,json.dumps({**source,'Claim Number':'TEMP-C2','Claim Status':'Repudiated','Date Of Claim Paid':'','Net amount paid via bank':'0','Amount net off against claim':'0'})))
                publish_claims(store,stage,'claims.xlsx',900000001,'Net amount paid via bank',['2026-09'])
            from claim_insights import analyse
            insights=analyse(store,dict(date_from='2026-09-01',date_to='2026-09-30'))
            assert insights['reported_count']==2 and insights['paid_count']==1
            assert {r['label']:r['count'] for r in insights['tables']['status']}=={'Approved':1,'Repudiated':1}
            claim=next(store.claims());assert claim['amount']=='70' and claim['net_off']=='30' and claim['bank_paid']=='70'
            publish_ledger(store,'ledger',900000001,'2026-09','control.xlsx',[dict(franchise='MARTINS FUNERALS ERMELO',opening='20',payment='10',system_control='400',account='TEST',paid_on='2026-09-30')])
            batch=store.publications()[0];files=catalog(store,batch)
            for kind in ('reconciliation','payover','claims','product_claims','commissions','internal','combine'):
                item=next(i for i in files if i['kind']==kind);generate(store,batch,item,root/item['name'])
            print('PASS: TEMP-only PostgreSQL publications, immutable versions, all source rows, System REF snapshot, claims components, unpaid status analytics, balance controls and separate report exports.')
    finally:engine.dispose()

if __name__=='__main__':run()
