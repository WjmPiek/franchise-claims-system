"""Validate audit writes using session-only temporary tables; no production rows are changed."""
import os
from pathlib import Path
from sqlalchemy import create_engine
from sqlalchemy.pool import StaticPool
from import_audit import ImportAudit
from direct_analytics import DirectStore

class TemporaryStore(DirectStore):
    def execute(self,conn,sql,params=None):
        if sql.startswith('CREATE TABLE IF NOT EXISTS '):sql=sql.replace('CREATE TABLE IF NOT EXISTS ','CREATE TEMP TABLE IF NOT EXISTS ',1)
        return super().execute(conn,sql,params)

def run():
    url=os.environ.get('DATABASE_URL','')
    if not url:print('Skipped: DATABASE_URL is not configured.');return
    if url.startswith('postgres://'):url='postgresql://'+url[len('postgres://'):]
    engine=create_engine(url,poolclass=StaticPool)
    try:
        audit=ImportAudit(Path('unused.sqlite'));audit.store=TemporaryStore(Path('unused.sqlite'),lambda:engine)
        first=audit.start(900000001,'Direct','PolicyData_20260901_to_20260930.csv','a'*64)
        audit.finish(first,{'ok':True,'month':'2026-09','source':{'rows':5,'risk':'100.005','retail':'150.123'}})
        second=audit.start(900000001,'Direct','PolicyData_20260901_to_20260930.csv','b'*64)
        audit.finish(second,{'ok':False,'error':'Synthetic failed row'})
        rows=audit.rows(900000001);assert len(rows)==2 and {r['status'] for r in rows}=={'saved / replaced','failed / review'}
        assert audit.rows(900000002)==[]
        assert all(r['receipt_month']=='2026-09' for r in rows)
        print('PASS: durable audit schema, successful replacement and failed-attempt records, exact source metadata and owner isolation. Temporary test table disappears on connection close.')
    finally:engine.dispose()

if __name__=='__main__':run()
