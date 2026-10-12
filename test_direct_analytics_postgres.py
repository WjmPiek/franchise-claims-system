"""Optional PostgreSQL integration check using session-only temporary tables.

Run in the deployed service shell. No production rows or tables are modified.
"""
import os
import sqlite3
from decimal import Decimal
from pathlib import Path
from sqlalchemy import create_engine
from sqlalchemy.pool import StaticPool
from direct_analytics import DirectStore, FIELDS, encode


class TemporaryStore(DirectStore):
    def execute(self, conn, sql, params=None):
        if sql.startswith('CREATE TABLE IF NOT EXISTS '):
            sql = sql.replace('CREATE TABLE IF NOT EXISTS ', 'CREATE TEMP TABLE IF NOT EXISTS ', 1)
        return super().execute(conn, sql, params)


def run():
    url = os.environ.get('DATABASE_URL', '')
    if not url:
        print('Skipped: DATABASE_URL is not configured.')
        return
    if url.startswith('postgres://'):
        url = 'postgresql://' + url[len('postgres://'):]
    engine = create_engine(url, poolclass=StaticPool)
    stage = sqlite3.connect(':memory:')
    stage.execute('CREATE TABLE prepared (' + ','.join(f'{field} TEXT' for field in FIELDS) + ')')
    def fill(amount):
        stage.execute('DELETE FROM prepared')
        data=['ERMELO','Martins','000011','Test, quoted "name"','Example','0012312345678','1980-03-02','','MEM','2020-01-03','50000','Family R 50 000',amount,'150.123','Address','0821234567','0012345','1','','','0001','Agent']
        values = [1,'AUL','Plan 2','ERMELO','000011','MEM',amount,'150.123','0','1','2.08','96.93','0','0','0821234567',encode(data),encode([])]
        stage.execute('INSERT INTO prepared VALUES ('+','.join('?' for _ in FIELDS)+')',values)
        stage.commit()
    try:
        store = TemporaryStore(Path('unused.sqlite'), lambda: engine)
        fill('100.005')
        store.save(stage,'synthetic-a',900000001,'2026-09','synthetic.csv')
        assert store.overview(900000001)['categories']['source']['original_risk'] == Decimal('100.005')
        overview, rows, count, page = store.detail(900000001,None,'plan2')
        assert count == 1 and rows[0]['policy'] == '000011'
        assert overview['detail_totals']['original_risk'] == Decimal('100.005')
        assert store.row(900000001,'synthetic-a',1)['data'][5] == '0012312345678'
        assert store.row(900000002,'synthetic-a',1) is None
        before=store.overview(900000001)['categories']
        products=store.product_insights(900000001)
        assert products['top'][0]['value']==1 and products['products']==['Family R 50 000']
        assert store.overview(900000001)['categories']==before
        assert store.product_insights(900000002)['rows']==[]
        fill('200.005')
        store.save(stage,'synthetic-b',900000001,'2026-09','synthetic.csv')
        assert store.overview(900000001)['categories']['source']['rows'] == 1
        assert store.overview(900000001)['categories']['source']['original_risk'] == Decimal('200.005')
        assert store.row(900000001,'synthetic-a',1) is None
        fresh = TemporaryStore(Path('unused.sqlite'), lambda: engine)
        assert fresh.detail(900000001,None,'source')[2] == 1
        restored=sqlite3.connect(':memory:')
        batch=fresh.restore_stage(900000001,'synthetic-b',restored)
        assert batch['month']=='2026-09'
        assert restored.execute('SELECT policy,original_risk,payover FROM saved').fetchone()==('000011','200.005','96.93')
        restored.close()
        print('PASS: PostgreSQL COPY, exact decimals, detail totals, owner isolation, atomic monthly replacement saved report restoration and product rankings. Temporary test tables disappear when the connection closes.')
    finally:
        stage.close()
        engine.dispose()


if __name__ == '__main__':
    run()
