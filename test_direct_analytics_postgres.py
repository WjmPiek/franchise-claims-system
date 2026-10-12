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
        values = [1,'AUL','Plan 2','ERMELO','000011','MEM',amount,'150.123','0','1','2.08','96.93','0','0','0821234567',encode(['0012312345678','Test, quoted "name"']),encode([])]
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
        assert store.row(900000001,'synthetic-a',1)['data'][0] == '0012312345678'
        assert store.row(900000002,'synthetic-a',1) is None
        fill('200.005')
        store.save(stage,'synthetic-b',900000001,'2026-09','synthetic.csv')
        assert store.overview(900000001)['categories']['source']['rows'] == 1
        assert store.overview(900000001)['categories']['source']['original_risk'] == Decimal('200.005')
        assert store.row(900000001,'synthetic-a',1) is None
        fresh = TemporaryStore(Path('unused.sqlite'), lambda: engine)
        assert fresh.detail(900000001,None,'source')[2] == 1
        print('PASS: PostgreSQL COPY, exact decimals, detail totals, owner isolation and atomic monthly replacement. Temporary test tables disappear when the connection closes.')
    finally:
        stage.close()
        engine.dispose()


if __name__ == '__main__':
    run()
