"""Focused regression checks, loading reporting functions without web startup."""
import ast
from pathlib import Path
from datetime import datetime
from types import SimpleNamespace
from unittest.mock import patch
import os
import re
import pandas as pd

source = Path(__file__).parent / 'app.py'
if not source.exists():
    source = Path(__file__).parent / 'franchise_claims_system' / 'app.py'
tree = ast.parse(source.read_text(encoding='utf-8'))
names = {'display_source_filename', 'month_from_filename', '_dashboard_month_filter',
         'load_raw_from_postgres', 'reload_dashboard_from_postgres', 'load_logged_in_user'}
nodes = [n for n in tree.body if isinstance(n, ast.FunctionDef) and n.name in names]
for n in nodes:
    n.decorator_list = []
ns = dict(pd=pd, os=os, re=re, datetime=datetime)
exec(compile(ast.Module(body=nodes, type_ignores=[]), str(source), 'exec'), ns)

parse = ns['month_from_filename']
prefix = '20260401-aaaa-4bbb-8ccc-000000000000_'
for month in range(5, 10):
    assert parse(prefix + f'PolicyData_2026{month:02d}01_to_2026{month:02d}28.xlsx') == pd.Timestamp(2026, month, 1)
assert parse('PolicyData_September_2026.xlsx') == pd.Timestamp('2026-09-01')
assert parse('PolicyData_2026_September.xlsx') == pd.Timestamp('2026-09-01')
assert parse('PolicyData_2026-09-01_to_2026-09-30.xlsx') == pd.Timestamp('2026-09-01')
for invalid in ['PolicyData.xlsx', 'PolicyData_20261301.xlsx']:
    try:
        parse(invalid)
    except ValueError:
        pass
    else:
        raise AssertionError('An unknown report month must not be guessed')

monthly = pd.DataFrame({'Month': pd.date_range('2026-01-01', periods=9, freq='MS'),
                        'Retail Premium': 100, 'Risk Premium': 80})
filtered, label = ns['_dashboard_month_filter'](monthly, 'month')
assert label == 'Sep 2026' and len(filtered) == 1
filtered, label = ns['_dashboard_month_filter'](monthly, 'six_months')
assert label == 'Apr 2026 - Sep 2026' and len(filtered) == 6
claims_only = monthly.copy()
claims_only.loc[claims_only.Month > '2026-04-01', ['Retail Premium', 'Risk Premium']] = 0
assert ns['_dashboard_month_filter'](claims_only, 'month')[1] == 'Apr 2026'

ns.update(get_db_engine=lambda: object(), get_import_history_summary=lambda: {},
          merge_claims_into_raw=lambda policy, claims: claims if policy.empty else policy,
          analyse=lambda raw, rates, book: (raw.rename(columns={'month':'Month'}), raw.copy(), {}),
          DEFAULT_RATES={}, DEFAULT_BOOK_VALUE={}, LAST_RESULT={'raw': pd.DataFrame({'month':[pd.Timestamp('2026-04-01')]})},
          LAST_CLAIMS_DF=pd.DataFrame({'claims':[999]}))
policy = pd.DataFrame({'franchise':['Example'], 'month':[pd.Timestamp('2026-09-01')]})
empty_claims = pd.DataFrame(columns=['franchise', 'month', 'claims'])
with patch.object(pd, 'read_sql', side_effect=[policy, empty_claims]):
    assert ns['reload_dashboard_from_postgres'](strict=True)
assert ns['LAST_RESULT']['monthly'].Month.max() == pd.Timestamp('2026-09-01')
assert ns['LAST_CLAIMS_DF'].empty
with patch.object(pd, 'read_sql', side_effect=[policy.iloc[:0], empty_claims]):
    assert not ns['reload_dashboard_from_postgres'](strict=True)
assert ns['LAST_RESULT']['monthly'].empty
claims = pd.DataFrame({'franchise':['Example'], 'month':[pd.Timestamp('2026-09-01')], 'claims':[100]})
with patch.object(pd, 'read_sql', side_effect=[policy.iloc[:0], claims]):
    assert ns['reload_dashboard_from_postgres'](strict=True)
assert not ns['LAST_CLAIMS_DF'].empty
with patch.object(pd, 'read_sql', side_effect=RuntimeError('connection lost')):
    try:
        ns['reload_dashboard_from_postgres'](strict=True)
    except RuntimeError:
        pass
    else:
        raise AssertionError('Database failure must not return cached reports')

user = {'id':1, 'is_active':True, 'role':'admin'}
ns.update(g=SimpleNamespace(), session={'user_id':1}, get_user_by_id=lambda _:user,
          MAINTENANCE_MODE=False, PUBLIC_ENDPOINTS={'healthz','static'}, ADMIN_ENDPOINTS=set(),
          record_user_activity=lambda _:None, DATABASE_URL='configured')
calls=[]
ns['reload_dashboard_from_postgres'] = lambda **kw:calls.append(kw) or True
for endpoint in ['dashboard', 'workspace_page', 'export', 'export_payover', 'board_report', 'client_heatmap_page']:
    ns['request']=SimpleNamespace(endpoint=endpoint, method='GET')
    assert ns['load_logged_in_user']() is None
assert len(calls)==6 and all(c == {'strict':True} for c in calls)
def fail_refresh(**kw):
    raise RuntimeError('unavailable')
ns['reload_dashboard_from_postgres']=fail_refresh
assert ns['load_logged_in_user']()[1] == 503
print('PASS: filename dates, September/6-month filters, stale cache refresh, empty database, claims-only load, reporting routes and database failure handling.')

