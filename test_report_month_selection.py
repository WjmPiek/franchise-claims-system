import ast,re
from pathlib import Path
import pandas as pd
source=Path(__file__).with_name('app.py').read_text(encoding='utf-8')
tree=ast.parse(source)
names={'_report_month_options','_select_report_months','_dashboard_month_filter'}
ns={'pd':pd,'re':re,'load_franchise_config':lambda:{'groups':{'A':'GROUP','B':'GROUP'},'use_groups':True}}
exec(compile(ast.Module(body=[n for n in tree.body if isinstance(n,ast.FunctionDef) and n.name in names],type_ignores=[]),'report_helpers','exec'),ns)
df=pd.DataFrame([{'Franchise':f,'Month':f'2026-{m:02d}-01','Retail Premium':m*100,'Risk Premium':m*60,'Claims':m*10} for f in ('A','B') for m in range(1,10)])
select=ns['_select_report_months']
row,label=select(df,{'report_period':'month','report_month':'2026-09','franchise':'A'})
assert len(row)==1 and row['Retail Premium'].sum()==900 and label=='Sep 2026'
row,label=select(df,{'report_period':'range','report_start':'2026-05','report_end':'2026-09'})
assert len(row)==10 and row['Retail Premium'].sum()==7000 and label=='May 2026 - Sep 2026'
row,label=select(df,{'report_period':'all'});assert len(row)==len(df)
row,label=select(df,{},'six_months'); assert len(row)==12 and label=='Apr 2026 - Sep 2026'
for args in ({'report_period':'range','report_start':'2026-09','report_end':'2026-05'},{'report_period':'month','report_month':'2026-13'},{'report_period':'month','report_month':'2025-12'},{'franchise':'UNAUTHORIZED','report_period':'all'}):
    try: select(df,args)
    except ValueError: pass
    else: raise AssertionError(args)
print('PASS: single month, inclusive range, all data, six months, invalid ranges/months and scoped franchise.')


row,label=select(df,{'report_period':'month','report_month':'2026-09','franchise':'GROUP'}); assert len(row)==2
