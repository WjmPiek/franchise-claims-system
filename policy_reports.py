"""Relation premium reports with independently entered external controls."""
import csv
import io
import json
from decimal import Decimal, InvalidOperation
import pandas as pd
from flask import request, render_template, send_file

MONEY = ['original_risk_premium', 'retail_premium', 'r1_policy_fee', 'adv_fund_2_1_fee', 'new_risk_premium']

def safe_csv(value):
    return "'" + value if isinstance(value, str) and value.startswith(('=', '+', '-', '@', '\t', '\r')) else value

def register_policy_reports(app, get_engine, sql_text, get_memory, scope_fn=None):
    @app.route('/policy_reports')
    def policy_reports():
        engine = get_engine()
        memory = get_memory()
        if engine is not None:
            with engine.begin() as conn:
                months = [str(r[0])[:7] for r in conn.execute(sql_text('SELECT DISTINCT import_month FROM policydata_detail_raw ORDER BY import_month DESC'))]
                franchises = [r[0] for r in conn.execute(sql_text('SELECT DISTINCT franchise_name FROM policydata_detail_raw ORDER BY franchise_name'))]
        else:
            months = sorted(set(pd.to_datetime(memory.get('import_month', pd.Series(dtype=object))).dt.strftime('%Y-%m')), reverse=True)
            franchises = sorted(set(memory.get('franchise', pd.Series(dtype=str))))
        if scope_fn is not None:
            scoped = scope_fn(pd.DataFrame({'Franchise':franchises}))
            franchises = scoped['Franchise'].tolist()
        month = request.args.get('month') or (months[0] if months else '')
        franchise = request.args.get('franchise', 'All')
        relation = request.args.get('relation', 'All')
        if month and month not in months:
            return 'Choose an available imported month.', 400
        if franchise != 'All' and franchise not in franchises:
            return 'Choose an available franchise.', 400
        params = {'month': month + '-01'}
        where = ['import_month = :month']
        if scope_fn is not None:
            for i,f in enumerate(franchises): params[f'scope_{i}'] = f
            where.append('franchise_name IN (' + ','.join(f':scope_{i}' for i in range(len(franchises))) + ')' if franchises else '1=0')
        if franchise != 'All':
            where.append('franchise_name = :franchise'); params['franchise'] = franchise
        if relation != 'All':
            where.append("COALESCE(NULLIF(UPPER(TRIM(relation)), ''), '(Blanks)') = :relation"); params['relation'] = relation
        if engine is not None and month:
            with engine.begin() as conn:
                query = "SELECT COALESCE(NULLIF(UPPER(TRIM(relation)), ''), '(Blanks)') AS relation, COUNT(*) AS rows, "
                query += ', '.join(f'COALESCE(SUM({c}),0) AS {c}' for c in MONEY)
                query += ' FROM policydata_detail_raw WHERE ' + ' AND '.join(where) + ' GROUP BY 1 ORDER BY 1'
                groups = [dict(r) for r in conn.execute(sql_text(query), params).mappings()]
        else:
            frame = memory.copy()
            if not frame.empty:
                frame = frame[pd.to_datetime(frame.import_month).dt.strftime('%Y-%m') == month]
                if scope_fn is not None: frame = frame[frame.franchise.isin(franchises)]
                if franchise != 'All': frame = frame[frame.franchise == franchise]
                frame['relation'] = frame.relation.fillna('').str.strip().str.upper().replace('', '(Blanks)')
                if relation != 'All': frame = frame[frame.relation == relation]
            groups = []
            if not frame.empty:
                for rel, rows in frame.groupby('relation'):
                    groups.append(dict(relation=rel, rows=len(rows), **{c:sum(Decimal(str(v)) for v in rows[c]) for c in MONEY}))
        totals = {c:sum(Decimal(str(g[c])) for g in groups) for c in MONEY}
        totals['rows'] = sum(g['rows'] for g in groups)
        controls = {}
        for field, key in [('external_risk', 'original_risk_premium'), ('external_retail', 'retail_premium')]:
            value = request.args.get(field, '').strip()
            if value:
                try:
                    external = Decimal(value.replace(',', ''))
                    if not external.is_finite() or external != external.quantize(Decimal('.01')): raise InvalidOperation
                except InvalidOperation:
                    return 'External totals must be finite amounts with at most two decimals.', 400
                delta = external - totals[key]
                controls[field] = {'value':external, 'difference':delta if groups else None, 'status':('Balanced' if delta == 0 else 'Difference to investigate') if groups else 'No imported rows; comparison unavailable'}
        download = request.args.get('download')
        if download:
            output = io.StringIO()
            writer = csv.writer(output)
            if download == 'detail':
                writer.writerow(['Source file','Excel row','Franchise','Month','Relation','Policy number','Risk before fees','Retail','R1 fee','ADV fee','Net risk'])
                def write_record(r):
                    raw = r.get('raw_data') or {}
                    if isinstance(raw,str): raw=json.loads(raw)
                    policy = raw.get('Policy Number') or raw.get('Policy No') or raw.get('Policy') or ''
                    # Neutralize formulas when opening identifiers from untrusted workbooks in Excel.
                    def safe(v):
                        return "'"+v if isinstance(v,str) and v.startswith(('=','+','-','@','\t','\r')) else v
                    writer.writerow([safe(r.get('source_file','')),r.get('row_number'),safe(r.get('franchise_name',r.get('franchise',''))),month,safe(r.get('relation') or '(Blanks)'),safe(str(policy)),*[r.get(c,0) for c in MONEY]])
                if engine is not None and month:
                    with engine.connect() as conn:
                        query = 'SELECT source_file, row_number, franchise_name, relation, raw_data, ' + ', '.join(MONEY) + ' FROM policydata_detail_raw WHERE ' + ' AND '.join(where) + ' ORDER BY franchise_name, row_number'
                        for r in conn.execution_options(stream_results=True).execute(sql_text(query),params).mappings(): write_record(dict(r))
                elif not memory.empty:
                    for r in frame.to_dict('records'): write_record(r)
            elif download == 'summary':
                writer.writerow(['Month','Franchise','Relation','Rows',*MONEY])
                for g in groups: writer.writerow([month,safe_csv(franchise),safe_csv(g['relation']),g['rows'],*[g[c] for c in MONEY]])
                writer.writerow([month,safe_csv(franchise),'TOTAL',totals['rows'],*[totals[c] for c in MONEY]])
                for k,v in controls.items():writer.writerow([k,'Other system',v['value'],'Other system minus imported',v['difference'],v['status']])
            else: return 'Unknown report format.',400
            return send_file(io.BytesIO(output.getvalue().encode('utf-8-sig')),as_attachment=True,download_name=f'policy_{month}_{download}.csv',mimetype='text/csv')
        return render_template('policy_reports.html', groups=groups,totals=totals,controls=controls,months=months,franchises=franchises,month=month,franchise=franchise,relation=relation,
            relations=['DEP','DEPEXTRACHILD','EXT','MEM','MEMBEN','SPS','(Blanks)','EXTBEN','SPBEN','DEPBEN'],args=request.args)
