"""Read-only verification of archived uploads against their saved versions."""
import calendar,json,sqlite3,tempfile
from collections import defaultdict,Counter
from contextlib import closing
from datetime import datetime
from decimal import Decimal
from pathlib import Path
from flask import g,request,render_template,jsonify
from sqlalchemy import text

def ensure(store):
    store.ensure()
    with store.direct.connection() as conn:
        store.direct.execute(conn,'CREATE TABLE IF NOT EXISTS central_import_verification (batch_id TEXT PRIMARY KEY,kind TEXT NOT NULL,checked_at TEXT NOT NULL,result TEXT NOT NULL)')

def source_job(store,batch):
    with store.direct.connection() as conn:
        rows=store.direct.records(store.direct.execute(conn,'SELECT id,args FROM central_jobs WHERE id=:id AND source_bytes>0',{'id':batch}))
    if not rows:raise ValueError('Original uploaded file is not linked to this version. Re-import the original source if full verification is required.')
    job=rows[0];job['args']=json.loads(job['args']);return job

def streamed(store,sql,params):
    with store.direct.connection() as conn:
        rows=conn.execution_options(stream_results=True).execute(text(sql),params) if store.direct.get_engine else conn.execute(sql,params)
        for row in rows:yield row

def policy_check(store,queue,batch):
    from direct_data import stage_source
    path=queue.restore_source(source_job(store,batch))
    with tempfile.TemporaryDirectory(dir=queue.root) as directory:
        with closing(sqlite3.connect(str(Path(directory)/'source.sqlite'))) as stage:source=stage_source(path,stage)
    count=review=0;risk=retail=Decimal(0);reasons=Counter()
    for row in streamed(store,'SELECT record_json FROM central_records WHERE batch_id=:id',{'id':batch}):
        record=json.loads(row[0]);count+=1;risk+=Decimal(str(record['original_risk']));retail+=Decimal(str(record['retail']));review+=bool(record['issues'])
        reasons.update(record['issues'])
    controls=[dict(label='Source records',source=str(source['rows']),stored=str(count),difference=str(count-int(source['rows']))),
      dict(label='Original Risk',source=str(source['risk']),stored=str(risk),difference=str(risk-Decimal(str(source['risk'])))),
      dict(label='Total Retail',source=str(source['retail']),stored=str(retail),difference=str(retail-Decimal(str(source['retail']))))]
    return dict(matched=all(Decimal(c['difference'])==0 for c in controls),controls=controls,review=review,review_reasons=dict(reasons.most_common()),method='Original uploaded bytes checked by SHA-256; source file re-read and compared with individual stored records.')

def claims_check(store,queue,batch):
    from central_claims import read_transactions,verified_annual_sheet,VERIFIED_AMOUNT_HEADER,component
    path=queue.restore_source(source_job(store,batch));controls={};source=defaultdict(lambda:dict(rows=0,bank=Decimal(0),netoff=Decimal(0)))
    for r in read_transactions(path,VERIFIED_AMOUNT_HEADER,verified_annual_sheet(path),controls):
        m=source[r['month']];m['rows']+=1;m['bank']+=Decimal(r['amount']);m['netoff']+=component(r['source'],'Amount net off against claim')
    stored=defaultdict(lambda:dict(rows=0,bank=Decimal(0),netoff=Decimal(0)))
    reviews=0;reasons=Counter()
    for row in streamed(store,'SELECT c.month,c.amount,x.net_off,c.issues FROM central_claims c LEFT JOIN central_claim_components x ON x.batch_id=c.batch_id AND x.n=c.n WHERE c.batch_id=:id',{'id':batch}):
        m=stored[row[0]];m['rows']+=1;m['bank']+=Decimal(row[1]);m['netoff']+=Decimal(row[2] or '0')
        reviews+=bool(json.loads(row[3]))
        reasons.update(json.loads(row[3]))
    checks=[]
    for month in sorted(set(source)|set(stored)):
        for key,label in [('rows','Financial records'),('bank','Bank-paid claims (AM)'),('netoff','Dated claim net-offs (AL)')]:
            a=source[month][key];b=stored[month][key];checks.append(dict(month=month,label=label,source=str(a),stored=str(b),difference=str(b-a)))
    with store.direct.connection() as conn:
        events=store.direct.execute(conn,'SELECT COUNT(*) FROM central_claim_events WHERE batch_id=:id',{'id':batch}).fetchone()[0]
    all_rows=int(controls.get('source_rows',0));checks.insert(0,dict(label='All source claim cases, including unpaid / repudiated',source=str(all_rows),stored=str(events),difference=str(events-all_rows)))
    return dict(matched=all(Decimal(c['difference'])==0 for c in checks),controls=checks,review=reviews+controls.get('undated_netoff_rows',0),
      method='Original uploaded bytes checked by SHA-256; all claim cases counted and financial rows compared per actual paid month. Undated net-offs stay outside paid-month totals.',source_notes=controls,review_reasons=dict(reasons.most_common()))

def verify_targets(store,queue,targets,job_id):
    ensure(store);results=[]
    for index,target in enumerate(targets):
        queue.update(job_id,'Checking '+target['kind']+' '+target['month'],index,len(targets))
        try:result=(policy_check if target['kind']=='policy' else claims_check)(store,queue,target['batch'])
        except Exception as exc:
            queue.app.logger.exception('Import verification failed')
            result=dict(matched=False,unavailable=True,error=str(exc)[:1000],controls=[])
        stamp=datetime.utcnow().isoformat()
        with store.direct.connection() as conn:
            store.direct.execute(conn,'INSERT INTO central_import_verification VALUES (:batch,:kind,:stamp,:result) ON CONFLICT(batch_id) DO UPDATE SET kind=excluded.kind,checked_at=excluded.checked_at,result=excluded.result',dict(batch=target['batch'],kind=target['kind'],stamp=stamp,result=json.dumps(result,default=str)))
        results.append(dict(batch=target['batch'],matched=result['matched']))
    return dict(ok=True,checked=len(results),matched=sum(r['matched'] for r in results),needs_review=sum(not r['matched'] for r in results))

def inventory(store):
    ensure(store);d=store.direct
    with d.connection() as conn:
        policies=d.records(d.execute(conn,'SELECT b.*,p.batch_id AS current_id FROM central_batches b LEFT JOIN central_publications p ON p.month=b.month ORDER BY b.month DESC,b.created_at DESC'))
        claims=d.records(d.execute(conn,'SELECT b.*,c.month,p.batch_id AS current_id FROM central_claim_batches b LEFT JOIN (SELECT DISTINCT batch_id,month FROM central_claims UNION SELECT batch_id,month FROM central_claim_publications) c ON c.batch_id=b.id LEFT JOIN central_claim_publications p ON p.month=c.month ORDER BY c.month DESC,b.created_at DESC'))
        checks={r['batch_id']:r for r in d.records(d.execute(conn,'SELECT * FROM central_import_verification'))}
        historical=[]
        if d.get_engine:
            exists=d.execute(conn,"SELECT to_regclass('public.policy_monthly_raw')").fetchone()[0]
            if exists:historical=d.records(d.execute(conn,"SELECT to_char(import_month,'YYYY-MM') AS month,COUNT(*) AS branch_rows,SUM(retail_premium) AS retail,SUM(risk_premium) AS risk FROM policy_monthly_raw GROUP BY 1 ORDER BY 1 DESC"))
    for row in policies+claims:
        row['check']=checks.get(row['id'])
        if row['check']:row['check']={**row['check'],'result':json.loads(row['check']['result'])}
    return policies,claims,historical

def register(app,store,queue):
    def admin():return (getattr(g,'user',None) or {}).get('role')=='admin'
    @app.get('/imports/verification')
    def import_verification_page():
        if not admin():return 'Administrator access required.',403
        queue.ensure();policies,claims,historical=inventory(store)
        months=sorted({r['month'] for r in policies+claims+historical if r.get('month')})
        start=request.args.get('date_from') or (months[0]+'-01' if months else datetime.utcnow().strftime('%Y-%m-01'))
        end=request.args.get('date_to') or (months[-1]+'-'+str(calendar.monthrange(int(months[-1][:4]),int(months[-1][5:]))[1]) if months else datetime.utcnow().date().isoformat())
        from period_selection import month_bounds
        try:first,last=month_bounds(dict(date_from=start,date_to=end))
        except ValueError as exc:return str(exc),400
        policies=[r for r in policies if first<=r['month']<=last];claims=[r for r in claims if r['month'] and first<=r['month']<=last]
        historical=[r for r in historical if first<=r['month']<=last and not any(p['month']==r['month'] and p['current_id']==p['id'] for p in policies)]
        found={r['month'] for r in policies+historical}
        import pandas as pd
        missing=[str(m) for m in pd.period_range(first,last,freq='M') if str(m) not in found]
        return render_template('import_verification.html',policies=policies,claims=claims,historical=historical,missing=missing,date_from=start,date_to=end)
    @app.post('/imports/verification/run')
    def import_verification_run():
        if not admin():return jsonify(error='Administrator access required.'),403
        from period_selection import month_bounds
        try:first,last=month_bounds(request.form)
        except ValueError as exc:return jsonify(error=str(exc)),400
        if not first:return jsonify(error='Choose start and end dates.'),400
        queue.ensure();policies,claims,_=inventory(store)
        targets=[dict(kind='policy',month=p['month'],batch=p['id']) for p in policies if p['id']==p['current_id'] and first<=p['month']<=last]
        seen=set()
        for r in claims:
            if r['month'] and r['status']!='Confirmed no paid claims' and r['id']==r['current_id'] and first<=r['month']<=last and r['id'] not in seen:targets.append(dict(kind='claims',month=r['month'],batch=r['id']));seen.add(r['id'])
        if not targets:return jsonify(error='No current source files in this date range.'),400
        identifier=queue.create(g.user['id'],'verification',dict(targets=targets))
        return jsonify(job=identifier)
