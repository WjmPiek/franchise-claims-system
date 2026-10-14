"""Audited metadata corrections using approved rules; source and amounts stay intact."""
import json,sqlite3,tempfile
from pathlib import Path
from contextlib import closing
from collections import defaultdict
from datetime import datetime
from decimal import Decimal
from central_calculations import mappings
from central_validation import plan_assignments

UNRESOLVED='Unresolved franchise allocation; transaction retained once in reconciliation.'

def ensure(store):
    store.ensure()
    with store.direct.connection() as conn:
        if store.direct.get_engine:store.direct.execute(conn,"SELECT pg_advisory_xact_lock(hashtext('central-allocation-audit-schema'))")
        store.direct.execute(conn,'CREATE TABLE IF NOT EXISTS central_allocation_audit (kind TEXT NOT NULL,batch_id TEXT NOT NULL,n INTEGER NOT NULL,changed_at TEXT NOT NULL,owner_id INTEGER NOT NULL,before_json TEXT NOT NULL,after_json TEXT NOT NULL,reason TEXT NOT NULL)')
        store.direct.execute(conn,'CREATE INDEX IF NOT EXISTS central_allocation_batch ON central_allocation_audit(batch_id,changed_at)')

def aliases(names,key,alias):
    result=defaultdict(set)
    for name in names:result[alias(key(name))].add(name)
    return result

def repair(store,owner,queue=None,job_id=None):
    import app as application
    ensure(store);d=store.direct;stamp=datetime.utcnow().isoformat();policy_changes=claims_changes=events_changes=0
    known={r['franchise'] for r in store.summaries()}|set(mappings().get('claims_only_branches',[]))
    known_aliases=aliases(known,application.franchise_match_key,application.apply_franchise_alias)
    def resolved(name):
        matches=known_aliases.get(application.apply_franchise_alias(application.franchise_match_key(name)),set())
        return next(iter(matches)) if len(matches)==1 else None
    def audit(conn,kind,batch,n,before,after,reason):
        d.execute(conn,'INSERT INTO central_allocation_audit VALUES (:kind,:batch,:n,:stamp,:owner,:before,:after,:reason)',dict(kind=kind,batch=batch,n=n,stamp=stamp,owner=owner,before=json.dumps(before,default=str),after=json.dumps(after,default=str),reason=reason))
    for batch in store.publications():
        with d.connection() as conn:
            unresolved=d.execute(conn,"SELECT COUNT(*) FROM central_records WHERE batch_id=:id AND plan='Unclassified'",{'id':batch['id']}).fetchone()[0]
        if not unresolved:continue
        if queue:queue.update(job_id,'Applying approved product plans '+batch['month'])
        with tempfile.TemporaryDirectory(dir=queue.root if queue else None) as temporary, closing(sqlite3.connect(str(Path(temporary)/'classification.sqlite'))) as stage:
            stage.execute('CREATE TABLE rows(n INTEGER PRIMARY KEY,branch TEXT,policy TEXT,product TEXT,destination TEXT,relation TEXT,voucher INTEGER,values_json TEXT)')
            stage.execute('CREATE TABLE source_raw(n INTEGER PRIMARY KEY,data TEXT)')
            targets=[]
            for r in store.records(batch['id']):
                stage.execute('INSERT INTO rows VALUES (?,?,?,?,?,?,?,?)',(r['n'],r['franchise'],r['policy'],r['data'][11],r['destination'],r['relation'],int(r['destination']=='Vouchers'),json.dumps(r['data'])))
                stage.execute('INSERT INTO source_raw VALUES (?,?)',(r['n'],json.dumps(r['calculation']['source'])))
                if r['plan']=='Unclassified':targets.append(r)
                if queue and r['n']%5000==0:queue.update(job_id,'Checking policy membership '+batch['month'],r['n'],batch['source_rows'])
            plans,_=plan_assignments(stage)
        with d.connection() as conn:
            if d.get_engine:d.execute(conn,'SELECT pg_advisory_xact_lock(hashtext(:key))',{'key':'allocation:'+batch['id']})
            current=d.execute(conn,'SELECT batch_id FROM central_publications WHERE month=:m',{'m':batch['month']}).fetchone()
            if not current or current[0]!=batch['id']:continue
            for old in targets:
                latest=d.execute(conn,'SELECT plan FROM central_records WHERE batch_id=:batch AND n=:n',{'batch':batch['id'],'n':old['n']}).fetchone()
                if not latest or latest[0]!='Unclassified':continue
                plan=plans.get(old['n'],'Unclassified')
                if plan=='Unclassified':continue
                new={**old,'plan':plan,'issues':[s for s in old['issues'] if not s.startswith('[BLOCK] AUL product or related policy plan is unclassified')]}
                assert old['calculation']==new['calculation'] and old['data']==new['data']
                audit(conn,'policy',batch['id'],old['n'],old,new,'Approved MEM cover / product rule; classification only, no amount or source changes.')
                d.execute(conn,'UPDATE central_records SET plan=:plan,record_json=:record WHERE batch_id=:batch AND n=:n',dict(plan=plan,record=json.dumps(new,default=str),batch=batch['id'],n=old['n']))
                row=d.execute(conn,'SELECT summary FROM central_monthly WHERE batch_id=:batch AND franchise=:f',{'batch':batch['id'],'f':old['franchise']}).fetchone();summary=json.loads(row[0])
                group='plan2' if plan=='Plan 2' else 'historic';summary['aul_'+group+'_rows']=str(Decimal(str(summary.get('aul_'+group+'_rows',0)))+1)
                summary['aul_'+group+'_payover']=str(Decimal(str(summary.get('aul_'+group+'_payover',0)))+Decimal(old['payover']))
                removed=int(bool(old['issues']))-int(bool(new['issues']));summary['review_rows']=int(summary['review_rows'])-removed
                d.execute(conn,'UPDATE central_monthly SET summary=:s WHERE batch_id=:batch AND franchise=:f',dict(s=json.dumps(summary),batch=batch['id'],f=old['franchise']))
                d.execute(conn,'UPDATE central_batches SET review_rows=review_rows-:removed WHERE id=:batch',dict(removed=removed,batch=batch['id']))
                policy_changes+=1
            d.execute(conn,"UPDATE central_batches SET status=CASE WHEN review_rows=0 THEN 'Published' ELSE 'Published with review' END WHERE id=:id",{'id':batch['id']})
            # Product-level aggregates use the corrected plan, with amounts unchanged.
            groups=d.records(d.execute(conn,"SELECT * FROM central_product_monthly WHERE batch_id=:id AND plan='Unclassified'",{'id':batch['id']}))
            for product in groups:
                splits=defaultdict(lambda:dict(rows=Decimal(0),retail=Decimal(0),payover=Decimal(0)))
                for r in targets:
                    if r['franchise']!=product['franchise'] or r['data'][11]!=product['product']:continue
                    v=splits[plans.get(r['n'],'Unclassified')];v['rows']+=1;v['retail']+=Decimal(r['calculation']['retail']);v['payover']+=Decimal(r['payover'])
                params=dict(batch=batch['id'],f=product['franchise'],product=product['product'],dest=product['destination'])
                d.execute(conn,"DELETE FROM central_product_monthly WHERE batch_id=:batch AND franchise=:f AND product=:product AND destination=:dest AND plan='Unclassified'",params)
                for plan,values in splits.items():
                    existing=d.execute(conn,'SELECT summary FROM central_product_monthly WHERE batch_id=:batch AND franchise=:f AND product=:product AND destination=:dest AND plan=:plan',{**params,'plan':plan}).fetchone()
                    if existing:
                        other=json.loads(existing[0]);values={k:values[k]+Decimal(str(other[k])) for k in values}
                    d.execute(conn,'INSERT INTO central_product_monthly VALUES (:batch,:f,:product,:dest,:plan,:summary) ON CONFLICT(batch_id,franchise,product,destination,plan) DO UPDATE SET summary=excluded.summary',{**params,'plan':plan,'summary':json.dumps(values,default=str)})
    with d.connection() as conn:
        if d.get_engine:d.execute(conn,"SELECT pg_advisory_xact_lock(hashtext('claims-allocation-refresh'))")
        rows=d.records(d.execute(conn,'SELECT c.* FROM central_claims c JOIN central_claim_publications p ON p.batch_id=c.batch_id AND p.month=c.month'))
        for row in rows:
            name=resolved(row['franchise'])
            if name is None:continue
            issues=[s for s in json.loads(row['issues']) if s!=UNRESOLVED and not s.startswith('Franchise resolved by the existing normalised alias')]
            if name==row['franchise'] and issues==json.loads(row['issues']):continue
            after={**row,'franchise':name,'issues':json.dumps(issues)};assert row['amount']==after['amount'] and row['source_json']==after['source_json']
            audit(conn,'claims',row['batch_id'],row['n'],row,after,'Approved branch alias or approved own claims-only branch. Original source, dates and payment amounts retained.')
            d.execute(conn,'UPDATE central_claims SET franchise=:f,issues=:issues WHERE batch_id=:batch AND n=:n',dict(f=name,issues=json.dumps(issues),batch=row['batch_id'],n=row['n']));claims_changes+=1
        events=d.records(d.execute(conn,'SELECT e.* FROM central_claim_events e WHERE e.batch_id IN (SELECT DISTINCT batch_id FROM central_claim_publications)'))
        for row in events:
            name=resolved(row['franchise'])
            if name is None or name==row['franchise']:continue
            after={**row,'franchise':name};audit(conn,'claims_source',row['batch_id'],row['n'],row,after,'Apply approved branch name for source analytics; raw claim case retained.')
            d.execute(conn,'UPDATE central_claim_events SET franchise=:f WHERE batch_id=:batch AND n=:n',dict(f=name,batch=row['batch_id'],n=row['n']));events_changes+=1
    return dict(ok=True,policy_classifications=policy_changes,claim_allocations=claims_changes,claim_cases=events_changes,amounts_changed=False)
