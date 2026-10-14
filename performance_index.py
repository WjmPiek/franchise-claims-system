"""Derived cover lookups; original policy rows and calculations remain immutable."""
import json
from collections import defaultdict
from datetime import datetime
from decimal import Decimal,InvalidOperation

def cover(value):
    try:
        number=Decimal(str(value));return str(number) if number.is_finite() else None
    except (ValueError,InvalidOperation):return None

def write(store,conn,batch,values,member_rows):
    for franchise,product,benefit in sorted(values):
        store.direct.execute(conn,'INSERT INTO central_product_benefits VALUES (:batch,:f,:p,:c) ON CONFLICT(batch_id,franchise,product,cover) DO NOTHING',dict(batch=batch,f=franchise,p=product,c=benefit))
    store.direct.execute(conn,'INSERT INTO central_benefit_versions VALUES (:batch,:rows,:time) ON CONFLICT(batch_id) DO UPDATE SET member_rows=excluded.member_rows,built_at=excluded.built_at',dict(batch=batch,rows=member_rows,time=datetime.utcnow().isoformat()))

def backfill(store,queue=None,job_id=None):
    store.ensure();built=rows=0;publications=store.publications()
    for batch in publications:
        with store.direct.connection() as conn:
            if store.direct.execute(conn,'SELECT batch_id FROM central_benefit_versions WHERE batch_id=:id',{'id':batch['id']}).fetchone():continue
        if queue:queue.update(job_id,'Preparing benefit lookups '+batch['month'],built,len(publications))
        values=set();count=0
        for r in store.records(batch['id'],voucher=False,relation='MEM'):
            c=cover(r['data'][10]);count+=1
            if c is not None:values.add((r['franchise'],r['data'][11],c))
        with store.direct.connection() as conn:write(store,conn,batch['id'],values,count)
        built+=1;rows+=count
    return dict(ok=True,indexed_months=built,member_rows=rows,financial_data_changed=False)

def products(store,start,end,allowed=None):
    from central_calculations import normal
    premium=defaultdict(lambda:Decimal(0));covers=defaultdict(set);names={}
    with store.direct.connection() as conn:
        for r in store.direct.records(store.direct.execute(conn,"SELECT x.franchise,x.product,x.summary FROM central_product_monthly x JOIN central_publications p ON p.batch_id=x.batch_id WHERE p.month BETWEEN :s AND :e AND x.destination<>'Vouchers'",dict(s=start,e=end))):
            if allowed is not None and r['franchise'] not in allowed:continue
            k=normal(r['product']);names[k]=r['product'];premium[k]+=Decimal(json.loads(r['summary'])['payover'])
        ready={r[0] for r in store.direct.execute(conn,'SELECT batch_id FROM central_benefit_versions')}
        for r in store.direct.records(store.direct.execute(conn,'SELECT x.franchise,x.product,x.cover FROM central_product_benefits x JOIN central_publications p ON p.batch_id=x.batch_id WHERE p.month BETWEEN :s AND :e',dict(s=start,e=end))):
            if allowed is None or r['franchise'] in allowed:covers[normal(r['product'])].add(Decimal(r['cover']))
    # Unindexed old versions retain the exact MEM cover fallback until prepared.
    for batch in store.publications():
        if batch['id'] in ready or not start<=batch['month']<=end:continue
        for r in store.records(batch['id'],allowed,voucher=False,relation='MEM'):
            value=cover(r['data'][10])
            if value is not None:covers[normal(r['data'][11])].add(Decimal(value))
    return premium,covers,names
