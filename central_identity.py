"""Latest published selected month, distinct franchise/System REF MEM identities."""
def member_snapshot(store,start,end,allowed=None):
    publications=[b for b in store.publications() if start<=b['month']<=end]
    latest=max((b['month'] for b in publications),default='')
    branches={};products={};identities=set();missing=0
    for batch in publications:
        if batch['month']!=latest:continue
        with store.direct.connection() as conn:
            rows=store.direct.execute(conn,'SELECT franchise,system_ref,product FROM central_member_index WHERE batch_id=:batch',{'batch':batch['id']})
            for franchise,reference,product in rows:
                if allowed is not None and franchise not in allowed:continue
                if not reference:missing+=1;continue
                identity=(franchise,reference)
                identities.add(identity)
                branches.setdefault(franchise,set()).add(identity)
                products.setdefault(product,set()).add(identity)
    def ranked(values):
        return [{'name':name,'count':len(refs)} for name,refs in sorted(values.items(),key=lambda item:(-len(item[1]),item[0].casefold()))]
    return {'month':latest,'count':len(identities),'missing_refs':missing,'franchises':ranked(branches),'products':ranked(products)}

def products_by_retail(store,start,end,allowed=None):
    import json
    from decimal import Decimal
    totals={}
    with store.direct.connection() as conn:
        rows=store.direct.execute(conn,"SELECT r.franchise,r.product,r.summary FROM central_product_monthly r JOIN central_publications p ON p.batch_id=r.batch_id WHERE p.month BETWEEN :start AND :end AND r.destination<>'Vouchers'",{'start':start,'end':end})
        for franchise,product,summary in rows:
            if allowed is not None and franchise not in allowed:continue
            totals[product]=totals.get(product,Decimal(0))+Decimal(json.loads(summary)['retail'])
    return [{'name':name,'retail':value} for name,value in sorted(totals.items(),key=lambda item:(-item[1],item[0].casefold()))]
