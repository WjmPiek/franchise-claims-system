"""One combined franchise total and one product row per selected period."""
import json
from decimal import Decimal
from flask import g,request,render_template
from central_calculations import normal

def grouped_products(store,franchise,dates):
    from period_selection import month_bounds
    from claim_insights import populations
    start,end=month_bounds(dates);groups={}
    def product(name):
        return groups.setdefault(normal(name),dict(label=name,rows=0,retail=Decimal(0),payover=Decimal(0),claim_count=0,bank=Decimal(0),netoff=Decimal(0)))
    with store.direct.connection() as conn:
        sql='SELECT r.product,r.summary FROM central_product_monthly r JOIN central_publications p ON p.batch_id=r.batch_id WHERE r.franchise=:f AND p.month BETWEEN :s AND :e'
        for name,raw in store.direct.execute(conn,sql,dict(f=franchise,s=start,e=end)):
            values=json.loads(raw);item=product(name)
            item['rows']+=int(values['rows']);item['retail']+=Decimal(values['retail']);item['payover']+=Decimal(values['payover'])
    _,paid,_=populations(store,dates,{franchise})
    for row in paid:
        item=product(row['product']);item['claim_count']+=int(row['bank']!=0);item['bank']+=row['bank'];item['netoff']+=row['netoff']
    return sorted(groups.values(),key=lambda r:(-r['retail'],-r['claim_count'],r['label'].casefold()))

def register(app,store,access):
    @app.get('/franchise-summary')
    def franchise_summary():
        user=getattr(g,'user',None) or {}
        if not user:return 'Authentication required.',401
        if user.get('role') in ('underwriter','voucher_supplier'):return 'Franchise summary is outside your report scope.',403
        from period_selection import selected_dates,month_bounds
        dates=selected_dates(request.args);franchise=request.args.get('franchise','All');allowed=access()
        if franchise=='All':return 'Select an individual franchise holder.',400
        if allowed is not None and franchise not in allowed:return 'Franchise access required.',403
        rows=grouped_products(store,franchise,dates);start,end=month_bounds(dates);totals={}
        for row in store.summaries(start,end,{franchise}):
            for key,value in row['summary'].items():totals[key]=totals.get(key,Decimal(0))+Decimal(str(value))
        totals['claim_count']=sum(r['claim_count'] for r in rows);totals['claims']=sum((r['bank'] for r in rows),Decimal(0));totals['netoff']=sum((r['netoff'] for r in rows),Decimal(0))
        section=request.args.get('section','risk')
        if section not in ('risk','commissions','policies','top_franchise'):section='risk'
        return render_template('franchise_summary.html',franchise=franchise,dates=dates,rows=rows,totals=totals,section=section)
