from collections import defaultdict
from decimal import Decimal
from flask import url_for,g,request,session,render_template,redirect
import calendar
from datetime import date

SPECS=[('original_risk','Original Risk','source','money'),('source_retail','Retail','source','money'),('payover','Insurer Payover','insurers','money'),('aul_payover','AUL Payover','aul','money'),('aul_plan2_rows','AUL Plan 2 rows','plan2','rows'),('aul_historic_rows','AUL Historic rows','historic','rows'),('r1','R1 Fees','fees','money'),('adv','2.1% Fees','fees','money'),('faithnet','Faithnet Deductions','fees','money'),('supplier','Voucher Supplier Payable','vouchers','money'),('retained','Voucher Retained','vouchers','money'),('policy_qty','Paid policy months / MPIA','fees','rows'),('verified_claims','Claim Amount Comparison Over Time','claims','money')]

def default_dates(publications):
    latest=max((b['month'] for b in publications),default=date.today().strftime('%Y-%m'))
    year,month=map(int,latest.split('-'));index=year*12+month-1-11;first_year,first_month=divmod(index,12)
    return dict(date_from=f'{first_year:04d}-{first_month+1:02d}-01',date_to=f'{year:04d}-{month:02d}-{calendar.monthrange(year,month)[1]:02d}')

def chart_geometry(points):
    """Render published values without relying on dashboard JavaScript."""
    values=[float(p['value']) for p in points]
    low=min([0]+values);high=max([1]+values);span=high-low
    y=lambda v:20+(high-v)/span*170
    plotted=[dict(x=300 if len(points)==1 else 70+i*500/(len(points)-1),
                  y=y(v),label=p['label'],value=v) for i,(p,v) in enumerate(zip(points,values))]
    path=' '.join(('M' if i==0 else 'L')+f"{p['x']:.2f} {p['y']:.2f}" for i,p in enumerate(plotted))
    return dict(points=plotted,path=path,baseline=y(0),ticks=[dict(y=y(low+span*i/3),value=low+span*i/3) for i in range(4)],
                label_step=max(1,(len(points)+5)//6))

def graphs(store,start,end,allowed=None,keys=None,franchise='All'):
    months=defaultdict(dict)
    for row in store.summaries(start,end,allowed):
        for key,value in row['summary'].items():months[row['month']][key]=months[row['month']].get(key,Decimal(0))+Decimal(str(value))
    claims=defaultdict(lambda:Decimal(0))
    specs=[s for s in SPECS if keys is None or s[0] in keys]
    result=[]
    for key,label,category,unit in specs:
        if key=='verified_claims':
            for row in store.claims(start,end,allowed):claims[row['month']]+=Decimal(row['amount'])
            verified=store.verified_claim_months()
            values=[(month,claims[month]) for month in sorted(set(months)|set(claims)) if month in verified]
        else:values=[(month,v.get(key,0)) for month,v in sorted(months.items())]
        points=[dict(label=month,value=float(value),url=url_for('central_calculation_detail',category=category,franchise=franchise,date_from=month+'-01',date_to=month+'-'+str(calendar.monthrange(*map(int,month.split('-')))[1]))) for month,value in values]
        result.append(dict(key=key,label=label,category=category,unit=unit,points=points))
    for graph in result:graph['chart']=chart_geometry(graph['points'])
    return result

def graph_scope(app):
    user=getattr(g,'user',None) or {}
    if not user:return None,None,('Authentication required.',401)
    if user.get('role') in ('underwriter','voucher_supplier'):return None,None,('Graphs are outside your report scope.',403)
    allowed=app.extensions['central_page_access']();franchise=request.args.get('franchise','All')
    if franchise!='All':allowed={franchise} if allowed is None or franchise in allowed else set()
    return allowed,franchise,None

def render_page(app):
    from period_selection import month_bounds
    g.graph_page=True
    allowed,franchise,error=graph_scope(app)
    if error:return error
    store=app.extensions['central_store'];publications=store.publications();defaults=default_dates(publications)
    periods=dict(session.get('graph_periods',{}));key=request.args.get('graph')
    if key:
        if key not in {s[0] for s in SPECS}:return 'Unknown graph.',404
        dates={k:request.args.get(k,'') for k in ('date_from','date_to')}
        try:month_bounds(dates)
        except ValueError as exc:return str(exc),400
        if not all(dates.values()):return 'Choose start and end dates.',400
        periods[key]=dates;session['graph_periods']=periods
    grouped=defaultdict(list)
    for key,*_ in SPECS:
        dates=periods.get(key,defaults);grouped[(dates['date_from'],dates['date_to'])].append(key)
    result={}
    for (first,last),keys in grouped.items():
        dates=dict(date_from=first,date_to=last);start,end=month_bounds(dates)
        for graph in graphs(store,start,end,allowed,keys,franchise):result[graph['key']]={**graph,'dates':dates}
    # Do not read policy member indexes, report catalogues or legacy workflows.
    names=sorted({r['franchise'] for r in store.summaries(allowed=app.extensions['central_page_access']())})
    sidebar_dates=session.get('period_dates') or dict(date_from=defaults['date_to'][:7]+'-01',date_to=defaults['date_to'])
    return render_template('dashboard.html',workspace='home',home_section='graphs',home_title='Graphs',graph_page=True,
        central_ready=True,central_graphs=[result[s[0]] for s in SPECS],executive_dashboard={'period_label':'Latest 12 months · dates can be changed per graph'},
        selected=franchise,franchises=names,search_text='',period_dates=sidebar_dates,period_view='month',traffic_filter='All')

def register(app):
    @app.get('/graphs/<key>')
    def update_graph(key):
        from period_selection import month_bounds
        g.graph_page=True;allowed,franchise,error=graph_scope(app)
        if error:return error
        if key not in {s[0] for s in SPECS}:return 'Unknown graph.',404
        dates={k:request.args.get(k,'') for k in ('date_from','date_to')}
        try:start,end=month_bounds(dates)
        except ValueError as exc:return str(exc),400
        if not start:return 'Choose start and end dates.',400
        graph=graphs(app.extensions['central_store'],start,end,allowed,[key],franchise)[0];graph['dates']=dates
        periods=dict(session.get('graph_periods',{}));periods[key]=dates;session['graph_periods']=periods
        response=app.make_response(render_template('central_graph_card.html',graph=graph,selected=franchise));response.headers['Cache-Control']='no-store';return response
