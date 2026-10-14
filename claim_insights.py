"""Claims source cohorts and paid transactions, with explicit dates and provenance."""
import json,sqlite3,time
from collections import defaultdict
from decimal import Decimal,InvalidOperation
from datetime import date
from flask import g,request,render_template,jsonify,url_for
from central_calculations import normal
from central_claims import component,paid_date,read_transactions,VERIFIED_AMOUNT_HEADER

VERSION='claims-insights-20261013-v1'
DAYS='Number of days between Incident Date/Entry Date'
_CACHE={}
_SELECTED_SCOPE=object()

def sales_period(dates):
    """Twelve complete receipt/payment months ending in the selected end month."""
    import calendar
    from dateutil.relativedelta import relativedelta
    end=date.fromisoformat(dates['date_to']).replace(day=1)
    start=end-relativedelta(months=11)
    return dict(date_from=start.isoformat(),date_to=end.replace(day=calendar.monthrange(end.year,end.month)[1]).isoformat())

def sales_selection(args,dates):
    from werkzeug.exceptions import BadRequest
    if args.get('product_scenario','ratio') not in ('ratio','count'):raise BadRequest('Choose a valid product scenario.')
    if not args.get('product_from') and not args.get('product_to'):return sales_period(dates)
    try:
        start=date.fromisoformat(args.get('product_from',''));end=date.fromisoformat(args.get('product_to',''))
        if start>end:raise ValueError()
    except ValueError:raise BadRequest('Enter valid product From and To dates, with From before To.')
    return dict(date_from=start.isoformat(),date_to=end.isoformat())

def branch_selection(args,dates):
    from werkzeug.exceptions import BadRequest
    if args.get('branch_scenario','name') not in ('name','count','amount'):raise BadRequest('Choose a valid branch scenario.')
    if not args.get('branch_from') and not args.get('branch_to'):return sales_period(dates)
    try:
        start=date.fromisoformat(args.get('branch_from',''));end=date.fromisoformat(args.get('branch_to',''))
        if start>end:raise ValueError()
    except ValueError:raise BadRequest('Enter valid branch From and To dates, with From before To.')
    return dict(date_from=start.isoformat(),date_to=end.isoformat())

def rank_branches(rows,scenario='name'):
    return sorted(rows,key=(lambda r:(-r['count'],-r['claims'],r['franchise'].casefold())) if scenario=='count' else (lambda r:(-r['claims'],-r['count'],r['franchise'].casefold())) if scenario=='amount' else lambda r:r['franchise'].casefold())
SPECS={'relation':'Relationships of life assured','age':'Age categories','product':'Top 10 claimed products','product_age':'Products by age category',
 'cover':'Product cover claimed','gender':'Policyholder gender','cause':'Incident causes','status':'Approved and repudiated',
 'payee':'Paid to franchise holder or client','netoff':'Net-offs against claims','duration':'Average product duration to incident',
 'branch_month':'Paid claims by branch and month','risk':'Product sales flags','risk_franchise':'Franchises claiming red products',
 'early':'Early incidents within 180 days','delay':'Days from reported to paid','quality':'Claims data to review'}

class FastFields(dict):
    def __init__(self,source):
        super().__init__(source);self.normalised={}
        for k,v in source.items():self.normalised.setdefault(normal(k),str(v).strip())

def field(source,name):
    if isinstance(source,FastFields):return source.normalised.get(normal(name),'')
    return next((str(v).strip() for k,v in source.items() if normal(k)==normal(name)),'')

def source_date(source,name):
    raw=field(source,name)
    if not raw:return ''
    try:return paid_date(raw,0).strftime('%Y-%m-%d')
    except ValueError:return ''

def number(value):
    try:
        result=Decimal(str(value).strip().replace(',',''))
        return result if result.is_finite() else None
    except InvalidOperation:return None

def classify(source):
    source=FastFields(source)
    rel=field(source,'Relationship of Life Assured/Member');r=normal(rel)
    relation='Main Life' if r in ('MAINLIFE','MEM','MEMBER','MEMBERMAINLIFE') else 'Extended Family' if r in ('EXT','EXTENDEDFAMILY') else 'Spouse' if r in ('SPS','SPOUSE') else 'Children' if r.startswith('CHILD') or r in ('DEP','DEPEXTRACHILD') else rel or 'Unknown'
    status=field(source,'Claim Status');s=status.casefold()
    outcome='Repudiated' if 'repudiat' in s else 'Approved' if 'approved' in s else 'Awaiting' if 'awaiting' in s else 'Other / Unknown'
    gender=field(source,'Policyholder Gender');gender={'M':'Male','MALE':'Male','F':'Female','FEMALE':'Female','OTHER':'Other'}.get(normal(gender),'Other / Unknown')
    payee={'ENTITY':'Franchise holder','INDIVIDUAL':'Client / beneficiary'}.get(normal(field(source,'PayableTo')),'Unknown')
    days=number(field(source,DAYS));days=days if days is not None and days>=0 else None
    return dict(relation=relation,relation_source=rel,age=field(source,'Age Category') or 'Unknown',
      gender=gender,cause=field(source,'Incident Cause') or 'Unknown',status=outcome,status_source=status,
      payee=payee,payee_source=field(source,'PayableTo'),days=days,years=days/Decimal('365.25') if days is not None else None,
      reported_on=source_date(source,'Claim Reported Date'),paid_on=source_date(source,'Date Of Claim Paid'),
      incident_on=source_date(source,'Date of Incident'),product=field(source,'Product') or 'Unallocated',
      administrator=field(source,'Administrator'),claim=field(source,'Claim Number') or field(source,'ClaimNumber_Administrator'),
      bank=component(source,VERIFIED_AMOUNT_HEADER),netoff=component(source,'Amount net off against claim'),
      reported_amount=component(source,'Amount of Claim Reported'))

def persist_events(store,conn,batch,rows):
    """Publish source analysis inside the import transaction without altering paid amounts."""
    aliases=defaultdict(set)
    for item in store.direct.execute(conn,'SELECT DISTINCT m.franchise FROM central_monthly m JOIN central_publications p ON p.batch_id=m.batch_id'):aliases[normal(item[0])].add(item[0])
    paid={}
    for row in store.direct.execute(conn,'SELECT n,franchise,source_json FROM central_claims WHERE batch_id=:batch',{'batch':batch}):
        paid[row[0]]=row[1];original=json.loads(row[2])
        aliases[normal(field(original,'Participating Group') or field(original,'Franchise') or field(original,'Branch'))].add(row[1])
    count=0
    for n,raw in rows:
        source=json.loads(raw);franchise=field(source,'Participating Group') or field(source,'Franchise') or field(source,'Branch')
        matches=aliases.get(normal(franchise),set())
        franchise=paid.get(n,next(iter(matches)) if len(matches)==1 else franchise)
        store.direct.execute(conn,'INSERT INTO central_claim_events VALUES (:batch,:n,:franchise,:source) ON CONFLICT(batch_id,n) DO NOTHING',dict(batch=batch,n=n,franchise=franchise,source=raw));count+=1
    store.direct.execute(conn,'INSERT INTO central_claim_event_versions VALUES (:batch,:version,:rows) ON CONFLICT(batch_id) DO NOTHING',dict(batch=batch,version=VERSION,rows=count))
    return count

def backfill(store,queue,job_id):
    """Read archived original workbooks. Never replace financial publications."""
    store.ensure();queue.ensure();d=store.direct
    with d.connection() as conn:
        jobs=d.records(d.execute(conn,"SELECT DISTINCT j.id,j.args FROM central_jobs j JOIN central_claim_publications p ON p.batch_id=j.id LEFT JOIN central_claim_event_versions e ON e.batch_id=j.id WHERE j.kind='claims' AND e.batch_id IS NULL"))
    count=0
    for job in jobs:
        job['args']=json.loads(job['args']);path=queue.restore_source(job)
        try:
            with sqlite3.connect(':memory:') as stage:
                stage.execute('CREATE TABLE events(n INTEGER PRIMARY KEY,source TEXT)')
                for _ in read_transactions(path,VERIFIED_AMOUNT_HEADER,job['args'].get('sheet',''),event_sink=lambda n,s:stage.execute('INSERT INTO events VALUES (?,?)',(n,json.dumps(s)))):pass
                with d.connection() as conn:count+=persist_events(store,conn,job['id'],stage.execute('SELECT n,source FROM events ORDER BY n'))
            queue.update(job_id,'Analysing archived claims sources',count)
        finally:path.unlink(missing_ok=True)
    return dict(ok=True,source={'rows':count},analysed_files=len(jobs),financial_data_changed=False)

def _source_populations(store):
    store.ensure();d=store.direct;paid=[];seen={};missing=[]
    with d.connection() as conn:
        batches=d.records(d.execute(conn,'SELECT DISTINCT b.id,b.source_file,b.created_at,e.version FROM central_claim_batches b JOIN central_claim_publications p ON p.batch_id=b.id LEFT JOIN central_claim_event_versions e ON e.batch_id=b.id ORDER BY b.created_at,b.id'))
        for batch in batches:
            if not batch['version']:missing.append(batch['source_file']);continue
            for raw in d.records(d.execute(conn,'SELECT n,franchise,source_json FROM central_claim_events WHERE batch_id=:id ORDER BY n',{'id':batch['id']})):
                source=json.loads(raw['source_json']);item={**classify(source),'n':raw['n'],'franchise':raw['franchise'],'source_file':batch['source_file'],'batch':batch['id']}
                # Deduplicate overlapping exports by the source claim identifier; newest snapshot wins.
                key=(normal(item['administrator']),normal(item['claim'])) if item['claim'] else (batch['id'],raw['n'])
                seen[key]=item
        # Financial rows follow the existing monthly publications, never the reported cohort.
        query='SELECT c.n,c.month,c.franchise,c.source_json,b.source_file,c.batch_id FROM central_claims c JOIN central_claim_publications p ON p.batch_id=c.batch_id AND p.month=c.month JOIN central_claim_batches b ON b.id=c.batch_id'
        for raw in d.records(d.execute(conn,query)):
            item={**classify(json.loads(raw['source_json'])),'n':raw['n'],'franchise':raw['franchise'],'source_file':raw['source_file'],'batch':raw['batch_id'],'month':raw['month']}
            paid.append(item)
    return list(seen.values()),paid,missing

def populations(store,dates,allowed=None):
    from published_cache import CLAIM_SOURCE,signature
    events,paid,missing=CLAIM_SOURCE.get(signature(store),lambda:_source_populations(store))
    start,end=dates['date_from'],dates['date_to']
    def within(r,field):return (allowed is None or r['franchise'] in allowed) and r[field] and start<=r[field]<=end
    return ([dict(r) for r in events if within(r,'reported_on')],[dict(r) for r in paid if within(r,'paid_on')],list(missing))

def grouped(records,key):
    groups={}
    for r in records:
        label=key(r);item=groups.setdefault(label,dict(label=label,count=0,bank=Decimal(0),netoff=Decimal(0),days_total=Decimal(0),days_count=0))
        item['count']+=1;item['bank']+=r['bank'];item['netoff']+=r['netoff']
        if r['days'] is not None:item['days_total']+=r['days'];item['days_count']+=1
    for row in groups.values():
        row['days']=row['days_total']/row['days_count'] if row['days_count'] else None
        row['years']=row['days']/Decimal('365.25') if row['days'] is not None else None
    return sorted(groups.values(),key=lambda r:(-r['count'],r['label'].casefold()))

def analyse(store,dates,allowed=None,sales_allowed=_SELECTED_SCOPE,sales_dates=None,branch_dates=None):
    store.ensure()
    if sales_allowed is _SELECTED_SCOPE:sales_allowed=allowed
    default_dates=sales_period(dates);branch_dates=branch_dates or default_dates;risk_dates=sales_dates or default_dates
    from published_cache import signature
    key=(signature(store),tuple(sorted(dates.items())),None if allowed is None else tuple(sorted(allowed)),None if sales_allowed is None else tuple(sorted(sales_allowed)),tuple(sorted(risk_dates.items())),tuple(sorted(branch_dates.items())))
    if key in _CACHE and time.monotonic()-_CACHE[key][0]<120:return _CACHE[key][1]
    events,paid,missing=populations(store,dates,allowed);bank=[r for r in paid if r['bank']!=0];offset=[r for r in paid if r['netoff']!=0]
    start,end=dates['date_from'][:7],dates['date_to'][:7]
    # Product cover is sourced from actual Direct MEM benefits, not claim payout amounts.
    from performance_index import products
    _,covers,_=products(store,start,end,allowed)
    for r in events:
        options=covers.get(normal(r['product']),set())
        r['cover']='R'+format(next(iter(options)),',.2f') if len(options)==1 else 'Unknown / conflicting Direct cover'
    risk_start,risk_end=risk_dates['date_from'][:7],risk_dates['date_to'][:7]
    _,risk_paid,_=populations(store,risk_dates,sales_allowed)
    _,branch_paid,_=populations(store,branch_dates,allowed)
    branch_bank=[r for r in branch_paid if r['bank']!=0]
    risk_bank=[r for r in risk_paid if r['bank']!=0]
    premium,_,names=products(store,risk_start,risk_end,sales_allowed)
    risk_pubs=[b for b in store.publications() if risk_start<=b['month']<=risk_end]
    totals=defaultdict(lambda:dict(amount=Decimal(0),count=0,branches=defaultdict(lambda:[0,Decimal(0)])))
    for r in risk_bank:
        k=normal(r['product']);names[k]=r['product'];totals[k]['amount']+=r['bank'];totals[k]['count']+=1
        totals[k]['branches'][r['franchise']][0]+=1;totals[k]['branches'][r['franchise']][1]+=r['bank']
    expected=[];y,m=map(int,risk_start.split('-'));last=tuple(map(int,risk_end.split('-')))
    while (y,m)<=last:
        expected.append(f'{y:04d}-{m:02d}');m+=1
        if m==13:y+=1;m=1
    verified=store.verified_claim_months();complete=all(m in verified and m in {b['month'] for b in risk_pubs} for m in expected)
    whole_months=risk_dates['date_from'].endswith('-01') and date.fromisoformat(risk_dates['date_to']).day==__import__('calendar').monthrange(*map(int,risk_end.split('-')))[1]
    risk=[];red=set()
    for k in totals:
        t=totals[k];den=premium[k];value=t['amount']/den*100 if den>0 and complete and whole_months else None
        colour='red' if value is not None and value>=70 else 'green' if value is not None and value>=50 else 'blue' if value is not None else 'unknown'
        action='Must not be sold' if colour=='red' else 'May be sold — green band' if colour=='green' else 'Lower claim ratio — blue band' if colour=='blue' else 'Needs complete claims and monthly premiums'
        if colour=='red':red.add(k)
        risk.append(dict(label=names.get(k,'Unallocated'),count=t['count'],bank=t['amount'],premium=den,remaining=den-t['amount'],franchises=len(t['branches']),ratio=value,colour=colour,action=action))
    risk.sort(key=lambda r:(r['ratio'] is None,-(r['ratio'] or 0),-r['bank'],r['label']))
    tables={k:grouped(events,lambda r,k=k:r[k]) for k in ('relation','age','gender','cause','status','product')}
    tables.update(product_age=grouped(events,lambda r:r['product']+' / '+r['age']),cover=grouped(events,lambda r:r['cover']),
      payee=grouped(bank,lambda r:r['payee']),netoff=grouped(offset,lambda r:r['franchise']),duration=grouped(events,lambda r:r['product']),
      branch_month=grouped(branch_bank,lambda r:r['franchise']+' / '+r['month']),risk=risk,
      risk_franchise=grouped([r for r in risk_bank if normal(r['product']) in red],lambda r:r['franchise']),
      early=grouped([r for r in events if r['days'] is not None and r['days']<=180],lambda r:r['product']),
      quality=grouped([r for r in events if 'Unknown' in (r['age'],r['cause'],r['payee']) or r['days'] is None or not r['claim']],lambda r:r['product']))
    delay=[]
    for r in bank:
        if r['reported_on']:
            days=(date.fromisoformat(r['paid_on'])-date.fromisoformat(r['reported_on'])).days
            if days>=0:delay.append({**r,'days':Decimal(days)})
    tables['delay']=grouped(delay,lambda r:r['franchise'])
    cards=[]
    for key,label in SPECS.items():
        rows=tables[key];count=sum(r['count'] for r in rows)
        value=len(red) if key=='risk' else sum(r['netoff'] for r in rows) if key=='netoff' else count
        if key=='risk' and (not complete or not whole_months):value=None
        unit='money' if key=='netoff' else 'count'
        if key=='duration':
            valid=sum(r['days_count'] for r in rows);value=sum((r['days_total'] for r in rows),Decimal(0))/valid if valid else None;unit='days'
        incomplete=bool(missing) and key in ('relation','age','product','product_age','cover','gender','cause','status','duration','early','quality')
        cards.append(dict(key=key,label=label,value=value,unit=unit,count=count,top=rows[:10],incomplete=incomplete))
    maximum=max((r['ratio'] for r in risk if r['ratio'] is not None and r['ratio']>0),default=Decimal(0))
    for r in risk:r['scale']=max(0,float(r['ratio']/maximum*100)) if maximum and r['ratio'] is not None else 0
    risk_groups=[]
    for colour,title in [('red','Red ≥70%: Must not be sold'),('green','Green 50–69.99%: May be sold'),('blue','Blue <50%: lower claim ratio')]:
        rows=[r for r in risk if r['colour']==colour]
        risk_groups.append(dict(colour=colour,title=title,count=len(rows),rows=rows[:7]))
    most_groups=[]
    for colour,title in [('green','Green 50–69.99%: May be sold'),('blue','Blue <50%: lower claim ratio')]:
        rows=sorted([r for r in risk if r['colour']==colour],key=lambda r:(-r['count'],-r['bank'],r['label'].casefold()))
        maximum=max((r['count'] for r in rows),default=0)
        most_groups.append(dict(colour=colour,title=title,count=len(rows),rows=[{**r,'scale':r['count']/maximum*100} for r in rows[:7]]))
    branch_summaries={}
    for r in branch_paid:
        item=branch_summaries.setdefault(r['franchise'],dict(franchise=r['franchise'],claims=Decimal(0),net_off=Decimal(0),count=0))
        item['claims']+=r['bank'];item['net_off']+=r['netoff'];item['count']+=int(r['bank']!=0)
    branch_totals=dict(claims=sum((r['bank'] for r in branch_paid),Decimal(0)),claim_net_off=sum((r['netoff'] for r in branch_paid),Decimal(0)),claim_count=len(branch_bank))
    from dateutil.relativedelta import relativedelta
    branch_months=[];month=date.fromisoformat(branch_dates['date_from']).replace(day=1);last=date.fromisoformat(branch_dates['date_to']).replace(day=1)
    while month<=last:branch_months.append(month.strftime('%Y-%m'));month+=relativedelta(months=1)
    result=dict(cards=cards,tables=tables,events=events,paid=paid,risk_paid=risk_paid,risk_dates=risk_dates,branch_dates=branch_dates,branch_paid=branch_paid,branch_totals=branch_totals,branch_summaries=sorted(branch_summaries.values(),key=lambda r:r['franchise'].casefold()),branch_complete=all(m in verified for m in branch_months),risk_groups=risk_groups,most_groups=most_groups,risk_unknown=sum(r['ratio'] is None for r in risk),missing=missing,red=red,complete=complete,whole_months=whole_months,dates=dates,
      reported_count=len(events),paid_count=len(bank),bank=sum((r['bank'] for r in bank),Decimal(0)),netoff=sum((r['netoff'] for r in paid),Decimal(0)),netoff_count=len(offset))
    if len(_CACHE)>=16:_CACHE.pop(next(iter(_CACHE)))
    _CACHE[key]=(time.monotonic(),result)
    return result

def register(app,store,access):
    def selection():
        user=getattr(g,'user',None) or {}
        if not user:return None,None,None,('Authentication required.',401)
        if user.get('role') in ('underwriter','voucher_supplier'):return None,None,None,('Claims analytics are outside your report scope.',403)
        from period_selection import selected_dates
        dates=selected_dates(request.args);allowed=access();franchise=request.args.get('franchise','All')
        if franchise!='All':allowed={franchise} if allowed is None or franchise in allowed else set()
        return dates,allowed,franchise,None
    @app.get('/claims/analytics/<metric>')
    def claim_insight_detail(metric):
        dates,allowed,franchise,error=selection()
        if error:return error
        if metric not in SPECS:return 'Unknown claims analysis.',404
        insights=analyse(store,dates,allowed,sales_allowed=access(),sales_dates=sales_selection(request.args,dates),branch_dates=branch_selection(request.args,dates));rows=insights['tables'][metric];group=request.args.get('group','');page=max(1,request.args.get('page',1,type=int))
        most_claims=metric=='risk' and request.args.get('product_scenario')=='count'
        if most_claims:rows=sorted([r for r in rows if r['colour'] in ('green','blue')],key=lambda r:(-r['count'],-r['bank'],r['label'].casefold()))
        if metric in ('risk','risk_franchise'):records=insights['risk_paid']
        elif metric=='branch_month':records=insights['branch_paid']
        elif metric in ('payee','netoff','branch_month','delay'):records=insights['paid']
        else:records=insights['events']
        if metric=='netoff':records=[r for r in records if r['netoff']!=0]
        elif metric in ('payee','branch_month','risk','risk_franchise','delay'):records=[r for r in records if r['bank']!=0]
        if most_claims:
            products={normal(r['label']) for r in rows};records=[r for r in records if normal(r['product']) in products]
        if metric=='delay':
            records=[{**r,'days':Decimal((date.fromisoformat(r['paid_on'])-date.fromisoformat(r['reported_on'])).days),'years':Decimal((date.fromisoformat(r['paid_on'])-date.fromisoformat(r['reported_on'])).days)/Decimal('365.25')} for r in records if r['reported_on'] and r['paid_on']>=r['reported_on']]
        if metric=='risk_franchise':records=[r for r in records if normal(r['product']) in insights['red']]
        if metric=='early':records=[r for r in records if r['days'] is not None and r['days']<=180]
        if metric=='quality':records=[r for r in records if 'Unknown' in (r['age'],r['cause'],r['payee']) or r['days'] is None or not r['claim']]
        def label(r):
            return r['product'] if metric in ('risk','duration','early','quality') else r['product']+' / '+r['age'] if metric=='product_age' else r.get('cover','Unknown') if metric=='cover' else r['franchise']+' / '+r['month'] if metric=='branch_month' else r['franchise'] if metric in ('risk_franchise','netoff','delay') else r.get(metric,'Unknown')
        if group:records=[r for r in records if (normal(label(r))==normal(group) if metric=='risk' else label(r)==group)]
        product=request.args.get('product','') if metric=='product_age' else ''
        if product:
            records=[r for r in records if normal(r['product'])==normal(product)]
            if metric=='product_age':rows=grouped(records,label)
        count=len(records)
        return render_template('claim_insight_detail.html',insights=insights,title=SPECS[metric],metric=metric,rows=rows,records=records[(page-1)*50:page*50],count=count,page=page,group=group,franchise=franchise,dates=dates)
    @app.get('/claims/branch-products')
    def branch_claim_products():
        dates,allowed,franchise,error=selection()
        if error:return error
        if allowed==set():return 'Franchise access required.',403
        period=branch_selection(request.args,dates)
        _,paid,_=populations(store,period,allowed)
        groups={}
        for r in paid:
            item=groups.setdefault(normal(r['product']),dict(label=r['product'],count=0,bank=Decimal(0),netoff=Decimal(0)))
            item['count']+=int(r['bank']!=0);item['bank']+=r['bank'];item['netoff']+=r['netoff']
        rows=sorted(groups.values(),key=lambda r:(-r['count'],-r['bank'],r['label'].casefold()))
        product=request.args.get('product','')
        records=[r for r in paid if not product or normal(r['product'])==normal(product)]
        page=max(1,request.args.get('page',1,type=int))
        totals=dict(count=sum(r['count'] for r in rows),bank=sum((r['bank'] for r in rows),Decimal(0)),netoff=sum((r['netoff'] for r in rows),Decimal(0)))
        return render_template('branch_claim_products.html',franchise=franchise,dates=dates,period=period,rows=rows,totals=totals,records=records[(page-1)*50:page*50],count=len(records),page=page,product=product)
    @app.post('/admin/claims/analyse-saved')
    def claim_insight_backfill():
        if (getattr(g,'user',None) or {}).get('role')!='admin':return 'Administrator access required.',403
        identifier=app.extensions['central_jobs'].create(g.user['id'],'claim_analysis',{'filename':'Archived claims analytics'})
        return render_template('claim_analysis_wait.html',identifier=identifier),202
