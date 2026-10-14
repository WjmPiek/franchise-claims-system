"""Lightweight reported-claim graphs: no policy-row scans or financial changes."""
from collections import Counter,defaultdict
from central_calculations import normal
from claim_insights import populations
from math import ceil
import hashlib,re

SPECS=[('claim_relationship_products','Top 5 claimed products by relationship','claims','rows'),
       ('claim_product_age','Product claims by age category','claims','rows'),
       ('claim_gender','Claims by policyholder gender','claims','rows'),
       ('claim_cause','Natural and unnatural incidents','claims','rows')]
COLOURS=['#2563eb','#d97706','#059669','#9333ea','#dc2626','#0891b2','#be185d','#4f46e5','#65a30d','#a16207','#0f766e','#7c3aed']
RELATIONS=['Main Life','Extended Family','Spouse','Children']

def months(dates):
    first_y,first_m=map(int,dates['date_from'][:7].split('-'));last_y,last_m=map(int,dates['date_to'][:7].split('-'))
    return [f'{i//12:04d}-{i%12+1:02d}' for i in range(first_y*12+first_m-1,last_y*12+last_m)]

def geometry(series,labels):
    high=ceil(max([1]+[p['value'] for s in series for p in s['points']])/4)*4;y=lambda v:20+(high-v)/high*170
    for s in series:
        points=[{**p,'x':300 if len(labels)==1 else 70+i*500/(len(labels)-1),'y':y(p['value'])} for i,p in enumerate(s['points'])]
        s['chart_points']=points;s['path']=' '.join(('M' if i==0 else 'L')+f"{p['x']:.2f} {p['y']:.2f}" for i,p in enumerate(points))
    return dict(ticks=[dict(y=y(high*i/4),value=high*i/4) for i in range(5)],
        labels=[dict(x=300 if len(labels)==1 else 70+i*500/(len(labels)-1),label=label,index=i) for i,label in enumerate(labels)],label_step=max(1,(len(labels)+5)//6))

def cause(value):
    n=normal(value)
    if n.startswith('UNNATURAL'):return 'Unnatural'
    if n.startswith('NATURAL'):return 'Natural'
    return 'Other / Unknown'

def age_colour(label):
    if normal(label) in ('UNKNOWN','OTHERUNKNOWN'):return '#64748b'
    band=re.sub(r'[^0-9+]','',label)
    fixed={'065':'#2563eb','6675':'#d97706','7685':'#059669','86+':'#9333ea'}
    if band in fixed:return fixed[band]
    hue=int(hashlib.sha256(normal(label).encode()).hexdigest()[:8],16)%36000/100
    return f'hsl({hue},70%,35%)'

def graphs(store,dates,allowed,keys,product=''):
    events,_,missing=populations(store,dates,allowed)
    products={}
    for r in events:products.setdefault(normal(r['product']),r['product'])
    options=sorted(products.values(),key=str.casefold);result=[]
    for key,title,category,unit in SPECS:
        if key not in keys:continue
        records=events;counts=defaultdict(Counter);labels=months(dates);axis='Reported month';detail='';rows=[];selected_product=''
        if key=='claim_relationship_products':
            totals=Counter(normal(r['product']) for r in records)
            top=sorted(totals,key=lambda p:(-totals[p],products[p].casefold()))[:5]
            records=[r for r in records if normal(r['product']) in top]
            labels=['P'+str(i+1) for i in range(len(top))];position={p:labels[i] for i,p in enumerate(top)}
            series_labels=RELATIONS.copy();axis='Top 5 products, ranked by reported claim count';detail='relation'
            for r in records:
                label=r['relation'] if r['relation'] in RELATIONS else 'Other / Unknown'
                counts[label][position[normal(r['product'])]]+=1
            if counts['Other / Unknown']:series_labels.append('Other / Unknown')
            rows=[dict(label=position[p]+' · '+products[p],values=[counts[s][position[p]] for s in series_labels],total=totals[p]) for p in top]
        else:
            if key=='claim_product_age':
                selected_product=product;records=[r for r in records if not product or normal(r['product'])==normal(product)]
                series_labels=sorted({r['age'] for r in records},key=lambda a:('Unknown' in a,a.casefold()));detail='product_age'
                bucket=lambda r:r['age']
            elif key=='claim_gender':
                series_labels=['Male','Female','Other'];detail='gender';bucket=lambda r:r['gender']
                if any(r['gender']=='Other / Unknown' for r in records):series_labels.append('Other / Unknown')
            else:
                series_labels=['Natural','Unnatural'];detail='cause';bucket=lambda r:cause(r['cause'])
                if any(cause(r['cause'])=='Other / Unknown' for r in records):series_labels.append('Other / Unknown')
            for r in records:counts[bucket(r)][r['reported_on'][:7]]+=1
            rows=[dict(label=m,values=[counts[s][m] for s in series_labels],total=sum(counts[s][m] for s in series_labels)) for m in labels]
        colours={'Male':'#2563eb','Female':'#d97706','Other':'#9333ea','Natural':'#059669','Unnatural':'#dc2626','Other / Unknown':'#64748b'}
        series=[dict(label=s,colour=age_colour(s) if key=='claim_product_age' else colours.get(s,COLOURS[i%len(COLOURS)]),dash='' if i<len(COLOURS) else '6 3',
            points=[dict(label=label,value=counts[s][label]) for label in labels]) for i,s in enumerate(series_labels)]
        product_age_counts=Counter((r['product'],r['age']) for r in records) if key=='claim_product_age' else {}
        result.append(dict(key=key,label=title,category=category,unit=unit,points=[],series=series,
            chart=geometry(series,labels) if records else geometry([],[]),axis=axis,detail_metric=detail,rows=rows,
            product_options=options if key=='claim_product_age' else [],product_filter=selected_product,
            product_age_rows=[dict(product=p,age=a,count=n) for (p,a),n in sorted(product_age_counts.items())],
            count=len(records),population_count=len(events),missing=missing,note='Reported claims · Claim Reported Date · includes approved, repudiated and other recorded statuses.'))
    return result
