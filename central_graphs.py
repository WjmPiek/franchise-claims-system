from collections import defaultdict
from decimal import Decimal
from flask import url_for

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

def graphs(store,start,end,allowed=None):
    months=defaultdict(dict)
    for row in store.summaries(start,end,allowed):
        for key,value in row['summary'].items():months[row['month']][key]=months[row['month']].get(key,Decimal(0))+Decimal(str(value))
    claims=defaultdict(lambda:Decimal(0))
    for row in store.claims(start,end,allowed):claims[row['month']]+=Decimal(row['amount'])
    specs=[('original_risk','Original Risk','source','money'),('source_retail','Retail','source','money'),('payover','Insurer Payover','insurers','money'),('aul_payover','AUL Payover','aul','money'),('aul_plan2_rows','AUL Plan 2 rows','plan2','rows'),('aul_historic_rows','AUL Historic rows','historic','rows'),('r1','R1 Fees','fees','money'),('adv','2.1% Fees','fees','money'),('faithnet','Faithnet Deductions','fees','money'),('supplier','Voucher Supplier Payable','vouchers','money'),('retained','Voucher Retained','vouchers','money'),('policy_qty','Paid policy months / MPIA','fees','rows')]
    result=[]
    for key,label,category,unit in specs:
        points=[dict(label=month,value=float(v.get(key,0)),url=url_for('central_calculation_detail',category=category,date_from=month+'-01',date_to=month+'-28')) for month,v in sorted(months.items())]
        result.append(dict(key=key,label=label,category=category,unit=unit,points=points))
    verified=store.verified_claim_months()
    points=[dict(label=month,value=float(claims[month]),url=url_for('central_calculation_detail',category='claims',date_from=month+'-01',date_to=month+'-28')) for month in sorted(months) if month in verified]
    result.append(dict(key='verified_claims',label='Claim Amount Comparison Over Time',category='claims',unit='money',points=points))
    for graph in result:graph['chart']=chart_geometry(graph['points'])
    return result
