"""Paid main-member product mix; source premiums and report calculations stay unchanged."""
import json
from decimal import Decimal

def product_summary(records):
    groups={};seen=set()
    for row in records:
        data=json.loads(row['data']) if isinstance(row['data'],str) else row['data']
        product=str(data[11] or '(Unspecified product)')
        from direct_data import norm,branch_key
        key=(row['franchise'],norm(product));policy=(branch_key(row['franchise']),row['policy'],norm(product))
        item=groups.setdefault(key,{'franchise':row['franchise'],'product':product,'policies':0,'paid_months':Decimal(0),'retail':Decimal(0),'risk':Decimal(0),'payover':Decimal(0)})
        if policy not in seen:item['policies']+=1;seen.add(policy)
        item['paid_months']+=Decimal(str(data[17]))
        for field,source in [('retail','retail'),('risk','original_risk'),('payover','payover')]:item[field]+=Decimal(str(row[source]))
    return list(groups.values())

def insights(rows,month,product=''):
    converted=[{**row,**{key:Decimal(str(row[key])) for key in ('paid_months','retail','risk','payover')}} for row in rows]
    products=sorted(set(row['product'] for row in converted))
    chosen=[row for row in converted if not product or row['product']==product]
    franchises={};mix={}
    for row in chosen:
        item=franchises.setdefault(row['franchise'],{'label':row['franchise'],'value':0,'retail':Decimal(0),'products':set()})
        item['value']+=row['policies'];item['retail']+=row['retail'];item['products'].add(row['product'])
        mix[row['product']]=mix.get(row['product'],0)+row['policies']
    top=sorted(franchises.values(),key=lambda row:(-row['value'],row['label']))[:10]
    for row in top:row['products']=len(row['products'])
    return {'month':month,'products':products,'product':product,'top':top,'mix':[{'label':key,'value':value} for key,value in sorted(mix.items(),key=lambda pair:(-pair[1],pair[0]))[:10]],'rows':sorted(chosen,key=lambda row:(row['franchise'],row['product']))}
