"""Business overview display totals from published records; no recalculation."""
from decimal import Decimal

def branch_totals(summaries, claims):
    branches={}
    for row in summaries:
        item=branches.setdefault(row['franchise'],dict(franchise=row['franchise'],retail=Decimal(0),risk=Decimal(0),payover=Decimal(0),claims=Decimal(0),rows=0))
        s=row['summary']
        for key,source in [('retail','source_retail'),('risk','original_risk'),('payover','payover')]:
            item[key]+=Decimal(str(s.get(source,0)))
        item['rows']+=int(s.get('rows',0))
    for row in claims:
        if row['franchise'] in branches:branches[row['franchise']]['claims']+=Decimal(str(row['claims']))
    return sorted(branches.values(),key=lambda r:(-r['retail'],r['franchise']))
