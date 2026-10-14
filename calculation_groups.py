"""Exact, scoped franchise totals without loading source JSON into Python."""
from decimal import Decimal
from published_cache import signature
from report_cache import ResultCache
GROUP_CACHE=ResultCache(ttl=3600,max_entries=32)

FIELDS = ('original_risk','retail','faithnet','r1','adv','payover','paid_commissions','supplier','retained')

def money_expression(postgres, field):
    # Extract only generated decimal scalars from TEXT. Whole-document PostgreSQL
    # JSON casts reject escaped NULs in unrelated original customer fields.
    if postgres:
        pattern = '"' + field + '"[[:space:]]*:[[:space:]]*"([-+0-9.eE]+)"'
        return "SUM(CAST(substring(r.record_json FROM '" + pattern + "') AS NUMERIC))"
    path = '$.calculation.paid_commissions' if field == 'paid_commissions' else '$.' + field
    return "decimal_sum(json_extract(r.record_json,'" + path + "'))"

def policy_groups(store, base, params):
    key=(signature(store),base,tuple(sorted(params.items())))
    return [dict(row) for row in GROUP_CACHE.get(key,lambda: _policy_groups(store,base,params))]

def _policy_groups(store, base, params):
    columns = ['r.franchise','COUNT(*) AS rows']
    columns += ['COALESCE(' + money_expression(bool(store.direct.get_engine), field) + ',0) AS ' + field for field in FIELDS]
    with store.direct.connection() as conn:
        return store.direct.records(store.direct.execute(conn, 'SELECT ' + ','.join(columns) + base + ' GROUP BY r.franchise ORDER BY LOWER(r.franchise),r.franchise', params))

def claim_groups(store, base, params):
    def total(field):
        return 'SUM(CAST(' + field + ' AS NUMERIC))' if store.direct.get_engine else 'decimal_sum(' + field + ')'
    joined = base.replace(' WHERE ', ' LEFT JOIN central_claim_components x ON x.batch_id=c.batch_id AND x.n=c.n WHERE ')
    columns = ['c.franchise','COUNT(*) AS rows', 'SUM(CASE WHEN CAST(x.bank_paid AS NUMERIC)<>0 THEN 1 ELSE 0 END) AS claim_count']
    columns += ['COALESCE(' + total(field) + ',0) AS ' + label for field,label in [('c.amount','amount'),('x.bank_paid','bank_paid'),('x.net_off','net_off')]]
    with store.direct.connection() as conn:
        return store.direct.records(store.direct.execute(conn, 'SELECT ' + ','.join(columns) + joined + ' GROUP BY c.franchise ORDER BY LOWER(c.franchise),c.franchise', params))

def summary_groups(store,start,end,allowed,category):
    groups={}
    for row in store.summaries(start,end,allowed):
        summary=row['summary'];insurance=category!='source'
        item=groups.setdefault(row['franchise'],dict(franchise=row['franchise'],rows=0,**{f:Decimal(0) for f in FIELDS}))
        item['rows']+=int(summary['rows'])-(int(summary['voucher_rows']) if insurance else 0)
        for field in FIELDS:
            source='source_retail' if field=='retail' else 'insurance_original_risk' if insurance and field=='original_risk' else field
            value=Decimal(summary[source])
            if insurance and field=='retail':value-=Decimal(summary['voucher_retail'])
            if insurance and field in ('supplier','retained'):value=Decimal(0)
            item[field]+=value
    return [r for r in sorted(groups.values(),key=lambda r:r['franchise'].casefold()) if r['rows']]
