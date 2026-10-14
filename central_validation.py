"""Classify against supplied product examples; keep unresolved rows explicit."""
import json,re
from collections import defaultdict
from central_calculations import mappings,normal,effective_mpia

def member_plan(product,row,products):
    explicit=products.get(normal(product),'Unclassified')
    if explicit!='Unclassified':return explicit
    from direct_data import number
    try:cover=number(row[10],'Main Cover',0)
    except ValueError:return 'Unclassified'
    if cover>0:return 'Plan 2' if cover in (40000,50000) else 'Historic'
    return 'Unclassified'

def plan_assignments(conn):
    registry=mappings();products=registry['plans'];known={normal(n) for n in registry['franchises']}
    has_fingerprints=conn.execute("SELECT COUNT(*) FROM sqlite_master WHERE name='source_fingerprints'").fetchone()[0]
    duplicates={r[0] for r in conn.execute('SELECT n FROM source_fingerprints WHERE digest IN (SELECT digest FROM source_fingerprints GROUP BY digest HAVING COUNT(*)>1)')} if has_fingerprints else set()
    members=defaultdict(set);policies=defaultdict(set)
    for branch,policy,product,data in conn.execute("SELECT branch,policy,product,values_json FROM rows WHERE destination='AUL' AND voucher=0 AND relation='MEM'"):
        row=json.loads(data);plan=member_plan(product,row,products)
        members[(branch,policy,row[16])].add(plan);policies[(branch,policy)].add(plan)
    assignments={};issues={}
    has_extra=conn.execute("SELECT COUNT(*) FROM sqlite_master WHERE name='source_extra'").fetchone()[0]
    for n,branch,policy,rel,product,voucher,destination,data in conn.execute('SELECT n,branch,policy,relation,product,voucher,destination,values_json FROM rows ORDER BY n'):
        row=json.loads(data);messages=[]
        if has_extra:
            extras=conn.execute('SELECT data FROM source_extra WHERE n=?',(n,)).fetchone()
            aul=json.loads(extras[0]).get('AULRISK') if extras else None
            if aul not in (None,''):
                from direct_data import number
                from decimal import Decimal
                if number(aul,'AUL Risk',n)!=Decimal(row[12]):messages.append('Risk / AUL Risk conflict: calculation uses Risk; review the AUL Risk source value.')
        if n in duplicates:messages.append('Identical source fields appear on separate source rows; all rows were retained and calculated. Verify their source provenance.')
        if not policy:
            raise ValueError(f'Source row {n}: missing policy number. Previous published month has been preserved.')
        if row[5] and not re.fullmatch(r'\d{13}',row[5]) and not row[7]:
            messages.append('ID number does not contain 13 digits; verify ID/passport before sending personal data.')
        if not product:messages.append('[BLOCK] Missing product; classify before final insurer export.')
        raw=conn.execute('SELECT data FROM source_raw WHERE n=?',(n,)).fetchone()
        if raw:
            original=json.loads(raw[0])
            if str(original[17]).strip().upper() in ('','NULL') or str(original[17])!=str(effective_mpia(original[17])):
                try:
                    from decimal import Decimal
                    altered=not str(original[17]).strip() or Decimal(str(original[17]))<=0
                except Exception:altered=True
                if altered:messages.append('Source MPIA was blank/nonpositive; calculation MPIA normalised to 1. Original value retained.')
        if destination=='AUL' and not voucher:
            approved_covers=registry.get('legacy_main_covers',{}).get(normal(product))
            if rel=='MEM' and approved_covers:
                from decimal import Decimal
                if Decimal(row[10]) not in {Decimal(v) for v in approved_covers}:messages.append('[BLOCK] Legacy product main cover differs from the supplied October example; confirm its approved plan mapping before export.')
            if normal(row[0]) not in known:messages.append('Franchise is not in the supplied September AUL examples; confirm its AUL allocation.')
            group=members.get((branch,policy,row[16]),set()) if row[16] else set()
            if not group:group=policies.get((branch,policy),set())
            explicit=products.get(normal(product),'Unclassified')
            if not group and explicit!='Unclassified':
                plan=explicit
                messages.append('No MEM row for this policy in the source; approved product mapping used. Verify the missing main-member source record. All related source rows retained.')
            else:plan=next(iter(group)) if len(group)==1 else explicit if explicit in group else 'Unclassified'
            if rel=='MEM':plan=member_plan(product,row,products)
            if plan=='Unclassified':messages.append('[BLOCK] AUL product or related policy plan is unclassified; record retained in AUL Unclassified.')
            assignments[n]=plan
        if messages:issues[n]=messages
    return assignments,issues
