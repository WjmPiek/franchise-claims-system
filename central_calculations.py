"""Versioned transaction rules; reports may aggregate these values but never reapply fees."""
import json
import re
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path

VERSION = 'claims-word-20261013-v1'
ZERO = Decimal(0)
def cents(value):
    return Decimal(str(value)).quantize(Decimal('.01'), rounding=ROUND_HALF_UP)

def effective_mpia(value):
    text = str(value or '').strip()
    if text.upper() in ('', 'NULL', 'NONE'):
        return Decimal(1)
    try:
        number = Decimal(text.replace(',', '.'))
        if not number.is_finite(): raise InvalidOperation
    except InvalidOperation:
        raise ValueError('MPIA must be a finite number.') from None
    return number if number > 0 else Decimal(1)

def fees(row):
    from direct_data import branch_key, norm, VOUCHER_PRODUCT
    risk = Decimal(str(row[12]))
    mpia = effective_mpia(row[17])
    if norm(row[11]) == norm(VOUCHER_PRODUCT):
        supplier = cents(Decimal('6.50') * mpia)
        return ZERO, ZERO, ZERO, ZERO, supplier, cents(risk-supplier)
    mem = str(row[8]).upper() == 'MEM'
    faithnet = ZERO
    if mem and branch_key(row[0]) in ('FAITHNET', 'FAITHNETSA'):
        faithnet = Decimal(40 if re.search(r'\b1\s*\+',str(row[11])) else 10) * mpia
    basis = cents(risk-faithnet)
    r1 = mpia if mem else ZERO
    after = max(ZERO,basis-r1)
    adv = cents(after*Decimal('.021')) if mem else ZERO
    return faithnet,r1,adv,after-adv,ZERO,ZERO

def normal(value):
    return re.sub(r'[^A-Z0-9]','',str(value or '').upper())

def mappings():
    return json.loads(Path(__file__).with_name('approved_products.json').read_text(encoding='utf-8'))

def snapshot(row, prepared, source=None):
    """Complete stored financial components from already calculated fee values."""
    from direct_data import VOUCHER_PRODUCT,norm
    voucher = norm(row[11])==norm(VOUCHER_PRODUCT)
    mem = str(row[8]).upper()=='MEM' and not voucher
    mpia = effective_mpia(row[17])
    retail = Decimal(str(prepared['retail'])) if not voucher else ZERO
    risk = Decimal(str(prepared['original_risk']))
    components={key:Decimal(str(prepared[key])) for key in ('faithnet','r1','adv','payover','supplier','retained')}
    commissions = {key:retail*rate for key,rate in (
        ('brightrock',Decimal('.10')),('mkhulu',Decimal('.025')),('mff',Decimal('.025')))}
    percentage=sum(commissions.values(),ZERO)
    paid=percentage+components['r1']+components['adv']
    mff_book=franchise_book=retail*Decimal('.025')
    result={
        'version':VERSION,'source_mpia':(source or row)[17],
        'calculation_mpia':str(mpia),'policy_qty':str(mpia if mem else ZERO),
        'single_monthly_premium':str(risk/mpia),'retail':str(retail),
        'source_retail':str(prepared['retail']),'original_risk':str(risk),
        'insurance_original_risk':str(risk if not voucher else ZERO),
        'voucher_gross':str(risk if voucher else ZERO),
        'voucher_retail':str(prepared['retail'] if voucher else ZERO),
        'reconciliation_adjustment':str(risk-sum(components.values(),ZERO)),
        'risk_after_r1':str(max(ZERO,cents(risk-components['faithnet'])-components['r1']) if not voucher else ZERO),
        **{key:str(value) for key,value in components.items()},
        **{key:str(value) for key,value in commissions.items()},
        'percentage_commission':str(percentage),'paid_commissions':str(paid),
        'admin':str(retail-components['payover']-paid),
        'mff_book_value':str(mff_book),'franchise_book_value':str(franchise_book),
        'book_value':str(mff_book+franchise_book),'remaining_book_value':str(retail-mff_book-franchise_book),
        'retail_less_payover':str(retail-components['payover']),
        'voucher_units':str(mpia if voucher else ZERO),
        'source_row':prepared['n'], 'source':source or row,
    }
    return result

def ratio(claims, risk):
    claims,risk=Decimal(str(claims)),Decimal(str(risk))
    if risk==0:
        return {'value':None,'status':'Undefined — investigate claims without net Risk' if claims else 'Not applicable — no premium basis','colour':'grey'}
    value=claims/risk*100
    return {'value':value,'status':'Calculated','colour':'blue' if value<50 else 'green' if value<70 else 'red'}

def retail_rank(current,previous):
    """Rank the complete population; alphabetical franchise name breaks equal Retail totals."""
    def order(items):return {name:n+1 for n,(name,_) in enumerate(sorted(items.items(),key=lambda x:(-Decimal(str(x[1])),x[0].casefold(),x[0])))}
    now,before=order(current),order(previous)
    return [{'position':position,'franchise':name,'change':before[name]-position if name in before else None}
            for name,position in sorted(now.items(),key=lambda x:x[1])]
