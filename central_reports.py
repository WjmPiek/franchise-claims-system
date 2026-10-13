"""Individual report schemas over stored values, with no fee recalculation."""
import calendar,json,re,hashlib
from collections import defaultdict
from decimal import Decimal
from datetime import datetime
from pathlib import Path
import xlsxwriter
from central_calculations import ratio,normal

PLAN2_COLUMNS=[0,1,2,3,4,5,6,7,8,9,10,11,12,13,17,18,19,20,21]
COMBINE={'FAITNETSA':"Martin’s Funerals Brokers",'FAITHNETSA':"Martin’s Funerals Brokers",'MARTINSFUNERALSBROKERS':"Martin’s Funerals Brokers",
 'WESTONAREA':"Martin’s Randfontein",'WESTONARIA':"Martin’s Randfontein",'MARTINSFUNERALSWESTONARIA':"Martin’s Randfontein",'MARTINSRAND FONTEIN'.replace(' ',''):"Martin’s Randfontein",'MARTINSFUNERALSRANDFONTEIN':"Martin’s Randfontein"}

def safe(value):return re.sub(r'[^\w.-]+','_',str(value),flags=re.UNICODE).strip('_')[:130] or 'All'
def period(month):
    y,m=map(int,month.split('-'));return f'{y}{m:02d}01',f'{y}{m:02d}{calendar.monthrange(y,m)[1]:02d}'

def catalog_all(store,batch,allowed=None):
    start,end=period(batch['month']);stamp=batch['month'].replace('-','')
    sums=[r for r in store.summaries(batch['month'],batch['month'],allowed) if r['batch_id']==batch['id']]
    franchises=[r['franchise'] for r in sums];totals={r['franchise']:r['summary'] for r in sums}
    claim_total=Decimal(0);claim_rows=0
    for row in store.claims(batch['month'],batch['month'],allowed):claim_total+=Decimal(row['amount']);claim_rows+=1
    files=[]
    def add(kind,name,title,franchise='',destination='',plan='',metric='payover',internal=False):
        from direct_data import allocation
        subset=[totals[f] for f in franchises if (not franchise or f==franchise) and (not destination or allocation(f)==destination)]
        count_key='rows';amount_key=metric
        if destination=='AUL':
            count_key='aul_plan2_rows' if plan=='Plan 2' else 'aul_historic_rows' if plan=='Historic' else 'aul_rows'
            amount_key='aul_plan2_payover' if plan=='Plan 2' else 'aul_historic_payover' if plan=='Historic' else 'aul_payover'
        if kind in ('supplier','vouchers_internal'):count_key='voucher_rows'
        if kind=='review':count_key='review_rows'
        rows=sum(int(x.get(count_key,0)) for x in subset)
        if destination and destination!='AUL':rows-=sum(int(x.get('voucher_rows',0)) for x in subset)
        if plan=='Unclassified':rows=sum(int(x.get('aul_rows',0))-int(x.get('aul_plan2_rows',0))-int(x.get('aul_historic_rows',0)) for x in subset)
        amount=sum((Decimal(x.get(amount_key,0)) for x in subset),Decimal(0))
        if plan=='Unclassified':amount=sum((Decimal(x.get('aul_payover',0))-Decimal(x.get('aul_plan2_payover',0))-Decimal(x.get('aul_historic_payover',0)) for x in subset),Decimal(0))
        if kind=='claims':rows=claim_rows;amount=claim_total;metric='Paid claims (bank)'
        files.append({'kind':kind,'name':name,'title':title,'franchise':franchise,'destination':destination,'plan':plan,
          'batch_id':batch['id'],'month':batch['month'],'version':batch['version'],'status':'Generate on request',
          'source_rows':rows,'total':str(amount),
          'total_label':metric.replace('_',' ').title(),'internal':internal})
    for franchise in franchises:
        label=safe(franchise)
        add('payover',f'Monthly_Payover_{label}_{stamp}.xlsx','Monthly Pay-over',franchise,internal=True)
        add('product_claims',f'Eiseverhouding_Per_Produk_{label}_{stamp}.xlsx','Claims ratio per product',franchise,metric='payover',internal=True)
        if totals[franchise]['voucher_rows']:
            add('vouchers_internal',f'Monthly_Vouchers_{label}_{stamp}.xlsx','Monthly vouchers — internal',franchise,metric='supplier',internal=True)
    from direct_data import EXCEPTIONS,GROUP_NAME
    for destination in EXCEPTIONS+[GROUP_NAME]:
        name='PretoriaNorth_Soshanguve_Hammanskraal' if destination==GROUP_NAME else destination
        add('policy',f'Policy Data_{start}_to_{end}_{name}.xlsx',destination+' policy data',destination=destination)
    add('policy',f'PolicyData_{start}_to_{end}_AUL.xlsx','AUL Standard',destination='AUL')
    add('complete',f'AUL_Complete_Data_{stamp}.xlsx','AUL Complete Data',destination='AUL',internal=True)
    add('combine',f'PolicyData_{start}_to_{end}_AUL_Complete_Data_Combine.xlsx','AUL Complete Data Combine',destination='AUL',internal=True)
    add('policy',f'PolicyData_{start}_to_{end}_Martins Funeral Plan 2.xlsx','Martins Funeral Plan 2',destination='AUL',plan='Plan 2')
    add('policy',f'PolicyData_{start}_to_{end}_Martins Funeral Plan (Historic).xlsx','Martins Funeral Plan Historic',destination='AUL',plan='Historic')
    add('unclassified',f'AUL_Unclassified_{stamp}.xlsx','AUL unclassified review',destination='AUL',plan='Unclassified',internal=True)
    add('supplier',f'Grocery_Vouchers_Supplier_{stamp}.xlsx','Voucher supplier',metric='supplier')
    add('claims',f'Claims_Summary_{stamp}.xlsx','Claims summary',internal=True)
    add('commissions',f'Commission_Summary_{stamp}.xlsx','Commission summary',metric='paid_commissions',internal=True)
    add('internal',f'Internal_Commissions_Claims_Scenarios_Book_Value_{stamp}.xlsx','Internal financial report',internal=True)
    add('review',f'Import_Review_{stamp}.xlsx','Source and calculation review',internal=True)
    add('reconciliation',f'Reconciliation_{stamp}.xlsx','Monthly reconciliation',metric='original_risk',internal=True)
    add('source',f'Converted_Source_{stamp}.xlsx','Original source values',metric='original_risk',internal=True)
    # Report-specific counts/totals are replaced with actual generated-file metadata.
    return files

APPROVED_BRANCHES=('Katlehong','Alberton','Pretoria North Group','Carletonville','Rustenburg')
REPORT_FORMAT_VERSION='approved-two-sources-20261013-v1'

def catalog(store,batch,allowed=None):
    """Only the user's named report families are offered to users."""
    files=catalog_all(store,batch,allowed);start,end=period(batch['month'])
    result=[]
    def combined(kind,name,title):
        items=[i for i in files if i['kind']==kind]
        if not items:return
        result.append({**items[0],'name':name,'title':title,'franchise':'','source_rows':sum(i['source_rows'] for i in items),'total':str(sum((Decimal(i['total']) for i in items),Decimal(0)))})
    combined('payover',f'Monthly Pay-over per Franchise (New) - {batch["month"]}.xlsx','Monthly Pay-over per Franchise')
    result.extend(i for i in files if i['kind']=='product_claims')
    for item in files:
        if item['kind']!='policy' or item['destination'] not in (*APPROVED_BRANCHES,'AUL'):continue
        if item['plan']=='Plan 2':item={**item,'name':f'PolicyData_{start}_to_{end}_Martins Funeral Plan 2 (New).xlsx'}
        result.append(item)
    voucher=next((i for i in files if i['kind']=='supplier'),None)
    if voucher:result.append({**voucher,'name':f'{calendar.month_name[int(batch["month"][5:])]} {batch["month"][:4]} Vouchers.xlsx','title':'Monthly vouchers'})
    from direct_data import allocation
    sums=store.summaries(batch['month'],batch['month'],allowed)
    for destination in APPROVED_BRANCHES:
        subset=[r['summary'] for r in sums if allocation(r['franchise'])==destination]
        label='PretoriaNorth_Soshanguve_Hammanskraal' if destination=='Pretoria North Group' else destination
        result.append(dict(kind='branch_calculation',name=f'{label} - Calculation - {batch["month"]}.xlsx',title=label+' calculation',franchise='',destination=destination,plan='',batch_id=batch['id'],month=batch['month'],version=batch['version'],status='Generate on request',source_rows=sum(r['rows'] for r in subset),total=str(sum((Decimal(r['payover']) for r in subset),Decimal(0))),total_label='Net insurer Risk',internal=True))
    return result

def financial_rows(store,batch,allowed=None,franchise=''):
    summaries=store.summaries(batch['month'],batch['month'],allowed)
    claims=defaultdict(lambda:[Decimal(0),0])
    historical_claims=defaultdict(lambda:Decimal(0))
    for row in store.claims(allowed=allowed):
        if row['month']>batch['month']:continue
        historical_claims[(row['franchise'],row['month'])]+=Decimal(row['amount'])
        if row['month']==batch['month']:
            claims[row['franchise']][0]+=Decimal(row['amount']);claims[row['franchise']][1]+=int(Decimal(row['amount'])!=0)
    history=defaultdict(list)
    for item in store.summaries(allowed=allowed):
        if item['month']<=batch['month']:history[item['franchise']].append(item)
    verified=store.verified_claim_months()
    from central_store import SUM_FIELDS
    for name in set(claims)-{r['franchise'] for r in summaries}:
        summaries.append(dict(franchise=name,summary={**{key:'0' for key in SUM_FIELDS},'rows':0,'mem_rows':0,'review_rows':1}))
    for item in summaries:
        if franchise and item['franchise']!=franchise:continue
        values=item['summary'];amount,count=claims[item['franchise']];r=ratio(amount,values['payover'])
        previous=history[item['franchise']];missing=[h['month'] for h in previous if h['month'] not in verified]
        running='Needs verified claims: '+', '.join(missing) if missing else sum((Decimal(h['summary']['retail'])-historical_claims[(item['franchise'],h['month'])]-Decimal(h['summary']['paid_commissions']) for h in previous),Decimal(0))
        result={'Franchise':item['franchise'],'Month':batch['month'],'Source Version':batch['version'],
          'Source Rows':values['rows'],'MEM Source Rows':values['mem_rows'],'Paid Policy Months':values['policy_qty'],
          'Retail':values['retail'],'Original Risk':values['insurance_original_risk'],'Faithnet Deduction':values['faithnet'],
          'Net Risk / Insurer Payover':values['payover'],'R1':values['r1'],'ADV':values['adv'],
          'BrightRock':values['brightrock'],'Mkhulu':values['mkhulu'],'MFF':values['mff'],
          'Percentage Commissions':values['percentage_commission'],'Paid Commissions':values['paid_commissions'],
          'Admin / Scenario 1':values['admin'],'MFF Book Value':values['mff_book_value'],
          'Franchise Book Value':values['franchise_book_value'],'Total Book Value':values['book_value'],
          'Remaining Book Value':values['remaining_book_value'],'Claims':str(amount),'Claim Count':count,
          'Average Claim':str(amount/count) if count else 'Not applicable',
          'Claim Ratio':str(r['value']/100) if r['value'] is not None else r['status'],
          'Scenario 2 / BrightRock Month Value':str(Decimal(values['retail'])-amount-Decimal(values['paid_commissions'])),
          'Insurer Remaining After Commissions And Claims':str(Decimal(values['payover'])-Decimal(values['percentage_commission'])-amount) if batch['month'] in verified else 'Unconfirmed paid claims',
          'BrightRock Running Balance':running,'Running Balance History':'Available published receipt months through '+batch['month']+'; no external opening balance included',
          'Voucher Supplier':values['supplier'],'Voucher Retained':values['retained'],'Review Rows':values['review_rows']}
        for key,value in result.items():
            if key not in ('Franchise','Month','Source Version') and isinstance(value,str) and re.fullmatch(r'-?\d+(?:\.\d+)?',value):result[key]=Decimal(value)
        yield result

def generate(store,batch,item,path,allowed=None):
    from direct_data import HEADERS,VOUCHER_HEADERS,SheetWriter
    kind=item['kind'];franchise=item['franchise'];dest=item['destination'];plan=item['plan']
    if kind in ('product_claims','claims','internal','branch_calculation') and batch['month'] not in store.verified_claim_months():
        raise ValueError('Paid claims for this month have not been verified. Import the complete paid-claims source or explicitly confirm a month with no paid claims in Admin.')
    allowed_scope={franchise} if franchise and (allowed is None or franchise in allowed) else allowed
    if franchise and allowed is not None and franchise not in allowed:raise ValueError('Franchise is outside your authorised scope.')
    path=Path(path);path.parent.mkdir(parents=True,exist_ok=True)
    # Generate into a temporary file; only complete validated output becomes downloadable.
    temporary=path.with_suffix('.pending.xlsx')
    book=xlsxwriter.Workbook(str(temporary),{'constant_memory':True,'strings_to_formulas':False,'strings_to_urls':False})
    writers=[];count=0;total=Decimal(0);issues=0
    def sheet(name,headers):
        w=SheetWriter(temporary,name,headers,book=book);writers.append(w);return w
    def summary(notes):
        w=sheet('Summary',['Item','Value'])
        for label,value in [('Report',item['title']),('Receipt month',batch['month']),('Franchise / group',franchise or dest or 'All authorised'),('Source file',batch['source_file']),('Source version',batch['version']),('Source batch',batch['id']),*notes]:w.append([label,value])
    try:
        if kind=='branch_calculation':
            from direct_data import allocation
            rows=[r for r in financial_rows(store,batch,allowed_scope) if allocation(r['Franchise'])==dest]
            fields=('Retail','BrightRock','Mkhulu','MFF','Original Risk','R1','ADV','Voucher Supplier','Voucher Retained','Claims')
            sums={key:sum((Decimal(r[key]) for r in rows),Decimal(0)) for key in fields}
            gross=sums['Original Risk']+sums['Voucher Supplier']+sums['Voucher Retained'];brokers=sums['Mkhulu']+sums['MFF'];fees=sums['R1']+sums['ADV'];vouchers=sums['Voucher Supplier']+sums['Voucher Retained']
            due=sums['Claims']+brokers+sums['BrightRock']+fees+vouchers
            headers=['Month','Retail','Bright Rock Fee (10%)','Brokers & Mkhulu','System Risk premium before COMS','Advertising Fund (2.1%) and MFF (R1.00) on System Risk premium','Voucher Coupons','Risk premium after all fees','Claims','Total fees and Claims payable to Brokers','Insurance pay-over money kept after claims and fees paid on Pay-over premium']
            w=sheet(dest,headers);w.sheet.set_row(0,105);w.sheet.set_column(5,5,55);w.sheet.set_column(9,10,55);w.append([batch['month']+'-01',sums['Retail'],sums['BrightRock'],brokers,gross,fees,vouchers,gross-sums['BrightRock']-brokers,sums['Claims'],due,gross-due]);count=1;total=gross-due
            summary([('Included franchises',', '.join(r['Franchise'] for r in rows)),('Source Risk',gross),('Paid claims use AM',sums['Claims']),('Reference column H','As in the example: E minus C minus D, before R1/ADV/vouchers. Final column K deducts those stored fees once.'),('Final calculation','K = E - (I + D + C + F + G). No external example amounts are used.')])
        elif kind=='reconciliation':
            summary([('Note','Exact source conservation. Rounding and zero-floor adjustments are explicit, never extra fees.')])
            headers=['Franchise','Source Rows','Source Risk','Insurance Net Risk','Faithnet','R1','ADV','Voucher Supplier','Voucher Retained','Rounding / Zero Floor','Risk Difference','Source Retail','Insurance Retail','Voucher Retail','Retail Difference','Source Version'];w=sheet('Reconciliation',headers)
            for source in store.summaries(batch['month'],batch['month'],allowed_scope):
                v=source['summary'];amounts={k:Decimal(str(v.get(k,0))) for k in ('original_risk','payover','faithnet','r1','adv','supplier','retained','reconciliation_adjustment','source_retail','retail','voucher_retail')}
                diff=amounts['original_risk']-sum((amounts[k] for k in ('payover','faithnet','r1','adv','supplier','retained','reconciliation_adjustment')),Decimal(0))
                retail_diff=amounts['source_retail']-amounts['retail']-amounts['voucher_retail']
                if diff or retail_diff:raise ValueError('Source conservation failed; final reports are blocked.')
                w.append([source['franchise'],v['rows'],amounts['original_risk'],amounts['payover'],amounts['faithnet'],amounts['r1'],amounts['adv'],amounts['supplier'],amounts['retained'],amounts['reconciliation_adjustment'],diff,amounts['source_retail'],amounts['retail'],amounts['voucher_retail'],retail_diff,batch['version']]);count+=v['rows'];total+=amounts['original_risk']
        elif kind in ('payover','commissions','internal'):
            rows=list(financial_rows(store,batch,allowed_scope,franchise))
            from central_ledger import controls
            ledger=controls(store,batch['month']) if kind=='payover' else {}
            summary([('Note','All amounts are stored calculations; R1/ADV are not deducted again.'),('Outstanding balances','Imported balance controls are compared with source totals; review each row status.' if ledger else 'Unavailable: external balances and bank payments have not been imported.')])
            if kind=='payover':
                headers=['Franchise','Account Number / Reference Number','System Pay-over','MFF Comission','Advertising Fund','Brokers & Mkhulu','Vouchers','Voucher COMS','Voucher Payment','AUL Pay-over','','CURRENT PAYOVER','OUTSTANDING PAYOVER','TOTAL DUE','DATUM BETAAL','BETALING','TOTAL OUTSTANDING']
                w=sheet('Monthly Pay-over',headers)
                for row in rows:
                    # Gross insurance Risk and voucher source Risk reconcile to System Pay-over;
                    # net insurer amount is read directly, never calculated by subtracting again.
                    vals=next((x['summary'] for x in store.summaries(batch['month'],batch['month'],allowed_scope) if x['franchise']==row['Franchise']),{'supplier':'0','retained':'0','original_risk':'0'})
                    gross_voucher=Decimal(vals['supplier'])+Decimal(vals['retained'])
                    gross_insurance=Decimal(vals['original_risk'])-gross_voucher
                    system=Decimal(vals['original_risk']);control=ledger.get(normal(row['Franchise']))
                    opening=Decimal(control['opening']) if control else None;payment=Decimal(control['payment']) if control else None
                    due=system+opening if opening is not None else None;outstanding=due-payment if due is not None else None
                    external=Decimal(control['system_control']) if control and control.get('system_control') not in ('',None) else None;difference=system-external if external is not None else None
                    status='Draft — import opening balances and payment controls' if not control else 'Review — external system differs' if difference else 'Reconciled'
                    w.append([row['Franchise'],control['account'] if control else 'Unconfirmed',system,row['R1'],row['ADV'],row['Mkhulu']+row['MFF'],gross_voucher,Decimal(vals['retained']),Decimal(vals['supplier']),row['Net Risk / Insurer Payover'],'',system,opening if opening is not None else 'Unconfirmed',due if due is not None else 'Unconfirmed',control['paid_on'] if control else '',payment if payment is not None else 'Unconfirmed',outstanding if outstanding is not None else 'Unconfirmed']);count+=1;total+=Decimal(row['Net Risk / Insurer Payover'])
                    control_notes=sheet('Control Review',['Franchise','Source Version','External System Control','Source / External Difference','Status']) if count==1 else control_notes
                    control_notes.append([row['Franchise'],batch['version'],external if external is not None else 'Unconfirmed',difference if difference is not None else 'Unconfirmed',status if external is not None or not control else 'Manual balances — no external System control'])
                detail=sheet('Insurer Calculation',['Franchise','Retail','BrightRock 10%','Brokers & Mkhulu 5%','Original Insurance Risk','Faithnet','R1 + ADV','Voucher Risk (separate)','Net Insurance Risk','Claims Paid Via Bank','Commissions + Bank Claims','Insurer Remaining After Commissions And Claims','Claims Verification'])
                for row in rows:
                    ready=batch['month'] in store.verified_claim_months()
                    detail.append([row['Franchise'],row['Retail'],row['BrightRock'],row['Mkhulu']+row['MFF'],row['Original Risk'],row['Faithnet Deduction'],row['R1']+row['ADV'],row['Voucher Supplier']+row['Voucher Retained'],row['Net Risk / Insurer Payover'],row['Claims'] if ready else 'Unconfirmed',row['Percentage Commissions']+row['Claims'] if ready else 'Unconfirmed',row['Insurer Remaining After Commissions And Claims'],'Verified paid month' if ready else 'Import complete paid month or confirm no paid claims in Admin'])
            else:
                headers=list(rows[0]) if rows else ['Franchise','Month']
                if kind=='commissions':
                    headers=[h for h in headers if h not in ('Claims','Claim Count','Average Claim','Claim Ratio','Scenario 2 / BrightRock Month Value','Insurer Remaining After Commissions And Claims','BrightRock Running Balance','Running Balance History')]
                w=sheet(item['title'][:31],headers)
                for row in rows:w.append([row.get(k,'') for k in headers]);count+=1;total+=Decimal(row['Paid Commissions'])
        elif kind=='claims':
            summary([('Note','Paid claims use Net amount paid via bank. Net-offs are separate. Nonzero bank payments count once; dated net-off-only rows do not increase the paid-claim count.')]);headers=['Source Row','Franchise','Paid Month','Date Of Claim Paid','Product','Net amount paid via bank','Amount net off against claim','Final Amount (source reference)','Issue','Source Version'];w=sheet('Claims',headers)
            for row in store.claims(batch['month'],batch['month'],allowed_scope):
                w.append([row['n'],row['franchise'],row['month'],row['paid_on'],row['product'],Decimal(row['amount']),Decimal(row['net_off']),Decimal(row['final_reference']),row['issues'],batch['version']]);count+=1;total+=Decimal(row['amount'])
        elif kind=='product_claims':
            premiums=defaultdict(lambda:Decimal(0));claim_totals=defaultdict(lambda:Decimal(0));names={}
            for record in store.records(batch['id'],allowed_scope,voucher=False):
                product=record['data'][11];key=normal(product);names[key]=product;premiums[key]+=Decimal(record['payover'])
            for row in store.claims(batch['month'],batch['month'],allowed_scope):
                key=normal(row['product'])
                if not key or key not in premiums:key='__unallocated';names[key]='Unallocated claims — resolve product mapping'
                else:names[key]=row['product']
                claim_totals[key]+=Decimal(row['amount'])
            summary([('Note','Ratio is paid claims divided by stored net Risk. Unallocated claims remain in the total.')]);w=sheet('Claims per product',['Product','Claims','Premiums','Claims Ratio'])
            for key in sorted(set(premiums)|set(claim_totals)):
                r=ratio(claim_totals[key],premiums[key]);w.append([names[key],claim_totals[key],premiums[key],r['value']/100 if r['value'] is not None else r['status']]);count+=1;total+=premiums[key]
            r=ratio(sum(claim_totals.values()),sum(premiums.values()));w.append(['TOTAL',sum(claim_totals.values()),sum(premiums.values()),r['value']/100 if r['value'] is not None else r['status']])
        else:
            policy=kind in ('policy','complete','combine','unclassified')
            if kind=='supplier':headers=['Franchise','Policy Number','Surname','Full Names','Date of Birth','Passport Nr','Relation','Inseption Date','Product','Fee x MPIA','Contact Number','MPIA']
            elif kind=='vouchers_internal':headers=['Source Row','Franchise','Policy Number','Product','Calculation MPIA','Voucher Units','Risk / Gross','Martins Retained','Supplier Fee','Contact Number','Issue','Source Version']
            elif kind=='review':headers=['Source Row','Franchise','Policy Number','Issue','Source Version']
            elif kind in ('complete','combine','unclassified'):headers=HEADERS+['Original Franchise','Original Risk','Calculation MPIA','Faithnet Deduction','R1','ADV','Net Risk','BrightRock','Mkhulu','MFF','Paid Commissions','Admin / Scenario 1','Book Value','Plan','Issue','Source Version']
            elif kind=='policy' and plan=='Plan 2':headers=[HEADERS[i] for i in PLAN2_COLUMNS]
            else:headers=HEADERS
            # The insurer load schema remains header-first, as in the supplied examples.
            w=sheet('Sheet1',headers)
            for record in store.records(batch['id'],allowed_scope,destination=dest,plan=plan,voucher=True if kind in ('supplier','vouchers_internal') else False if policy else None):
                data=record['data'].copy();calc=record['calculation'];notes=record['issues']
                if kind=='review':
                    for message in notes:w.append([record['n'],record['franchise'],record['policy'],message,batch['version']]);count+=1
                    continue
                if kind=='supplier':
                    if not record['contact'] or any('contact' in s.lower() or 'phone' in s.lower() or 'voucher risk differs' in s.lower() for s in notes):raise ValueError('Voucher contacts or fee reconciliation require review before supplier export.')
                    values=[data[0],data[2],data[3],data[4],data[6],'',data[8],data[9],data[11],Decimal(record['supplier']),record['contact'],Decimal(calc['calculation_mpia'])];total+=Decimal(record['supplier'])
                elif kind=='vouchers_internal':values=[record['n'],data[0],data[2],data[11],Decimal(calc['calculation_mpia']),Decimal(calc['voucher_units']),Decimal(record['original_risk']),Decimal(record['retained']),Decimal(record['supplier']),record['contact'],'; '.join(notes),batch['version']];total+=Decimal(record['supplier'])
                elif kind=='source':
                    values=calc['source'].copy();values[6]=data[6];values[9]=data[9];total+=Decimal(record['original_risk'])
                else:
                    if kind=='policy' and any(s.startswith('[BLOCK]') for s in notes):raise ValueError('This insurer dataset has unclassified records. Resolve its Review report before final export.')
                    data[12]=record['payover']
                    if kind=='combine':data[0]=COMBINE.get(normal(data[0]),data[0])
                    if kind in ('complete','combine','unclassified'):
                        values=data+[record['franchise'],Decimal(record['original_risk']),Decimal(calc['calculation_mpia']),Decimal(record['faithnet']),Decimal(record['r1']),Decimal(record['adv']),Decimal(record['payover']),Decimal(calc['brightrock']),Decimal(calc['mkhulu']),Decimal(calc['mff']),Decimal(calc['paid_commissions']),Decimal(calc['admin']),Decimal(calc['book_value']),record['plan'],'; '.join(notes),batch['version']]
                    elif plan=='Plan 2':values=[data[i] for i in PLAN2_COLUMNS]
                    else:values=data
                    total+=Decimal(record['payover'])
                w.append(values);count+=1;issues+=bool(notes)
            summary([('Exported records',count),('Relevant total',total),('Records with review notes',issues),('Note','Voucher Passport column is blank: the supplier receives the approved limited fields.' if kind=='supplier' else 'Original franchise retained; all amounts come from stored source calculations.')])
        for w in writers:w.close()
        book.close();temporary.replace(path)
        return {'rows':count,'total':str(total),'review_rows':issues,'version':batch['version'],'name':path.name}
    except Exception:
        book.close()
        temporary.unlink(missing_ok=True)
        raise
