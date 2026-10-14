"""Atomic monthly publications and immutable financial/source snapshots."""
import json
from collections import defaultdict
from datetime import datetime
from decimal import Decimal
from sqlalchemy import text
from central_calculations import VERSION,snapshot,ratio

SUM_FIELDS=('retail','source_retail','original_risk','faithnet','r1','adv','payover','supplier','retained',
 'policy_qty','brightrock','mkhulu','mff','percentage_commission','paid_commissions','admin',
 'risk_after_r1','mff_book_value','franchise_book_value','book_value','remaining_book_value','retail_less_payover','voucher_units',
 'insurance_original_risk','voucher_gross','voucher_retail','reconciliation_adjustment')

class CentralStore:
    def __init__(self,direct):self.direct=direct;self.ready=False
    def ensure(self):
        if self.ready:return
        with self.direct.connection() as conn:
            if self.direct.get_engine:self.direct.execute(conn,'SELECT pg_advisory_xact_lock(hashtext(:key))',{'key':'claims-central-schema-v1'})
            for sql in (
                'CREATE TABLE IF NOT EXISTS central_batches (id TEXT PRIMARY KEY,owner_id INTEGER NOT NULL,month TEXT NOT NULL,source_file TEXT NOT NULL,version TEXT NOT NULL,created_at TEXT NOT NULL,status TEXT NOT NULL,source_rows INTEGER NOT NULL,review_rows INTEGER NOT NULL)',
                'CREATE TABLE IF NOT EXISTS central_records (batch_id TEXT NOT NULL,n INTEGER NOT NULL,franchise TEXT NOT NULL,destination TEXT NOT NULL,plan TEXT NOT NULL,product TEXT NOT NULL,relation TEXT NOT NULL,record_json TEXT NOT NULL,PRIMARY KEY(batch_id,n))',
                'CREATE INDEX IF NOT EXISTS central_record_scope ON central_records(batch_id,franchise,destination,plan)',
                'CREATE TABLE IF NOT EXISTS central_member_index (batch_id TEXT NOT NULL,n INTEGER NOT NULL,franchise TEXT NOT NULL,system_ref TEXT NOT NULL,product TEXT NOT NULL,PRIMARY KEY(batch_id,n))',
                'CREATE TABLE IF NOT EXISTS central_product_monthly (batch_id TEXT NOT NULL,franchise TEXT NOT NULL,product TEXT NOT NULL,destination TEXT NOT NULL,plan TEXT NOT NULL,summary TEXT NOT NULL,PRIMARY KEY(batch_id,franchise,product,destination,plan))',
                'CREATE TABLE IF NOT EXISTS central_product_benefits (batch_id TEXT NOT NULL,franchise TEXT NOT NULL,product TEXT NOT NULL,cover TEXT NOT NULL,PRIMARY KEY(batch_id,franchise,product,cover))',
                'CREATE TABLE IF NOT EXISTS central_benefit_versions (batch_id TEXT PRIMARY KEY,member_rows INTEGER NOT NULL,built_at TEXT NOT NULL)',
                'CREATE TABLE IF NOT EXISTS central_monthly (batch_id TEXT NOT NULL,franchise TEXT NOT NULL,summary TEXT NOT NULL,PRIMARY KEY(batch_id,franchise))',
                'CREATE TABLE IF NOT EXISTS central_publications (month TEXT PRIMARY KEY,batch_id TEXT NOT NULL,version TEXT NOT NULL,published_at TEXT NOT NULL)',
                'CREATE TABLE IF NOT EXISTS central_claim_batches (id TEXT PRIMARY KEY,source_file TEXT NOT NULL,owner_id INTEGER NOT NULL,created_at TEXT NOT NULL,amount_header TEXT NOT NULL,status TEXT NOT NULL)',
                'CREATE TABLE IF NOT EXISTS central_claims (batch_id TEXT NOT NULL,n INTEGER NOT NULL,month TEXT NOT NULL,franchise TEXT NOT NULL,product TEXT NOT NULL,amount TEXT NOT NULL,source_json TEXT NOT NULL,issues TEXT NOT NULL,PRIMARY KEY(batch_id,n))',
                'CREATE INDEX IF NOT EXISTS central_claim_scope ON central_claims(month,franchise,product)',
                'CREATE TABLE IF NOT EXISTS central_claim_components (batch_id TEXT NOT NULL,n INTEGER NOT NULL,net_off TEXT NOT NULL,bank_paid TEXT NOT NULL,PRIMARY KEY(batch_id,n))',
                'CREATE TABLE IF NOT EXISTS central_claim_publications (month TEXT PRIMARY KEY,batch_id TEXT NOT NULL)',
                'CREATE TABLE IF NOT EXISTS central_claim_events (batch_id TEXT NOT NULL,n INTEGER NOT NULL,franchise TEXT NOT NULL,source_json TEXT NOT NULL,PRIMARY KEY(batch_id,n))',
                'CREATE TABLE IF NOT EXISTS central_claim_event_versions (batch_id TEXT PRIMARY KEY,version TEXT NOT NULL,source_rows INTEGER NOT NULL)',
                'CREATE TABLE IF NOT EXISTS central_recipient_access (user_id INTEGER NOT NULL,kind TEXT NOT NULL,value TEXT NOT NULL,PRIMARY KEY(user_id,kind,value))',
                'CREATE TABLE IF NOT EXISTS central_legacy_archive (batch_id TEXT NOT NULL,n INTEGER NOT NULL,record_json TEXT NOT NULL,PRIMARY KEY(batch_id,n))',
                'CREATE TABLE IF NOT EXISTS central_legacy_batches (batch_id TEXT PRIMARY KEY,metadata TEXT NOT NULL)',
            ):self.direct.execute(conn,sql)
        self.ready=True

    def archive_legacy(self,conn,identifier):
        """Keep the pre-specification source/fee snapshot before replacing its legacy table."""
        from direct_analytics import encode
        if self.direct.execute(conn,'SELECT COUNT(*) FROM central_batches WHERE id=:id',{'id':identifier}).fetchone()[0]:return
        metadata=self.direct.records(self.direct.execute(conn,'SELECT * FROM direct_data_batches WHERE id=:id',{'id':identifier}))
        if not metadata:return
        self.direct.execute(conn,'INSERT INTO central_legacy_batches VALUES (:id,:metadata) ON CONFLICT(batch_id) DO NOTHING',dict(id=identifier,metadata=encode(metadata[0])))
        sql='SELECT * FROM direct_data_rows WHERE batch_id=:id ORDER BY n'
        result=conn.execute(text(sql).execution_options(stream_results=True),{'id':identifier}) if self.direct.get_engine else conn.execute(sql,{'id':identifier})
        records=result.mappings() if self.direct.get_engine else result
        for row in records:
            row=dict(row);self.direct.execute(conn,'INSERT INTO central_legacy_archive VALUES (:id,:n,:record) ON CONFLICT(batch_id,n) DO NOTHING',dict(id=identifier,n=row['n'],record=encode(row)))
        result.close() if hasattr(result,'close') else None

    def reject_older_job(self,conn,identifier,month,publication_table):
        """Newest received import wins even when multiple workers finish out of order."""
        if publication_table not in ('central_publications','central_claim_publications','central_ledger_publications'):raise ValueError('Invalid publication stream.')
        exists=self.direct.execute(conn,"SELECT to_regclass('central_jobs')" if self.direct.get_engine else "SELECT name FROM sqlite_master WHERE name='central_jobs'").fetchone()
        if not exists or not exists[0]:return
        incoming=self.direct.execute(conn,'SELECT created_at FROM central_jobs WHERE id=:id',{'id':identifier}).fetchone()
        previous=self.direct.execute(conn,'SELECT j.created_at FROM '+publication_table+' p JOIN central_jobs j ON j.id=p.batch_id WHERE p.month=:month',{'month':month}).fetchone()
        if incoming and previous and previous[0]>incoming[0]:raise ValueError('A newer received import for this month is already published. This older task was superseded; no current data was replaced.')

    def publish(self,stage,conn,batch_id,owner,month,filename):
        """Called inside the same transaction as Direct.save; old central versions remain archived."""
        from direct_analytics import FIELDS,encode
        if self.direct.get_engine:
            self.direct.execute(conn,'SELECT pg_advisory_xact_lock(hashtext(:key))',{'key':'central:'+month})
        self.reject_older_job(conn,batch_id,month,'central_publications')
        totals={};rows=[];count=review=0;member_rows=[];product_totals={};benefits=set();member_count=0
        expected=stage.execute('SELECT COUNT(*) FROM prepared').fetchone()[0]
        has_raw=stage.execute("SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='source_raw'").fetchone()[0]
        has_extra=stage.execute("SELECT COUNT(*) FROM sqlite_master WHERE name='source_extra'").fetchone()[0]
        for raw in stage.execute('SELECT * FROM prepared ORDER BY n'):
            prepared=dict(raw) if hasattr(raw,'keys') else dict(zip(FIELDS,raw))
            data=json.loads(prepared['data'])
            source=json.loads(stage.execute('SELECT data FROM source_raw WHERE n=?',(prepared['n'],)).fetchone()[0]) if has_raw else data
            calc=snapshot(data,prepared,source)
            record={**prepared,'data':data,'issues':json.loads(prepared['issues']),'calculation':calc}
            if has_extra:
                extra=stage.execute('SELECT data FROM source_extra WHERE n=?',(prepared['n'],)).fetchone()
                if extra:record['source_extra']=json.loads(extra[0])
            item=totals.setdefault(prepared['franchise'],{'rows':0,'review_rows':0,'mem_rows':0,'voucher_rows':0,**{k:Decimal(0) for k in SUM_FIELDS}})
            for key in ('aul_rows','aul_plan2_rows','aul_historic_rows','aul_payover','aul_plan2_payover','aul_historic_payover'):item.setdefault(key,Decimal(0))
            if prepared['destination']=='AUL':
                item['aul_rows']+=1;item['aul_payover']+=Decimal(calc['payover'])
                group='plan2' if prepared['plan']=='Plan 2' else 'historic' if prepared['plan']=='Historic' else None
                if group:item['aul_'+group+'_rows']+=1;item['aul_'+group+'_payover']+=Decimal(calc['payover'])
            item['rows']+=1;count+=1
            relation_key='relation_'+prepared['relation'] if prepared['relation'] else 'relation_BLANK'
            item[relation_key]=item.get(relation_key,0)+1
            if record['issues']:item['review_rows']+=1;review+=1
            if prepared['destination']=='Vouchers':item['voucher_rows']+=1
            elif prepared['relation']=='MEM':
                item['mem_rows']+=1
                from performance_index import cover
                benefit=cover(data[10]);member_count+=1
                if benefit is not None:benefits.add((prepared['franchise'],data[11],benefit))
                member_rows.append(dict(batch=batch_id,n=prepared['n'],franchise=prepared['franchise'],ref=str(data[16] or '').strip(),product=data[11]))
            for key in SUM_FIELDS:item[key]+=Decimal(calc[key])
            product_key=(prepared['franchise'],data[11],prepared['destination'],prepared['plan'])
            product_values=product_totals.setdefault(product_key,{'rows':0,'retail':Decimal(0),'payover':Decimal(0)})
            product_values['rows']+=1;product_values['retail']+=Decimal(calc['retail']);product_values['payover']+=Decimal(calc['payover'])
            rows.append({'batch':batch_id,'n':prepared['n'],'franchise':prepared['franchise'],
                'destination':prepared['destination'],'plan':prepared['plan'],'product':data[11],
                'relation':prepared['relation'],'record':encode(record)})
            if len(rows)>=1000:
                self._write(conn,rows);rows=[]
                self._write_members(conn,member_rows);member_rows=[]
                from import_progress import report
                if self.direct.get_engine:report('Saving calculated records',count,expected)
        if rows:self._write(conn,rows)
        if member_rows:self._write_members(conn,member_rows)
        if not count:raise ValueError('No financial source records; previous publication is preserved.')
        if count!=stage.execute('SELECT COUNT(*) FROM rows').fetchone()[0]:raise ValueError('Not every source record reached the calculation dataset. Previous publication preserved.')
        for values in totals.values():
            if values['original_risk']!=sum((values[k] for k in ('payover','faithnet','r1','adv','supplier','retained','reconciliation_adjustment')),Decimal(0)):
                raise ValueError('Risk reconciliation failed; previous publication preserved.')
            if values['source_retail']!=values['retail']+values['voucher_retail']:
                raise ValueError('Retail reconciliation failed; previous publication preserved.')
            if values['percentage_commission']!=values['retail']*Decimal('.15'):
                raise ValueError('Commission reconciliation failed; previous publication preserved.')
            if values['r1']!=values['policy_qty']:
                raise ValueError('MEM R1 does not balance to paid policy months; previous publication preserved.')
        stamp=datetime.utcnow().isoformat()
        self.direct.execute(conn,'INSERT INTO central_batches VALUES (:id,:owner,:month,:source,:version,:created,:status,:count,:review)',
            dict(id=batch_id,owner=owner,month=month,source=filename,version=VERSION,created=stamp,status='Published with review' if review else 'Published',count=count,review=review))
        for franchise,values in totals.items():
            self.direct.execute(conn,'INSERT INTO central_monthly VALUES (:batch,:franchise,:summary)',dict(batch=batch_id,franchise=franchise,summary=encode(values)))
        for (franchise,product,destination,plan),values in product_totals.items():
            self.direct.execute(conn,'INSERT INTO central_product_monthly VALUES (:batch,:franchise,:product,:dest,:plan,:summary)',dict(batch=batch_id,franchise=franchise,product=product,dest=destination,plan=plan,summary=encode(values)))
        from performance_index import write as write_benefits
        write_benefits(self,conn,batch_id,benefits,member_count)
        from import_progress import report
        if self.direct.get_engine:report('Committing reconciled month',count,count)
        self.direct.execute(conn,'INSERT INTO central_publications VALUES (:month,:batch,:version,:created) ON CONFLICT(month) DO UPDATE SET batch_id=excluded.batch_id,version=excluded.version,published_at=excluded.published_at',dict(month=month,batch=batch_id,version=VERSION,created=stamp))

    def _write(self,conn,rows):
        sql='INSERT INTO central_records VALUES (:batch,:n,:franchise,:destination,:plan,:product,:relation,:record)'
        if self.direct.get_engine:conn.execute(text(sql),rows)
        else:conn.executemany(sql,rows)

    def _write_members(self,conn,rows):
        if not rows:return
        sql='INSERT INTO central_member_index VALUES (:batch,:n,:franchise,:ref,:product)'
        if self.direct.get_engine:conn.execute(text(sql),rows)
        else:conn.executemany(sql,rows)

    def publications(self):
        self.ensure()
        with self.direct.connection() as conn:
            return self.direct.records(self.direct.execute(conn,'SELECT b.* FROM central_batches b JOIN central_publications p ON p.batch_id=b.id ORDER BY b.month'))

    def existing_batch(self,identifier):
        self.ensure()
        with self.direct.connection() as conn:
            batches=self.direct.records(self.direct.execute(conn,'SELECT * FROM central_batches WHERE id=:id',{'id':identifier}))
            if not batches:return None
            current=self.direct.execute(conn,'SELECT COUNT(*) FROM central_publications WHERE batch_id=:id',{'id':identifier}).fetchone()[0]
            rows=self.direct.execute(conn,'SELECT summary FROM central_monthly WHERE batch_id=:id',{'id':identifier})
            risk=retail=Decimal(0)
            for row in rows:
                values=json.loads(row[0]);risk+=Decimal(values['original_risk']);retail+=Decimal(values['source_retail'])
            batch=batches[0]
            return {'ok':True,'batch_id':identifier,'month':batch['month'],'version':batch['version'],
              'source':{'rows':batch['source_rows'],'risk':str(risk),'retail':str(retail)},'issues':batch['review_rows'],'superseded':not bool(current)}

    def summaries(self,start='',end='',allowed=None):
        self.ensure();params={};where=''
        if start:where=' WHERE b.month BETWEEN :start AND :end';params=dict(start=start,end=end)
        with self.direct.connection() as conn:
            rows=self.direct.records(self.direct.execute(conn,'SELECT b.month,b.id AS batch_id,b.version,m.franchise,m.summary FROM central_monthly m JOIN central_batches b ON b.id=m.batch_id JOIN central_publications p ON p.batch_id=b.id'+where+' ORDER BY b.month,m.franchise',params))
        rows=[r for r in rows if allowed is None or r['franchise'] in allowed]
        for row in rows:row['summary']=json.loads(row['summary'])
        return rows

    def records(self,batch_id,allowed=None,destination='',plan='',voucher=None,relation=''):
        self.ensure();params={'batch':batch_id};where=['r.batch_id=:batch']
        for key,value in [('destination',destination),('plan',plan),('relation',relation)]:
            if value:where.append('r.'+key+'=:'+key);params[key]=value
        if voucher is not None:where.append("r.destination"+("=" if voucher else "<>")+"'Vouchers'")
        with self.direct.connection() as conn:
            sql='SELECT r.record_json FROM central_records r JOIN central_publications p ON p.batch_id=r.batch_id WHERE '+' AND '.join(where)+' ORDER BY r.n'
            result=conn.execution_options(stream_results=True).execute(text(sql),params) if self.direct.get_engine else conn.execute(sql,params)
            for raw in result:
                record=json.loads(raw[0])
                if allowed is None or record['franchise'] in allowed:yield record
            result.close() if hasattr(result,'close') else None

    def claims(self,start='',end='',allowed=None):
        self.ensure();params={};where=''
        if start:where=' WHERE c.month BETWEEN :start AND :end';params=dict(start=start,end=end)
        with self.direct.connection() as conn:
            sql="SELECT c.n,c.month,c.franchise,c.product,c.amount,c.issues,c.source_json,COALESCE(x.net_off,'0') AS net_off,COALESCE(x.bank_paid,'0') AS bank_paid FROM central_claims c JOIN central_claim_publications p ON p.batch_id=c.batch_id AND p.month=c.month LEFT JOIN central_claim_components x ON x.batch_id=c.batch_id AND x.n=c.n"+where+' ORDER BY c.month,c.n'
            result=conn.execution_options(stream_results=True).execute(text(sql),params) if self.direct.get_engine else conn.execute(sql,params)
            records=result.mappings() if self.direct.get_engine else result
            for raw in records:
                row=dict(raw)
                from central_claims import source_paid_on,component
                source=json.loads(row.pop('source_json'));row['paid_on']=source_paid_on(source);row['final_reference']=str(component(source,'Final Amount'))
                if allowed is None or row['franchise'] in allowed:yield row
            result.close() if hasattr(result,'close') else None

    def verified_claim_months(self):
        self.ensure()
        with self.direct.connection() as conn:return {r[0] for r in self.direct.execute(conn,'SELECT month FROM central_claim_publications')}

    def frame(self,legacy_claims=None):
        """Read stored monthly financial values; claims and ratios are monthly aggregation only."""
        import pandas as pd
        summaries=self.summaries()
        claims=defaultdict(lambda:[Decimal(0),0])
        claim_months=set()
        for row in self.claims():
            claims[(row['franchise'],row['month'])][0]+=Decimal(row['amount']);claims[(row['franchise'],row['month'])][1]+=int(Decimal(row['amount'])!=0);claim_months.add(row['month'])
        # Legacy month aggregates remain visible until a verified transaction source replaces them.
        if not summaries and legacy_claims is not None and not legacy_claims.empty:
            for row in legacy_claims.to_dict('records'):
                month=pd.Timestamp(row['Month']).strftime('%Y-%m')
                if month not in claim_months:
                    claims[(row['Franchise'],month)]=[Decimal(str(row.get('Claims',0))),int(row.get('Claim Count',0) or 0)]
        present={(r['franchise'],r['month']) for r in summaries}
        versions={r['month']:r['version'] for r in summaries}
        for (franchise,month),values in claims.items():
            if (franchise,month) not in present and month in versions:
                summaries.append({'franchise':franchise,'month':month,'version':versions[month],
                  'summary':{**{key:'0' for key in SUM_FIELDS},'rows':0,'review_rows':1,'mem_rows':0,'voucher_rows':0}})
        labels={'retail':'Retail Premium','insurance_original_risk':'Original Risk Premium','payover':'Risk Premium',
          'risk_after_r1':'Risk After R1','r1':'R1 Policy Fee','adv':'Underwriter 2.1% Fee','policy_qty':'Policy Qty',
          'brightrock':'BrightRock Amount','mkhulu':'Inkulu Amount','mff':'MFF Amount',
          'percentage_commission':'Total Commission','paid_commissions':'Total Paid Commissions','admin':'Admin Fee',
          'mff_book_value':'MFF Book Value 2.5%','franchise_book_value':'Franchise Book Value 2.5%',
          'book_value':'Total Book Value','remaining_book_value':'Remaining Book Value','faithnet':'Faithnet Deduction',
          'supplier':'Voucher Supplier Payable','retained':'Voucher Retained','voucher_units':'Total Vouchers'}
        rows=[];seen=set()
        for source in summaries:
            values=source['summary'];franchise=source['franchise'];month=source['month'];amount,count=claims[(franchise,month)];seen.add((franchise,month))
            result={label:float(Decimal(str(values[key]))) for key,label in labels.items()}
            result['BrightRock Month Total']=result['Retail Premium']-float(amount)-result['Total Paid Commissions']
            risk=Decimal(values['payover']);r=ratio(amount,risk)
            result.update({'Franchise':franchise,'Month':pd.Timestamp(month+'-01'),'Claims':float(amount),'Claim Count':count,
              'Average Claim':float(amount/count) if count else 0,'Claim Ratio':float(r['value']) if r['value'] is not None else float('nan'),
              'Claim Ratio Status':r['colour'].title(),'Claim Ratio Label':r['status'],'Source Version':source['version'],
              'MEM Source Rows':values['mem_rows'],'Review Rows':values['review_rows'],
              'Scenario 1 Value':result['Admin Fee'],'Scenario 2 Value':result['Retail Premium']-float(amount)-result['Total Paid Commissions'],
              'BrightRock Month Total':result['Retail Premium']-float(amount)-result['Total Paid Commissions'],
              'Separate Paid Commissions':result['R1 Policy Fee']+result['Underwriter 2.1% Fee'],
              'Percentage Commission':result['Total Commission'],'Current Scenario':'100% Claim Ratio',
              'Total Commission Rate':15,'Recommendation':'Review required — recommendation rule not confirmed',
              'Scenario 1 Book Value':result['Total Book Value'],'Scenario 2 Book Value':result['Total Book Value'],
              'Claim Paid to Franchise':0,'Claim Paid to Client':0,'Repudiated / Pending':0,'Grand Total Claims':float(amount),
              '100% Move Status':'Review','100% Move Traffic':'Grey','100% Traffic Light':'Grey',
              '100% Traffic Result':'Review','100% Traffic Reason':'Recommendation is not a final decision',
              'BrightRock Traffic Light':'Green' if result['BrightRock Month Total']>=0 else 'Red',
              'BrightRock Money Result':'Positive' if result['BrightRock Month Total']>=0 else 'Negative',
              'BrightRock Eligible':False,'Months Below 75% Streak':0})
            rows.append(result)
        frame=pd.DataFrame(rows)
        if not frame.empty:
            frame=frame.sort_values(['Franchise','Month'])
            frame['BrightRock Running Balance']=frame.groupby('Franchise')['BrightRock Month Total'].cumsum()
            frame['Period']=frame['Month'].dt.strftime('%b %Y');frame['Period Index']=1
            frame['Previous Month Claims']=frame.groupby('Franchise')['Claims'].shift(1).fillna(0)
            frame.attrs['central_dataset']=True
        return frame
