"""Durable database-backed jobs and chunked private source storage."""
import csv,hashlib,json,sqlite3,threading,time,uuid,re
from contextlib import closing
from datetime import datetime,timedelta
from pathlib import Path
from decimal import Decimal
from flask import g,request,jsonify
from central_calculations import VERSION

class JobQueue:
    def __init__(self,app,store,root,audit,resolver=None):
        self.app=app;self.store=store;self.root=Path(root)/'central_jobs';self.root.mkdir(parents=True,exist_ok=True)
        self.audit=audit;self.resolver=resolver;self.ready=False;self.lock=threading.Lock();self.worker=None
    def ensure(self):
        if self.ready:return
        d=self.store.direct
        with d.connection() as conn:
            if not d.get_engine:conn.execute('BEGIN IMMEDIATE')
            else:d.execute(conn,'SELECT pg_advisory_xact_lock(hashtext(:key))',{'key':'claims-central-jobs-schema-v1'})
            d.execute(conn,'CREATE TABLE IF NOT EXISTS central_jobs (id TEXT PRIMARY KEY,owner_id INTEGER NOT NULL,kind TEXT NOT NULL,args TEXT NOT NULL,state TEXT NOT NULL,stage TEXT NOT NULL,current_rows INTEGER NOT NULL,total_rows INTEGER NOT NULL,source_bytes INTEGER NOT NULL,created_at TEXT NOT NULL,updated_at TEXT NOT NULL,error TEXT NOT NULL,result TEXT NOT NULL)')
            payload='BYTEA' if d.get_engine else 'BLOB'
            d.execute(conn,'CREATE TABLE IF NOT EXISTS central_job_sources (job_id TEXT NOT NULL,chunk_no INTEGER NOT NULL,payload '+payload+' NOT NULL,PRIMARY KEY(job_id,chunk_no))')
            d.execute(conn,'CREATE INDEX IF NOT EXISTS central_job_state ON central_jobs(state,updated_at)')
        self.ready=True
    def create(self,owner,kind,args,upload=None):
        self.ensure();identifier=uuid.uuid4().hex;stamp=datetime.utcnow().isoformat();digest=hashlib.sha256();size=0;d=self.store.direct
        with d.connection() as conn:
            d.execute(conn,'INSERT INTO central_jobs VALUES (:id,:owner,:kind,:args,:state,:stage,0,0,0,:created,:updated,\'\',\'{}\')',dict(id=identifier,owner=owner,kind=kind,args=json.dumps(args),state='receiving',stage='Receiving source',created=stamp,updated=stamp))
            if upload:
                i=0
                while True:
                    chunk=upload.stream.read(1024*1024)
                    if not chunk:break
                    digest.update(chunk);size+=len(chunk)
                    d.execute(conn,'INSERT INTO central_job_sources VALUES (:job,:chunk,:payload)',dict(job=identifier,chunk=i,payload=chunk));i+=1
                args['sha256']=digest.hexdigest();args['audit_id']=self.audit.start(owner,{'direct':'Central Direct','ledger':'Pay-over balances'}.get(kind,'Paid claims'),args['filename'],digest.hexdigest())
            elif kind=='migrate':args['audit_id']=self.audit.start(owner,'Saved Direct conversion',args['filename'],'')
            d.execute(conn,"UPDATE central_jobs SET args=:args,state='queued',stage='Queued for processing',source_bytes=:bytes,updated_at=:time WHERE id=:id",dict(args=json.dumps(args),bytes=size,time=datetime.utcnow().isoformat(),id=identifier))
        self.kick();return identifier
    def status(self,identifier,owner):
        self.ensure();d=self.store.direct
        with d.connection() as conn:
            rows=d.records(d.execute(conn,'SELECT id,owner_id,kind,state,stage,current_rows,total_rows,source_bytes,created_at,updated_at,error,result FROM central_jobs WHERE id=:id AND owner_id=:owner',dict(id=identifier,owner=owner)))
        if not rows:return None
        row=rows[0];row['result']=json.loads(row['result']);row.pop('owner_id');row['elapsed_seconds']=int((datetime.utcnow()-datetime.fromisoformat(row['created_at'])).total_seconds())
        self.kick();return row
    def update(self,identifier,stage='',current=None,total=None,state=None,error=None,result=None):
        fields=['updated_at=:time'];params={'time':datetime.utcnow().isoformat(),'id':identifier}
        for key,value in [('stage',stage or None),('current_rows',current),('total_rows',total),('state',state),('error',error),('result',json.dumps(result,default=str) if result is not None else None)]:
            if value is not None:fields.append(key+'=:'+key);params[key]=value
        d=self.store.direct
        with d.connection() as conn:d.execute(conn,'UPDATE central_jobs SET '+','.join(fields)+' WHERE id=:id',params)
    def kick(self):
        with self.lock:
            if self.worker and self.worker.is_alive():return
            self.worker=threading.Thread(target=self.run,daemon=True,name='Claims durable job worker');self.worker.start()
    def claim(self):
        self.ensure();d=self.store.direct;cutoff=(datetime.utcnow()-timedelta(minutes=15)).isoformat()
        with d.connection() as conn:
            if not d.get_engine:conn.execute('BEGIN IMMEDIATE')
            # A lease allows recovery after a process restart without stealing a live worker.
            d.execute(conn,"UPDATE central_jobs SET state='queued',stage='Recovering interrupted job' WHERE state='processing' AND updated_at<:cutoff",{'cutoff':cutoff})
            sql="SELECT * FROM central_jobs WHERE state='queued' ORDER BY created_at LIMIT 1"+(' FOR UPDATE SKIP LOCKED' if d.get_engine else '')
            rows=d.records(d.execute(conn,sql))
            if not rows:return None
            job=rows[0];d.execute(conn,"UPDATE central_jobs SET state='processing',stage='Starting',updated_at=:now WHERE id=:id",dict(now=datetime.utcnow().isoformat(),id=job['id']))
            job['args']=json.loads(job['args']);return job
    def restore_source(self,job):
        path=self.root/(job['id']+Path(job['args']['filename']).suffix.lower());d=self.store.direct;digest=hashlib.sha256()
        with path.open('wb') as output,d.connection() as conn:
            sql='SELECT payload FROM central_job_sources WHERE job_id=:id ORDER BY chunk_no'
            from sqlalchemy import text
            rows=conn.execution_options(stream_results=True).execute(text(sql),{'id':job['id']}) if d.get_engine else conn.execute(sql,{'id':job['id']})
            for row in rows:chunk=bytes(row[0]);digest.update(chunk);output.write(chunk)
        if digest.hexdigest()!=job['args']['sha256']:raise ValueError('Stored source checksum mismatch; previous published month was preserved.')
        return path
    def store_export(self,identifier,path):
        d=self.store.direct;digest=hashlib.sha256()
        with d.connection() as conn,path.open('rb') as source:
            d.execute(conn,'DELETE FROM central_job_sources WHERE job_id=:id',{'id':identifier})
            number=0
            while True:
                chunk=source.read(1024*1024)
                if not chunk:break
                digest.update(chunk);d.execute(conn,'INSERT INTO central_job_sources VALUES (:job,:chunk,:payload)',dict(job=identifier,chunk=number,payload=chunk));number+=1
        return digest.hexdigest()
    def export_path(self,identifier,result):
        path=self.root/(identifier+'.xlsx')
        if path.exists():return path
        d=self.store.direct;digest=hashlib.sha256();temp=path.with_suffix('.restoring')
        with d.connection() as conn,temp.open('wb') as target:
            from sqlalchemy import text
            sql='SELECT payload FROM central_job_sources WHERE job_id=:id ORDER BY chunk_no'
            records=conn.execution_options(stream_results=True).execute(text(sql),{'id':identifier}) if d.get_engine else conn.execute(sql,{'id':identifier})
            for row in records:chunk=bytes(row[0]);digest.update(chunk);target.write(chunk)
        if digest.hexdigest()!=result['sha256']:temp.unlink(missing_ok=True);raise ValueError('Report checksum mismatch. Regenerate this report.')
        temp.replace(path);return path
    def run(self):
        with self.app.app_context():
            while True:
                job=self.claim()
                if job is None:return
                identifier=job['id'];stop=threading.Event()
                def heartbeat():
                    while not stop.wait(30):
                        try:self.update(identifier)
                        except Exception:self.app.logger.exception('Job lease refresh failed')
                lease=threading.Thread(target=heartbeat,daemon=True);lease.start()
                path=None
                try:
                    user=self.resolver(job['owner_id']) if self.resolver else {'id':job['owner_id'],'role':'admin','is_active':True}
                    if not user or not user.get('is_active',True) or user.get('role')!='admin' and job['kind']!='export':raise ValueError('Import owner is no longer an active administrator.')
                    g.user=user
                    from import_progress import _state
                    previous=0
                    def progress(event):
                        nonlocal previous
                        now=time.monotonic()
                        if now-previous>1 or event.get('total')==event.get('current'):
                            self.update(identifier,event['stage'],event.get('current'),event.get('total'));previous=now
                    _state.emit=progress
                    if job['kind']=='direct':
                        path=self.restore_source(job)
                        from direct_data import stage_source,prepare_outputs
                        from direct_analytics import data_month
                        month=data_month(job['args']['filename'],job['args'].get('month',''))
                        result=self.store.existing_batch(identifier)
                        if result is None:
                            with closing(sqlite3.connect(self.root/(identifier+'.sqlite'))) as stage:
                                source=stage_source(path,stage)
                                _,_,_,issues=prepare_outputs(stage,self.root/identifier,job['args']['filename'],'mpia','mpia','policy',generate_files=False)
                                self.store.direct.save(stage,identifier,job['owner_id'],month,job['args']['filename'],central_only=True)
                            result={'ok':True,'batch_id':identifier,'month':month,'source':source,'issues':len(issues),'version':VERSION}
                    elif job['kind']=='claims':
                        path=self.restore_source(job)
                        from central_claims import read_transactions,publish_claims,verified_annual_sheet,VERIFIED_AMOUNT_HEADER
                        annual_sheet=verified_annual_sheet(path)
                        with closing(sqlite3.connect(self.root/(identifier+'.sqlite'))) as stage:
                            stage.execute('CREATE TABLE claims(n INTEGER PRIMARY KEY,month TEXT,franchise TEXT,product TEXT,amount TEXT,source TEXT,issues TEXT)')
                            stage.execute('CREATE TABLE claim_events(n INTEGER PRIMARY KEY,source TEXT)')
                            known={r['franchise'] for r in self.store.summaries()};aliases={}
                            from app import franchise_match_key
                            for name in known:aliases.setdefault(franchise_match_key(name),set()).add(name)
                            count=0;amount=Decimal(0);paid_count=0
                            claims_controls={}
                            for record in read_transactions(path,VERIFIED_AMOUNT_HEADER,annual_sheet,claims_controls,lambda n,source:stage.execute('INSERT INTO claim_events VALUES (?,?)',(n,json.dumps(source)))):
                                issues=[];franchise=record['franchise'];matches=aliases.get(franchise_match_key(franchise),set())
                                if franchise not in known:
                                    if len(matches)==1:franchise=next(iter(matches));issues.append('Franchise resolved by the existing normalised alias; original source retained.')
                                    else:issues.append('Unresolved franchise allocation; transaction retained once in reconciliation.')
                                if not record['product']:issues.append('Unresolved product allocation; no policy-number multiplication was performed.')
                                stage.execute('INSERT INTO claims VALUES (?,?,?,?,?,?,?)',(record['n'],record['month'],franchise,record['product'],record['amount'],json.dumps(record['source']),json.dumps(issues)))
                                count+=1;amount+=Decimal(record['amount'])
                                paid_count+=int(Decimal(record['amount'])!=0)
                                if count%1000==0:self.update(identifier,'Validating paid claim transactions',count)
                            stage.commit();months=[r[0] for r in stage.execute('SELECT DISTINCT month FROM claims')]
                            chosen=job['args'].get('replace_months') or months
                            result=publish_claims(self.store,stage,job['args']['filename'],job['owner_id'],VERIFIED_AMOUNT_HEADER,chosen,batch_id=identifier)
                            result.update(ok=True,issues=claims_controls.get('undated_netoff_rows',0),source={'rows':count,'claims':str(amount),'claim_count':paid_count,**claims_controls},month=','.join(result['months']))
                    elif job['kind']=='claim_analysis':
                        from claim_insights import backfill
                        result=backfill(self.store,self,identifier)
                    elif job['kind']=='ledger':
                        from central_ledger import read_ledger,publish
                        path=self.restore_source(job)
                        records=list(read_ledger(path,job['args']['sheet']))
                        publish(self.store,identifier,job['owner_id'],job['args']['month'],job['args']['filename'],records)
                        result={'ok':True,'month':job['args']['month'],'source':{'rows':len(records)},'batch_id':identifier}
                    elif job['kind']=='migrate':
                        from direct_data import prepare_outputs
                        args=job['args'];result=self.store.existing_batch(identifier)
                        if result is None:
                            source_batch=next((b for b in self.store.direct.batches(args['source_owner']) if b['id']==args['source_batch']),None)
                            if source_batch is None:raise ValueError('The saved source has been replaced. Select its current month version.')
                            with closing(sqlite3.connect(self.root/(identifier+'.sqlite'))) as stage:
                                self.store.direct.restore_stage(args['source_owner'],args['source_batch'],stage)
                                stage.execute('CREATE TABLE source_raw(n INTEGER PRIMARY KEY,data TEXT NOT NULL)')
                                for row in stage.execute('SELECT n,values_json FROM rows'):stage.execute('INSERT INTO source_raw VALUES (?,?)',(row[0],row[1]))
                                _,_,_,issues=prepare_outputs(stage,self.root/identifier,source_batch['source_file'],'mpia','mpia','policy',generate_files=False)
                                self.store.direct.save(stage,identifier,args['source_owner'],source_batch['month'],source_batch['source_file'],central_only=True)
                        result=self.store.existing_batch(identifier);result.update(migrated_from=args['source_batch'],source_provenance='Saved source fields; original file bytes were not available for this older import.')
                    elif job['kind']=='verification':
                        from import_verification import verify_targets
                        result=verify_targets(self.store,self,job['args']['targets'],identifier)
                    elif job['kind']=='export':
                        from central_reports import generate,catalog
                        args=job['args'];batch=next((b for b in self.store.publications() if b['id']==args['batch']),None)
                        validator=self.app.extensions.get('central_export_access')
                        if validator and not validator(args):raise ValueError('Your report permission or source version changed. Generate a fresh authorised report.')
                        if batch is None:raise ValueError('This report source has been superseded. Choose the current published month.')
                        allowed=set(args['allowed']) if args['allowed'] is not None else None
                        item=next((i for i in catalog(self.store,batch,allowed) if i['name']==args['name']),None)
                        if item is None:raise ValueError('Report is outside authorised scope.')
                        out=self.root/(identifier+'.xlsx');metadata=generate(self.store,batch,item,out,allowed)
                        if validator and not validator(args):raise ValueError('Source, claims or balance controls changed while generating this report. Generate the current report again.')
                        digest=self.store_export(identifier,out)
                        result={'ok':True,'name':item['name'],'path':out.name,'batch_id':batch['id'],'metadata':metadata,'sha256':digest}
                    else:raise ValueError('Unknown processing task.')
                    if job['args'].get('audit_id'):self.audit.finish(job['args']['audit_id'],{**result,'source':{k:str(v) for k,v in result.get('source',{}).items()}})
                    self.update(identifier,'Checks complete' if job['kind']=='verification' else 'Published' if job['kind']!='export' else 'Ready for download',state='completed',result=result)
                except Exception as exc:
                    self.app.logger.exception('Central task failed')
                    message=str(exc)[:1500] if isinstance(exc,ValueError) else 'Processing failed. Previous valid publications remain available. See the import log.'
                    if job['args'].get('audit_id'):self.audit.finish(job['args']['audit_id'],{'ok':False,'error':message})
                    self.update(identifier,'Needs correction',state='failed',error=message,result={'ok':False,'error':message})
                finally:
                    stop.set()
                    from import_progress import _state
                    if hasattr(_state,'emit'):del _state.emit
                    if path:path.unlink(missing_ok=True)
                    (self.root/(identifier+'.sqlite')).unlink(missing_ok=True)

def register_jobs(app,store,export_dir,audit,resolver=None):
    jobs=JobQueue(app,store,export_dir,audit,resolver);app.extensions['central_jobs']=jobs
    from import_verification import register as register_verification
    register_verification(app,store,jobs)
    @app.before_request
    def resume_central_jobs():
        # Any authenticated request restarts durable queued work after a web-worker restart.
        if getattr(g,'user',None):jobs.kick()
    @app.post('/imports/central')
    def central_upload():
        user=getattr(g,'user',None) or {}
        if user.get('role')!='admin':return jsonify(ok=False,error='Administrator access required.'),403
        upload=request.files.get('file');kind=request.form.get('kind','direct')
        if kind not in ('direct','claims') or not upload or not upload.filename:return jsonify(ok=False,error='Select a monthly Direct file or a paid-claims workbook.'),400
        suffix=Path(upload.filename).suffix.lower()
        if suffix not in ('.xlsx','.csv') or kind=='claims' and suffix!='.xlsx':return jsonify(ok=False,error='Unsupported source format.'),400
        args={'filename':Path(upload.filename).name,'month':request.form.get('month','')}
        if kind=='direct' and not re.match(r'^PolicyData_DTSTAMP_RECEIPTSFOR_\d{8}_to_\d{8}',args['filename'],re.I):return jsonify(ok=False,error='Use the Direct PolicyData_DTSTAMP_RECEIPTSFOR monthly source, in the October example format. Insurer exports and older policy/premium files are not import sources.'),400
        if kind=='direct':
            match=re.match(r'^PolicyData_DTSTAMP_RECEIPTSFOR_(\d{8})_to_(\d{8})',args['filename'],re.I)
            try:
                first,last=(datetime.strptime(v,'%Y%m%d') for v in match.groups())
                import calendar
                if first.day!=1 or (first.year,first.month)!=(last.year,last.month) or last.day!=calendar.monthrange(last.year,last.month)[1]:raise ValueError
            except ValueError:return jsonify(ok=False,error='The Direct source must contain one complete named calendar month.'),400
        if kind=='claims' and not re.match(r'^Martins Funeral All Claims\b',args['filename'],re.I):return jsonify(ok=False,error='Use the Martins Funeral All Claims annual source workbook.'),400
        if kind=='claims':
            from central_claims import VERIFIED_AMOUNT_HEADER
            args.update(amount_header=VERIFIED_AMOUNT_HEADER,sheet='',replace_months=[])
            from central_calculations import normal
            if normal(args['amount_header'])!=normal(VERIFIED_AMOUNT_HEADER):return jsonify(ok=False,error='Paid claims must use the approved Net amount paid via bank heading.'),400
            # The approved annual source retains every actual paid month, including
            # payments beyond its reported-date window; no month is forced from its filename.
        try:identifier=jobs.create(user['id'],kind,args,upload)
        except Exception:
            app.logger.exception('Could not save source job');return jsonify(ok=False,error='The source could not be stored safely. No monthly data was replaced.'),500
        return jsonify(ok=True,job=identifier),202
    @app.get('/jobs/<identifier>')
    def central_job_status(identifier):
        user=getattr(g,'user',None) or {}
        if not user:return jsonify(ok=False,error='Authentication required.'),401
        status=jobs.status(identifier,user['id'])
        if status and status['kind']!='export' and user.get('role')!='admin':return jsonify(ok=False,error='Administrator access required.'),403
        if status and status['kind']=='export':
            validator=app.extensions.get('central_export_access')
            with store.direct.connection() as conn:raw=store.direct.execute(conn,'SELECT args FROM central_jobs WHERE id=:id AND owner_id=:owner',dict(id=identifier,owner=user['id'])).fetchone()
            if validator and not validator(json.loads(raw[0])):return jsonify(ok=False,error='Report permission or source version changed.'),403
        return jsonify(status) if status else (jsonify(ok=False,error='Job not found.'),404)
    return jobs
