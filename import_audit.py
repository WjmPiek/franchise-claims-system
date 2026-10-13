"""Durable metadata-only import attempts; source records are never appended by this module."""
import csv,hashlib,io,json,re,uuid
from datetime import datetime,timezone
from pathlib import Path
from functools import wraps
from flask import request,g,jsonify,render_template,Response
from direct_analytics import DirectStore

class ImportAudit:
    def __init__(self,path,get_engine=None):self.store=DirectStore(path,get_engine)
    def ensure(self):
        with self.store.connection() as conn:self.store.execute(conn,'CREATE TABLE IF NOT EXISTS import_audit_log (id TEXT PRIMARY KEY, owner_id INTEGER NOT NULL, source_kind TEXT NOT NULL, source_file TEXT NOT NULL, receipt_month TEXT NOT NULL, sha256 TEXT NOT NULL, started_at TEXT NOT NULL, finished_at TEXT NOT NULL, status TEXT NOT NULL, details TEXT NOT NULL)')
    def start(self,owner,kind,name,digest):
        self.ensure();identifier=uuid.uuid4().hex
        found=re.search(r'(\d{4})(\d{2})\d{2}_to_',name)
        month=found[1]+'-'+found[2] if found else ''
        with self.store.connection() as conn:self.store.execute(conn,'INSERT INTO import_audit_log VALUES (:id,:owner,:kind,:name,:month,:hash,:started,:finished,:status,:details)',dict(id=identifier,owner=owner,kind=kind,name=name,month=month,hash=digest,started=datetime.now(timezone.utc).isoformat(),finished='',status='processing',details=''))
        return identifier
    def finish(self,identifier,result,status=None):
        source=result.get('source') or {};details={key:source[key] for key in ('rows','source_rows','nonfinancial_rows','risk','retail','claims','claim_count','unpaid_rows','unpaid_final_amount','netoff_only_rows','undated_netoff_rows','undated_netoff_amount','undated_netoff_source_rows') if key in source}
        for key in ('version','months','issues'):
            if key in result:details[key]=result[key]
        if result.get('error'):details['error']=str(result['error'])[:2000]
        if result.get('month'):details['saved_month']=result['month']
        with self.store.connection() as conn:self.store.execute(conn,"UPDATE import_audit_log SET status=:status,finished_at=:time,details=:details,receipt_month=COALESCE(NULLIF(:month,''),receipt_month) WHERE id=:id",dict(month=result.get('month',''),status=status or ('saved / replaced' if result.get('ok') else 'failed / review'),time=datetime.now(timezone.utc).isoformat(),details=json.dumps(details,ensure_ascii=False),id=identifier))
    def rows(self,owner):
        self.ensure()
        with self.store.connection() as conn:return self.store.records(self.store.execute(conn,'SELECT * FROM import_audit_log WHERE owner_id=:owner ORDER BY started_at DESC',{'owner':owner}))

def register_import_audit(app,export_dir,get_engine=None):
    audit=ImportAudit(Path(export_dir)/'import_audit.sqlite',get_engine)
    for endpoint,kind,fields in [('direct_data','Direct',['direct_file']),('dashboard','Policy / claims',['excel_file','excel_files','claims_file','claims_files'])]:
        original=app.view_functions[endpoint]
        def decorate(original,kind,fields):
            @wraps(original)
            def logged(*args,**kwargs):
                if request.method!='POST':return original(*args,**kwargs)
                files=[file for field in fields for file in request.files.getlist(field) if file.filename]
                if not files:return original(*args,**kwargs)
                user=getattr(g,'user',None) or {}
                if user.get('role')!='admin':return original(*args,**kwargs)
                ids=[]
                for file in files:
                    position=file.stream.tell();digest=hashlib.sha256()
                    while True:
                        chunk=file.stream.read(1024*1024)
                        if not chunk:break
                        digest.update(chunk)
                    file.stream.seek(position)
                    actual_kind=kind if kind=='Direct' else 'Claims paid' if file.name in ('claims_file','claims_files') else 'Policy / premium'
                    ids.append(audit.start(user['id'],actual_kind,file.filename,digest.hexdigest()))
                def finish(result,status=None):
                    for identifier in ids:audit.finish(identifier,result,status)
                try:response=app.make_response(original(*args,**kwargs))
                except Exception:finish({'error':'Import raised an exception. Check this month before retrying.'},'failed / review');raise
                if 'application/x-ndjson' in response.content_type:
                    stream=response.response
                    def observed():
                        pending='';confirmed=False
                        try:
                            for chunk in stream:
                                pending+=chunk.decode('utf-8') if isinstance(chunk,bytes) else chunk
                                lines=pending.split('\n');pending=lines.pop()
                                for line in lines:
                                    try:event=json.loads(line)
                                    except ValueError:continue
                                    if event.get('type')=='done':finish(event);confirmed=True
                                yield chunk
                        finally:
                            if hasattr(stream,'close'):stream.close()
                            if not confirmed:finish({'error':'No final server confirmation. Verify the saved month before retrying.'},'unconfirmed / review')
                    response.response=observed();return response
                result=response.get_json(silent=True) or {'ok':False,'error':'No JSON confirmation. Verify this month.'}
                finish(result);return response
            return logged
        app.view_functions[endpoint]=decorate(original,kind,fields)
    @app.get('/imports/log')
    def import_log():
        user=getattr(g,'user',None) or {}
        if user.get('role')!='admin':return 'Admin access required.',403
        rows=audit.rows(user['id'])
        if request.args.get('format')=='csv':
            output=io.StringIO();writer=csv.writer(output);keys=['id','source_kind','source_file','receipt_month','sha256','started_at','finished_at','status','details'];writer.writerow(keys)
            for row in rows:writer.writerow(["'"+str(row[key]) if str(row[key]).startswith(('=','+','-','@')) else row[key] for key in keys])
            return Response(output.getvalue(),mimetype='text/csv',headers={'Content-Disposition':'attachment; filename=import_log.csv'})
        return render_template('import_log.html',rows=rows)
    return audit
