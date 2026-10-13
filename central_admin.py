"""Explicit month controls; never infer a missing claims import means zero claims."""
import json,re,uuid
from datetime import datetime
from flask import g,request,redirect,url_for,abort

def register_admin(app,store):
    @app.post('/admin/recipients/set-scope')
    def central_set_recipient_scope():
        if (getattr(g,'user',None) or {}).get('role')!='admin':abort(403)
        identifier=request.form.get('user_id',type=int);kind=request.form.get('kind','');value=request.form.get('value','').strip()
        from app import get_user_by_id,log_audit
        recipient=get_user_by_id(identifier) if identifier else None
        expected={'underwriter':'underwriter','supplier':'voucher_supplier'}.get(kind)
        if not recipient or recipient.get('role')!=expected:abort(400,'Choose an existing account with the matching recipient role.')
        from direct_data import EXCEPTIONS,GROUP_NAME
        if kind=='underwriter' and value not in EXCEPTIONS+[GROUP_NAME,'AUL']:abort(400,'Choose a recognised insurer group.')
        if kind=='supplier' and value not in ['All vouchers']+[r['franchise'] for r in store.summaries()]:abort(400,'Choose a recognised voucher scope.')
        store.ensure()
        with store.direct.connection() as conn:
            store.direct.execute(conn,'DELETE FROM central_recipient_access WHERE user_id=:id',{'id':identifier})
            if request.form.get('revoke')!='yes':store.direct.execute(conn,'INSERT INTO central_recipient_access VALUES (:id,:kind,:value)',dict(id=identifier,kind=kind,value=value))
        log_audit('recipient_scope_updated',f'Claims recipient {identifier}; kind {kind}; scope {value}; revoke {request.form.get("revoke")=="yes"}')
        return redirect(url_for('overview_page',section='admin'))
    @app.post('/admin/payover/import-balances')
    def central_import_balances():
        user=getattr(g,'user',None) or {}
        if user.get('role')!='admin':abort(403)
        upload=request.files.get('file');month=request.form.get('month','');sheet=request.form.get('sheet','').strip()
        if not upload or not upload.filename.lower().endswith('.xlsx') or not sheet or not re.fullmatch(r'\d{4}-(0[1-9]|1[0-2])',month):abort(400,'Select an Excel workbook, its exact worksheet and the applicable month.')
        from pathlib import Path
        app.extensions['central_jobs'].create(user['id'],'ledger',{'filename':Path(upload.filename).name,'month':month,'sheet':sheet},upload)
        return redirect(url_for('overview_page',section='admin'))
    @app.post('/admin/payover/save-controls')
    def central_save_balance_controls():
        user=getattr(g,'user',None) or {}
        if user.get('role')!='admin':abort(403)
        month=request.form.get('month','');franchise=request.form.get('franchise','')
        if not re.fullmatch(r'\d{4}-(0[1-9]|1[0-2])',month) or franchise not in {r['franchise'] for r in store.summaries(month,month)}:abort(400,'Choose an imported month and franchise.')
        from direct_data import number
        from central_calculations import normal
        from central_ledger import controls,publish
        from app import log_audit
        try:
            record={key:str(number(request.form.get(key,''),key,0)) for key in ('opening','payment')}
            if any(request.form.get(key,'').strip()=='' for key in ('opening','payment')):raise ValueError('Enter verified opening balance and payment, including zero when applicable.')
            raw=request.form.get('system_control','').strip();record['system_control']=str(number(raw,'External System control',0)) if raw else None
            record.update(franchise=franchise,account=request.form.get('account','').strip(),paid_on=request.form.get('paid_on','').strip())
            if not record['account']:raise ValueError('Enter the verified account/reference number.')
            if record['paid_on']:datetime.strptime(record['paid_on'],'%Y-%m-%d')
        except ValueError as exc:abort(400,str(exc))
        previous=controls(store,month);previous[normal(franchise)]=record
        publish(store,uuid.uuid4().hex,user['id'],month,'Manual administrator balance/payment controls',list(previous.values()))
        log_audit('payover_controls_saved',f'Claims monthly control: {month}, {franchise}; existing franchises and archived versions retained')
        return redirect(url_for('overview_page',section='admin'))
    @app.context_processor
    def monthly_controls():
        if (getattr(g,'user',None) or {}).get('role')!='admin':return {}
        if request.endpoint not in ('overview_page','dashboard','workspace_page'):return {}
        if (request.view_args or {}).get('section')!='admin':return {}
        pubs=store.publications();verified=store.verified_claim_months()
        legacy=[b for b in store.direct.batches(g.user['id']) if b['id'] not in {p['id'] for p in pubs}]
        with store.direct.connection() as conn:
            legacy=[b for b in legacy if store.direct.execute(conn,'SELECT COUNT(*) FROM direct_data_rows WHERE batch_id=:id',{'id':b['id']}).fetchone()[0]]
        return {'central_admin_months':pubs,'central_verified_claim_months':verified,'central_legacy_convertible':legacy,'central_control_franchises':sorted({r['franchise'] for r in store.summaries()})}

    @app.post('/admin/claims/confirm-zero')
    def central_confirm_zero_claims():
        user=getattr(g,'user',None) or {}
        if user.get('role')!='admin':abort(403)
        month=request.form.get('month','')
        if not re.fullmatch(r'\d{4}-(0[1-9]|1[0-2])',month) or request.form.get('confirm')!='yes':abort(400)
        store.ensure();identifier=uuid.uuid4().hex;stamp=datetime.utcnow().isoformat()
        with store.direct.connection() as conn:
            if store.direct.get_engine:store.direct.execute(conn,'SELECT pg_advisory_xact_lock(hashtext(:key))',{'key':'claims:'+month})
            existing=store.direct.execute(conn,'SELECT COUNT(*) FROM central_claims c JOIN central_claim_publications p ON p.batch_id=c.batch_id AND p.month=c.month WHERE p.month=:month',{'month':month}).fetchone()[0]
            if existing:abort(409,'This month has published claims. A zero confirmation cannot replace paid transactions.')
            store.direct.execute(conn,'INSERT INTO central_claim_batches VALUES (:id,:source,:owner,:created,:amount,:status)',dict(id=identifier,source='Explicit administrator zero-claims confirmation',owner=user['id'],created=stamp,amount='Net amount paid via bank',status='Confirmed no paid claims'))
            store.direct.execute(conn,'INSERT INTO central_claim_publications VALUES (:month,:batch) ON CONFLICT(month) DO UPDATE SET batch_id=excluded.batch_id',dict(month=month,batch=identifier))
        return redirect(url_for('overview_page',section='admin'))

    @app.post('/admin/direct/convert-saved')
    def central_convert_saved():
        user=getattr(g,'user',None) or {}
        if user.get('role')!='admin':abort(403)
        batch=next((b for b in store.direct.batches(user['id']) if b['id']==request.form.get('batch')),None)
        if not batch:abort(404)
        if any(p['month']==batch['month'] for p in store.publications()):abort(409,'This month already has a current calculation version. Import its original source to replace it.')
        with store.direct.connection() as conn:
            if not store.direct.execute(conn,'SELECT COUNT(*) FROM direct_data_rows WHERE batch_id=:id',{'id':batch['id']}).fetchone()[0]:abort(409,'No archived source records to convert.')
        app.extensions['central_jobs'].create(user['id'],'migrate',{'source_batch':batch['id'],'source_owner':user['id'],'filename':batch['source_file'],'month':batch['month']})
        return redirect(url_for('overview_page',section='admin'))
