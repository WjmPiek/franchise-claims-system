"""Shared financial pages and report catalogue with backend scope checks."""
import calendar,json,uuid
from pathlib import Path
from flask import g,request,render_template,url_for,send_file,jsonify,has_request_context,redirect
from central_calculations import VERSION,retail_rank
from central_reports import catalog,generate

def financial_result(app,legacy):
    if not has_request_context() or not getattr(g,'user',None):return legacy
    if hasattr(g,'central_financial_result'):return g.central_financial_result
    store=app.extensions.get('central_store')
    if store is None:return legacy
    frame=store.frame(legacy.get('monthly'))
    if frame.empty:return legacy
    import pandas as pd
    frame=frame.sort_values(['Franchise','Month']);frame['BrightRock Running Balance']=frame.groupby('Franchise')['BrightRock Month Total'].cumsum()
    frame.attrs['central_dataset']=True
    import app as application
    periods,portfolio=application.rebuild_periods_and_portfolio(frame)
    for target in (frame,periods):
        if not target.empty:target['Recommendation']='Review — recommendation rule requires confirmation'
    g.central_financial_result={**legacy,'monthly':frame,'periods':periods,'portfolio':portfolio,
      'rates':{'BrightRock':10.0,'Inkulu':2.5,'MFF':2.5},'book_rates':{'MFF Book Value':2.5,'Franchise Book Value':2.5}}
    return g.central_financial_result

def register_central(app,export_dir,scope):
    store=app.extensions['central_store'];root=Path(export_dir)/'central_reports';root.mkdir(parents=True,exist_ok=True)
    def access():
        user=getattr(g,'user',None) or {}
        if not user:return None
        if user.get('role')=='admin':return None
        if user.get('role') in ('underwriter','voucher_supplier'):
            store.ensure()
            with store.direct.connection() as conn:
                grants=store.direct.records(store.direct.execute(conn,'SELECT kind,value FROM central_recipient_access WHERE user_id=:id',{'id':user['id']}))
                if user['role']=='underwriter':
                    destinations={r['value'] for r in grants if r['kind']=='underwriter'}
                    names=set()
                    for destination in destinations:
                        rows=store.direct.execute(conn,'SELECT DISTINCT r.franchise FROM central_records r JOIN central_publications p ON p.batch_id=r.batch_id WHERE r.destination=:dest',{'dest':destination})
                        names.update(row[0] for row in rows)
                    return names
                values={r['value'] for r in grants if r['kind']=='supplier'}
                return {r['franchise'] for r in store.summaries()} if 'All vouchers' in values else values
        import pandas as pd
        names=sorted({r['franchise'] for r in store.summaries()})
        frame=scope(pd.DataFrame({'Franchise':names}))
        return set(frame['Franchise']) if 'Franchise' in frame else set()
    app.extensions['central_scope']=access
    from claim_insights import register as register_claim_insights
    register_claim_insights(app,store,access)

    def export_arguments(batch,item,allowed):
        from central_reports import REPORT_FORMAT_VERSION
        if app.extensions.get('allocation_audit_path')!=str(store.direct.path):
            from allocation_repair import ensure
            ensure(store);app.extensions['allocation_audit_path']=str(store.direct.path)
        args={'batch':batch['id'],'name':item['name'],'allowed':sorted(allowed) if allowed is not None else None,'report_format_version':REPORT_FORMAT_VERSION}
        with store.direct.connection() as conn:
            if item['kind'] in ('claims','product_claims','internal','payover','branch_calculation'):
                row=store.direct.execute(conn,'SELECT batch_id FROM central_claim_publications WHERE month=:month',{'month':batch['month']}).fetchone()
                args['claims_version']=row[0] if row else None
            if item['kind']=='payover':
                from central_ledger import ensure
                ensure(store)
                row=store.direct.execute(conn,'SELECT batch_id FROM central_ledger_publications WHERE month=:month',{'month':batch['month']}).fetchone()
                args['ledger_version']=row[0] if row else None
            row=store.direct.execute(conn,"SELECT COALESCE(MAX(changed_at),'') FROM central_allocation_audit WHERE batch_id IN (:policy,:claims)",{'policy':batch['id'],'claims':args.get('claims_version','')}).fetchone()
            args['allocation_revision']=row[0]
        return args

    def allowed_catalog(batch,allowed):
        user=getattr(g,'user',None) or {};items=catalog(store,batch,allowed)
        jobs=app.extensions.get('central_jobs')
        if jobs:
            jobs.ensure()
            with store.direct.connection() as conn:
                rows=store.direct.records(store.direct.execute(conn,"SELECT args,state,result,error FROM central_jobs WHERE owner_id=:owner AND kind='export' ORDER BY created_at",{'owner':user.get('id',0)}))
            states={}
            for row in rows:
                args=json.loads(row['args'])
                item=next((i for i in items if i['name']==args.get('name')),None)
                if item and args==export_arguments(batch,item,allowed):states[args.get('name')]=row
            for item in items:
                if item['name'] in states:
                    state=states[item['name']];item['status']={'completed':'Ready for download','failed':'Needs correction','queued':'Queued','processing':'Generating'}.get(state['state'],state['state'])
                    if state['state']=='completed':item['generated_rows']=json.loads(state['result']).get('metadata',{}).get('rows')
                    if state['state']=='failed':item['error']=state['error']
        if user.get('role')=='voucher_supplier':return [i for i in items if i['kind']=='supplier']
        if user.get('role')=='underwriter':
            with store.direct.connection() as conn:
                destinations={r[0] for r in store.direct.execute(conn,"SELECT value FROM central_recipient_access WHERE user_id=:id AND kind='underwriter'",{'id':user['id']})}
            return [i for i in items if i['kind']=='policy' and i['destination'] in destinations]
        return items
    def permitted_export(args):
        allowed=access()
        if args.get('allowed')!=(sorted(allowed) if allowed is not None else None):return False
        batch=next((b for b in store.publications() if b['id']==args.get('batch')),None)
        return bool(batch and any(i['name']==args.get('name') and args==export_arguments(batch,i,allowed) for i in allowed_catalog(batch,allowed)))
    app.extensions['central_export_access']=permitted_export
    from report_workflow import register as register_report_workflow
    register_report_workflow(app)

    def selected():
        from period_selection import selected_dates,month_bounds
        return month_bounds(selected_dates(request.args))

    @app.context_processor
    def central_context():
        if not getattr(g,'user',None) or request.endpoint not in ('overview_page','dashboard','workspace_page'):return {}
        start,end=selected();allowed=access();pubs=[b for b in store.publications() if start<=b['month']<=end]
        selected_franchise=request.args.get('franchise','All')
        if selected_franchise!='All':allowed={selected_franchise} if allowed is None or selected_franchise in allowed else set()
        summaries=store.summaries(start,end,allowed)
        insights={}
        if (request.view_args or {}).get('section')=='claims' or request.args.get('section')=='claims':
            from claim_insights import analyse
            from period_selection import selected_dates
            insights={'claim_insights':analyse(store,selected_dates(request.args),allowed)}
        import app as application
        from premium_comparison import comparison
        current=application.apply_user_franchise_scope(application.effective_result()['monthly'])
        historical=application.apply_user_franchise_scope(application.LAST_RESULT.get('monthly',__import__('pandas').DataFrame()))
        if allowed is not None:
            if not current.empty:current=current[current['Franchise'].isin(allowed)]
            if not historical.empty:historical=historical[historical['Franchise'].isin(allowed)]
        premium_comparison=comparison(current,historical,start,end,selected_franchise,application.build_premium_parlour_analysis)
        if not summaries:return {'central_batches':pubs,'central_ready':False,'central_premium_comparison':premium_comparison,**insights}
        from decimal import Decimal
        totals={}
        for row in summaries:
            for key,value in row['summary'].items():totals[key]=totals.get(key,Decimal(0))+Decimal(str(value))
        totals['claims']=Decimal(0);totals['claim_count']=0;totals['claim_net_off']=Decimal(0);totals['claim_bank_paid']=Decimal(0)
        claims_by_franchise={}
        for claim in store.claims(start,end,allowed):
            totals['claims']+=Decimal(claim['amount']);totals['claim_count']+=int(Decimal(claim['amount'])!=0)
            totals['claim_net_off']+=Decimal(claim['net_off']);totals['claim_bank_paid']+=Decimal(claim['bank_paid'])
            item=claims_by_franchise.setdefault(claim['franchise'],dict(franchise=claim['franchise'],claims=Decimal(0),net_off=Decimal(0),bank_paid=Decimal(0),count=0))
            item['claims']+=Decimal(claim['amount']);item['net_off']+=Decimal(claim['net_off']);item['bank_paid']+=Decimal(claim['bank_paid']);item['count']+=int(Decimal(claim['amount'])!=0)
        verified_claims=store.verified_claim_months()
        totals['scenario2']=totals.get('retail',Decimal(0))-totals['claims']-totals.get('paid_commissions',Decimal(0))
        from datetime import datetime
        from dateutil.relativedelta import relativedelta
        previous_start=(datetime.strptime(start,'%Y-%m')-relativedelta(years=1)).strftime('%Y-%m');previous_end=(datetime.strptime(end,'%Y-%m')-relativedelta(years=1)).strftime('%Y-%m')
        def retail(rows):
            output={}
            for row in rows:output[row['franchise']]=output.get(row['franchise'],Decimal(0))+Decimal(row['summary']['source_retail'])
            return output
        rankings=retail_rank(retail(store.summaries(start,end)),retail(store.summaries(previous_start,previous_end)))
        if allowed is not None:rankings=[r for r in rankings if r['franchise'] in allowed]
        expected=[];year,month=map(int,start.split('-'));last=tuple(map(int,end.split('-')))
        while (year,month)<=last:
            expected.append(f'{year:04d}-{month:02d}');month+=1
            if month==13:year+=1;month=1
        present={b['month'] for b in pubs}
        from central_reports import COMBINE
        from central_calculations import normal
        combined=request.args.get('aul_view')=='combined';aul_groups={}
        for row in summaries:
            value=row['summary']
            if not value.get('aul_rows',0):continue
            name=COMBINE.get(normal(row['franchise']),row['franchise']) if combined else row['franchise']
            item=aul_groups.setdefault(name,dict(franchise=name,rows=0,payover=Decimal(0)))
            item['rows']+=int(value['aul_rows']);item['payover']+=Decimal(value['aul_payover'])
        return {**insights,'central_ready':True,'central_totals':totals,'central_batches':pubs,
          'central_premium_comparison':premium_comparison,
          'business_branches':__import__('business_overview').branch_totals(summaries,claims_by_franchise.values()),
          'central_rankings':rankings,'central_missing_months':[m for m in expected if m not in present],
          'central_catalog':[{'batch':{**b,'source_rows':sum(r['summary']['rows'] for r in summaries if r['month']==b['month'])},'files':allowed_catalog(b,allowed)} for b in pubs],
          'central_source_version':VERSION,'central_identity_pending':False,
          'central_members':__import__('central_identity').member_snapshot(store,start,end,allowed),
          'central_product_retail':__import__('central_identity').products_by_retail(store,start,end,allowed),
          'central_voucher_ranking':sorted([dict(franchise=r['franchise'],units=sum(Decimal(v['summary']['voucher_units']) for v in summaries if v['franchise']==r['franchise'])) for r in {v['franchise']:v for v in summaries}.values()],key=lambda r:(-r['units'],r['franchise'].casefold())),
          'central_graphs':__import__('central_graphs').graphs(store,start,end,allowed) if request.args.get('section')=='graphs' or (request.view_args or {}).get('section')=='graphs' else [],
          'central_claims_verified':all(b['month'] in verified_claims for b in pubs),
          'central_claim_branches':sorted(claims_by_franchise.values(),key=lambda row:row['franchise'].casefold()),
          'central_aul_groups':sorted(aul_groups.values(),key=lambda row:row['franchise'].casefold()),'central_aul_combined':combined,
          'central_franchise_months':summaries}

    @app.get('/calculations')
    def central_calculation_detail():
        user=getattr(g,'user',None) or {}
        if not user:return 'Authentication required.',401
        if user.get('role') in ('underwriter','voucher_supplier'):return redirect(url_for('central_catalog_page'))
        start,end=selected();allowed=access();franchise=request.args.get('franchise','All')
        if franchise!='All':allowed={franchise} if allowed is None or franchise in allowed else set()
        category=request.args.get('category','source');page=max(1,request.args.get('page',1,type=int));size=50
        if category=='active':
            months=[b['month'] for b in store.publications() if start<=b['month']<=end]
            if months:start=end=max(months)
        store.ensure();params={'start':start,'end':end};clauses=['p.month BETWEEN :start AND :end']
        review_predicate=r'''r.record_json !~ '"issues"[[:space:]]*:[[:space:]]*\[[[:space:]]*\]' ''' if store.direct.get_engine else "json_array_length(json_extract(r.record_json,'$.issues'))>0"
        predicates={'review':review_predicate,'aul':"r.destination='AUL'",'plan2':"r.plan='Plan 2'",'historic':"r.plan='Historic'",'unclassified':"r.plan='Unclassified'",'vouchers':"r.destination='Vouchers'",'insurers':"r.destination<>'Vouchers'",'fees':"r.relation='MEM' AND r.destination<>'Vouchers'"}
        if category in predicates:clauses.append(predicates[category])
        if category=='commissions':clauses.append("r.destination<>'Vouchers'")
        if category=='faithnet':clauses.append("r.relation='MEM' AND LOWER(r.franchise) LIKE '%faithnet%'")
        if category=='active':clauses.append("r.relation='MEM' AND r.destination<>'Vouchers'")
        if allowed is not None:
            if not allowed:clauses.append('1=0')
            else:
                keys=[]
                for i,name in enumerate(sorted(allowed)):key='f'+str(i);keys.append(':'+key);params[key]=name
                clauses.append('r.franchise IN ('+','.join(keys)+')')
        query=request.args.get('q','').strip()[:150]
        if category=='claims':
            claim_params={k:v for k,v in params.items() if not k.startswith('limit') and not k.startswith('offset')}
            claim_clauses=[c.replace('p.month','c.month').replace('r.franchise','c.franchise') for c in clauses if c not in predicates.values()]
            if query:claim_clauses.append('LOWER(c.source_json) LIKE :query');claim_params['query']='%'+query.lower()+'%'
            claim_base=' FROM central_claims c JOIN central_claim_publications p ON p.batch_id=c.batch_id AND p.month=c.month WHERE '+' AND '.join(claim_clauses)
            with store.direct.connection() as conn:
                aggregate='SUM(CAST(c.amount AS NUMERIC))' if store.direct.get_engine else 'decimal_sum(c.amount)'
                control=store.direct.execute(conn,'SELECT COUNT(*),COALESCE('+aggregate+',0),SUM(CASE WHEN CAST(c.amount AS NUMERIC)<>0 THEN 1 ELSE 0 END)'+claim_base,claim_params).fetchone()
                claim_params.update(limit=size,offset=(page-1)*size)
                records=store.direct.records(store.direct.execute(conn,"SELECT c.n,c.month,c.franchise,c.product,c.amount,c.issues,c.source_json,COALESCE(x.net_off,'0') AS net_off,COALESCE(x.bank_paid,'0') AS bank_paid"+claim_base.replace(' WHERE ', ' LEFT JOIN central_claim_components x ON x.batch_id=c.batch_id AND x.n=c.n WHERE ')+' ORDER BY c.month DESC,c.n LIMIT :limit OFFSET :offset',claim_params))
            from central_claims import source_paid_on,component
            for record in records:
                source=json.loads(record.pop('source_json'));record['paid_on']=source_paid_on(source);record['final_reference']=str(component(source,'Final Amount'))
            from decimal import Decimal
            net_off=Decimal(0);bank_paid=Decimal(0)
            for c in store.claims(start,end,allowed):net_off+=Decimal(c['net_off']);bank_paid+=Decimal(c['bank_paid'])
            if query:
                # Keep displayed control totals subject to the same search predicate.
                with store.direct.connection() as conn:
                    aggregate='SUM(CAST(x.{field} AS NUMERIC))' if store.direct.get_engine else 'decimal_sum(x.{field})'
                    joined=claim_base.replace(' WHERE ',' LEFT JOIN central_claim_components x ON x.batch_id=c.batch_id AND x.n=c.n WHERE ')
                    amounts=store.direct.execute(conn,'SELECT COALESCE('+aggregate.format(field='net_off')+',0),COALESCE('+aggregate.format(field='bank_paid')+',0)'+joined,claim_params).fetchone();net_off,bank_paid=amounts
            return render_template('central_claim_detail.html',records=records,count=control[0],claim_count=control[2] or 0,total=control[1],net_off=net_off,bank_paid=bank_paid,page=page,date_from=g.period_dates['date_from'],date_to=g.period_dates['date_to'],franchise=franchise,query=query)
        if query:clauses.append('LOWER(r.record_json) LIKE :query');params['query']='%'+query.lower()+'%'
        base=' FROM central_records r JOIN central_publications p ON p.batch_id=r.batch_id WHERE '+' AND '.join(clauses)
        with store.direct.connection() as conn:
            count=store.direct.execute(conn,'SELECT COUNT(*)'+base,params).fetchone()[0]
            params.update(limit=size,offset=(page-1)*size)
            result=store.direct.execute(conn,'SELECT r.record_json,p.month'+base+' ORDER BY p.month DESC,r.n LIMIT :limit OFFSET :offset',params)
            records=[{**json.loads(row[0]),'month':row[1]} for row in result]
        return render_template('central_calculations.html',records=records,count=count,page=page,category=category,
            date_from=g.period_dates['date_from'],date_to=g.period_dates['date_to'],franchise=franchise,query=query)

    @app.get('/reports/catalog')
    def central_catalog_page():
        if not getattr(g,'user',None):return 'Authentication required.',401
        start,end=selected();allowed=access()
        summaries=store.summaries(start,end,allowed)
        return render_template('central_reports.html',entries=[{'batch':{**b,'source_rows':sum(r['summary']['voucher_rows'] if g.user.get('role')=='voucher_supplier' else r['summary']['rows'] for r in summaries if r['month']==b['month'])},'files':allowed_catalog(b,allowed)} for b in store.publications() if start<=b['month']<=end])

    @app.get('/reports/central/<batch_id>/<int:index>')
    def central_report(batch_id,index):
        if not getattr(g,'user',None):return 'Authentication required.',401
        allowed=access();batch=next((b for b in store.publications() if b['id']==batch_id),None)
        if not batch:return 'This source version is superseded or unavailable. Select the current published month.',409
        items=allowed_catalog(batch,allowed)
        if not 0<=index<len(items):return 'Report not available in your scope.',404
        item=items[index]
        if allowed is not None and not allowed:return 'No authorised franchise data.',403
        # A scoped key prevents another user's generated file from being reused.
        key=uuid.uuid5(uuid.NAMESPACE_URL,json.dumps([batch_id,item['name'],sorted(allowed) if allowed is not None else 'admin'],ensure_ascii=False)).hex
        folder=root/key;path=folder/item['name'];meta=folder/'metadata.json'
        jobs=app.extensions.get('central_jobs')
        if jobs:
            args=export_arguments(batch,item,allowed)
            jobs.ensure()
            with store.direct.connection() as conn:
                existing=store.direct.records(store.direct.execute(conn,"SELECT id,state,result FROM central_jobs WHERE owner_id=:owner AND kind='export' AND args=:args AND state<>'failed' ORDER BY created_at DESC LIMIT 1",dict(owner=g.user['id'],args=json.dumps(args))))
            if existing:
                job=existing[0]
                if job['state']=='completed':return redirect(url_for('central_report_download',identifier=job['id'],format=request.args.get('format','excel')))
                identifier=job['id']
            else:identifier=jobs.create(g.user['id'],'export',args)
            if request.args.get('prepare')=='json':return jsonify(ok=True,job=identifier),202
            return render_template('central_job_wait.html',identifier=identifier,title=item['title'],format=request.args.get('format','excel'))
        if not path.exists():
            try:
                details=generate(store,batch,item,path,allowed);meta.write_text(json.dumps(details),encoding='utf-8')
            except ValueError as exc:return str(exc),409
        if request.args.get('format')=='view':
            from report_files import workbook_preview
            preview=workbook_preview(path,request.args.get('sheet',''),request.args.get('page',1,type=int))
            def page_url(n):return url_for('central_report',batch_id=batch_id,index=index,format='view',sheet=preview['sheet'],page=n)
            return render_template('workbook_report.html',title=item['title'],preview=preview,
                excel_url=url_for('central_report',batch_id=batch_id,index=index),
                pdf_url=url_for('central_report',batch_id=batch_id,index=index,format='pdf'),
                previous=page_url(preview['page']-1) if preview['page']>1 else '',
                next_page=page_url(preview['page']+1) if preview['page']*50<preview['count'] else '')
        if request.args.get('format')=='pdf':
            from report_files import workbook_pdf
            return send_file(workbook_pdf(path,item['title']),as_attachment=True,download_name=path.stem+'.pdf',mimetype='application/pdf')
        return send_file(path,as_attachment=True,download_name=path.name,max_age=0)

    @app.get('/jobs/<identifier>/download')
    def central_report_download(identifier):
        user=getattr(g,'user',None) or {}
        if not user:return 'Authentication required.',401
        jobs=app.extensions['central_jobs'];status=jobs.status(identifier,user['id'])
        if not status or status['kind']!='export':return 'Report not found.',404
        if status['state']!='completed':return 'Report is not ready.',409
        result=status['result'];batch=next((b for b in store.publications() if b['id']==result['batch_id']),None)
        if batch is None:return 'This source version has been replaced. Download the current published report.',409
        allowed=access()
        with store.direct.connection() as conn:
            raw=store.direct.execute(conn,'SELECT args FROM central_jobs WHERE id=:id AND owner_id=:owner',dict(id=identifier,owner=user['id'])).fetchone()
        args=json.loads(raw[0])
        if not permitted_export(args):return 'Your access scope changed. Generate a fresh authorised report.',403
        path=jobs.export_path(identifier,result);format=request.args.get('format','excel')
        if format=='view':
            from report_files import workbook_preview
            preview=workbook_preview(path,request.args.get('sheet',''),request.args.get('page',1,type=int))
            def page_url(n):return url_for('central_report_download',identifier=identifier,format='view',sheet=preview['sheet'],page=n)
            return render_template('workbook_report.html',title=result['name'],preview=preview,
              excel_url=url_for('central_report_download',identifier=identifier),pdf_url=url_for('central_report_download',identifier=identifier,format='pdf'),
              previous=page_url(preview['page']-1) if preview['page']>1 else '',next_page=page_url(preview['page']+1) if preview['page']*50<preview['count'] else '')
        if format=='pdf':
            from report_files import workbook_pdf
            return send_file(workbook_pdf(path,result['name']),as_attachment=True,download_name=Path(result['name']).stem+'.pdf',mimetype='application/pdf')
        return send_file(path,as_attachment=True,download_name=result['name'],max_age=0)
