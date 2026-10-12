"""Scoped, paginated drill-down for current policy relation counts."""
import pandas as pd
from sqlalchemy import inspect
from flask import g,request,render_template,url_for

def register_current_members(app,get_engine,sql_text,scope_fn):
    @app.get('/policy-members')
    def current_relation_members():
        user=getattr(g,'user',None) or {}
        if user.get('role')!='admin':return 'Admin access required for member details.',403
        engine=get_engine();selected=request.args.get('franchise','All');relation=request.args.get('relation','All').upper();query=request.args.get('q','')[:100]
        rows=[];count=0;page=max(1,request.args.get('page',1,type=int));franchises=[]
        unavailable=engine is not None and not inspect(engine).has_table('app_policydata_current_members')
        if engine is not None and not unavailable:
            with engine.begin() as conn:
                franchises=[r[0] for r in conn.execute(sql_text('SELECT DISTINCT franchise_name FROM app_policydata_current_members WHERE active=true ORDER BY franchise_name'))]
                franchises=scope_fn(pd.DataFrame({'Franchise':franchises}))['Franchise'].tolist()
                if selected!='All' and selected not in franchises:return 'Choose an available franchise.',400
                params={f'f{i}':name for i,name in enumerate(franchises)}
                where=['active=true','franchise_name IN ('+','.join(':'+key for key in params)+')' if params else '1=0']
                if selected!='All':where.append('franchise_name=:franchise');params['franchise']=selected
                if relation!='ALL':where.append("UPPER(TRIM(COALESCE(relation,'')))=:relation");params['relation']=relation
                if query:where.append('(LOWER(member_name) LIKE :query OR LOWER(policy_number) LIKE :query)');params['query']='%'+query.lower()+'%'
                condition=' AND '.join(where)
                count=conn.execute(sql_text('SELECT COUNT(*) FROM app_policydata_current_members WHERE '+condition),params).scalar() or 0
                page=min(page,max(1,(count+49)//50));params.update(limit=50,offset=(page-1)*50)
                rows=[dict(row) for row in conn.execute(sql_text('SELECT franchise_name,relation,member_name,policy_number,id_number,age,age_band,import_month FROM app_policydata_current_members WHERE '+condition+' ORDER BY franchise_name,policy_number,member_name LIMIT :limit OFFSET :offset'),params).mappings()]
        args=request.args.to_dict()
        return render_template('current_members.html',rows=rows,count=count,page=page,selected=selected,relation=relation,franchises=franchises,unavailable=unavailable,
            previous=url_for('current_relation_members',**{**args,'page':page-1}) if page>1 else '',next_page=url_for('current_relation_members',**{**args,'page':page+1}) if page*50<count else '')
