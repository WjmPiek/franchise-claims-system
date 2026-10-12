"""Durable Direct-only monthly data, dashboard summaries and paginated details."""
import csv
import io
import json
import re
import sqlite3
import threading
from contextlib import contextmanager
from datetime import datetime
from decimal import Decimal
from pathlib import Path
from flask import g, request, render_template, url_for
from sqlalchemy import text

FIELDS = ['n','destination','plan','franchise','policy','relation','original_risk','retail',
          'faithnet','r1','adv','payover','supplier','retained','contact','data','issues']
MONEY = FIELDS[6:14]
CATEGORIES = {'source':'All Direct source rows', 'insurers':'Insurer payover', 'aul':'AUL data',
              'plan2':'AUL Plan 2', 'historic':'AUL Historic', 'faithnet':'Faithnet MEM adjustments',
              'fees':'MEM fees', 'vouchers':'Vouchers', 'review':'Records to review'}

def data_month(filename, chosen=''):
    match = re.search(r'(\d{8})_to_(\d{8})', filename)
    inferred = ''
    if match:
        start, end = [datetime.strptime(v,'%Y%m%d') for v in match.groups()]
        if start.strftime('%Y-%m') != end.strftime('%Y-%m'):
            raise ValueError('Use a source file for one month only.')
        inferred = start.strftime('%Y-%m')
    if chosen and not re.fullmatch(r'\d{4}-(0[1-9]|1[0-2])', chosen):
        raise ValueError('Choose a valid Direct data month.')
    if chosen:
        datetime.strptime(chosen+'-01','%Y-%m-%d')
    if chosen and inferred and chosen != inferred:
        raise ValueError('The selected month does not match the source filename.')
    if not (chosen or inferred):
        raise ValueError('Choose the Direct data month; it is not present in the filename.')
    return chosen or inferred

def metric():
    return {'rows':0, **{key:Decimal(0) for key in MONEY}}

def add_metric(target, row):
    target['rows'] += 1
    for key in MONEY:
        target[key] += Decimal(str(row[key]))

def encode(value):
    return json.dumps(value, default=str, ensure_ascii=False)

class DirectStore:
    def __init__(self, path, get_engine=None):
        self.path = Path(path)
        self.get_engine = get_engine
        self.ready = False
        self.schema_lock = threading.Lock()

    @contextmanager
    def connection(self):
        if self.get_engine is not None:
            engine = self.get_engine()
            if engine is None:
                raise RuntimeError('Direct analytics database is unavailable.')
            with engine.begin() as conn:
                yield conn
        else:
            conn = sqlite3.connect(self.path)
            conn.row_factory = sqlite3.Row
            class DecimalSum:
                def __init__(self):self.total=Decimal(0)
                def step(self,value):
                    if value is not None:self.total+=Decimal(str(value))
                def finalize(self):return str(self.total)
            conn.create_aggregate('decimal_sum',1,DecimalSum)
            try:
                yield conn
                conn.commit()
            except Exception:
                conn.rollback()
                raise
            finally:
                conn.close()

    def execute(self, conn, sql, params=None):
        return conn.execute(text(sql),params or {}) if self.get_engine else conn.execute(sql,params or {})

    def records(self, result):
        return [dict(r) for r in (result.mappings() if self.get_engine else result)]

    def ensure(self):
        with self.schema_lock:
            self._ensure_schema()

    def _ensure_schema(self):
        if self.ready:
            return
        with self.connection() as conn:
            self.execute(conn, 'CREATE TABLE IF NOT EXISTS direct_data_batches (id TEXT PRIMARY KEY, owner_id INTEGER NOT NULL, month TEXT NOT NULL, source_file TEXT NOT NULL, created_at TEXT NOT NULL, summary TEXT NOT NULL, UNIQUE(owner_id,month))')
            money_type = 'NUMERIC' if self.get_engine else 'TEXT'
            columns = ','.join(f'{key} {money_type} NOT NULL' for key in MONEY)
            self.execute(conn, 'CREATE TABLE IF NOT EXISTS direct_data_rows (batch_id TEXT NOT NULL REFERENCES direct_data_batches(id) ON DELETE CASCADE, n INTEGER NOT NULL, destination TEXT NOT NULL, plan TEXT NOT NULL, franchise TEXT NOT NULL, policy TEXT NOT NULL, relation TEXT NOT NULL, '+columns+', contact TEXT NOT NULL, data TEXT NOT NULL, issues TEXT NOT NULL, PRIMARY KEY(batch_id,n))')
            self.execute(conn, 'CREATE INDEX IF NOT EXISTS direct_data_rows_scope ON direct_data_rows(batch_id,destination,plan)')
            self.execute(conn, 'CREATE INDEX IF NOT EXISTS direct_data_rows_franchise ON direct_data_rows(batch_id,franchise)')
        self.ready = True

    def save(self, stage, batch_id, owner, month, filename):
        self.ensure()
        summaries = {'categories':{k:metric() for k in CATEGORIES}, 'destinations':{}, 'franchises':{}}
        stage.row_factory = sqlite3.Row
        for raw in stage.execute('SELECT * FROM prepared ORDER BY n'):
            row = dict(raw)
            categories = ['source']
            if row['destination'] == 'Vouchers': categories.append('vouchers')
            else: categories.append('insurers')
            if row['destination'] == 'AUL':
                categories.extend(['aul', 'plan2' if row['plan']=='Plan 2' else 'historic'])
            if Decimal(row['faithnet']) != 0: categories.append('faithnet')
            if row['relation']=='MEM' and row['destination']!='Vouchers': categories.append('fees')
            if json.loads(row['issues']): categories.append('review')
            for key in categories: add_metric(summaries['categories'][key],row)
            add_metric(summaries['destinations'].setdefault(row['destination'],metric()),row)
            add_metric(summaries['franchises'].setdefault(row['franchise'],metric()),row)
        with self.connection() as conn:
            if self.get_engine:
                self.execute(conn, 'SELECT pg_advisory_xact_lock(hashtext(:key))', {'key':f'direct:{owner}:{month}'})
            old = self.records(self.execute(conn,'SELECT id FROM direct_data_batches WHERE owner_id=:owner AND month=:month',{'owner':owner,'month':month}))
            for batch in old:
                self.execute(conn,'DELETE FROM direct_data_rows WHERE batch_id=:id',batch)
                self.execute(conn,'DELETE FROM direct_data_batches WHERE id=:id',batch)
            self.execute(conn,'INSERT INTO direct_data_batches VALUES (:id,:owner,:month,:source,:created,:summary)',
                         {'id':batch_id,'owner':owner,'month':month,'source':filename,
                          'created':datetime.utcnow().isoformat(),'summary':encode(summaries)})
            cursor = stage.execute('SELECT '+','.join(FIELDS)+' FROM prepared ORDER BY n')
            while True:
                rows = cursor.fetchmany(2000)
                if not rows: break
                values = [(batch_id,*tuple(r)) for r in rows]
                if self.get_engine:
                    buffer = io.StringIO()
                    csv.writer(buffer, quoting=csv.QUOTE_ALL).writerows(values)
                    buffer.seek(0)
                    with conn.connection.cursor() as raw_cursor:
                        raw_cursor.copy_expert('COPY direct_data_rows (batch_id,'+','.join(FIELDS)+') FROM STDIN WITH (FORMAT CSV)',buffer)
                else:
                    conn.executemany('INSERT INTO direct_data_rows VALUES ('+','.join('?' for _ in range(len(FIELDS)+1))+')',values)
        return summaries

    def batches(self, owner):
        self.ensure()
        with self.connection() as conn:
            rows = self.records(self.execute(conn,'SELECT id,month,source_file,created_at,summary FROM direct_data_batches WHERE owner_id=:owner ORDER BY month DESC',{'owner':owner}))
        for row in rows: row['summary'] = json.loads(row['summary'])
        return rows

    def overview(self, owner, month=None):
        batches = self.batches(owner)
        months = [r['month'] for r in batches]
        selected = month or (months[0] if months else '')
        if selected not in months and selected not in ('all',''):
            raise ValueError('Choose an available Direct data month.')
        chosen = [r for r in batches if selected=='all' or r['month']==selected]
        categories = {k:metric() for k in CATEGORIES}
        destinations,franchises = {},{}
        for batch in chosen:
            for collection,target in [('categories',categories),('destinations',destinations),('franchises',franchises)]:
                for key,values in batch['summary'][collection].items():
                    total = target.setdefault(key,metric())
                    total['rows'] += values['rows']
                    for field in MONEY: total[field] += Decimal(values[field])
        return {'months':months,'month':selected,'batches':chosen,'categories':categories,
                'destinations':destinations,'franchises':franchises,'available':bool(chosen)}

    def detail(self, owner, month, category, destination='', franchise='', query='', page=1):
        if category not in CATEGORIES: raise ValueError('Choose a valid detail category.')
        overview = self.overview(owner,month)
        params = {'owner':owner}
        where = ['b.owner_id=:owner']
        if overview['month']!='all':
            where.append('b.month=:month');params['month']=overview['month']
        predicates = {'insurers':"r.destination<>'Vouchers'",'vouchers':"r.destination='Vouchers'",
                      'aul':"r.destination='AUL'",'plan2':"r.destination='AUL' AND r.plan='Plan 2'",
                      'historic':"r.destination='AUL' AND r.plan='Historic'",
                      'faithnet':'CAST(r.faithnet AS NUMERIC)<>0',
                      'fees':"r.relation='MEM' AND r.destination<>'Vouchers'",'review':"r.issues<>'[]'"}
        if category in predicates:where.append('('+predicates[category]+')')
        for key,value in [('destination',destination),('franchise',franchise)]:
            if value:
                where.append('r.'+key+'=:'+key);params[key]=value
        if query:
            where.append('(LOWER(r.policy) LIKE :query OR LOWER(r.data) LIKE :query)')
            params['query']='%'+query.lower()[:100]+'%'
        base = ' FROM direct_data_rows r JOIN direct_data_batches b ON b.id=r.batch_id WHERE '+' AND '.join(where)
        with self.connection() as conn:
            aggregate='SUM' if self.get_engine else 'decimal_sum'
            columns=', '.join('COALESCE('+aggregate+'(r.'+field+'),0) AS '+field for field in MONEY)
            totals=self.records(self.execute(conn,'SELECT COUNT(*) AS rows, '+columns+base,params))[0]
            count=int(totals['rows'])
            overview['detail_totals']={key:Decimal(str(totals[key])) for key in MONEY}
            page = max(1,min(page,max(1,(count+49)//50)))
            params.update({'limit':50,'offset':(page-1)*50})
            rows = self.records(self.execute(conn,'SELECT r.*,b.month,b.source_file'+base+' ORDER BY b.month DESC,r.n LIMIT :limit OFFSET :offset',params))
        for row in rows:
            row['data']=json.loads(row['data']);row['issues']=json.loads(row['issues'])
        return overview,rows,count,page

    def row(self, owner, batch_id, row_number):
        self.ensure()
        with self.connection() as conn:
            rows = self.records(self.execute(conn,'SELECT r.*,b.month,b.source_file FROM direct_data_rows r JOIN direct_data_batches b ON b.id=r.batch_id WHERE b.owner_id=:owner AND r.batch_id=:batch AND r.n=:n',{'owner':owner,'batch':batch_id,'n':row_number}))
        if not rows:return None
        row=rows[0];row['data']=json.loads(row['data']);row['issues']=json.loads(row['issues'])
        return row

def register_direct_analytics(app, store):
    @app.template_filter('direct_number')
    def direct_number(value):
        whole, dot, fraction = format(Decimal(str(value)),',f').partition('.')
        return whole+'.'+fraction.rstrip('0').ljust(2,'0')

    def owner():
        user = getattr(g,'user',None) or {}
        return user.get('id') if user.get('role')=='admin' else None

    @app.context_processor
    def direct_dashboard_context():
        if not owner() or request.endpoint not in ('dashboard','workspace_page','direct_data'):
            return {}
        try:return {'direct_analytics_data':store.overview(owner(),request.args.get('direct_month'))}
        except Exception as exc:
            app.logger.warning('Direct analytics unavailable (%s)',type(exc).__name__)
            return {'direct_analytics_data':{'available':False,'error':'Direct analytics could not be loaded. Please retry.'}}

    @app.get('/direct-data/analytics')
    def direct_analytics():
        if not owner():return 'Admin access required.',403
        try:
            category=request.args.get('category','source')
            overview,rows,count,page=store.detail(owner(),request.args.get('month'),category,
                request.args.get('destination',''),request.args.get('franchise',''),request.args.get('q',''),request.args.get('page',1,type=int))
        except ValueError as exc:return str(exc),400
        args=request.args.to_dict();args.update(month=overview['month'],category=category)
        def page_url(n):return url_for('direct_analytics',**{**args,'page':n})
        return render_template('direct_analytics.html',overview=overview,rows=rows,count=count,page=page,
            category=category,title=CATEGORIES[category],categories=CATEGORIES,
            previous=page_url(page-1) if page>1 else '',next_page=page_url(page+1) if page*50<count else '')

    @app.get('/direct-data/analytics/record/<batch_id>/<int:row_number>')
    def direct_record(batch_id,row_number):
        if not owner():return 'Admin access required.',403
        row=store.row(owner(),batch_id,row_number)
        if row is None:return 'Record not available.',404
        from direct_data import HEADERS
        return render_template('direct_record.html',row=row,fields=list(zip(HEADERS,row['data'])))
