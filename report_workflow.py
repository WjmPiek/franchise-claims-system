"""Single approved import/export workflow; archived data remains intact."""
from html.parser import HTMLParser
from urllib.parse import urlsplit,parse_qs
from flask import request,g,redirect,url_for,jsonify,session

LEGACY_EXPORTS={'export','export_payover','board_report','policy_reports','direct_reports','direct_recreate','direct_download','direct_report_view','sample','claims_report'}
LEGACY_IMPORTS={'central_import_balances','central_convert_saved','load_default_import_file'}
FILE_PATHS=('/export','/export_payover','/board_report','/policy_reports','/direct-data/reports','/direct-data/download/','/direct-data/report/','/sample','/claims/report')

def retired_link(value):
    parts=urlsplit(value);params=parse_qs(parts.query)
    return parts.path.startswith(FILE_PATHS) or 'download' in params or params.get('format',[''])[0] in ('xlsx','csv','zip') and not parts.path.startswith(('/reports/central/','/jobs/'))

class RemoveOldLinks(HTMLParser):
    def __init__(self):super().__init__(convert_charrefs=False);self.output=[];self.skip=0;self.skip_tag=''
    def handle_starttag(self,tag,attrs):
        if self.skip:
            if tag==self.skip_tag:self.skip+=1
            return
        attributes=dict(attrs)
        retired=(tag=='a' and retired_link(attributes.get('href',''))) or (tag=='form' and (retired_link(attributes.get('action','')) or attributes.get('id') in ('importForm','reportDownloadForm','directForm'))) or (tag=='button' and retired_link(attributes.get('formaction','')))
        if retired:self.skip=1;self.skip_tag=tag;return
        self.output.append(self.get_starttag_text())
    def handle_startendtag(self,tag,attrs):
        if not self.skip:self.output.append(self.get_starttag_text())
    def handle_endtag(self,tag):
        if self.skip:
            if tag==self.skip_tag:self.skip-=1
            return
        self.output.append('</'+tag+'>')
    def handle_data(self,data):
        if not self.skip:self.output.append(data)
    def handle_entityref(self,name):
        if not self.skip:self.output.append('&'+name+';')
    def handle_charref(self,name):
        if not self.skip:self.output.append('&#'+name+';')
    def handle_comment(self,data):
        if not self.skip:self.output.append('<!--'+data+'-->')
    def handle_decl(self,data):self.output.append('<!'+data+'>')

def register(app):
    @app.before_request
    def approved_file_workflow():
        if not getattr(g,'user',None):return
        if request.files and request.endpoint!='central_upload':return jsonify(ok=False,error='Use Imports: monthly Direct policy data or the annual claims workbook.'),410
        if request.endpoint in LEGACY_IMPORTS:return 'This import has been retired. Use the two approved source formats in Imports.',410
        if request.endpoint in LEGACY_EXPORTS or request.args.get('download'):
            return redirect(url_for('overview_page',section='downloads'))
    app.extensions['approved_file_workflow']=approved_file_workflow
    @app.after_request
    def reset_period_after_login(response):
        if request.endpoint=='login' and response.status_code==302 and session.get('user_id'):
            session.pop('period_dates',None)
            session['latest_period_pending']=True
        return response
    @app.after_request
    def remove_retired_download_links(response):
        if not response.is_streamed and response.mimetype=='text/html':
            parser=RemoveOldLinks();parser.feed(response.get_data(as_text=True));response.set_data(''.join(parser.output))
        return response
