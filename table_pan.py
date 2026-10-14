"""Load the shared table interaction on every Claims HTML page."""
from flask import url_for

def register(app):
    if app.extensions.get('table_pan_registered'):return
    app.extensions['table_pan_registered']=True
    @app.after_request
    def add_table_pan_assets(response):
        if response.mimetype!='text/html' or response.direct_passthrough or not 200<=response.status_code<300:return response
        html=response.get_data(as_text=True)
        if 'table-pan.js' in html:return response
        index=html.lower().find('</head>')
        if index<0:return response
        assets='<link rel="stylesheet" href="'+url_for('static',filename='table-pan.css',v='20261014')+'"><script defer src="'+url_for('static',filename='table-pan.js',v='20261014')+'"></script>'
        response.set_data(html[:index]+assets+html[index:])
        return response
