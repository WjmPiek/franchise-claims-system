"""Render published sidebar pages from stored totals instead of legacy analysis."""
from collections import defaultdict
from decimal import Decimal
import calendar
from flask import g,request,session,render_template,flash

def render_if_published(app,section):
    if section not in ('risk','commissions','policies','top_franchise','claims','downloads','imports','admin'):return None
    store=app.extensions.get('central_store')
    if store is None:return None
    pubs=store.publications()
    if not pubs:return None
    from period_selection import selected_dates,month_bounds
    latest=max(p['month'] for p in pubs);year,month=map(int,latest.split('-'))
    fallback=dict(date_from=min(p['month'] for p in pubs)+'-01',date_to=latest+'-'+str(calendar.monthrange(year,month)[1]))
    try:dates=selected_dates(request.args,fallback);start,end=month_bounds(dates)
    except ValueError as exc:return str(exc),400
    if not any(start<=p['month']<=end for p in pubs):return None
    base=app.extensions['central_page_access']();names={r['franchise'] for r in store.summaries(allowed=base)}
    with store.direct.connection() as conn:
        names.update(r[0] for r in store.direct.execute(conn,'SELECT DISTINCT c.franchise FROM central_claims c JOIN central_claim_publications p ON p.month=c.month AND p.batch_id=c.batch_id') if base is None or r[0] in base)
    names=sorted(names);selected=request.args.get('franchise','All');search=request.args.get('q','').strip()
    if selected!='All':search=''
    if search:
        matches=[n for n in names if n.casefold()==search.casefold()] or [n for n in names if search.casefold() in n.casefold()]
        if matches:selected=matches[0]
        else:flash('No franchise found matching "'+search+'".','warning')
    allowed=base if selected=='All' else {selected} if base is None or selected in base else set()
    rows=store.summaries(start,end,allowed)
    # Historical-only / claims-only selections retain their established fallback.
    if not rows:return None
    g.sidebar_franchise=selected
    grouped=defaultdict(lambda:defaultdict(Decimal))
    labels={'retail':'Retail Premium','percentage_commission':'Total Commission','paid_commissions':'Total Paid Commissions','book_value':'Total Book Value'}
    for row in rows:
        for key,label in labels.items():grouped[row['franchise']][label]+=Decimal(str(row['summary'][key]))
    topic=[dict(franchise=name,**values) for name,values in sorted(grouped.items())]
    ed=dict(period_label=start+' to '+end+' (receipt months)')
    if section=='policies':
        import app as application
        ed.update(age_notifications=application.get_policy_age_notifications(selected,12),
            age_notification_summary=application.get_policy_age_notification_summary(selected,100),
            age_notification_totals=application.get_policy_age_notification_totals(selected),
            age_notification_age_bands=application.get_policy_age_notification_age_bands(selected))
    titles={'risk':'Business Overview','commissions':'Commissions & Fees','policies':'Policies & Products','top_franchise':'Top Franchise','claims':'Claims','downloads':'Reports & Downloads','imports':'Imports','admin':'Admin'}
    return render_template('dashboard.html',workspace='home',home_section=section,home_title=titles[section],
        executive_dashboard=ed,topic_franchises=topic,topic_totals={},topic_top=[],topic_monthly=[],
        selected=selected,franchises=names,search_text=search,period_view='month',traffic_filter='All',
        dashboard_period_view='custom',period_dates=dates,commission_range_args=dates,
        premium_analysis_month='',policy_import_summary={},claims_import_summary={})
