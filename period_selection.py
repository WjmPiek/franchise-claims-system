"""Receipt-month periods for monthly imports, with inclusive calendar date bounds."""
from datetime import date

def month_bounds(args):
    start=args.get('date_from','');end=args.get('date_to','')
    if not start and not end:return None,None
    try:
        first=date.fromisoformat(start);last=date.fromisoformat(end)
    except (ValueError,TypeError):raise ValueError('Choose valid start and end dates.')
    if first>last:raise ValueError('The start date must be on or before the end date.')
    return first.strftime('%Y-%m'),last.strftime('%Y-%m')


def selected_dates(args, fallback=None):
    from flask import session,current_app,g
    from datetime import date
    if getattr(g,'login_period_reset',False):return g.period_dates
    explicit=any(args.get(k) for k in ('date_from','date_to','report_from_date','report_to_date'))
    reset=bool(fallback) and (session.get('latest_period_pending') or
          (not explicit and session.get('period_default_version') != 2))
    if reset:
        import calendar
        last=date.fromisoformat(fallback['date_to'])
        dates={'date_from':last.replace(day=1).isoformat(),
               'date_to':last.replace(day=calendar.monthrange(last.year,last.month)[1]).isoformat()}
        session.pop('latest_period_pending',None)
        g.login_period_reset=True
    elif args.get('date_from') or args.get('date_to'):
        dates={'date_from':args.get('date_from',''),'date_to':args.get('date_to','')}
    elif args.get('report_from_date') or args.get('report_to_date'):
        dates={'date_from':args.get('report_from_date',''),'date_to':args.get('report_to_date','')}
    else:
        dates=getattr(g,'period_dates',None) or session.get('period_dates')
        if not dates and fallback:
            # A fresh login clears the session. Open the last available receipt
            # month, while retaining explicitly selected periods on later pages.
            import calendar
            last=date.fromisoformat(fallback['date_to'])
            dates={'date_from':last.replace(day=1).isoformat(),
                   'date_to':last.replace(day=calendar.monthrange(last.year,last.month)[1]).isoformat()}
        if not dates:
            today=date.today();dates={'date_from':date(today.year-1,1,1).isoformat(),'date_to':today.isoformat()}
    month_bounds(dates)
    g.period_dates=dates
    if current_app.secret_key:
        session['period_dates']=dates
        if fallback or explicit:session['period_default_version']=2
    return dates
