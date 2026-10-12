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
