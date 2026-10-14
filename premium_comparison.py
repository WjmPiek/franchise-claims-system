"""Read current publications and earlier saved monthly imports for comparison."""
import pandas as pd

def comparison(current, historical, start, end, selected, builder):
    source=current.copy()
    old=historical.copy()
    for frame in (source,old):
        if 'Month' in frame:frame['Month']=pd.to_datetime(frame['Month'],errors='coerce')
    published=set(source['Month'].dt.to_period('M')) if not source.empty else set()
    if not old.empty:old=old[~old['Month'].dt.to_period('M').isin(published)]
    source=pd.concat([source,old],ignore_index=True)
    if selected!='All' and not source.empty:source=source[source['Franchise']==selected]
    first=pd.Period(start,freq='M');last=pd.Period(end,freq='M')
    months=source['Month'].dt.to_period('M') if not source.empty else pd.Series(dtype=object)
    selected_frame=source[months.between(first,last)] if not source.empty else source
    result=builder(source,selected_frame,selected)
    previous_first=first-12;previous_last=last-12
    expected=list(pd.period_range(previous_first,previous_last,freq='M'))
    available=set(months)
    missing=[str(m) for m in expected if m not in available]
    result['missing_previous_months']=missing
    result['historical_previous_months']=[str(m) for m in expected if m in available and m not in published]
    result['current_period']=str(first) if first==last else f'{first} to {last}'
    result['last_year_period']=str(previous_first) if previous_first==previous_last else f'{previous_first} to {previous_last}'
    if missing:
        result['has_previous']=False
        for row in result['comparison_rows']:row['previous']=row['difference']=None
    return result
