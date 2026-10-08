"""Read-only reconciliation of report scope against stored Excel detail rows."""
from decimal import Decimal
from xml.sax.saxutils import escape
import pandas as pd
from reportlab.platypus import Paragraph, Spacer, Table, TableStyle, PageBreak
from reportlab.lib import colors

FIELDS = ['original_risk_premium', 'retail_premium', 'r1_policy_fee',
          'adv_fund_2_1_fee', 'new_risk_premium']
HEADERS = ['Month', 'Franchise', 'Excel rows', 'Risk Premium', 'Retail Premium',
           'R1 fee', 'ADV fee', 'Payover Less Comm.', 'Risk difference', 'Retail difference',
           'Payover difference']

def amount(value):
    return Decimal(str(value if pd.notna(value) else 0)).quantize(Decimal('.01'))

def reconciliation(monthly, engine=None, sql_text=None, memory=None, config=None):
    """Monthly is scoped *before* grouping; never fetch outside those exact pairs."""
    config = config or {}
    work = monthly.copy()
    work = work[~work.Franchise.isin(config.get('excluded', []))]
    mapping = config.get('groups', {}) if config.get('use_groups', False) else {}
    mapping = {k: v for k, v in mapping.items() if str(v).strip()}
    pairs = {(str(r.Franchise), pd.Timestamp(r.Month).strftime('%Y-%m'))
             for r in work.itertuples()}
    detail = []
    if pairs and engine is not None:
        names = sorted({f for f, m in pairs})
        months = sorted({m for f, m in pairs})
        params = {f'f{i}': f for i, f in enumerate(names)}
        params.update({f'm{i}': m + '-01' for i, m in enumerate(months)})
        query = ('SELECT franchise_name, import_month, source_file, COUNT(*) AS rows, '
                 + ', '.join(f'COALESCE(SUM({c}),0) AS {c}' for c in FIELDS)
                 + ' FROM policydata_detail_raw WHERE franchise_name IN ('
                 + ','.join(f':f{i}' for i in range(len(names)))
                 + ') AND import_month IN ('
                 + ','.join(f':m{i}' for i in range(len(months)))
                 + ') GROUP BY franchise_name, import_month, source_file')
        with engine.connect() as conn:
            detail = [dict(r) for r in conn.execute(sql_text(query), params).mappings()]
    elif memory is not None and not memory.empty:
        frame = memory.rename(columns={'franchise': 'franchise_name'}).copy()
        for keys, group in frame.groupby(['franchise_name', 'import_month', 'source_file'], dropna=False):
            detail.append(dict(zip(['franchise_name', 'import_month', 'source_file'], keys),
                               rows=len(group), **{c: sum(amount(v) for v in group[c]) for c in FIELDS}))
    imported, sources = {}, set()
    for row in detail:
        franchise, month = str(row['franchise_name']), pd.Timestamp(row['import_month']).strftime('%Y-%m')
        if (franchise, month) not in pairs:
            continue
        sources.add((month, franchise, str(row['source_file'])))
        key = (month, mapping.get(franchise, franchise))
        total = imported.setdefault(key, {'rows': 0, **{c: Decimal(0) for c in FIELDS}})
        total['rows'] += int(row['rows'])
        for c in FIELDS:
            total[c] += amount(row[c])
    reports = {}
    for row in work.to_dict('records'):
        key = (pd.Timestamp(row['Month']).strftime('%Y-%m'), mapping.get(row['Franchise'], row['Franchise']))
        total = reports.setdefault(key, [Decimal(0)] * 3)
        for i, c in enumerate(['Original Risk Premium', 'Retail Premium', 'Risk Premium']):
            total[i] += amount(row.get(c, 0))
    rows = []
    missing = []
    for key, report in sorted(reports.items()):
        imp = imported.get(key)
        if imp is None:
            missing.append(' / '.join(key))
            rows.append([*key, 'Unavailable', *([None] * 8)])
        else:
            rows.append([*key, imp['rows'], *[imp[c] for c in FIELDS],
                         report[0] - imp[FIELDS[0]], report[1] - imp[FIELDS[1]],
                         report[2] - imp[FIELDS[4]]])
    totals = [sum(row[i] for row in rows if row[i] is not None and row[2] != 'Unavailable')
              for i in range(3, 11)]
    has_difference = any(row[i] != 0 for row in rows if row[2] != 'Unavailable' for i in (8, 9, 10))
    status = ('Detail unavailable for some months; re-import those source files to verify.' if missing
              else 'Differences require review.' if has_difference
              else 'Imported detail totals match this report.' if rows else 'No imported data in this report.')
    return dict(rows=rows, totals=totals, count=sum(v['rows'] for v in imported.values()),
                sources=sorted(sources), status=status, missing=missing)

NOTE = ('All applicable relation rows, including blanks and negative adjustments, are included. '
        'Risk Premium and Retail Premium are the amounts to compare with the source Excel sheet. '
        'Payover Less Comm. is the existing risk amount after R1 and ADV fees. Retail minus Payover Less Comm. is Admin fee. Differences are report totals minus imported detail. '
        'A match confirms internal consistency; compare the totals and row counts with Excel to confirm the source.')

def add_pdf_summary(story, styles, result, width):
    story.append(Paragraph('Import reconciliation summary', styles['Heading1']))
    story.append(Paragraph(escape(result['status']), styles['Normal']))
    story.append(Paragraph(f"Imported Excel rows: {result['count']:,} | Source files: {len({s[2] for s in result['sources']})}", styles['Normal']))
    story.append(Spacer(1, 8))
    labels = ['Risk Premium', 'Retail Premium', 'R1 fee', 'ADV fee', 'Payover Less Comm.',
              'Risk difference', 'Retail difference', 'Payover difference']
    overview = [['Imported totals' + (' (available detail only)' if result['missing'] else ''), 'Rand']]
    overview += [[label, f'{value:,.2f}'] for label, value in zip(labels, result['totals'])]
    table = Table(overview, colWidths=[width * .6, width * .4], repeatRows=1)
    table.setStyle(TableStyle([('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#e8edf5')),
                               ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                               ('ALIGN', (1, 0), (1, -1), 'RIGHT'),
                               ('FONTSIZE', (0, 0), (-1, -1), 9),
                               ('BOTTOMPADDING', (0, 0), (-1, -1), 5)]))
    story.append(table)
    story.append(Spacer(1, 8))
    story.append(Paragraph(NOTE, styles['Normal']))
    story.append(Spacer(1, 8))
    by_month = {}
    for row in result['rows']:
        if row[2] == 'Unavailable':
            continue
        values = by_month.setdefault(row[0], [0, Decimal(0), Decimal(0), Decimal(0)])
        for index, col in enumerate([2, 3, 4, 7]):
            values[index] += row[col]
    if by_month:
        data = [['Month', 'Excel rows', 'Risk Premium', 'Retail Premium', 'Payover Less Comm.']]
        data += [[month, f'{v[0]:,}', *[f'{a:,.2f}' for a in v[1:]]]
                 for month, v in sorted(by_month.items())]
        table = Table(data, colWidths=[width * .14, width * .14, width * .24, width * .24, width * .24], repeatRows=1)
        table.setStyle(TableStyle([('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#e8edf5')),
                                   ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                                   ('ALIGN', (1, 0), (-1, -1), 'RIGHT'),
                                   ('FONTSIZE', (0, 0), (-1, -1), 8)]))
        story.append(table)
        story.append(Spacer(1, 8))
    for source in sorted({s[2] for s in result['sources']}):
        story.append(Paragraph(escape('Source file: ' + source), styles['Normal']))
    if result['missing']:
        story.append(Paragraph(escape('Detail unavailable: ' + '; '.join(result['missing'])), styles['Normal']))
    story.append(PageBreak())

def add_excel_summary(writer, result):
    ws = writer.book.add_worksheet('Import Summary')
    writer.sheets['Import Summary'] = ws
    bold = writer.book.add_format({'bold': True, 'bg_color': '#e8edf5'})
    money = writer.book.add_format({'num_format': '#,##0.00'})
    wrap = writer.book.add_format({'text_wrap': True, 'valign': 'top'})
    ws.set_column(0, 0, 13)
    ws.set_column(1, 1, 25)
    ws.set_column(2, 2, 16)
    ws.set_column(3, 10, 19, money)
    ws.merge_range('A1:K1', 'Import reconciliation summary', bold)
    ws.merge_range('A2:K2', result['status'], wrap)
    ws.merge_range('A3:K4', NOTE, wrap)
    ws.set_row(2, 24)
    ws.write_row(5, 0, HEADERS, bold)
    for index, row in enumerate(result['rows'], 6):
        for col, value in enumerate(row):
            if value is not None:
                if isinstance(value, str):
                    ws.write_string(index, col, value)
                else:
                    ws.write_number(index, col, float(value), money if col >= 3 else None)
    index = len(result['rows']) + 6
    ws.write_string(index, 1, 'AVAILABLE TOTAL' if result['missing'] else 'TOTAL', bold)
    ws.write_number(index, 2, result['count'])
    for col, value in enumerate(result['totals'], 3):
        ws.write_number(index, col, float(value), money)
    ws.write_row(index + 3, 0, ['Month', 'Source Excel file'], bold)
    for index, row in enumerate(sorted({(month, source) for month, franchise, source in result['sources']}), index + 4):
        for col, value in enumerate(row):
            ws.write_string(index, col, value)
    ws.set_landscape()
    ws.fit_to_pages(1, 0)
    ws.freeze_panes(6, 3)
