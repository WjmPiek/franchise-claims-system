"""Independent preparation of Direct exports; never writes Claims premium tables."""
import csv
import io
import json
import re
import sqlite3
import shutil
import threading
import uuid
import zipfile
from collections import defaultdict
from contextlib import closing
from datetime import datetime, date, timedelta
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path
from flask import redirect, g, request, render_template, send_file, jsonify, session, url_for
from openpyxl import load_workbook
import xlsxwriter
from import_progress import report, active, stream_import
from direct_analytics import DirectStore, FIELDS, data_month, register_direct_analytics

HEADERS = ['Franchise', 'Group', 'Policy Number', 'Surname', 'Full Names', 'ID Number',
           'Date of Birth', 'Passport Nr', 'Relation', 'Inception Date', 'Cover', 'Product',
           'Risk', 'Retail', 'Address', 'Contact Number', 'System REF', 'MPIA', 'Area',
           'Suburb', 'Code', 'Agent Name and Policy']
VOUCHER_PRODUCT = 'GROCERY DISCOUNT COUPONS FOR MARTINS CLIENTS'
VOUCHER_HEADERS = ['Franchise', 'Policy Number', 'Surname', 'Full Names', 'Date of Birth',
                   'Relation', 'Inception Date', 'Product', 'Fee', 'Contact Number', 'MPIA']
EXCEPTIONS = ['Alberton', 'Bethlehem', 'Carletonville', 'Frankfort', 'Reitz',
              'Rustenburg', 'Katlehong']
PRETORIA = {'PRETORIANOORD', 'PRETORIANORTH', 'HAMMANSKRAAL', 'SOSHANGUVE'}
GROUP_NAME = 'Pretoria North Group'
LOCK = threading.Lock()

def norm(value):
    return re.sub(r'[^A-Z0-9]', '', str(value or '').upper())

def branch_key(value):
    key = norm(value)
    for prefix in ('MARTINSFUNERALS', 'MARTINSFUNERAL', 'MARTINS'):
        if key.startswith(prefix):
            return key[len(prefix):]
    return key

def allocation(franchise):
    key = branch_key(franchise)
    if key in PRETORIA:
        return GROUP_NAME
    if key in {'CARLETONVILLEKEMPTONPARK', 'CARLETONVILLEANDKEMPTONPARK'}:
        return 'Carletonville'
    for name in EXCEPTIONS:
        if key == norm(name):
            return name
    return 'AUL'

def number(value, label, row_number, blank_zero=True):
    text = str(value if value is not None else '').strip().replace('\u00a0', '').replace(' ', '')
    if text.upper() in {'', 'NULL'} and blank_zero:
        return Decimal(0)
    text = re.sub(r'^[Rr]', '', text)
    if ',' in text and '.' not in text:
        text = text.replace(',', '.')
    elif ',' in text:
        text = text.replace(',', '')
    if text.startswith('(') and text.endswith(')'):
        text = '-' + text[1:-1]
    try:
        value = Decimal(text)
        if not value.is_finite():
            raise InvalidOperation
        return value
    except InvalidOperation:
        raise ValueError(f'Excel/source row {row_number}: {label} is not a valid number.') from None

def cents(value):
    return value.quantize(Decimal('.01'), rounding=ROUND_HALF_UP)

def short_date(value, label, row_number):
    if isinstance(value, (datetime, date)):
        return value.strftime('%Y-%m-%d')
    text = str(value or '').replace('\x00', '').strip()
    if text.upper() in {'', 'NULL'}:
        return ''
    for fmt in ('%Y-%m-%d', '%Y/%m/%d', '%d/%m/%Y', '%d-%m-%Y'):
        try:
            return datetime.strptime(text.split(' ')[0], fmt).strftime('%Y-%m-%d')
        except ValueError:
            pass
    raise ValueError(f'Excel/source row {row_number}: {label} is not a valid date.')

def text_value(value):
    if value is None:
        return ''
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    text = str(value).replace('\x00', '').strip()
    return '' if text.upper() == 'NULL' else text

def mobile_number(value):
    text = text_value(value)
    digits = re.sub(r'\D', '', text)
    if re.fullmatch(r'[678]\d{8}', digits):
        return '0' + digits
    if re.fullmatch(r'27[678]\d{8}', digits):
        return '0' + digits[2:]
    if re.fullmatch(r'0[678]\d{8}', digits):
        return digits
    return text

def source_rows(path, encoding='auto', sheet_name=None):
    """Headerless Direct text uses the verified 22-column example order."""
    path = Path(path)
    if path.suffix.lower() == '.xlsx':
        book = load_workbook(path, read_only=True, data_only=True)
        try:
            candidates = [s for s in book if s.max_column >= 22]
            if sheet_name:
                if sheet_name not in book.sheetnames:
                    raise ValueError('The selected Excel worksheet does not exist.')
                sheet = book[sheet_name]
            elif len(candidates) == 1:
                sheet = candidates[0]
            else:
                raise ValueError('Specify the original Direct worksheet name; more than one data sheet was found.')
            for n, values in enumerate(sheet.iter_rows(values_only=True), 1):
                yield n, list(values)
        finally:
            book.close()
    else:
        if encoding == 'auto':
            encoding = 'utf-8-sig'
            try:
                with path.open(encoding=encoding) as handle:
                    while handle.read(65536):
                        pass
            except UnicodeDecodeError:
                # Direct's OEM export contains bytes invalid in Windows-1252.
                encoding = 'cp850'
        with path.open(encoding=encoding, newline='') as handle:
            sample = handle.read(65536)
            handle.seek(0)
            if not sample.strip():
                raise ValueError('The source file is empty.')
            first = sample.splitlines()[0]
            delimiter = '\t' if '\t' in first else ';' if first.count(';') > first.count(',') else ','
            for n, values in enumerate(csv.reader(handle, delimiter=delimiter, strict=True), 1):
                yield n, values

def stage_source(path, conn, encoding='auto', sheet_name=None):
    conn.execute('CREATE TABLE rows (n INTEGER PRIMARY KEY, branch TEXT, policy TEXT, relation TEXT, product TEXT, voucher INTEGER, destination TEXT, plan INTEGER, risk TEXT, retail TEXT, mpia TEXT, values_json TEXT)')
    conn.execute('CREATE TABLE source_raw (n INTEGER PRIMARY KEY,data TEXT NOT NULL)')
    conn.execute('CREATE TABLE source_fingerprints(n INTEGER PRIMARY KEY,digest TEXT NOT NULL)')
    conn.execute('CREATE TABLE source_extra(n INTEGER PRIMARY KEY,data TEXT NOT NULL)')
    indices = list(range(22))
    header_keys=None
    count = 0
    totals = defaultdict(lambda: Decimal(0))
    report('Reading and converting Direct source rows')
    for n, values in source_rows(path, encoding, sheet_name):
        if not any(text_value(v) for v in values):
            continue
        if n == 1 and norm(values[0]) == 'FRANCHISE':
            keys = [norm(v).replace('INSEPTION', 'INCEPTION') for v in values]
            header_keys=keys
            if len(set(keys))!=len(keys):raise ValueError('Duplicate Direct source headings; resolve before import.')
            try:
                indices = [keys.index(norm(h)) for h in HEADERS]
            except ValueError:
                raise ValueError('The source headers do not match the October Direct example.') from None
            continue
        if len(values) != (len(header_keys) if header_keys else 22):
            raise ValueError(f'Source row {n} has {len(values)} columns; the Direct layout needs 22.')
        extra={header_keys[i]:values[i] for i in range(len(values)) if i not in indices} if header_keys else {}
        conn.execute('INSERT INTO source_extra VALUES (?,?)',(n,json.dumps(extra,ensure_ascii=False,default=str)))
        values = [values[i] for i in indices]
        row = [text_value(v) for v in values]
        conn.execute('INSERT INTO source_raw VALUES (?,?)',(n,json.dumps(values,ensure_ascii=False,default=str)))
        import hashlib
        conn.execute('INSERT INTO source_fingerprints VALUES (?,?)',(n,hashlib.sha256(json.dumps(values,ensure_ascii=False,default=str).encode('utf-8')).hexdigest()))
        if not row[0] or norm(row[0]) in {'TOTAL', 'GRANDTOTAL'}:
            raise ValueError(f'Source row {n} has no franchise or is a total row. Export source detail only.')
        for i in (6, 9):
            row[i] = short_date(values[i], HEADERS[i], n)
        risk, retail = (number(values[i], HEADERS[i], n) for i in (12, 13))
        from central_calculations import effective_mpia
        mpia = effective_mpia(values[17])
        cover = number(values[10], 'Cover', n)
        row[10], row[12], row[13], row[17] = str(cover), str(risk), str(retail), str(mpia)
        row[8] = row[8].upper()
        voucher = norm(row[11]) == norm(VOUCHER_PRODUCT)
        # Product amount takes precedence over a dependent's smaller individual Cover.
        product_plan = bool(re.search(r'(?<!\d)(?:40\s*000|50\s*000)(?!\d)', row[11]))
        plan = int(product_plan or (row[8] == 'MEM' and cover in (40000, 50000)))
        conn.execute('INSERT INTO rows VALUES (?,?,?,?,?,?,?,?,?,?,?,?)',
                     (n, branch_key(row[0]), row[2], row[8], row[11], int(voucher), allocation(row[0]),
                      plan, str(risk), str(retail), str(mpia), json.dumps(row, ensure_ascii=False)))
        totals['risk'] += risk
        totals['retail'] += retail
        count += 1
        if count % 2000 == 0:
            report('Reading and converting Direct source rows', count)
    if not count:
        raise ValueError('The source has no data rows.')
    conn.execute('CREATE INDEX policy_lookup ON rows(branch,policy)')
    conn.execute('CREATE INDEX source_fingerprint_lookup ON source_fingerprints(digest)')
    conn.commit()
    report('Direct source read', count, count)
    return {'rows': count, 'risk': totals['risk'], 'retail': totals['retail']}

def row_fees(row, faithnet_basis, voucher_basis):
    """Apply chosen timing only to this separate preparation workflow."""
    if faithnet_basis=='mpia' and voucher_basis=='mpia':
        from central_calculations import fees
        return fees(row)
    risk, mpia = Decimal(row[12]), Decimal(row[17])
    is_mem = row[8] == 'MEM'
    deduction = Decimal(0)
    if branch_key(row[0]) in {'FAITHNET', 'FAITHNETSA'} and is_mem:
        rate = Decimal(40 if re.search(r'\b1\s*\+', row[11]) else 10)
        multiplier = mpia if faithnet_basis == 'mpia' else Decimal(2) if faithnet_basis == 'over_two_double' and mpia > 2 else Decimal(1)
        deduction = rate * multiplier
    adjusted = cents(risk - deduction)
    r1 = cents(mpia) if is_mem else Decimal(0)
    adv = cents(max(adjusted - r1, Decimal(0)) * Decimal('.021')) if is_mem else Decimal(0)
    after_r1 = max(adjusted - r1, Decimal(0)) if is_mem else risk
    net = cents(after_r1 - adv) if is_mem else risk
    voucher_multiplier = mpia if voucher_basis == 'mpia' else Decimal(1)
    supplier = cents(Decimal('6.50') * voucher_multiplier)
    retained = cents(risk - supplier)
    return deduction, r1, adv, net, supplier, retained

class SinkWriter:
    """Calculation-only import; workbooks are generated later on demand."""
    def append(self,values):pass
    def note(self,value):pass
    def close(self):pass

class SheetWriter:
    def __init__(self, path, sheet, headers, book=None):
        self.owns_book = book is None
        self.book = book if book is not None else xlsxwriter.Workbook(str(path), {'constant_memory': True, 'strings_to_formulas': False, 'strings_to_urls': False})
        self.sheet = self.book.add_worksheet(sheet[:31])
        self.headers = headers
        self.date_fmt = self.book.add_format({'num_format': 'dd/mm/yyyy'})
        self.num_fmt = self.book.add_format({'num_format': '0.00########'})
        self.count_fmt = self.book.add_format({'num_format': '#,##0'})
        self.pct_fmt = self.book.add_format({'num_format': '0.00%'})
        self.text_fmt = self.book.add_format({'num_format': '@'})
        head = self.book.add_format({'bold': True, 'bg_color': '#e8edf5', 'text_wrap': True})
        self.sheet.write_row(0, 0, headers, head)
        self.sheet.set_row(0, 32)
        for i, name in enumerate(headers):
            self.sheet.set_column(i, i, {'Product': 60, 'Franchise': 48, 'Full Names': 36,
                                        'Surname': 26, 'Date of Birth': 16, 'Inception Date': 16}.get(name, 22))
        self.sheet.freeze_panes(1, 3)
        self.n = 1

    def append(self, values):
        for i, value in enumerate(values):
            if value is None or value == '':
                continue
            name = self.headers[i]
            if name in ('Date of Birth', 'Inception Date','Date Of Claim Paid'):
                self.sheet.write_datetime(self.n, i, datetime.strptime(str(value), '%Y-%m-%d'), self.date_fmt)
            elif name in ('ID Number','Passport Nr','Policy Number','System REF','Contact Number'):
                self.sheet.write_string(self.n,i,str(value),self.text_fmt)
            elif name in ('Risk', 'Retail', 'Cover', 'MPIA', 'Fee', 'Original Risk', 'Faithnet Deduction', 'R1 Fee', 'ADV Fee', 'Payover') or isinstance(value, (Decimal, int, float)):
                if str(value).strip().upper() in ('NULL','NONE'):
                    self.sheet.write_string(self.n,i,str(value),self.text_fmt)
                else:self.sheet.write_number(self.n, i, float(value), self.pct_fmt if name in ('Claims Ratio','Claim Ratio') else self.count_fmt if name in ('Rows','MPIA') else self.num_fmt)
            else:
                self.sheet.write_string(self.n, i, str(value), self.text_fmt)
        self.n += 1

    def close(self):
        self.sheet.autofilter(0, 0, max(self.n - 1, 1), len(self.headers) - 1)
        if self.owns_book:
            self.book.close()

    def note(self, text):
        fmt = self.book.add_format({'text_wrap': True, 'valign': 'top', 'font_color': '#596675'})
        self.sheet.merge_range(self.n, 0, self.n, len(self.headers) - 1, text, fmt)
        self.sheet.set_row(self.n, 32)
        self.n += 1

def separate_product_plans(conn):
    """Different insured products may legitimately share the same policy number."""
    assignments={}
    conflicts=conn.execute("SELECT branch,policy FROM rows WHERE destination='AUL' AND voucher=0 AND relation='MEM' GROUP BY branch,policy HAVING MIN(plan)<>MAX(plan)").fetchall()
    for branch,policy in conflicts:
        members=conn.execute("SELECT n,plan,product FROM rows WHERE branch=? AND policy=? AND relation='MEM' AND voucher=0 ORDER BY n",(branch,policy)).fetchall()
        products=defaultdict(set)
        for n,plan,product in members:products[norm(product)].add(plan)
        if any(not key or len(flags)>1 for key,flags in products.items()):
            description='; '.join(f'row {n}: {product} ({"Plan 2" if plan else "Historic"})' for n,plan,product in members[:10])
            raise ValueError(f'Source row {members[0][0]}: AUL policy {policy} at {branch} has conflicting MEM plans for the same product. {description}. Confirm the correct cover/product in the source. No rows were skipped or saved.')
        for n,rel,product,plan in conn.execute("SELECT n,relation,product,plan FROM rows WHERE branch=? AND policy=? AND voucher=0 ORDER BY n",(branch,policy)):
            flags=products.get(norm(product),set())
            if rel=='MEM':assignments[n]=plan
            elif len(flags)==1:assignments[n]=next(iter(flags))
            elif re.search(r'(?<!\d)(?:40\s*000|50\s*000)(?!\d)',product):assignments[n]=1
            else:
                raise ValueError(f'Source row {n}: policy {policy} at {branch} has separate MEM products, but {rel or "(Blank)"} product {product} cannot be linked to one of them. Confirm which insured product this row belongs to. No rows were skipped or saved.')
    return assignments

def prepare_outputs(conn, folder, source_name, faithnet_basis='over_two_double', voucher_basis='mpia', plan_basis='policy', saved=False, generate_files=True):
    if faithnet_basis not in ('row', 'mpia', 'over_two_double') or voucher_basis not in ('row', 'mpia') or plan_basis not in ('policy', 'row'):
        raise ValueError('Choose the fee timing and AUL policy split before preparing downloads.')
    product_plans={}
    central_plans={}
    central_issues={}
    if faithnet_basis=='mpia' and not saved:
        from central_validation import plan_assignments
        central_plans,central_issues=plan_assignments(conn)
    if plan_basis == 'policy' and not saved and faithnet_basis!='mpia':
        missing = conn.execute("SELECT MIN(n) FROM rows WHERE destination='AUL' AND voucher=0 AND TRIM(policy)=''").fetchone()[0]
        if missing:
            raise ValueError(f'Source row {missing}: missing policy number prevents the AUL policy split.')
        product_plans=separate_product_plans(conn)
    folder = Path(folder)
    folder.mkdir(exist_ok=True, parents=True)
    writers = {}
    paths = {}
    period = re.search(r'(\d{8}_to_\d{8})', source_name)
    prefix = 'PolicyData_' + period.group(1) if period else 'Direct'
    labels = {'AUL Plan 2 New': 'Martins Funeral Plan 2 (New)', 'AUL Historic': 'Martins Funeral Plan (Historic)'}
    summary_headers = ['Destination', 'Rows', 'Original Risk', 'Retail', 'Faithnet Deduction', 'R1 Fee', 'ADV Fee', 'Payover', 'Voucher Retained']
    combined_path = folder / (prefix + '_Insurer Sheets.xlsx')
    combined_book = xlsxwriter.Workbook(str(combined_path), {'constant_memory': True, 'strings_to_formulas': False, 'strings_to_urls': False}) if generate_files else None
    combined = {'AUL Totals': SheetWriter(combined_path, 'AUL Totals', summary_headers, book=combined_book) if generate_files else SinkWriter()}
    for destination in EXCEPTIONS + [GROUP_NAME, 'AUL']:
        combined[destination] = SheetWriter(combined_path, 'AUL All Data' if destination == 'AUL' else destination, HEADERS, book=combined_book) if generate_files else SinkWriter()
    if generate_files:paths['Insurer Sheets'] = combined_path
    def writer(key, headers=HEADERS):
        if key not in writers:
            path = folder / (prefix + '_' + labels.get(key, key) + '.xlsx')
            if generate_files:paths[key] = path
            writers[key] = SheetWriter(path, key, headers) if generate_files else SinkWriter()
        return writers[key]
    for key in EXCEPTIONS + [GROUP_NAME, 'AUL All Data', 'AUL Plan 2 New', 'AUL Historic', 'Vouchers', 'Converted Source']:
        writer(key, VOUCHER_HEADERS if key == 'Vouchers' else HEADERS)
    # Index only non-voucher MEM rows. Never match across franchises.
    policy_plan = defaultdict(set)
    available_plan = defaultdict(set)
    for branch, policy, is_voucher, rel, plan, data in conn.execute('SELECT branch,policy,voucher,relation,plan,values_json FROM rows'):
        row = json.loads(data)
        if faithnet_basis!='mpia' and not saved and not is_voucher and policy:
            available_plan[(branch, policy)].add(bool(plan))
        if faithnet_basis!='mpia' and not saved and not is_voucher and rel == 'MEM' and policy:
            policy_plan[(branch, policy)].add(bool(plan))
    totals = defaultdict(lambda: [0, *[Decimal(0) for _ in range(7)]])
    def voucher_contacts(branch,row):
        # Resolve the verified reference first; then the same person's exact source fields.
        # A shared policy number alone never selects a phone belonging to another person.
        matches=[]
        for raw in conn.execute("SELECT values_json FROM rows WHERE branch=? AND policy=? AND relation='MEM' AND voucher=0",(branch,row[2])):
            member=json.loads(raw[0]);same_person=all(norm(member[i])==norm(row[i]) for i in (3,4,5,6,7))
            if same_person:matches.append(member)
        exact=[m for m in matches if row[16] and m[16]==row[16]]
        selected=exact if exact else matches
        return {mobile_number(m[15]) for m in selected if m[15]}
    aul_details = defaultdict(lambda: [0, *[Decimal(0) for _ in range(7)]])
    issues = []
    source_count = conn.execute('SELECT COUNT(*) FROM rows').fetchone()[0]
    conn.execute('DROP TABLE IF EXISTS prepared')
    conn.execute('CREATE TABLE prepared (n INTEGER PRIMARY KEY,destination TEXT,plan TEXT,franchise TEXT,policy TEXT,relation TEXT,original_risk TEXT,retail TEXT,faithnet TEXT,r1 TEXT,adv TEXT,payover TEXT,supplier TEXT,retained TEXT,contact TEXT,data TEXT,issues TEXT)')
    processed = 0
    for n, branch, policy, rel, is_voucher, dest, plan, data in conn.execute('SELECT n,branch,policy,relation,voucher,destination,plan,values_json FROM rows ORDER BY n'):
        issue_start = len(issues)
        row = json.loads(data)
        for message in central_issues.get(n,[]):issues.append([n,row[0],policy,message])
        writer('Converted Source').append(row)
        risk, retail = Decimal(row[12]), Decimal(row[13])
        snapshot=None
        if saved:
            snapshot=dict(zip(FIELDS,conn.execute('SELECT * FROM saved WHERE n=?',(n,)).fetchone()))
            deduction,r1,adv,net,supplier,retained=[Decimal(snapshot[key]) for key in ('faithnet','r1','adv','payover','supplier','retained')]
        else:
            deduction, r1, adv, net, supplier, retained = row_fees(row, faithnet_basis, voucher_basis)
        plan_name = ''
        phone = row[15]
        if is_voucher:
            # Supplier receives precisely the eleven approved fields, never IDs/risk/audit data.
            contacts = voucher_contacts(branch,row)
            phone = next(iter(contacts)) if len(contacts) == 1 else mobile_number(row[15]) if len(contacts) == 0 else ''
            if len(contacts) > 1:
                issues.append([n, row[0], policy, 'Conflicting MEM phone numbers; voucher contact left blank.'])
            elif not phone:
                issues.append([n, row[0], policy, 'No contact number for voucher.'])
            elif not re.fullmatch(r'0[678]\d{8}', phone):
                issues.append([n, row[0], policy, 'Voucher contact is not a recognised South African mobile number.'])
            expected = cents(Decimal('8.50') * (Decimal(row[17]) if voucher_basis == 'mpia' else Decimal(1)))
            if risk != expected:
                issues.append([n, row[0], policy, 'Voucher risk differs from 8.50 under selected fee timing.'])
            if saved:phone=snapshot['contact']
            writer('Vouchers', VOUCHER_HEADERS).append([row[0], row[2], row[3], row[4], row[6], row[8], row[9], row[11], supplier, phone, row[17]])
            key = 'Vouchers'
            fees = (Decimal(0), Decimal(0), Decimal(0), supplier, retained)
        else:
            prepared = row.copy()
            prepared[12] = str(net)
            writer('AUL All Data' if dest == 'AUL' else dest).append(prepared)
            combined[dest].append(prepared)
            if dest == 'AUL':
                if plan_basis == 'row':
                    plan = int(Decimal(row[10]) in (40000, 50000))
                if faithnet_basis=='mpia' and not saved:
                    plan=central_plans.get(n,'Unclassified')=='Plan 2'
                elif saved:
                    plan=int(snapshot['plan']=='Plan 2')
                elif plan_basis == 'policy':
                    flags = policy_plan.get((branch, policy), set()) if policy else set()
                    if n in product_plans:
                        plan=product_plans[n]
                    elif flags:
                        plan = int(next(iter(flags)))
                    elif not policy:
                        raise ValueError(f'Source row {n}: missing policy number prevents the AUL policy split.')
                    else:
                        plan = int(True in available_plan.get((branch, policy), set()))
                        issues.append([n, row[0], policy, 'No MEM row in this source; policy split follows available product labels.'])
                plan_name = central_plans.get(n,'Unclassified') if faithnet_basis=='mpia' and not saved else snapshot['plan'] if saved else 'Plan 2' if plan else 'Historic'
                split = 'AUL Plan 2 New' if plan_name=='Plan 2' else 'AUL Historic' if plan_name=='Historic' else 'AUL Unclassified'
                writer(split).append(prepared)
                t = aul_details[split]
                t[0] += 1
                for i, value in enumerate((risk, retail, deduction, r1, adv, net, Decimal(0)), 1):
                    t[i] += value
            key = dest
            fees = (deduction, r1, adv, net, Decimal(0))
        t = totals[key]
        t[0] += 1
        for i, value in enumerate((risk, retail, *fees), 1):
            t[i] += value
        if saved:
            del issues[issue_start:]
            issues.extend([n,row[0],policy,message] for message in json.loads(snapshot['issues']))
        row_issues = [issue[3] for issue in issues[issue_start:]]
        analytics_values = (n,key,plan_name,row[0],policy,rel,str(risk),str(retail),
            str(deduction if not is_voucher else 0),str(r1 if not is_voucher else 0),
            str(adv if not is_voucher else 0),str(net if not is_voucher else 0),
            str(supplier if is_voucher else 0),str(retained if is_voucher else 0),phone,
            json.dumps(row,ensure_ascii=False),json.dumps(row_issues,ensure_ascii=False))
        conn.execute('INSERT INTO prepared VALUES ('+','.join('?' for _ in analytics_values)+')',analytics_values)
        processed += 1
        if processed % 2000 == 0:
            report('Preparing insurer and voucher reports', processed, source_count)
    summary = writer('Reconciliation', summary_headers)
    summary.note('Source file: ' + source_name)
    summary.note('Fee timing is defined by the import version. Reported values are final stored components; no extra deduction is applied.')
    summary.note('Voucher supplier fee: 6.50 x MPIA; retained amount equals source risk minus the supplier fee.' if voucher_basis == 'mpia' else 'Voucher supplier fee: 6.50 once per row.')
    summary.note('AUL split: confirmed products and related policy rows follow the MEM plan; unclassified records require review.')
    for key, values in sorted(totals.items()):
        summary.append([key, *values])
    summary.append(['TOTAL', *[sum(v[i] for v in totals.values()) for i in range(8)]])
    summary.note('Risk rounding / zero-floor adjustment: ' + str(sum(v[1] - sum(v[3:]) for v in totals.values())))
    summary.note('The risk difference shows fee rounding and the existing MEM zero floor. Original Risk and Retail retain source precision. All rows, including vouchers, are counted once in TOTAL. Review contains unmatched or ambiguous source records.')
    aul = writer('AUL Totals', summary_headers)
    aul.append(['AUL', *totals['AUL']])
    combined['AUL Totals'].append(['AUL', *totals['AUL']])
    for key in ('AUL Plan 2 New', 'AUL Historic','AUL Unclassified'):
        # The two loading files partition AUL All Data; counts are visible here.
        aul.append([key, *aul_details[key]])
        combined['AUL Totals'].append([key, *aul_details[key]])
    issue_writer = writer('Review', ['Source Row', 'Franchise', 'Policy Number', 'Issue'])
    for issue in issues:
        issue_writer.append(issue)
    for w in writers.values():
        w.close()
    for w in combined.values():
        w.close()
    if combined_book:combined_book.close()
    if not generate_files:
        conn.commit()
        return {},None,totals,issues
    report('Creating the download ZIP', source_count, source_count)
    archive = folder / 'Direct_reports.zip'
    with zipfile.ZipFile(archive, 'w', zipfile.ZIP_DEFLATED) as pack:
        for path in paths.values():
            pack.write(path, path.name)
    return paths, archive, totals, issues

def register_direct_data(app, export_dir, get_engine=None):
    root = Path(export_dir) / 'direct_preparation'
    root.mkdir(parents=True, exist_ok=True)
    store = DirectStore(Path(export_dir)/'direct_analytics.sqlite',get_engine)
    from central_store import CentralStore
    store.central=CentralStore(store)
    app.extensions['central_store']=store.central
    app.extensions['direct_store']=store
    register_direct_analytics(app,store)

    def cleanup():
        for folder in root.iterdir():
            if (folder.is_dir() and not folder.is_symlink() and re.fullmatch(r'[a-f0-9]{32}', folder.name)
                    and folder.resolve().parent == root.resolve()
                    and datetime.now() - datetime.fromtimestamp(folder.stat().st_mtime) > timedelta(hours=12)):
                shutil.rmtree(folder)

    @app.get('/direct-data/reports')
    def direct_reports():
        user=getattr(g,'user',None) or {}
        if user.get('role')!='admin':return 'Admin access required.',403
        return redirect(url_for('overview_page',section='downloads')+'#preparedDownloads')

    @app.post('/direct-data/reports/<batch_id>')
    def direct_recreate(batch_id):
        user=getattr(g,'user',None) or {}
        if user.get('role')!='admin':return 'Admin access required.',403
        if request.headers.get('X-Import-Progress')=='stream' and not active():
            authenticated_user=dict(user)
            def run():
                g.user=authenticated_user
                return direct_recreate(batch_id)
            return stream_import(run)
        batch=next((b for b in store.batches(user['id']) if b['id']==batch_id),None)
        if batch is None:return jsonify(ok=False,error='Saved month not found.'),400
        for candidate in root.iterdir():
            if not candidate.is_dir() or candidate.is_symlink() or not re.fullmatch(r'[a-f0-9]{32}',candidate.name):continue
            try:
                metadata=json.loads((candidate/'owner.json').read_text())
                if metadata.get('owner')!=user['id'] or metadata.get('batch',candidate.name)!=batch_id:continue
                if datetime.now()-datetime.fromisoformat(metadata['created'])>timedelta(hours=12):continue
                files=json.loads((candidate/'files.json').read_text())
                if files and all(Path(name).name==name and (candidate/name).is_file() for name in files):
                    return jsonify(ok=True,job=candidate.name,month=batch['month'],files=files)
            except (OSError,ValueError,KeyError):continue
        if not LOCK.acquire(blocking=False):return jsonify(ok=False,error='Another Direct preparation is running. Wait for it to finish.'),409
        folder=root/uuid.uuid4().hex
        folder.mkdir()
        try:
            cleanup()
            with closing(sqlite3.connect(folder/'rows.db')) as conn:
                batch=store.restore_stage(user['id'],batch_id,conn)
                paths,archive,totals,issues=prepare_outputs(conn,folder,batch['source_file'],saved=True)
            (folder/'rows.db').unlink()
            files=[p.name for p in paths.values()]+[archive.name]
            (folder/'owner.json').write_text(json.dumps({'owner':user['id'],'batch':batch_id,'created':datetime.now().isoformat()}))
            (folder/'files.json').write_text(json.dumps(files))
            return jsonify(ok=True,job=folder.name,month=batch['month'],files=files)
        except ValueError as exc:return jsonify(ok=False,error=str(exc)),400
        except Exception:
            app.logger.exception('Saved Direct report recreation failed')
            return jsonify(ok=False,error='Could not recreate these downloads. Please retry.'),500
        finally:LOCK.release()

    @app.route('/direct-data', methods=['GET', 'POST'])
    def direct_data():
        user = getattr(g, 'user', None) or {}
        if user.get('role') != 'admin':
            return 'Admin access required.', 403
        if app.extensions.get('central_jobs'):
            if request.method=='GET':return redirect(url_for('overview_page',section='imports'))
            return jsonify(ok=False,error='Use the consolidated Imports tab. It validates, stores and calculates Direct and paid-claims source files in the background.'),410
        if request.method == 'GET':
            return render_template('direct_data.html')
        if request.headers.get('X-Import-Progress') == 'stream' and not active():
            authenticated_user = dict(user)
            def run_preparation():
                g.user = authenticated_user
                return direct_data()
            return stream_import(run_preparation)
        if not LOCK.acquire(blocking=False):
            return jsonify(ok=False, error='Another Direct preparation is running. Wait for it to finish.'), 409
        folder = root / uuid.uuid4().hex
        folder.mkdir()
        try:
            cleanup()
            upload = request.files.get('direct_file')
            if not upload or not upload.filename:
                raise ValueError('Select a Direct CSV, text file or converted Excel workbook.')
            suffix = Path(upload.filename).suffix.lower()
            if suffix not in ('.csv', '.txt', '.xlsx'):
                raise ValueError('Use CSV, TXT or XLSX.')
            month = data_month(upload.filename,request.form.get('data_month',''))
            path = folder / ('source' + suffix)
            upload.save(path)
            with closing(sqlite3.connect(folder / 'rows.db')) as conn:
                source = stage_source(path, conn, request.form.get('encoding', 'auto'), request.form.get('excel_sheet') or None)
                paths, archive, totals, issues = prepare_outputs(conn, folder, Path(upload.filename).name,
                    'mpia', 'mpia', 'policy',generate_files=False)
                report('Saving Direct monthly analytics',source['rows'],source['rows'])
                store.save(conn,folder.name,user.get('id'),month,Path(upload.filename).name)
            path.unlink()
            (folder / 'rows.db').unlink()
            # Keep no identifiers in cookies. Downloads require this admin's session and opaque job ID.
            (folder / 'owner.json').write_text(json.dumps({'owner': user.get('id'), 'batch': folder.name, 'created': datetime.now().isoformat()}))
            (folder / 'files.json').write_text(json.dumps([p.name for p in paths.values()] + ([archive.name] if archive else [])))
            source_json = {k: str(v) if isinstance(v, Decimal) else v for k, v in source.items()}
            return jsonify(ok=True, source=source_json, issues=len(issues), job=folder.name, month=month,
                           files=[p.name for p in paths.values()] + ([archive.name] if archive else []))
        except (ValueError, csv.Error, UnicodeError) as exc:
            return jsonify(ok=False, error=str(exc)), 400
        except Exception:
            app.logger.exception('Direct data preparation failed')
            return jsonify(ok=False, error='Direct preparation failed; check the server logs.'), 500
        finally:
            LOCK.release()

    def owned_file(job,name):
        user = getattr(g, 'user', None) or {}
        if user.get('role') != 'admin' or not re.fullmatch(r'[a-f0-9]{32}', job):
            return None
        folder = root / job
        owner = folder / 'owner.json'
        if not owner.exists():
            return None
        metadata = json.loads(owner.read_text())
        if metadata['owner'] != user.get('id') or datetime.now() - datetime.fromisoformat(metadata['created']) > timedelta(hours=12):
            return None
        if not any(batch['id']==metadata.get('batch',job) for batch in store.batches(user['id'])):
            return None
        allowed = json.loads((folder / 'files.json').read_text()) if (folder / 'files.json').exists() else []
        if name not in allowed:
            return None
        path=folder/name
        return path if path.is_file() else None

    @app.get('/direct-data/download/<job>/<name>')
    def direct_download(job,name):
        path=owned_file(job,name)
        if path is None:return 'Open Saved Direct Reports to recreate expired downloads.',404
        return send_file(path,as_attachment=True,download_name=name,max_age=0)

    @app.get('/direct-data/report/<job>/<name>')
    def direct_report_view(job,name):
        path=owned_file(job,name)
        if path is None or path.suffix.lower()!='.xlsx':return 'Open Saved Direct Reports to recreate expired downloads.',404
        from report_files import workbook_preview,workbook_pdf
        title=path.stem
        if request.args.get('format')=='pdf':
            return send_file(workbook_pdf(path,title),as_attachment=True,download_name=title+'.pdf',mimetype='application/pdf',max_age=0)
        try:preview=workbook_preview(path,request.args.get('sheet',''),request.args.get('page',1,type=int))
        except ValueError as exc:return str(exc),400
        def page_url(page):return url_for('direct_report_view',job=job,name=name,sheet=preview['sheet'],page=page)
        return render_template('workbook_report.html',title=title,preview=preview,excel_url=url_for('direct_download',job=job,name=name),pdf_url=url_for('direct_report_view',job=job,name=name,format='pdf'),previous=page_url(preview['page']-1) if preview['page']>1 else '',next_page=page_url(preview['page']+1) if preview['page']*50<preview['count'] else '')
