"""Stock reconciliation: ERP stock report against either a 3PL inventory
report or a physical stock take count, with a suggested reason for every
difference in both directions and a net that ties back to the report totals.

Settings come from the "reconciliation" section of config.json.
"""
import collections
import csv
import datetime
import os
import re

from stocktools import config as _config

S = _config.load().get("reconciliation", {})
WAREHOUSE_OWNER = S.get("erp_warehouse_owner", {})
SITE_WAREHOUSE_OWNER = {(a, b): c for a, b, c in S.get("erp_site_warehouse_owner", [])}
BIN_OWNER = {(a, b.upper()): c for a, b, c in S.get("erp_bin_owner", [])}
BIN_WAREHOUSES = {a for a, _ in BIN_OWNER}
WHSE_CODE_OWNER = S.get("warehouse_code_owner", {})
BIN_CODE_OWNER = {(a, b.upper()): c for a, b, c in S.get("warehouse_code_bin_owner", [])}
NON_INVENTORY = set(S.get("non_inventory_parts", []))
HOLD_KEYWORD = S.get("hold_keyword", "").lower()
HOLD_PARTS = set(S.get("hold_parts", []))
RECENT_DAYS = S.get("recent_days", 3)
PENDING_BIN = S.get("pending_receipt_bin", "RCV-01").upper()
KPI_FOLDER = S.get("kpi_folder", "")
ACCURACY_TOLERANCE_PCT = S.get("accuracy_tolerance_pct", 0.0)

R_NONINV = 'Non-inventory'
R_HOLD = S.get('hold_reason', 'Product on hold, awaiting release')
R_INVOICED = 'Invoiced awaiting customer collection'
R_RECEIVED_3PL = 'Received at the 3PL not processed in ERP'
R_OTHER_SITE = 'ERP shows it at an onsite warehouse'
R_NOT_AT_3PL = 'Part not held at the 3PL in ERP'
R_TRANSIT = 'In transit'
R_OFFSET = 'Booked to a different lot or area'
R_DESPATCHED = 'Despatched by the 3PL, not yet moved in ERP'
R_RECEIVED_EPI = 'Received in ERP, not yet at the 3PL'
R_TBD = 'To Be Determined'
TPL_REASONS = [R_NONINV, R_INVOICED, R_RECEIVED_3PL, R_OTHER_SITE, R_NOT_AT_3PL, R_TRANSIT,
              R_OFFSET, R_HOLD, R_TBD]
EPI_REASONS = [R_NONINV, R_DESPATCHED, R_OFFSET, R_RECEIVED_EPI, R_TRANSIT, R_HOLD, R_TBD]


# ---------------------------------------------------------------- helpers
def clean_lot(v):
    """Normalise a lot: upper case, no spaces, blanks/N/A/NONE -> ''."""
    if v is None:
        return ''
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    s = str(v).strip().upper().replace(' ', '')
    if s.startswith('LOT') and s[3:].isdigit():
        s = s[3:]
    return '' if s in ('', 'N/A', 'NA', 'NULL', '-', 'NONE') else s


def clean_part(v):
    """Undo Excel damage to part codes: lost leading zeros, codes read as dates."""
    if isinstance(v, datetime.datetime):
        return '%02d-%d' % (v.month, v.year)
    if isinstance(v, (int, float)):
        return str(int(v)).zfill(6)
    return str(v or '').strip()


def nz(s):
    return str(s).lstrip('0')


def num(v):
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).replace(',', '').strip())
    except Exception:
        return None


def to_date(v):
    if isinstance(v, datetime.datetime):
        return v.date()
    if isinstance(v, datetime.date):
        return v
    for f in ('%d/%m/%Y', '%Y-%m-%d', '%d/%m/%y'):
        try:
            return datetime.datetime.strptime(str(v).strip()[:10], f).date()
        except Exception:
            pass
    return None


def header_map(row):
    return {str(h).strip().lower(): i for i, h in enumerate(row) if h is not None}


def col(hm, *names):
    for n in names:
        if n.lower() in hm:
            return hm[n.lower()]
    raise ValueError('Column not found: ' + ' / '.join(names))


def is_noninv(part):
    return part in NON_INVENTORY or not re.search(r'\d', part)


def is_on_hold(part, desc):
    """Parts held under one reason, e.g. a new product awaiting release."""
    return part in HOLD_PARTS or bool(HOLD_KEYWORD and HOLD_KEYWORD in str(desc or '').lower())


def erp_owner(site, whse, bin_):
    """Map an ERP site/warehouse/bin to a 3PL owner, or None if not at the 3PL."""
    if (site, whse) in SITE_WAREHOUSE_OWNER:
        return SITE_WAREHOUSE_OWNER[(site, whse)]
    if whse in BIN_WAREHOUSES:
        return BIN_OWNER.get((whse, str(bin_).strip().upper()))
    return WAREHOUSE_OWNER.get(whse)


# ---------------------------------------------------------------- loading
def load_erp(path):
    from openpyxl import load_workbook
    ws = load_workbook(path, read_only=True, data_only=True).worksheets[0]
    rows = ws.iter_rows(values_only=True)
    hm = header_map(next(rows))
    c = dict(site=col(hm, 'Site'), wh=col(hm, 'Warehouse'), part=col(hm, 'Part'),
             desc=col(hm, 'Description'), bin=col(hm, 'Bin'), lot=col(hm, 'Lot Number'),
             qty=col(hm, 'On Hand'), uom=col(hm, 'UOM'), val=col(hm, 'Value On Hand'))
    qty = collections.defaultdict(float)
    lots = collections.defaultdict(set)
    desc, other_site, val = {}, set(), collections.defaultdict(lambda: [0.0, 0.0])
    excluded = 0.0
    for r in rows:
        q = r[c['qty']] if len(r) > c['qty'] else None
        if not r[c['site']] or not isinstance(q, (int, float)):
            continue
        part = clean_part(r[c['part']])
        desc[part] = r[c['desc']]
        v = num(r[c['val']])
        if v is not None and q > 0:
            val[part][0] += v; val[part][1] += q
        g = erp_owner(r[c['site']], r[c['wh']], r[c['bin']])
        if g is None:
            excluded += q
            if q > 0 and r[c['wh']] != 'Ecomm Warehouse':
                other_site.add(part)
            continue
        lot = clean_lot(r[c['lot']])
        qty[(g, part, lot)] += q
        lots[(g, part)].add(lot)
    unit = {p: a[0] / a[1] for p, a in val.items() if a[1]}
    return qty, lots, desc, other_site, unit, excluded


def load_3pl(path, epi_lots):
    from openpyxl import load_workbook
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb['Full inventory'] if 'Full inventory' in wb.sheetnames else wb.worksheets[0]
    rows = ws.iter_rows(values_only=True)
    hm = header_map(next(rows))
    c = dict(owner=col(hm, 'Owner'), sku=col(hm, 'Sku', 'Stock Code'), desc=col(hm, 'Description'),
             qty=col(hm, 'Qty', 'Total'), alloc=col(hm, 'Allocated'), lot=col(hm, 'Batch', 'Batch_id'),
             rdate=col(hm, 'Receive date', 'In Date'))
    qty, alloc, method, rdate, desc = (collections.defaultdict(float), collections.defaultdict(float),
                                       collections.defaultdict(set), {}, {})
    for r in rows:
        g, q = r[c['owner']], r[c['qty']]
        if not g or not isinstance(q, (int, float)):
            continue
        g = str(g).strip(); part = clean_part(r[c['sku']]); lot = clean_lot(r[c['lot']])
        desc[part] = r[c['desc']]
        known = epi_lots.get((g, part), set())
        if known and known == {''}:
            key_lot, m = '', 'part level (ERP part not lot-tracked)'
        elif lot in known:
            key_lot, m = lot, 'exact'
        else:
            z = [x for x in known if lot and nz(x) == nz(lot)]
            if len(z) == 1:
                key_lot, m = z[0], 'ignoring leading zeros'
            elif '-' in lot and lot.rsplit('-', 1)[0] in known:
                key_lot, m = lot.rsplit('-', 1)[0], 'ignoring 3PL suffix'
            else:
                key_lot, m = lot, ('exact' if known else 'part not in ERP 3PL warehouses')
        k = (g, part, key_lot)
        qty[k] += q
        alloc[k] += num(r[c['alloc']]) or 0
        method[k].add(m)
        d = to_date(r[c['rdate']])
        if d and (k not in rdate or d > rdate[k]):
            rdate[k] = d
    return qty, alloc, method, rdate, desc


def load_transactions(path):
    """ERP transaction BAQ export (tab-delimited, PartTran layout)."""
    shipped, receipts = collections.defaultdict(float), collections.defaultdict(list)
    mtl = {}   # first positive Mtl Unit Cost per part: fallback value for parts with no cost on the stock report
    if not path:
        return shipped, receipts, mtl
    with open(path, encoding='utf-8-sig') as f:
        rows = list(csv.reader(f, delimiter='\t'))
    hdr = [h.strip() for h in rows[0]]
    idx = {}
    for i, h in enumerate(hdr):
        idx.setdefault(h, i)
    tcol = [i for i, h in enumerate(hdr) if h == 'Type'][-1]   # second Type = tran type
    for r in rows[1:]:
        if len(r) < len(hdr):
            continue
        if 'Mtl Unit Cost' in idx and (num(r[idx['Mtl Unit Cost']]) or 0) > 0:   # any site: cost is per part
            mtl.setdefault(r[idx['Part']].strip(), num(r[idx['Mtl Unit Cost']]))
        wc, bc = r[idx['Whse']].strip(), r[idx['Bin']].strip().upper()
        g = BIN_CODE_OWNER.get((wc, bc)) if any(w == wc for w, _ in BIN_CODE_OWNER) else WHSE_CODE_OWNER.get(wc)
        if not g:
            continue
        part, lot, tt = r[idx['Part']].strip(), clean_lot(r[idx['Lot Number']]), r[tcol]
        q = num(r[idx['Quantity']]) or 0
        if tt == 'STK-CUS':
            shipped[(g, part, nz(lot))] += q
        elif tt in ('MFG-STK', 'PUR-STK', 'PLT-STK', 'INS-STK') and q > 0:
            receipts[(g, part, nz(lot))].append((r[idx['Sys Date']], tt, q))
    return shipped, receipts, mtl


def load_3pl_orders(paths):
    desp, orders = collections.defaultdict(float), collections.defaultdict(set)
    for p in paths:
        with open(p, encoding='utf-8-sig') as f:
            rows = list(csv.reader(f))
        hm = header_map(rows[0])
        for r in rows[1:]:
            try:
                q = num(r[col(hm, 'Qty Shipped')])
                if not q:
                    continue
                k = (clean_part(r[col(hm, 'SKU ID')]), nz(clean_lot(r[col(hm, 'Batch ID')])))
                desp[k] += q
                orders[k].add(r[col(hm, 'Order ID')])
            except (IndexError, ValueError):
                continue
    return desp, orders


# ---------------------------------------------------------------- reasons
def tpl_reason(k, e, d, ctx):
    g, part, lot = k; diff = d - e; desc = ctx['desc'].get(part)
    if is_on_hold(part, desc): return R_HOLD, ''
    if is_noninv(part): return R_NONINV, ''
    s = ctx['shipped'].get((g, part, nz(lot)), 0)
    if s: return R_INVOICED, f'Shipped in ERP: {s:,.0f}'
    if 'part not in ERP 3PL warehouses' in ctx['method'].get(k, ()):
        if part in ctx['other_site']: return R_OTHER_SITE, 'ERP holds this part at an onsite warehouse, not at the 3PL'
        return R_NOT_AT_3PL, ''
    if abs(ctx['alloc'].get(k, 0) - diff) < 1e-6: return R_INVOICED, '3PL has the exact difference allocated to orders'
    rd = ctx['rdate'].get(k)
    if rd and ctx['report_date'] and (ctx['report_date'] - rd).days <= RECENT_DAYS:
        return R_RECEIVED_3PL, f'3PL receive date {rd:%d/%m/%Y}'
    rc = ctx['receipts'].get((g, part, nz(lot)), [])
    if any(t == 'PLT-STK' for _, t, _ in rc) and g != S.get('main_owner'):
        return R_TRANSIT, 'Transfer to this owner in ERP'
    off = [x for x in ctx['epi_higher'] if x[0][1] == part and abs(x[1] - diff) < 1e-6]
    if off:
        o = off[0][0]; return R_OFFSET, f'ERP holds the same qty under {o[0]} lot {o[2] or "(none)"}'
    return R_TBD, ''


def epi_reason(k, e, d, ctx):
    g, part, lot = k; diff = e - d; desc = ctx['desc'].get(part)
    if is_on_hold(part, desc): return R_HOLD, ''
    if is_noninv(part): return R_NONINV, ''
    off = [x for x in ctx['dt_higher'] if x[0][1] == part and abs(x[1] - diff) < 1e-6]
    if off:
        o = off[0][0]; return R_OFFSET, f'3PL holds the same qty under {o[0]} lot {o[2] or "(none)"}'
    dk = (part, nz(lot))
    if ctx['desp'].get(dk, 0) >= diff - 1e-6:
        return R_DESPATCHED, '3PL order ' + ', '.join(sorted(ctx['orders'][dk]))
    rc = ctx['receipts'].get((g, part, nz(lot)), [])
    if rc and sum(q for _, _, q in rc) >= diff - 1e-6:
        last = rc[-1]
        if last[1] == 'PLT-STK' and g != S.get('main_owner'):
            return R_TRANSIT, f'Transferred in ERP {last[0]}'
        return R_RECEIVED_EPI, f'{last[1]} {last[0]} qty {last[2]:,.0f}'
    if ctx['desp'].get(dk):
        return R_TBD, f'3PL shipped {ctx["desp"][dk]:,.0f}, less than the difference'
    return R_TBD, ''


# ---------------------------------------------------------------- workbook
def write_workbook(out_path, title, a_label, b_label, lines, rows_ahi, rows_bhi, reasons_ahi, reasons_bhi,
                   key_headers, notes, log=print):
    """lines: [(key tuple, a_qty, b_qty, match note)]; rows_*: [area, part, desc, lot, a, b, None, unit, None, reason, note]"""
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill
    from openpyxl.utils import get_column_letter as L
    F, B = Font(name='Arial', size=10), Font(name='Arial', size=10, bold=True)
    H = PatternFill('solid', fgColor='DDEBF7'); NUM = '#,##0;-#,##0;0'
    wb = Workbook(); summ = wb.active; summ.title = 'Summary'

    def sheet(name, hdr, rows, widths):
        ws = wb.create_sheet(name); ws.append(hdr)
        for r in rows: ws.append(r)
        for row in ws.iter_rows():
            for c in row:
                c.font = B if c.row == 1 else F
                if c.row == 1: c.fill = H
        for i, w in enumerate(widths, 1): ws.column_dimensions[L(i)].width = w
        ws.freeze_panes = 'A2'; ws.auto_filter.ref = ws.dimensions
        return ws

    def detail(name, rows, reasons, higher):
        rows.sort(key=lambda r: (reasons.index(r[9]), -abs(r[5] - r[4])))
        ws = sheet(name, ['Area', 'Part', 'Description', 'Batch', a_label + ' Qty', b_label + ' Qty',
                          higher + ' higher by', 'Unit cost (GBP)', 'Value (GBP)', 'Reason', 'Notes'],
                   rows, [16, 16, 46, 16, 13, 13, 15, 13, 13, 40, 60])
        for j in range(2, len(rows) + 2):
            ws.cell(j, 7).value = f'=E{j}-F{j}' if higher == a_label else f'=F{j}-E{j}'
            ws.cell(j, 9).value = f'=IF(H{j}="","",G{j}*H{j})'
            for cc, fm in ((5, NUM), (6, NUM), (7, NUM), (8, '0.00000'), (9, '#,##0.00')):
                ws.cell(j, cc).number_format = fm; ws.cell(j, cc).font = F
        return len(rows) + 1

    tab_b, tab_a = f'{b_label} higher', f'{a_label} higher'
    n1 = detail(tab_b, rows_bhi, reasons_bhi, b_label)
    n2 = detail(tab_a, rows_ahi, reasons_ahi, a_label)
    sheet('Comparison', key_headers + ['Description', a_label + ' Qty', b_label + ' Qty', f'{b_label} - {a_label}', 'Lot match'],
          [list(k) + [desc, a, b, b - a, m] for k, a, b, desc, m in lines],
          [14] * len(key_headers) + [46, 13, 13, 13, 34])
    summ['A1'] = title; summ['A1'].font = Font(name='Arial', size=14, bold=True)
    r = 3

    def block(tab, n, reasons):
        nonlocal r
        summ.cell(r, 1, tab + ' than ' + (a_label if tab.startswith(b_label) else b_label))
        summ.cell(r, 2, 'Qty'); summ.cell(r, 3, 'Value (GBP)')
        for c in range(1, 4): summ.cell(r, c).font = B; summ.cell(r, c).fill = H
        start = r + 1
        for rs in reasons:
            r += 1
            summ.cell(r, 1, rs)
            summ.cell(r, 2, f"=SUMIF('{tab}'!$J$2:$J${n},\"{rs}\",'{tab}'!$G$2:$G${n})")
            summ.cell(r, 3, f"=SUMIF('{tab}'!$J$2:$J${n},\"{rs}\",'{tab}'!$I$2:$I${n})")
        r += 1
        summ.cell(r, 1, 'Total ' + tab)
        summ.cell(r, 2, f'=SUM(B{start}:B{r - 1})'); summ.cell(r, 3, f'=SUM(C{start}:C{r - 1})')
        tot = r; r += 2
        return tot

    t1 = block(tab_b, n1, reasons_bhi)
    t2 = block(tab_a, n2, reasons_ahi)
    summ.cell(r, 1, f'Net difference ({b_label} higher less {a_label} higher)')
    summ.cell(r, 2, f'=B{t1}-B{t2}'); summ.cell(r, 3, f'=C{t1}-C{t2}'); net = r
    r += 2
    for txt in notes + ['Quantities are added as recorded in each system (no unit conversion).',
                        'Reasons are suggested by rules. Review every line before sharing.']:
        summ.cell(r, 1, txt); r += 1
    for row in summ.iter_rows(min_row=3):
        for c in row:
            if not (c.fill and c.fill.fgColor and c.fill.fgColor.rgb == '00DDEBF7'):
                c.font = B if c.row in (t1, t2, net) else F
            if c.column == 2: c.number_format = NUM
            if c.column == 3: c.number_format = '#,##0;-#,##0'
    summ.column_dimensions['A'].width = 62; summ.column_dimensions['B'].width = 15; summ.column_dimensions['C'].width = 15
    wb.save(out_path)
    log(f'{tab_b} lines: {len(rows_bhi)}  {tab_a} lines: {len(rows_ahi)}')
    log('Saved: ' + out_path)


def match_lot(lot, known):
    """Match a counted/3PL lot to the ERP lots for the same part and location."""
    if known and known == {''}:
        return '', 'part level (ERP part not lot-tracked)'
    if lot in known:
        return lot, 'exact'
    z = [x for x in known if lot and nz(x) == nz(lot)]
    if len(z) == 1:
        return z[0], 'ignoring leading zeros'
    if '-' in lot and lot.rsplit('-', 1)[0] in known:
        return lot.rsplit('-', 1)[0], 'ignoring suffix'
    return lot, ('exact' if known else 'part not in ERP here')


# ---------------------------------------------------------------- KPI logging
def append_csv(folder, name, header, rows, log=print):
    """Append rows to a CSV log, writing the header if the file is new. If the log is open in
    Excel (locked), the rows go to a dated side file instead so nothing is lost."""
    if not rows: return
    try:
        os.makedirs(folder, exist_ok=True)
    except OSError:
        log(f'KPI log folder {folder} could not be reached. Nothing logged for {name}: check the drive and run again.')
        return
    path = os.path.join(folder, name)
    try:
        new = not os.path.exists(path)
        with open(path, 'a', newline='', encoding='utf-8') as f:
            w = csv.writer(f)
            if new: w.writerow(header)
            w.writerows(rows)
    except PermissionError:
        alt = os.path.join(folder, name.replace('.csv', f'_{datetime.datetime.now():%Y%m%d_%H%M%S}.csv'))
        with open(alt, 'w', newline='', encoding='utf-8') as f:
            w = csv.writer(f); w.writerow(header); w.writerows(rows)
        log(f'KPI log {name} is open or locked. Rows saved to {os.path.basename(alt)}: add them to the main log later.')


def run_by():
    try:
        import getpass; return getpass.getuser()
    except Exception: return os.environ.get('USERNAME', '')


def log_kpis(tool, report_date, scope, files, lines, rows_other_hi, rows_epi_hi, area_of, kpi_folder, log=print):
    """lines: [(key, erp_qty, other_qty)]. rows_*: detail rows [area, part, desc, lot, erp, other, _, unit, _, reason, note]."""
    run_id = 'R' + datetime.datetime.now().strftime('%Y%m%d-%H%M%S-%f')[:-3]   # letters keep Excel from showing it as 2.02611E+16
    rd = report_date.strftime('%d/%m/%Y') if report_date else ''
    tol = ACCURACY_TOLERANCE_PCT / 100.0

    def ok(a, b): return abs(b - a) <= tol * max(abs(a), abs(b)) + 1e-9
    n = len(lines); matched = sum(1 for _, a, b in lines if ok(a, b))
    tot_e = sum(a for _, a, _ in lines); tot_o = sum(b for _, _, b in lines)
    val = lambda r, q: q * r[7] if r[7] else 0.0
    o_qty = sum(r[5] - r[4] for r in rows_other_hi); e_qty = sum(r[4] - r[5] for r in rows_epi_hi)
    o_val = sum(val(r, r[5] - r[4]) for r in rows_other_hi); e_val = sum(val(r, r[4] - r[5]) for r in rows_epi_hi)
    gross = o_qty + e_qty
    tbd = sum(r[5] - r[4] for r in rows_other_hi if r[9] == R_TBD) + sum(r[4] - r[5] for r in rows_epi_hi if r[9] == R_TBD)
    base = max(tot_e, tot_o) or 1
    r2 = lambda x: round(x, 4)
    append_csv(kpi_folder, 'recon_runs_log.csv',
               ['Run ID', 'Run date', 'Run by', 'Tool', 'Report date', 'Scope', 'Files',
                'Lines compared', 'Lines accurate', 'Line accuracy %', 'Tolerance %',
                'ERP qty', 'Other qty', 'Other higher qty', 'ERP higher qty', 'Gross variance qty',
                'Net variance qty', 'Qty accuracy %', 'Other higher value', 'ERP higher value',
                'Gross variance value', 'Net variance value', 'Explained %'],
               [[run_id, f'{datetime.datetime.now():%d/%m/%Y %H:%M}', run_by(), tool, rd, scope, files,
                 n, matched, r2(100 * matched / n) if n else '', ACCURACY_TOLERANCE_PCT,
                 r2(tot_e), r2(tot_o), r2(o_qty), r2(e_qty), r2(gross), r2(o_qty - e_qty),
                 r2(100 * (1 - gross / base)), round(o_val, 2), round(e_val, 2), round(o_val + e_val, 2),
                 round(o_val - e_val, 2), r2(100 * (1 - tbd / gross)) if gross else 100]], log)
    rs = collections.defaultdict(lambda: [0, 0.0, 0.0])
    for direction, rows, sign in (('Other higher', rows_other_hi, 1), ('ERP higher', rows_epi_hi, -1)):
        for r in rows:
            q = (r[5] - r[4]) * sign
            x = rs[(direction, r[9])]; x[0] += 1; x[1] += q; x[2] += val(r, q)
    append_csv(kpi_folder, 'recon_reasons_log.csv',
               ['Run ID', 'Tool', 'Report date', 'Direction', 'Reason', 'Lines', 'Qty', 'Value'],
               [[run_id, tool, rd, d, rsn, v[0], r2(v[1]), round(v[2], 2)] for (d, rsn), v in sorted(rs.items())], log)
    ar = collections.defaultdict(lambda: [0, 0, 0.0, 0.0])
    for k, a, b in lines:
        x = ar[area_of(k)]; x[0] += 1; x[1] += ok(a, b); x[2] += abs(b - a); x[3] += max(a, b)
    append_csv(kpi_folder, 'recon_area_log.csv',
               ['Run ID', 'Tool', 'Report date', 'Area', 'Lines compared', 'Lines accurate', 'Line accuracy %',
                'Gross variance qty', 'Qty accuracy %'],
               [[run_id, tool, rd, a, v[0], v[1], r2(100 * v[1] / v[0]) if v[0] else '', r2(v[2]),
                 r2(100 * (1 - v[2] / v[3])) if v[3] else 100] for a, v in sorted(ar.items())], log)
    log(f'KPIs logged to {kpi_folder} (line accuracy {100 * matched / n:.1f}%, explained {100 * (1 - tbd / gross) if gross else 100:.1f}%)' if n else 'KPIs logged')


# ---------------------------------------------------------------- mode 1: 3PL
def run_3pl(erp_path, tpl_path, out_path, txn_path=None, order_paths=(), report_date=None, log=print, kpi_folder=None):
    log('Reading ERP report...')
    eq, elots, edesc, other_site, unit, excluded = load_erp(erp_path)
    log('Reading 3PL report...')
    dq, alloc, method, rdate, ddesc = load_3pl(tpl_path, elots)
    shipped, receipts, mtl = load_transactions(txn_path)
    costed_from_txn = {p: c for p, c in mtl.items() if p not in unit}
    unit = {**costed_from_txn, **unit}
    if costed_from_txn:
        log(f'  {len(costed_from_txn)} part(s) not costed on the stock report: valued at Mtl Unit Cost from the transactions')
    desp, orders = load_3pl_orders(order_paths)
    desc = dict(edesc); desc.update({k: v for k, v in ddesc.items() if v})
    keys = sorted(set(eq) | set(dq))
    lines = [(k, round(eq.get(k, 0), 6), dq.get(k, 0)) for k in keys]
    dt_higher = [(k, d - e) for k, e, d in lines if d - e > 1e-9]
    epi_higher = [(k, e - d) for k, e, d in lines if e - d > 1e-9]
    ctx = dict(desc=desc, shipped=shipped, method=method, other_site=other_site, alloc=alloc,
               rdate=rdate, report_date=report_date, receipts=receipts, desp=desp, orders=orders,
               dt_higher=dt_higher, epi_higher=epi_higher)
    log('Classifying differences...')
    rows_tpl, rows_ep = [], []
    for k, e, d in lines:
        g, part, lot = k
        base = [g, part, desc.get(part, ''), lot, e, d, None, unit.get(part), None]
        if d - e > 1e-9:
            rows_tpl.append(base + list(tpl_reason(k, e, d, ctx)))
        elif e - d > 1e-9:
            rows_ep.append(base + list(epi_reason(k, e, d, ctx)))
    tot_e, tot_d = sum(eq.values()), sum(dq.values())
    notes = [f'3PL total: {tot_d:,.1f}', f'ERP total (3PL warehouses): {tot_e:,.1f}',
             f'Check: 3PL total less ERP total = {tot_d - tot_e:,.1f} (should equal the net difference)',
             f'ERP stock outside 3PL warehouses (not compared): {excluded:,.1f}',
             f'ERP report: {os.path.basename(erp_path)}', f'3PL report: {os.path.basename(tpl_path)}',
             f'Transactions: {os.path.basename(txn_path) if txn_path else "not supplied"}',
             'Value: ERP cost per unit from the stock report. Parts with no cost on the report use the first Mtl Unit Cost '
             'in the transactions file, if one was supplied; otherwise they have no value.',
             f'3PL order summaries: {len(order_paths)} file(s)']
    lines_out = [(k, e, d, desc.get(k[1], ''), '; '.join(sorted(method.get(k, [])))) for k, e, d in lines]
    write_workbook(out_path, '3PL vs ERP stock comparison', 'ERP', '3PL', lines_out, rows_ep, rows_tpl,
                   EPI_REASONS, TPL_REASONS, ['Owner', 'Part', 'Lot'], notes, log)
    log(f'3PL total {tot_d:,.1f}  ERP total {tot_e:,.1f}  Net {tot_d - tot_e:,.1f}')
    log_kpis('3PL vs ERP', report_date, '3PL', f'{os.path.basename(erp_path)}; {os.path.basename(tpl_path)}',
             lines, rows_tpl, rows_ep, lambda k: k[0], kpi_folder or KPI_FOLDER or os.path.dirname(os.path.abspath(out_path)), log)
    return out_path


# ---------------------------------------------------------------- mode 2: warehouse stock take
S_FOUND = 'Found stock (not in ERP anywhere)'
S_OTHER_WH = 'ERP shows it in another warehouse'
S_BIN = 'In a different bin in ERP'
S_LOT = 'Booked to a different lot'
S_ISSUED = 'Issued or shipped in ERP, still on the shelf'
S_RCV = 'Stranded in a receiving bin'
S_RECEIVED = 'Received in ERP, not found in count'
S_NOTFOUND = 'Not found in count'
COUNT_HI = [R_NONINV, S_LOT, S_BIN, S_ISSUED, S_OTHER_WH, S_FOUND, R_TBD]
EPI_HI = [R_NONINV, S_LOT, S_BIN, S_RCV, S_RECEIVED, S_NOTFOUND, R_TBD]


def read_table(path, sheet=None):
    """Return header list and rows from an .xlsx or .csv file."""
    if path.lower().endswith('.csv'):
        with open(path, encoding='utf-8-sig') as f:
            rows = [r for r in csv.reader(f) if any(x.strip() for x in r)]
    else:
        from openpyxl import load_workbook
        wb = load_workbook(path, read_only=True, data_only=True)
        ws = wb[sheet] if sheet else wb.worksheets[0]
        rows = [r for r in ws.iter_rows(values_only=True) if any(v not in (None, '') for v in r)]
    return [str(h).strip() if h is not None else '' for h in rows[0]], rows[1:]


def erp_warehouses(path):
    """List (site, warehouse) pairs in an ERP stock report, for the picker."""
    hdr, rows = read_table(path)
    hm = header_map(hdr); si, wi, qi = col(hm, 'Site'), col(hm, 'Warehouse'), col(hm, 'On Hand')
    return sorted({(str(r[si]), str(r[wi])) for r in rows if r[si] and isinstance(r[qi], (int, float))})


def load_stocktake_txns(path):
    """ERP transactions keyed by warehouse name, part and lot."""
    out = collections.defaultdict(list)
    if not path:
        return out
    with open(path, encoding='utf-8-sig') as f:
        rows = list(csv.reader(f, delimiter='\t'))
    hdr = [h.strip() for h in rows[0]]; idx = {}
    for i, h in enumerate(hdr): idx.setdefault(h, i)
    tcol = [i for i, h in enumerate(hdr) if h == 'Type'][-1]
    for r in rows[1:]:
        if len(r) < len(hdr): continue
        out[(r[idx['Warehouse']].strip(), r[idx['Part']].strip(), nz(clean_lot(r[idx['Lot Number']])))].append(
            (r[idx['Sys Date']], r[tcol], num(r[idx['Quantity']]) or 0, r[idx['Bin']].strip()))
    return out


def run_stocktake(erp_path, count_path, out_path, scope, mapping, by_bin=False, txn_path=None,
                  count_sheet=None, log=print, kpi_folder=None, report_date=None):
    """scope: set of (site, warehouse) counted. mapping: count column names for part, lot, qty, warehouse, bin."""
    log('Reading ERP report...')
    hdr, erows = read_table(erp_path)
    hm = header_map(hdr)
    c = dict(site=col(hm, 'Site'), wh=col(hm, 'Warehouse'), part=col(hm, 'Part'), desc=col(hm, 'Description'),
             bin=col(hm, 'Bin'), lot=col(hm, 'Lot Number'), qty=col(hm, 'On Hand'), val=col(hm, 'Value On Hand'))
    scope_wh = {w for _, w in scope}
    single_wh = len(scope_wh) == 1
    use_wh = bool(mapping.get('warehouse'))
    use_bin = by_bin and bool(mapping.get('bin'))
    eq = collections.defaultdict(float); lots = collections.defaultdict(set); desc = {}
    anywhere = collections.defaultdict(list); val = collections.defaultdict(lambda: [0.0, 0.0])

    def loc(wh, bn):
        return ((wh if use_wh else '*'), (bn if use_bin else '*'))

    for r in erows:
        q = r[c['qty']]
        if not r[c['site']] or not isinstance(q, (int, float)): continue
        part = clean_part(r[c['part']]); lot = clean_lot(r[c['lot']]); desc[part] = r[c['desc']]
        v = num(r[c['val']])
        if v is not None and q > 0: val[part][0] += v; val[part][1] += q
        wh, bn = str(r[c['wh']]), str(r[c['bin']] or '').strip()
        anywhere[(part, nz(lot))].append((str(r[c['site']]), wh, bn, q))
        if (str(r[c['site']]), wh) not in scope: continue
        lw, lb = loc(wh, bn)
        eq[(lw, lb, part, lot)] += q; lots[(lw, lb, part)].add(lot)
    unit = {p: a[0] / a[1] for p, a in val.items() if a[1]}
    log('Reading count...')
    chdr, crows = read_table(count_path, count_sheet)
    ch = {h: i for i, h in enumerate(chdr)}
    cq = collections.defaultdict(float); method = collections.defaultdict(set)
    for r in crows:
        q = num(r[ch[mapping['qty']]]) if ch[mapping['qty']] < len(r) else None
        if q is None: continue
        part = clean_part(r[ch[mapping['part']]])
        if not part: continue
        lot = clean_lot(r[ch[mapping['lot']]]) if mapping.get('lot') else ''
        wh = str(r[ch[mapping['warehouse']]]).strip() if use_wh else '*'
        bn = str(r[ch[mapping['bin']]] or '').strip() if use_bin else '*'
        lw, lb = loc(wh, bn)
        known = lots.get((lw, lb, part), set()) if mapping.get('lot') else {''}
        if not mapping.get('lot'):
            k_lot, m = '', 'part level (no lot column)'
        else:
            k_lot, m = match_lot(lot, known)
        cq[(lw, lb, part, k_lot)] += q; method[(lw, lb, part, k_lot)].add(m)
    if not mapping.get('lot'):   # compare at part level
        e2 = collections.defaultdict(float)
        for (lw, lb, p, l), q in eq.items(): e2[(lw, lb, p, '')] += q
        eq = e2
    txns = load_stocktake_txns(txn_path)
    keys = sorted(set(eq) | set(cq))
    lines = [(k, round(eq.get(k, 0), 6), cq.get(k, 0)) for k in keys]
    c_hi = [(k, b - a) for k, a, b in lines if b - a > 1e-9]
    e_hi = [(k, a - b) for k, a, b in lines if a - b > 1e-9]
    log('Classifying differences...')

    def tx_for(k):
        lw, lb, part, lot = k
        whs = [lw] if lw != '*' else list(scope_wh)
        return [t for w in whs for t in txns.get((w, part, nz(lot)), [])]

    def count_reason(k, diff):
        lw, lb, part, lot = k
        if is_noninv(part): return R_NONINV, ''
        same = [x for x in e_hi if x[0][2] == part and abs(x[1] - diff) < 1e-6]
        for x in same:
            if x[0][3] != lot and x[0][:2] == k[:2]: return S_LOT, f'ERP holds the same qty under lot {x[0][3] or "(none)"}'
        for x in same:
            if x[0][3] == lot and x[0][1] != lb: return S_BIN, f'ERP holds the same qty in bin {x[0][1]}'
        out = [t for t in tx_for(k) if t[1] in ('STK-CUS', 'STK-MTL', 'STK-PLT', 'STK-UKN', 'STK-ASM') and t[2] > 0]
        if out: t = out[-1]; return S_ISSUED, f'{t[1]} {t[0]} qty {t[2]:,.0f}'
        elsewhere = [a for a in anywhere.get((part, nz(lot)), []) if (a[0], a[1]) not in scope and a[3] > 0]
        if elsewhere and eq.get(k, 0) == 0:
            a = elsewhere[0]; return S_OTHER_WH, f'ERP: {a[3]:,.0f} at {a[0]} / {a[1]} bin {a[2]}'
        if not anywhere.get((part, nz(lot))): return S_FOUND, ''
        return R_TBD, ''

    def epi_reason_st(k, diff):
        lw, lb, part, lot = k
        if is_noninv(part): return R_NONINV, ''
        same = [x for x in c_hi if x[0][2] == part and abs(x[1] - diff) < 1e-6]
        for x in same:
            if x[0][3] != lot and x[0][:2] == k[:2]: return S_LOT, f'Counted the same qty under lot {x[0][3] or "(none)"}'
        for x in same:
            if x[0][3] == lot and x[0][1] != lb: return S_BIN, f'Counted the same qty in bin {x[0][1]}'
        if lb.upper() == PENDING_BIN or (lb == '*' and any(a[2].upper() == PENDING_BIN and (a[0], a[1]) in scope
                                                         for a in anywhere.get((part, nz(lot)), []))):
            return S_RCV, 'Receiving bin is not a physical location: relocate, do not just adjust'
        rec = [t for t in tx_for(k) if t[1] in ('PUR-STK', 'MFG-STK', 'PLT-STK', 'INS-STK', 'STK-STK') and t[2] > 0]
        if rec: t = rec[-1]; return S_RECEIVED, f'{t[1]} {t[0]} qty {t[2]:,.0f}'
        if cq.get(k, 0) == 0: return S_NOTFOUND, ''
        return R_TBD, ''

    rows_c, rows_e = [], []
    for k, a, b in lines:
        lw, lb, part, lot = k
        area = ' / '.join(x for x in (lw, lb) if x != '*') or ('Counted area' if not single_wh else next(iter(scope_wh)))
        base = [area, part, desc.get(part, ''), lot, a, b, None, unit.get(part), None]
        if b - a > 1e-9: rows_c.append(base + list(count_reason(k, b - a)))
        elif a - b > 1e-9: rows_e.append(base + list(epi_reason_st(k, a - b)))
    tot_e, tot_c = sum(eq.values()), sum(cq.values())
    notes = [f'Count total: {tot_c:,.1f}', f'ERP total (counted warehouses): {tot_e:,.1f}',
             f'Check: count total less ERP total = {tot_c - tot_e:,.1f} (should equal the net difference)',
             'Counted warehouses: ' + '; '.join(f'{s} / {w}' for s, w in sorted(scope)),
             'Compared by: ' + ', '.join(x for x, on in (('warehouse', use_wh), ('bin', use_bin), ('part', True),
                                                         ('lot', bool(mapping.get('lot')))) if on),
             f'ERP report: {os.path.basename(erp_path)}', f'Count file: {os.path.basename(count_path)}',
             f'Transactions: {os.path.basename(txn_path) if txn_path else "not supplied"}']
    lines_out = [(k, a, b, desc.get(k[2], ''), '; '.join(sorted(method.get(k, [])))) for k, a, b in lines]
    write_workbook(out_path, 'Stock take: physical count vs ERP', 'ERP', 'Count', lines_out, rows_e, rows_c,
                   EPI_HI, COUNT_HI, ['Warehouse', 'Bin', 'Part', 'Lot'], notes, log)
    log(f'Count total {tot_c:,.1f}  ERP total {tot_e:,.1f}  Net {tot_c - tot_e:,.1f}')
    log_kpis('Stock take', report_date or datetime.date.today(), '; '.join(f'{s} / {w}' for s, w in sorted(scope)),
             f'{os.path.basename(erp_path)}; {os.path.basename(count_path)}', lines, rows_c, rows_e,
             lambda k: k[0] if k[0] != '*' else ('; '.join(sorted(scope_wh))),
             kpi_folder or KPI_FOLDER or os.path.dirname(os.path.abspath(out_path)), log)
    return out_path


