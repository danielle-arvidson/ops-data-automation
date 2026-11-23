"""Single SKU reconciliation: 3PL vs ERP for one part.

Rebuilds the 3PL's and the ERP's balances per lot and owner from the 3PL's
inventory transaction trace(s) and the ERP part transaction history, pairs
the movements, and shows where the two systems stopped agreeing.
Settings come from the "reconciliation" and "sku_reconciliation" sections of config.json.
"""
import collections
import csv
import datetime
import os
import re

from stocktools import config as _config

_C = _config.load()
_R, _S = _C.get("reconciliation", {}), _C.get("sku_reconciliation", {})
WHSE_OWNER = {k.upper(): v for k, v in _R.get("warehouse_code_owner", {}).items()}
BIN_CODE_OWNER = {(a.upper(), b.upper()): c for a, b, c in _R.get("warehouse_code_bin_owner", [])}
WHSE_NAME_OWNER = dict(_R.get("erp_warehouse_owner", {}))
WHSE_NAME_OWNER.update({w: o for _, w, o in _R.get("erp_site_warehouse_owner", [])})
BIN_NAME_OWNER = {(a, b.upper()): c for a, b, c in _R.get("erp_bin_owner", [])}
# 3PL trace codes that change owner quantities. Allocate, Pick, Replenish, Vehicle Load,
# PreAdv Line, Stock Check and Repack net to zero at owner level, so they are ignored.
DT_MOVES = {'Receipt': 1, 'Shipment': -1, 'Owner Update': 1, 'Return': 1, 'Receipt Reverse': 1}
MATCH_DAYS = _S.get("match_days", 5)
NO_EFFECT = {'ADJ-CST', 'ADJ-PUR'}
KPI_FOLDER = _S.get("kpi_folder", "")
DEFAULT_OWNER = _S.get("default_owner", "")


def owner_mapping_text():
    parts = [f"{k} = {v}" for k, v in sorted(WHSE_OWNER.items())]
    parts += [f"{w} bin {b} = {o}" for (w, b), o in sorted(BIN_CODE_OWNER.items())]
    return ', '.join(parts)


# ---------------------------------------------------------------- helpers
def num(v):
    if isinstance(v, (int, float)): return float(v)
    try: return float(str(v).replace(',', '').strip())
    except Exception: return None


def clean_lot(v):
    if v is None: return ''
    if isinstance(v, float) and v.is_integer(): v = int(v)
    s = str(v).strip().upper().replace(' ', '')
    if s.startswith('LOT') and s[3:].isdigit(): s = s[3:]
    return '' if s in ('', 'N/A', 'NA', 'NULL', '-', 'NONE', 'NAN') else s


def lotkey(s):
    """Lot used for comparison: leading zeros and a 3PL -suffix ignored."""
    s = clean_lot(s)
    if re.fullmatch(r'\d+-[A-Z0-9]+', s): s = s.split('-')[0]
    return s.lstrip('0') or s


def to_date(v):
    if isinstance(v, datetime.datetime): return v
    if isinstance(v, datetime.date): return datetime.datetime(v.year, v.month, v.day)
    s = str(v or '').strip()
    for f in ('%d/%m/%Y %H:%M:%S', '%d/%m/%Y %H:%M', '%Y-%m-%d %H:%M:%S', '%d/%m/%Y', '%d-%b-%Y', '%d-%b-%y',
              '%Y-%m-%d', '%d/%m/%y'):
        for x in (s, s[:10], s[:11]):
            try: return datetime.datetime.strptime(x, f)
            except Exception: pass
    return None


def refbase(v):
    """Order/pack reference without commas or 3PL suffixes (288091BK3 -> 288091)."""
    s = str(v or '').replace(',', '').strip().upper()
    if s in ('', '0', 'NONE', 'NAN'): return ''
    m = re.match(r'^(\d{4,})', s)
    return m.group(1) if m else s


def read_rows(path, must=None):
    """All rows from .xlsx, .xls, .csv or tab-delimited .txt as lists. For .xlsx, the first tab whose
    header matches `must` is used (first tab if none match)."""
    p = path.lower()
    if p.endswith('.xlsx') or p.endswith('.xlsm'):
        from openpyxl import load_workbook
        wb = load_workbook(path, read_only=True, data_only=True)
        tabs = [[list(r) for r in ws.iter_rows(values_only=True, max_row=None)] for ws in wb.worksheets]
        if must:
            for rows in tabs:
                try:
                    find_header(rows, must); return rows
                except ValueError:
                    pass
        return tabs[0]
    if p.endswith('.xls'):
        import xlrd
        sh = xlrd.open_workbook(path).sheet_by_index(0)
        out = []
        for i in range(sh.nrows):
            row = []
            for c in sh.row(i):
                if c.ctype == xlrd.XL_CELL_DATE:
                    row.append(datetime.datetime(*xlrd.xldate_as_tuple(c.value, sh.book.datemode)))
                else:
                    row.append(c.value)
            out.append(row)
        return out
    with open(path, encoding='utf-8-sig', errors='replace') as f:
        text = f.read()
    delim = '\t' if text.count('\t') > text.count(',') else ','
    return [r for r in csv.reader(text.splitlines(), delimiter=delim)]


def find_header(rows, must):
    """must: list of column names, each a name or a tuple of accepted alternatives."""
    alts = [tuple(a.lower() for a in (m if isinstance(m, tuple) else (m,))) for m in must]
    for i, r in enumerate(rows[:30]):
        names = [str(x).strip().lower() for x in r]
        if all(any(a in names for a in m) for m in alts): return i
    first = next((r for r in rows[:30] if any(x not in (None, '') for x in r)), [])
    raise ValueError('Could not find a header row with columns: ' + ' and '.join('/'.join(m) for m in alts) +
                     '. First row in the file reads: ' + ', '.join(str(x) for x in first[:12] if x not in (None, '')) +
                     '. Check this is the ERP transaction BAQ export, not a 3PL file or a stock report.')


def pick(hdr, *names, last=False):
    low = [str(h).strip().lower() for h in hdr]
    for n in names:
        idx = [i for i, h in enumerate(low) if h == n.lower()]
        if idx: return idx[-1] if last else idx[0]
    return None


# ---------------------------------------------------------------- loading
def load_3pl_traces(paths, part):
    events = []
    for path in paths:
        rows = read_rows(path)
        h = find_header(rows, ['code', 'update qty'])
        hdr = rows[h]
        c = dict(code=pick(hdr, 'Code'), qty=pick(hdr, 'Update Qty'), lot=pick(hdr, 'Batch ID', 'Batch'),
                 owner=pick(hdr, 'Owner ID', 'Owner'), ref=pick(hdr, 'Reference'),
                 date=pick(hdr, 'Date', 'Transaction Date', 'Dstamp'), time=pick(hdr, 'Hours', 'Transaction Time'),
                 sku=pick(hdr, 'Sku ID', 'SKU ID', 'Sku', 'SKU'), notes=pick(hdr, 'Notes'),
                 cust=pick(hdr, 'Customer Shipment Number', 'Shipment Number'))
        for r in rows[h + 1:]:
            if len(r) <= c['qty']: continue
            code = str(r[c['code']] or '').strip()
            if code not in DT_MOVES: continue
            if c['sku'] is not None and str(r[c['sku']]).strip() and \
                    str(r[c['sku']]).strip().lstrip('0') != part.lstrip('0'):
                continue
            q = num(r[c['qty']])
            if q is None: continue
            d = to_date(r[c['date']]) if c['date'] is not None else None
            if d and c['time'] is not None and r[c['time']] not in (None, ''):
                t = r[c['time']]
                if isinstance(t, datetime.datetime): d = d.replace(hour=t.hour, minute=t.minute)
                elif isinstance(t, float) and t < 1: d = d + datetime.timedelta(days=t)
                else:
                    m = re.match(r'(\d{1,2}):(\d{2})', str(t))
                    if m: d = d.replace(hour=int(m.group(1)), minute=int(m.group(2)))
            events.append(dict(sys='3PL', owner=str(r[c['owner']] or '').strip() if c['owner'] is not None else '',
                               lot=lotkey(r[c['lot']]) if c['lot'] is not None else '',
                               lot_raw=str(r[c['lot']] or '') if c['lot'] is not None else '',
                               type=code, qty=DT_MOVES[code] * q, date=d,
                               ref=str(r[c['ref']] or '').strip() if c['ref'] is not None else '',
                               note=str(r[c['notes']] or '') if c['notes'] is not None else '',
                               src=os.path.basename(path)))
    return events


def erp_owner(code, name, bin_):
    code, name, b = str(code or '').strip().upper(), str(name or '').strip(), str(bin_ or '').strip().upper()
    # some exports (e.g. a lot tracker) put the warehouse code in a column called Warehouse
    if not code and (name.upper() in WHSE_OWNER or any(w == name.upper() for w, _ in BIN_CODE_OWNER)):
        code = name.upper()
    if any(w == code for w, _ in BIN_CODE_OWNER): return BIN_CODE_OWNER.get((code, b))
    if any(w == name for w, _ in BIN_NAME_OWNER): return BIN_NAME_OWNER.get((name, b))
    if code in WHSE_OWNER: return WHSE_OWNER[code]
    return WHSE_NAME_OWNER.get(name)


def onhand_owner(loc):
    u = loc.strip().upper()
    for (w, b), o in BIN_CODE_OWNER.items():
        if u in (w, b): return o
    return erp_owner(loc, loc, '') or u


def erp_sign(tt, cls, q):
    """Signed effect on stock. Uses the class column (R/S/I/A) when present."""
    if tt in NO_EFFECT: return 0
    if cls in ('R',): return q
    if cls in ('S', 'I'): return -q
    if cls == 'A': return q
    if tt in ('ADJ-QTY', 'STK-STK'): return q
    if tt.endswith('-STK'): return q
    if tt.startswith('STK-'): return -q
    return 0


def load_erp(path, part, default_lot=''):
    # Accepts the transaction BAQ (has Part and Lot columns) or the Lot Tracker PartTran export
    # (one part and one lot per file, so no Part or Lot column)
    must = [('quantity', 'tran qty', 'tranqty'), ('type', 'tran type', 'trantype')]
    rows = read_rows(path, must)
    h = find_header(rows, must)
    hdr = rows[h]
    types = [i for i, x in enumerate(hdr) if str(x).strip().lower() == 'type']
    c = dict(part=pick(hdr, 'Part', 'Part Num', 'PartNum'), qty=pick(hdr, 'Quantity', 'Tran Qty'),
             tt=(types[-1] if len(types) > 1 else pick(hdr, 'Tran Type', 'Tarn Type', 'TranType', 'Type')),
             cls=(types[0] if len(types) > 1 else None), lot=pick(hdr, 'Lot Number', 'Lot', 'LotNum'),
             whse=pick(hdr, 'Whse', 'Warehouse Code', 'WareHouseCode'), wname=pick(hdr, 'Warehouse'),
             bin=pick(hdr, 'Bin', 'Bin Num', 'BinNum'), sysd=pick(hdr, 'Sys Date', 'SysDate'),
             tdate=pick(hdr, 'Date', 'Tran Date', 'Date From'), syst=pick(hdr, 'Sys Time', 'SysTime'),
             tran=pick(hdr, 'Tran', 'Tran Num', 'TranNum'), order=pick(hdr, 'Order', 'Order Num'),
             pack=pick(hdr, 'Pack', 'Packing Slip'), po=pick(hdr, 'PO'), ref=pick(hdr, 'Reference'),
             rma=pick(hdr, 'RMA'), reason=pick(hdr, 'Reason Code'), user=pick(hdr, 'User', 'User ID', 'Entry Person'),
             invty=pick(hdr, 'Invty'))
    sys_dated = c['sysd'] is not None
    out, ctx = [], []   # ctx: ERP movements outside 3PL warehouses, shown for context only
    for r in rows[h + 1:]:
        if len(r) <= c['qty']: continue
        if c['part'] is not None:
            if not r[c['part']] or str(r[c['part']]).strip().lstrip('0') != part.lstrip('0'): continue
        if c['invty'] is not None and str(r[c['invty']]).strip().lower() == 'false': continue
        q = num(r[c['qty']])
        if q is None: continue
        tt = str(r[c['tt']] or '').strip().upper()
        cls = str(r[c['cls']] or '').strip().upper() if c['cls'] is not None else ''
        eff = erp_sign(tt, cls, q)
        if not eff: continue
        g = erp_owner(r[c['whse']] if c['whse'] is not None else '', r[c['wname']] if c['wname'] is not None else '',
                         r[c['bin']] if c['bin'] is not None else '')
        d = to_date(r[c['sysd']] if sys_dated else r[c['tdate']]) if (sys_dated or c['tdate'] is not None) else None
        if d and sys_dated and c['syst'] is not None:
            s = num(r[c['syst']])
            if s is not None and s < 86400: d = d + datetime.timedelta(seconds=s)
        g_ref = lambda k: str(r[c[k]] or '').strip() if c[k] is not None else ''
        refs = [x for x in (g_ref('order'), g_ref('pack'), g_ref('po'), g_ref('ref'), g_ref('rma')) if refbase(x)]
        (out if g is not None else ctx).append(dict(sys='ERP', owner=g,
                        lot=lotkey(r[c['lot']]) if c['lot'] is not None else lotkey(default_lot),
                        lot_raw=str(r[c['lot']] or '') if c['lot'] is not None else default_lot, type=tt, qty=eff, date=d,
                        ref=' / '.join(refs), refs={refbase(x) for x in refs},
                        tran=num(g_ref('tran')) or 0, reason=g_ref('reason'), user=g_ref('user'),
                        whse=g_ref('whse') or g_ref('wname'), src=os.path.basename(path)))
    # drop moves inside one owner (e.g. bin to bin): same tran group nets to zero
    grp = collections.defaultdict(list)
    for e in out: grp[(e['owner'], e['lot'], e['type'], e['date'], e['ref'])].append(e)
    kept = []
    for k, es in grp.items():
        if abs(sum(e['qty'] for e in es)) < 1e-9 and len(es) > 1: continue
        # cancel a posting and its reversal (e.g. STK-CUS -5, +5, -5 on one order: keep one -5)
        rest = list(es)
        for e in es:
            if e not in rest: continue
            opp = next((x for x in rest if x is not e and abs(x['qty'] + e['qty']) < 1e-9), None)
            if opp is not None: rest.remove(e); rest.remove(opp)
        kept.extend(rest)
    # transfers between two 3PL warehouses of the same owner (STK-PLT out, PLT-STK in on the same pack) net out
    tg = collections.defaultdict(list)
    for e in kept:
        if e['type'] in ('STK-PLT', 'PLT-STK') and e['ref']: tg[(e['owner'], e['lot'], e['ref'])].append(e)
    drop = set()
    for es in tg.values():
        if len({e['type'] for e in es}) == 2 and abs(sum(e['qty'] for e in es)) < 1e-9:
            drop.update(id(e) for e in es)
    kept = [e for e in kept if id(e) not in drop]
    kept.sort(key=lambda e: (e['date'] or datetime.datetime.min, e['tran']))
    if c['lot'] is None: sys_dated = 'lot tracker' if not sys_dated else sys_dated
    ctx.sort(key=lambda e: (e['date'] or datetime.datetime.min, e['tran']))
    return kept, sys_dated, ctx


def parse_onhand(text, is_erp, default_owner=None):
    default_owner = default_owner or DEFAULT_OWNER or next(iter(WHSE_OWNER.values()), '')
    """Typed or pasted lines: 'Lot Owner Qty' (or 'Lot Warehouse Qty' for ERP), or just 'Lot Qty' to use
    the default owner. Tabs, commas or spaces between values. Lines without a number at the end are skipped."""
    out = collections.defaultdict(float)
    for line in (text or '').splitlines():
        parts = [p for p in re.split(r'[\t,; ]+', line.strip()) if p]
        if len(parts) < 2: continue
        q = num(parts[-1])
        if q is None: continue
        lot = parts[0]
        if len(parts) == 2:
            owner = default_owner
        else:
            loc = ' '.join(parts[1:-1])
            owner = loc.upper()
            if is_erp:
                owner = onhand_owner(loc)
        out[(owner, lotkey(lot))] += q
    return out


# ---------------------------------------------------------------- matching
def match_events(dt_ev, ep_ev):
    """Pair 3PL and ERP movements per owner and lot: by reference first, then by qty and date."""
    for e in dt_ev: e['refb'] = refbase(e['ref']); e['pair'] = None
    for e in ep_ev: e['pair'] = None
    # 1. reference match, same owner, lot and direction. 3PL is grouped by its full reference
    #    (288091BK1, 288091BK3), ERP by the base reference (288091); 3PL groups whose qty adds up
    #    to ERP's pair, any extra 3PL group stays unmatched.
    from itertools import combinations
    dg, eg = collections.defaultdict(lambda: collections.defaultdict(list)), collections.defaultdict(list)
    for e in dt_ev:
        if e['refb']: dg[(e['owner'], e['lot'], e['refb'])][e['ref']].append(e)
    for e in ep_ev:
        for rb in e['refs']: eg[(e['owner'], e['lot'], rb)].append(e)
    notes = []
    for k, groups in dg.items():
        ees = [e for e in eg.get(k, []) if e['pair'] is None]
        if not ees: continue
        eq = sum(e['qty'] for e in ees)
        items = [(ref, evs, sum(e['qty'] for e in evs)) for ref, evs in groups.items()]
        chosen = None
        for n in range(len(items), 0, -1):
            for combo in combinations(items, n):
                if abs(sum(x[2] for x in combo) - eq) < 1e-9: chosen = combo; break
            if chosen or len(items) > 12: break
        tag = 'ref ' + k[2]
        if chosen:
            for _, evs, _ in chosen:
                for e in evs: e['pair'] = tag
            for e in ees: e['pair'] = tag
        else:
            dq = sum(x[2] for x in items)
            notes.append(dict(owner=k[0], lot=k[1], ref=k[2], dt=dq, erp=eq,
                              date=min((e['date'] for e in ees if e['date']), default=None), how='shared reference'))
            for _, evs, _ in items:
                for e in evs: e['pair'] = tag + ' (qty differs)'
            for e in ees: e['pair'] = tag + ' (qty differs)'
    # 2. same owner, lot, qty within MATCH_DAYS
    for d in dt_ev:
        if d['pair']: continue
        best = None
        for e in ep_ev:
            if e['pair'] or e['owner'] != d['owner'] or e['lot'] != d['lot'] or abs(e['qty'] - d['qty']) > 1e-9:
                continue
            gap = abs((e['date'] - d['date']).days) if (e['date'] and d['date']) else 999
            if gap <= MATCH_DAYS and (best is None or gap < best[0]): best = (gap, e)
        if best:
            d['pair'] = best[1]['pair'] = f'qty and date ({best[0]} day{"s" if best[0] != 1 else ""} apart)'
    # 3. same owner, lot and direction within MATCH_DAYS but a different qty: likely the same movement
    #    booked with different quantities (e.g. ERP moves 3,001 into 3PL, 3PL receipts 1)
    for d in dt_ev:
        if d['pair'] or not d['date']: continue
        best = None
        for e in ep_ev:
            if e['pair'] or not e['date'] or e['owner'] != d['owner'] or e['lot'] != d['lot'] or (e['qty'] > 0) != (d['qty'] > 0):
                continue
            gap = abs((e['date'] - d['date']).days)
            if gap <= MATCH_DAYS and (best is None or gap < best[0]): best = (gap, e)
        if best:
            e = best[1]
            d['pair'] = e['pair'] = f'likely same movement, qty differs ({best[0]} day{"s" if best[0] != 1 else ""} apart)'
            notes.append(dict(owner=d['owner'], lot=d['lot'], ref=d['ref'] or e['ref'] or 'no ref', dt=d['qty'],
                              erp=e['qty'], date=min(d['date'], e['date']), how='date and direction, no shared reference'))
    return notes


def explain(e):
    t, q, s = e['type'], e['qty'], e['sys']
    if s == '3PL':
        if t == 'Shipment': return '3PL shipped it; no matching shipment in ERP'
        if t == 'Receipt': return '3PL received it; no matching receipt in ERP'
        if t == 'Owner Update': return f'3PL moved it {"into" if q > 0 else "out of"} this owner; no matching transfer in ERP'
        if t == 'Return': return '3PL booked a return; no matching return in ERP'
        return '3PL movement with no match in ERP'
    if t == 'STK-CUS': return ('Shipped in ERP; no matching 3PL shipment' if q < 0 else
                              'Shipment reversed in ERP; no matching 3PL movement')
    if t in ('PUR-STK', 'MFG-STK', 'INS-STK'): return 'Received in ERP; no matching 3PL receipt'
    if t in ('PLT-STK', 'STK-PLT'): return f'Transferred {"into" if q > 0 else "out of"} this owner in ERP; no matching 3PL owner change'
    if t == 'STK-STK': return ('Moved into a 3PL warehouse in ERP; 3PL has no matching receipt' if q > 0 else
                              'Moved out of a 3PL warehouse in ERP; 3PL has no matching movement')
    if t in ('ADJ-QTY', 'STK-UKN', 'STK-MTL', 'UKN-STK'): return f'{t} in ERP (reason {e.get("reason") or "none"}); 3PL has no matching movement'
    return f'{t} in ERP with no match at 3PL'


def category(e):
    """Short cause category for KPI reporting."""
    t, q, s = e['type'], e['qty'], e['sys']
    if s == '3PL':
        return {'Shipment': '3PL shipment not in ERP', 'Receipt': '3PL receipt not in ERP',
                'Owner Update': 'Owner move not in ERP', 'Return': '3PL return not in ERP'}.get(t, 'Other 3PL movement')
    if t == 'STK-CUS': return 'ERP shipment not at 3PL' if q < 0 else 'ERP shipment reversal'
    if t in ('PUR-STK', 'MFG-STK', 'INS-STK'): return 'ERP receipt not at 3PL'
    if t in ('PLT-STK', 'STK-PLT'): return 'ERP transfer not at 3PL'
    if t == 'STK-STK': return 'ERP move into 3PL not received' if q > 0 else 'ERP move out of 3PL not at 3PL'
    if t in ('ADJ-QTY', 'STK-UKN', 'STK-MTL', 'UKN-STK'): return 'ERP adjustment or misc'
    return 'Other ERP movement'


def append_csv(folder, name, header, rows, log=print):
    """Append rows to a CSV log, writing the header if the file is new. If the log is locked
    (open in Excel), the rows go to a dated side file instead so nothing is lost."""
    if not rows: return
    os.makedirs(folder, exist_ok=True)
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


# ---------------------------------------------------------------- run
def run(part, trace_paths, erp_path, out_path, tpl_onhand_text='', erp_onhand_text='', log=print, kpi_folder=None,
        lots='', default_owner=None):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter as L
    part = part.strip().upper()
    log(f'Part {part}: reading 3PL traces...')
    dt_ev = load_3pl_traces(trace_paths, part)
    log(f'  {len(dt_ev)} 3PL movements (Receipt, Shipment, Owner Update, Return, Receipt Reverse)')
    log('Reading ERP history...')
    only_list = [x for x in re.split(r'[\s,;]+', lots or '') if x.strip()]
    stem = os.path.splitext(os.path.basename(erp_path))[0].split('-')[-1].strip()
    default_lot = only_list[0] if len(only_list) == 1 else stem
    ep_ev, sys_dated, ctx_ev = load_erp(erp_path, part, default_lot)
    if sys_dated == 'lot tracker':
        sys_dated = False
        log(f'  Lot Tracker export (no Part or Lot column): treated as part {part}, lot {default_lot} '
            f'({"from Lots to check" if len(only_list) == 1 else "from the file name"}). Sequenced on Tran Date.')
    log(f'  {len(ep_ev)} ERP stock movements at 3PL warehouses, {len(ctx_ev)} elsewhere (shown for context)')
    dt_oh, ep_oh = parse_onhand(tpl_onhand_text, False, default_owner), parse_onhand(erp_onhand_text, True, default_owner)
    only = {lotkey(x) for x in re.split(r'[\s,;]+', lots or '') if x.strip()}
    if not only:
        ep_lots = {e['lot'] for e in ep_ev}
        extra = {e['lot'] for e in dt_ev} - ep_lots
        if ep_lots and extra:
            log(f'Note: the 3PL trace has {len(extra)} lot(s) that are not in the ERP export. If the export was '
                f'filtered by lot, enter the lots to check so those are left out.')
    if only:
        dt_ev = [e for e in dt_ev if e['lot'] in only]; ep_ev = [e for e in ep_ev if e['lot'] in only]
        ctx_ev = [e for e in ctx_ev if e['lot'] in only]
        dt_oh = {k: v for k, v in dt_oh.items() if k[1] in only}; ep_oh = {k: v for k, v in ep_oh.items() if k[1] in only}
        log('Checking lots: ' + ', '.join(sorted(only)))
    qty_notes = match_events(dt_ev, ep_ev)

    ep_first = min((e['date'] for e in ep_ev if e['date']), default=None)
    dt_first = min((e['date'] for e in dt_ev if e['date']), default=None)
    dt_net, ep_net = collections.defaultdict(float), collections.defaultdict(float)
    for e in dt_ev: dt_net[(e['owner'], e['lot'])] += e['qty']
    for e in ep_ev: ep_net[(e['owner'], e['lot'])] += e['qty']
    keys = sorted(set(dt_net) | set(ep_net) | set(dt_oh) | set(ep_oh))
    dt_owners_typed, ep_owners_typed = {k[0] for k in dt_oh}, {k[0] for k in ep_oh}
    log('Owners in the 3PL trace: ' + (', '.join(sorted({e['owner'] for e in dt_ev})) or 'none'))
    log('Owners in the ERP history: ' + (', '.join(sorted({e['owner'] for e in ep_ev})) or 'none'))
    for nm, typed, net in (('3PL', dt_owners_typed, dt_net), ('ERP', ep_owners_typed, ep_net)):
        miss = sorted({k[0] for k in keys} - typed)
        if (dt_oh or ep_oh) and miss:
            log(f'  {nm} on hand not typed for {", ".join(miss)}: using the {nm} movement totals for those owners')
    lots_raw = {}
    for e in dt_ev + ep_ev + ctx_ev: lots_raw.setdefault(e['lot'], e['lot_raw'])

    summary, timeline, checks = [], [], []
    total_diff = 0.0
    for k in keys:
        owner, lot = k
        # typed on hand: a lot not listed for an owner you typed figures for is 0; an owner you typed
        # nothing for uses its trace/history total, and its tie check says so
        dto = dt_oh.get(k, 0.0) if owner in dt_owners_typed else None
        epo = ep_oh.get(k, 0.0) if owner in ep_owners_typed else None
        dtn, epn = dt_net.get(k, 0.0), ep_net.get(k, 0.0)
        dt_ref = dto if dto is not None else dtn
        ep_ref = epo if epo is not None else epn
        diff = dt_ref - ep_ref
        total_diff += diff
        evs = sorted([e for e in dt_ev + ep_ev if (e['owner'], e['lot']) == k],
                     key=lambda e: (e['date'] or datetime.datetime.min, 0 if e['sys'] == '3PL' else 1))
        # openings so each history ends at its on-hand figure
        dbal, ebal = dt_ref - dtn, ep_ref - epn
        last_agree = None if abs(dbal - ebal) > 1e-9 else 'start of history'
        real_agree, alone = last_agree, None   # alone: the one-sided movement that made the balances agree
        for e in evs:
            was = abs(dbal - ebal) < 1e-9
            if e['sys'] == '3PL': dbal += e['qty']
            else: ebal += e['qty']
            if abs(dbal - ebal) < 1e-9:
                last_agree = e['date']
                if not was and not e['pair']: alone = e
                else: alone = None; real_agree = e['date']
            elif was: alone = None
            timeline.append([owner, lots_raw.get(lot, lot), e['date'], e['sys'], e['type'], e['qty'], e['ref'],
                             e.get('reason', ''), e.get('user', ''), e['pair'] or 'NO MATCH',
                             (explain(e) if not e['pair'] else
                              'Same movement in both systems with different quantities: see the Qty differs tab'
                              if 'qty differs' in e['pair'] else ''), dbal, ebal, dbal - ebal])
        unmatched = [e for e in evs if not e['pair']]
        for e in unmatched:
            first = ep_first if e['sys'] == '3PL' else dt_first
            e['before'] = bool(e['date'] and first and e['date'] < first)
        cats = sorted({category(e) for e in unmatched})
        main = [e for e in unmatched if abs(abs(e['qty']) - abs(diff)) < 1e-9]
        cat = category(main[0]) if main else ('; '.join(cats) if cats else ('Agrees' if abs(diff) < 1e-9 else 'Paired, check on hand'))
        if abs(diff) < 1e-9 and not unmatched: cause = 'Agrees'
        elif unmatched:
            u = [e for e in unmatched if abs(abs(e['qty']) - abs(diff)) < 1e-9] or unmatched
            cause = '; '.join(f"{explain(e)} ({e['ref'] or 'no ref'}, {e['date']:%d/%m/%Y}, {e['qty']:+,.0f})"
                              + (' [before the other system\'s history starts: may be outside the export]' if e.get('before') else '')
                              if e['date'] else explain(e) for e in u[:3])
        else: cause = 'All movements pair up: check the on-hand figures and any qty-differs references'
        ties_dt = ('Not typed: trace total used' if (dt_oh or ep_oh) else '') if dto is None else ('Yes' if abs(dto - dtn) < 1e-9 else f'No (trace {dtn:,.0f})')
        ties_ep = ('Not typed: history total used' if (dt_oh or ep_oh) else '') if epo is None else ('Yes' if abs(epo - epn) < 1e-9 else f'No (history {epn:,.0f})')
        qd = [n for n in qty_notes if (n['owner'], n['lot']) == k]
        if qd:
            cause = '; '.join(f"Same movement booked with different quantities: 3PL {n['dt']:+,.0f}, ERP {n['erp']:+,.0f} "
                              f"({n['ref']}{', ' + format(n['date'], '%d/%m/%Y') if n.get('date') else ''}; matched on {n['how']})"
                              for n in qd) + ('' if cause.startswith('All movements pair up') else '; ' + cause)
            if not main: cat = 'Qty differs between systems'
        def fmt(v): return v.strftime('%d/%m/%Y') if isinstance(v, datetime.datetime) else (v or 'Not in this history')
        if alone is not None:
            agree_txt = (f"{fmt(last_agree)}, but only because {'3PL' if alone['sys'] == '3PL' else 'ERP'} moved it alone "
                         f"({alone['type']}, no matching movement in the other system). Before that: {fmt(real_agree)}")
        else:
            agree_txt = last_agree if isinstance(last_agree, datetime.datetime) else (last_agree or 'Not in this history')
        summary.append([owner, lots_raw.get(lot, lot), dt_ref, ep_ref, diff, ties_dt, ties_ep, agree_txt,
                        len(unmatched), cause, cat])
    for e in ctx_ev:
        timeline.append(['Outside 3PL', lots_raw.get(e['lot'], e['lot']), e['date'], 'ERP', e['type'], e['qty'], e['ref'],
                         e.get('reason', ''), e.get('user', ''), 'CONTEXT',
                         f"ERP movement at {e['whse'] or 'a non-3PL warehouse'}, outside 3PL. Does not affect the 3PL vs ERP balance",
                         None, None, None])
    timeline.sort(key=lambda t: (str(t[1]), t[2] or datetime.datetime.min, 0 if t[0] != 'Outside 3PL' else 1))
    # checks for each line that differs
    for s in summary:
        owner, lot, dt_ref, ep_ref, diff = s[:5]
        if abs(diff) < 1e-9: continue
        other = [x for x in summary if x[1] == lot and x[0] != owner and abs(x[4] + diff) < 1e-9]
        other_lot = [x for x in summary if x[0] == owner and x[1] != lot and abs(x[4] + diff) < 1e-9]
        if other: qa = f'Attribute only: the opposite difference is under owner {other[0][0]} (same lot). No quantity change'
        elif other_lot: qa = f'Attribute only: the opposite difference is on lot {other_lot[0][1]} (same owner). No quantity change'
        elif abs(total_diff) < 1e-9: qa = 'Net zero across all lots and owners for this part: likely attribute only. Confirm the pairing'
        else: qa = 'Quantity'
        ev_dt = [t for t in timeline if t[0] == owner and t[1] == lot and t[3] == '3PL' and t[9] == 'NO MATCH']
        ev_ep = [t for t in timeline if t[0] == owner and t[1] == lot and t[3] == 'ERP' and t[9] == 'NO MATCH']
        tie = s[5]
        if any((n['owner'], lots_raw.get(n['lot'], n['lot'])) == (owner, lot) for n in qty_notes) and not ev_dt and not ev_ep:
            which = ('Not established: both systems booked the movement, with different quantities. '
                     'Check the physical receipt paperwork')
        elif ev_dt and not ev_ep and tie != '' and tie.startswith('Yes'):
            which = 'Likely ERP: 3PL has movements ERP lacks, and 3PL\'s trace ties to its on hand'
        elif ev_ep and not ev_dt:
            which = 'Not established: ERP has movements 3PL lacks. Needs 3PL to confirm physically'
        else:
            which = 'Not established'
        phys = (f'3PL system shows {dt_ref:,.0f} under {owner}. Not physically confirmed' if dt_ref
                else f'3PL system shows none under {owner}. Not physically confirmed')
        checks.append([owner, lot, diff, phys, which, qa])

    # ---- workbook
    F, B = Font(name='Arial', size=10), Font(name='Arial', size=10, bold=True)
    H = PatternFill('solid', fgColor='DDEBF7'); RED = PatternFill('solid', fgColor='FCE4E4')
    NUM = '#,##0.##;-#,##0.##;0'; DT_FMT = 'dd/mm/yyyy hh:mm'
    wb = Workbook(); ws = wb.active; ws.title = 'Summary'

    def table(ws, hdr, rows, widths, start=1, red_col=None):
        for j, h in enumerate(hdr, 1):
            c = ws.cell(start, j, h); c.font = B; c.fill = H
        for i, r in enumerate(rows, start + 1):
            for j, v in enumerate(r, 1):
                c = ws.cell(i, j, v); c.font = F; c.alignment = Alignment(vertical='top', wrap_text=isinstance(v, str) and len(v) > 40)
                if isinstance(v, (int, float)): c.number_format = NUM
                if isinstance(v, datetime.datetime): c.number_format = DT_FMT
            if (red_col == 'diff' and abs(r[4]) > 1e-9) or (red_col == 'diff2' and abs(r[4]) > 1e-9) or (isinstance(red_col, int) and (r[red_col] == 'NO MATCH' or 'qty differs' in str(r[red_col]))):
                for j in range(1, len(r) + 1): ws.cell(i, j).fill = RED
        for j, w in enumerate(widths, 1): ws.column_dimensions[L(j)].width = w
        return start + len(rows) + 1

    ws['A1'] = f'Single SKU reconciliation: {part}'; ws['A1'].font = Font(name='Arial', size=14, bold=True)
    notes = [f'Run {datetime.datetime.now():%d/%m/%Y %H:%M}. 3PL traces: {", ".join(os.path.basename(p) for p in trace_paths)}. '
             f'ERP history: {os.path.basename(erp_path)}.',
             'Assumption: ERP warehouses map to 3PL owners as ' + owner_mapping_text() + '.',
             '3PL balances use Receipt, Shipment, Owner Update, Return and Receipt Reverse rows only. Allocate, Pick, Replenish, '
             'Vehicle Load, PreAdv Line, Stock Check and Repack net to zero at owner level.',
             ('ERP movements are sequenced on Sys Date and Sys Time.' if sys_dated else
              'ERP export has no Sys Date: sequenced on Tran Date, so the order is NOT verified.'),
             'On hand: ' + ('typed figures used. For an owner with no typed figures, that system\'s movement total is used.'
                            if (dt_oh or ep_oh) else
                            'none pasted, so balances are the trace/history totals and the tie checks are blank.')]
    r = 3
    for n in notes:
        ws.cell(r, 1, n).font = F; r += 1
    r += 1
    r = table(ws, ['Owner', 'Lot', '3PL on hand', 'ERP on hand', '3PL minus ERP', '3PL trace ties to on hand',
                   'ERP history ties to on hand', 'Last time both agreed', 'Unmatched movements', 'Likely cause'],
              [x[:10] for x in summary], [11, 16, 13, 14, 15, 18, 20, 20, 12, 90], r, red_col='diff')
    ws.cell(r, 1, f'Net difference for the part (3PL minus ERP): {total_diff:,.2f}').font = B
    r += 2
    ws.cell(r, 1, 'By owner').font = B; r += 1
    by = collections.OrderedDict()
    for x in summary:
        b = by.setdefault(x[0], [0.0, 0.0, 0.0, 0])
        b[0] += x[2]; b[1] += x[3]; b[2] += x[4]; b[3] += 1 if abs(x[4]) > 1e-9 else 0
    r = table(ws, ['Owner', 'Lots', '3PL on hand', 'ERP on hand', '3PL minus ERP', 'Lots that differ'],
              [[o, sum(1 for x in summary if x[0] == o), b[0], b[1], b[2], b[3]] for o, b in by.items()],
              [11, 16, 13, 14, 15, 18, 20, 20, 12, 90], r, red_col='diff2')
    r += 1
    ws.cell(r, 1, 'Before naming any correction').font = B; r += 1
    r = table(ws, ['Owner', 'Lot', '3PL minus ERP', '1. Does the stock physically exist, and where?',
                   '2. Which system holds the error?', '3. Quantity or only an attribute?'], checks,
              [11, 16, 13, 14, 15, 18, 20, 20, 12, 90], r)
    ws.cell(r, 1, 'Any quantity correction requires written approval on the adjustment approval form before posting. '
                  'A net-zero lot reclass does not. Nothing has been posted by this tool.').font = B
    r += 1
    ws.cell(r, 1, 'Never post a correction until root cause is confirmed and physically evidenced. '
                  'Adjustments are a last resort.').font = F
    for c in range(4, 7): ws.column_dimensions[L(c)].width = max(ws.column_dimensions[L(c)].width, 45)

    t = wb.create_sheet('Timeline')
    table(t, ['Owner', 'Lot', 'Date', 'System', 'Movement', 'Qty (signed)', 'Reference', 'Reason code', 'User',
              'Paired by', 'Why it matters', '3PL balance', 'ERP balance', '3PL minus ERP'],
          timeline, [10, 14, 17, 8, 14, 12, 26, 10, 18, 26, 55, 12, 13, 14], red_col=9)
    t.freeze_panes = 'A2'; t.auto_filter.ref = t.dimensions
    q = wb.create_sheet('Qty differs')
    table(q, ['Owner', 'Lot', 'Reference', 'Date', '3PL qty', 'ERP qty', '3PL minus ERP', 'Matched on'],
          [[n['owner'], lots_raw.get(n['lot'], n['lot']), n['ref'], n.get('date'), n['dt'], n['erp'],
            n['dt'] - n['erp'], n.get('how', '')] for n in qty_notes],
          [10, 14, 22, 17, 12, 12, 15, 34])
    raw = wb.create_sheet('3PL movements')
    table(raw, ['File', 'Owner', 'Batch', 'Code', 'Qty (signed)', 'Date', 'Reference', 'Notes'],
          [[e['src'], e['owner'], e['lot_raw'], e['type'], e['qty'], e['date'], e['ref'], e['note']] for e in dt_ev],
          [26, 10, 14, 14, 12, 17, 22, 40])
    raw = wb.create_sheet('ERP movements')
    table(raw, ['Tran', 'Owner', 'Warehouse', 'Lot', 'Type', 'Qty (signed)', 'Date', 'Reference', 'Reason', 'User'],
          [[e['tran'], e['owner'], e['whse'], e['lot_raw'], e['type'], e['qty'], e['date'], e['ref'], e['reason'],
            e['user']] for e in ep_ev] + [[e['tran'], 'Outside 3PL', e['whse'], e['lot_raw'], e['type'], e['qty'], e['date'],
            e['ref'], e['reason'], e['user']] for e in ctx_ev], [10, 10, 10, 14, 10, 12, 17, 26, 10, 18])
    wb.save(out_path)
    try:
        import getpass; who = getpass.getuser()
    except Exception:
        who = ''
    run_id = datetime.datetime.now().strftime('%Y%m%d%H%M%S%f')[:-3]
    chk = {(c[0], c[1]): c for c in checks}
    kpi_rows = []
    for x in summary:
        if abs(x[4]) < 1e-9: continue
        c = chk.get((x[0], x[1]), [None] * 6)
        kpi_rows.append([run_id, f'{datetime.datetime.now():%d/%m/%Y %H:%M}', who, part, x[0], x[1], x[2], x[3], x[4],
                         x[5], x[6], x[7].strftime('%d/%m/%Y') if isinstance(x[7], datetime.datetime) else x[7],
                         x[8], x[10], c[4], 'Attribute only' if str(c[5]).startswith('Attribute') else c[5],
                         x[9], '', '', '', ''])
    if not kpi_rows:
        kpi_rows.append([run_id, f'{datetime.datetime.now():%d/%m/%Y %H:%M}', who, part, '', '', '', '', 0, '', '', '', 0,
                         'No difference', '', '', 'All lots agree', '', '', '', ''])
    folder = kpi_folder or KPI_FOLDER or os.path.dirname(os.path.abspath(out_path))
    append_csv(folder, 'sku_investigations_log.csv',
               ['Run ID', 'Run date', 'Run by', 'Part', 'Owner', 'Lot', '3PL on hand', 'ERP on hand', '3PL minus ERP',
                '3PL trace ties', 'ERP history ties', 'Last agreed', 'Unmatched movements', 'Cause category',
                'Error side', 'Quantity or attribute', 'Likely cause', 'Reviewed by', 'Date resolved',
                'Approval form no', 'Notes'], kpi_rows, log)
    log(f'Logged {len(kpi_rows)} row(s) to sku_investigations_log.csv in {folder}')
    for s in summary:
        if abs(s[4]) > 1e-9:
            log(f'{s[0]} lot {s[1]}: 3PL {s[2]:,.0f} vs ERP {s[3]:,.0f} ({s[4]:+,.0f}). {s[9]}')
    log(f'Net difference: {total_diff:+,.2f}')
    log('Saved: ' + out_path)
    return summary, checks
