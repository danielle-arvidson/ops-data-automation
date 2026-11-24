"""Pack-out SO/Line/Release filler.

The 3PL sends back a packing list with part, batch and quantity but no
sales order reference. The ERP shipping exports (one per movement request)
carry the reference as "SO/Line/Release". One shipment can cover several
movement requests, and the same part and batch can sit on several sales
order lines and releases across them.

Within one part and batch the units are interchangeable, so the job is not
to work out which physical line belongs to which release, but to share the
packed quantity across the open releases so every quantity ties:

1. Check: total packed vs total released for each part and batch.
2. Exact: a packed line that equals one release's quantity takes it.
3. Combination: a packed line that equals the sum of several releases
   takes them all (smallest set, oldest first).
4. Oldest first: whatever is left fills releases oldest export first, then
   lowest SO/Line/Release, splitting packed lines across releases.

Packed quantity with no open release is flagged NO OPEN RELEASE; releases
left unfilled are listed in the check. A line that already has an SO keeps
it and uses up that release first.
"""
import os

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

from . import config as _config
from .codes import codes_match, normalize_code

_P = _config.load().get("packout", {})
# Customers on rolling orders: a release left off a shipment stays open and ships later.
# For every other customer, a release left off the shipment has to be ordered again.
ROLLING_ORDER_CUSTOMERS = {c.strip().upper() for c in _P.get("rolling_order_customers", [])}
CUSTOMER_COLUMNS = _P.get("customer_columns", ["Customer", "CustID", "CustNum", "Customer Name", "tbCustNameHidden"])
# Column names accepted for each field: the simple export, and the ERP's material queue export
ERP_ALIASES = {"part": ["PartNum", "txtPartNum"], "lot": ["LotNum", "txtLotNum"],
               "qty": ["Quantity", "txtQuantity"], "ref": ["Reference", "txtReference"], "uom": ["UOM", "txtIUM"]}

# Column names accepted in a packing list file from the 3PL
PACK_ALIASES = {"part": ["Part Number", "SKU", "Item No", "Part"], "batch": ["Batch", "Batch No", "Lot"],
                "qty": ["Quantity", "Qty"]}
PALLET_ID_DIGITS = 10   # a token this long and all digits starts a new row in a pasted packing list

PACK = {"part": "Part Number", "batch": "Batch", "qty": "Quantity",
        "so": "SO", "line": "Line", "rel": "Rel"}
ERP = {"part": "PartNum", "lot": "LotNum", "qty": "Quantity", "ref": "Reference"}
BORDERED_COLS = 11
PINK = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")


def customers(erp):
    """Distinct customer names in the loaded ERP exports (blank ignored)."""
    col = next((c for c in CUSTOMER_COLUMNS if c in erp.columns), None)
    if not col:
        return []
    return sorted({str(v).strip() for v in erp[col] if str(v).strip()})


def check_one_customer(erp):
    """Each shipment is for one customer. More than one means a wrong export was loaded."""
    names = customers(erp)
    if len(names) > 1:
        by_file = erp.groupby("_source")[next(c for c in CUSTOMER_COLUMNS if c in erp.columns)].first()
        detail = "; ".join(f"{os.path.basename(f)}: {c}" for f, c in by_file.items())
        raise ValueError(f"The exports loaded are for {len(names)} customers ({', '.join(names)}). "
                         f"A shipment is for one customer, so remove the export that doesn't belong. {detail}")
    return names[0] if names else ""


def normalise_packing(packing):
    """Rename the 3PL's packing list headers (SKU, Qty...) to the names this tool uses."""
    rename = {}
    for key, names in PACK_ALIASES.items():
        if PACK[key] in packing.columns:
            continue
        found = next((n for n in names if n in packing.columns), None)
        if found:
            rename[found] = PACK[key]
    packing = packing.rename(columns=rename)
    missing = [PACK[k] for k in ("part", "batch", "qty") if PACK[k] not in packing.columns]
    if missing:
        raise ValueError(f"The packing list has no column for {', '.join(missing)}. "
                         f"Columns found: {', '.join(map(str, packing.columns))}")
    return packing


def _is_pallet_id(token):
    return token.isdigit() and len(token) >= PALLET_ID_DIGITS


def parse_pasted_packing(text):
    """Read a packing list table copied from the body of the 3PL's email.

    Works whether the paste arrives tab separated (one row per line) or one cell
    per line, as Outlook's plain text does. Columns: Pallet ID, Pallet, SKU,
    Batch, Qty, Weight (KG), Height (CM). Only the first row of each pallet
    carries the pallet number, weight and height.
    """
    cols = ["Pallet ID", "Pallet", PACK["part"], PACK["batch"], PACK["qty"], "Weight (KG)", "Height (CM)"]
    lines = [l.rstrip("\r") for l in text.splitlines()]
    # tab separated only if a header row is tab separated (signatures can contain stray tabs)
    if any("\t" in l and "SKU" in l.upper() for l in lines):
        rows = []
        for l in lines:
            cells = [c.strip() for c in l.split("\t")]
            if cells and _is_pallet_id(cells[0]):
                rows.append((cells + [""] * 7)[:7])
        if not rows:
            raise ValueError("No rows starting with a pallet ID were found in the pasted text.")
        return pd.DataFrame(rows, columns=cols)
    tokens = [t for l in lines for t in l.split()]
    starts = [i for i, t in enumerate(tokens) if _is_pallet_id(t)]
    if not starts:
        raise ValueError("No pallet IDs were found in the pasted text. Copy the table from the email, header row included.")
    rows, prev = [], None
    for n, start in enumerate(starts):
        pid = tokens[start]
        seg = tokens[start + 1:starts[n + 1] if n + 1 < len(starts) else len(tokens)]
        # the first row of a pallet carries the pallet number; later rows on the same pallet don't
        first_row = pid != prev
        need = 4 if first_row else 3
        cells, extra = seg[:need], seg[need:]
        numbers = []
        for t in extra:
            if _qty(t) is None:
                break
            numbers.append(t)
        ended = len(numbers) < len(extra)   # text after the row: the table has finished
        if len(cells) < need or len(numbers) > 2:
            raise ValueError(f"Pallet {pid}: expected {'pallet number, ' if first_row else ''}SKU, batch and quantity "
                             f"but found {' '.join(seg[:need + 2]) or 'nothing'}. Check that row in the email.")
        pallet = cells.pop(0) if first_row else rows[-1][1]
        sku, batch, qty = cells
        weight, height = (numbers + ["", ""])[:2]
        rows.append([pid, pallet, sku, batch, qty, weight, height])
        prev = pid
        if ended:
            break
    return pd.DataFrame(rows, columns=cols)


def load_erp_exports(paths):
    """Combine ERP exports; list order is priority (oldest first)."""
    frames = []
    for priority, path in enumerate(paths, start=1):
        df = pd.read_csv(path, dtype=str, keep_default_na=False)
        rename = {}
        for key, names in ERP_ALIASES.items():
            found = next((n for n in names if n in df.columns), None)
            target = ERP.get(key, "UOM")
            if found and found != target and target not in df.columns:
                rename[found] = target
        df = df.rename(columns=rename)
        missing = [ERP[k] for k in ("part", "lot", "qty", "ref") if ERP[k] not in df.columns]
        if missing:
            raise ValueError(f"{os.path.basename(path)} has no column for {', '.join(missing)}. "
                             "Use the ERP shipping or material queue export for each movement request.")
        df["_source"] = path
        df["_priority"] = priority
        frames.append(df)
    combined = pd.concat(frames, ignore_index=True)
    combined["_part_norm"] = combined[ERP["part"]].str.strip().map(normalize_code)
    combined["_lot_norm"] = combined[ERP["lot"]].str.strip().map(normalize_code)
    return combined


def _so_number(ref):
    so = ref.split("/")[0].strip()
    return int(so) if so.isdigit() else 999999


MAX_COMBO_RELEASES = 12   # search combinations only among this many open releases per part and batch


def _qty(v):
    try:
        return float(str(v).replace(",", "").strip())
    except ValueError:
        return None


def _ref_key(ref):
    parts = [p.strip() for p in str(ref).split("/")]
    nums = [int(p) if p.isdigit() else 999999 for p in (parts + ["", "", ""])[:3]]
    return tuple(nums)


def _subset(releases, target):
    """Smallest set of releases whose remaining quantities sum exactly to target, oldest first."""
    from itertools import combinations
    pool = [r for r in releases if r["left"] > 0][:MAX_COMBO_RELEASES]
    for size in range(2, len(pool) + 1):
        for combo in combinations(pool, size):
            if abs(sum(r["left"] for r in combo) - target) < 1e-9:
                return list(combo)
    return None


def allocate(packing, erp):
    """Share each packed line's quantity across open releases. Returns (rows, checks, summary).

    rows: one dict per output line (packing columns plus Match and Pack line).
    checks: one dict per part and batch with packed and released totals.
    """
    for key in ("so", "line", "rel"):
        if PACK[key] not in packing.columns:
            packing[PACK[key]] = ""
    cols = list(packing.columns)

    releases = {}
    cust_col = next((c for c in CUSTOMER_COLUMNS if c in erp.columns), None)
    for _, r in erp.iterrows():
        q = _qty(r[ERP["qty"]])
        if q is None:
            continue
        releases.setdefault((r["_part_norm"], r["_lot_norm"]), []).append(dict(
            ref=str(r[ERP["ref"]]).strip(), left=q, qty=q, order=(r["_priority"], _ref_key(r[ERP["ref"]])),
            source=r["_source"], part=str(r[ERP["part"]]).strip(), lot=str(r[ERP["lot"]]).strip(),
            customer=str(r[cust_col]).strip() if cust_col else ""))
    for rel in releases.values():
        rel.sort(key=lambda x: x["order"])

    summary = {"total": len(packing), "matched": 0, "not_found": 0, "skipped": 0, "split": 0, "by_file": {}}
    out, todo = {}, {}
    for idx, row in packing.iterrows():
        part, batch = str(row.get(PACK["part"], "")).strip(), str(row.get(PACK["batch"], "")).strip()
        key = (normalize_code(part), normalize_code(batch))
        q = _qty(row.get(PACK["qty"], ""))
        if str(row.get(PACK["so"], "")).strip():
            summary["skipped"] += 1
            out[idx] = [dict(row, **{"Match": "Already filled", "Pack line": idx + 2})]
            ref = "/".join(str(row.get(PACK[k], "")).strip() for k in ("so", "line", "rel"))
            for r in releases.get(key, []):
                if _ref_key(r["ref"]) == _ref_key(ref) and q:
                    r["left"] = max(0.0, r["left"] - q)
            continue
        if not part or not batch or q is None:
            out[idx] = [dict(row, **{"Match": "Missing part, batch or quantity", "Pack line": idx + 2})]
            continue
        todo.setdefault(key, []).append([idx, q, []])

    def take(item, r, amount, how):
        item[1] -= amount
        r["left"] -= amount
        item[2].append((r, amount, how))

    checks = []
    for key, items in todo.items():
        rel = releases.get(key, [])
        packed = sum(i[1] for i in items)
        released = sum(r["left"] for r in rel)
        # 1 exact
        for item in items:
            r = next((r for r in rel if r["left"] > 0 and abs(r["left"] - item[1]) < 1e-9 and r["left"] == r["qty"]), None)
            if r: take(item, r, item[1], "Exact")
        # 2 combination
        for item in items:
            if item[1] > 0:
                combo = _subset(rel, item[1])
                if combo:
                    for r in combo: take(item, r, r["left"], "Combination")
        # 3 oldest first
        for item in items:
            for r in rel:
                if item[1] <= 0: break
                if r["left"] > 0: take(item, r, min(item[1], r["left"]), "Oldest first")
        unfilled = [r for r in rel if r["left"] > 1e-9]
        over = sum(i[1] for i in items)
        status = ("Not on any export" if not rel else "Ties" if abs(packed - released) < 1e-9
                  else "Packed more than released" if packed > released else "Packed less than released")
        first = packing.loc[items[0][0]]
        checks.append({"Part": str(first[PACK["part"]]).strip(), "Batch": str(first[PACK["batch"]]).strip(), "Packed": packed, "Released": released,
                       "Difference": packed - released, "Status": status,
                       "Releases not fully packed": ", ".join(f"{r['ref']} ({r['left']:g})" for r in unfilled),
                       "Packed with no release": over})
        for idx, left, takes in items:
            base = packing.loc[idx].to_dict()
            rows = []
            for r, amount, how in takes:
                so, line, rl = ([p.strip() for p in r["ref"].split("/")] + ["", "", ""])[:3]
                rows.append(dict(base, **{PACK["qty"]: f"{amount:g}", PACK["so"]: so, PACK["line"]: line,
                                          PACK["rel"]: rl, "Match": how, "Pack line": idx + 2}))
                summary["by_file"][r["source"]] = summary["by_file"].get(r["source"], 0) + 1
            if left > 1e-9:
                rows.append(dict(base, **{PACK["qty"]: f"{left:g}", PACK["so"]: "NOT FOUND" if not rel else "NO OPEN RELEASE",
                                          PACK["line"]: "", PACK["rel"]: "", "Match": "", "Pack line": idx + 2}))
                summary["not_found"] += 1
            if takes:
                summary["matched"] += 1
            if len(takes) > 1:
                summary["split"] += 1
            out[idx] = rows
    ordered = [r for idx in packing.index for r in out.get(idx, [])]
    unfilled = []
    for rel in releases.values():
        for r in rel:
            if r["left"] > 1e-9:
                unfilled.append({"SO/Line/Rel": r["ref"], "Part": r["part"], "Batch": r["lot"],
                                 "Customer": r["customer"] or "Not in export", "Released": r["qty"],
                                 "Packed": r["qty"] - r["left"], "Left off shipment": r["left"],
                                 "Next step": release_action(r["customer"])})
    unfilled.sort(key=lambda u: (u["Next step"], u["SO/Line/Rel"]))
    summary["unfilled"] = unfilled
    return pd.DataFrame(ordered, columns=cols + ["Match", "Pack line"]), pd.DataFrame(checks), summary


def release_action(customer):
    """What happens to a release left off this shipment."""
    if not customer:
        return "Customer unknown: check whether it ships later or needs a new order"
    if customer.strip().upper() in ROLLING_ORDER_CUSTOMERS:
        return "Stays open: ships on a later shipment (rolling orders)"
    return "Not shipped: needs a new order"


def fill(packing, erp):
    """Backwards-compatible wrapper: returns the summary of allocate()."""
    return allocate(packing, erp)[2]


def save(packing, output_path, checks=None, unfilled=None):
    """Write the packing list with borders, centred cells and failed rows highlighted,
    plus a Check sheet of packed vs released totals per part and batch."""
    with pd.ExcelWriter(output_path) as xw:
        packing.to_excel(xw, index=False, sheet_name="Packing list")
        if checks is not None and len(checks):
            checks.to_excel(xw, index=False, sheet_name="Check")
        if unfilled:
            pd.DataFrame(unfilled).to_excel(xw, index=False, sheet_name="Left off shipment")
    wb = load_workbook(output_path)
    ws = wb["Packing list"]
    side = Side(style="thin", color="000000")
    border = Border(left=side, right=side, top=side, bottom=side)
    centre = Alignment(horizontal="center", vertical="center")
    headers = [c.value for c in ws[1]]
    so_col = headers.index(PACK["so"]) + 1
    qty_col = headers.index(PACK["qty"]) + 1 if PACK["qty"] in headers else None

    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = centre
            if cell.column <= BORDERED_COLS:
                cell.border = border
            if cell.row == 1:
                cell.font = Font(bold=True)
            elif cell.column == qty_col and str(cell.value or "").strip():
                try:
                    cell.value = int(float(cell.value))
                    cell.number_format = "0"
                except (TypeError, ValueError):
                    pass
        if row[0].row > 1 and ws.cell(row[0].row, so_col).value in ("NOT FOUND", "NO OPEN RELEASE", "INVALID FORMAT"):
            for c in range(1, BORDERED_COLS + 1):
                ws.cell(row[0].row, c).fill = PINK

    for sheet in wb.worksheets:
        for column in sheet.columns:
            width = max((len(str(c.value)) for c in column if c.value is not None), default=0)
            sheet.column_dimensions[column[0].column_letter].width = min(width + 2, 50)
        if sheet.title == "Left off shipment":
            for c in sheet[1]: c.font = Font(bold=True)
            for row in sheet.iter_rows(min_row=2):
                if not str(row[7].value).startswith("Stays open"):
                    for c in row: c.fill = PINK
        if sheet.title == "Check":
            for c in sheet[1]: c.font = Font(bold=True)
            for row in sheet.iter_rows(min_row=2):
                if row[5].value != "Ties":
                    for c in row: c.fill = PINK
    wb.save(output_path)


def run(erp_paths, packing_path, output_path, log=print, packing=None):
    """packing_path: the 3PL packing list as .xlsx/.csv. Or pass packing as a DataFrame
    (for example from parse_pasted_packing) and leave packing_path as None."""
    erp = load_erp_exports(erp_paths)
    customer = check_one_customer(erp)
    if packing is None:
        packing = (pd.read_csv(packing_path, dtype=str, keep_default_na=False) if str(packing_path).lower().endswith(".csv")
                   else pd.read_excel(packing_path, dtype=str, keep_default_na=False))
    packing = normalise_packing(packing.reset_index(drop=True))
    if customer:
        log(f"Customer: {customer} ({'rolling orders' if customer.upper() in ROLLING_ORDER_CUSTOMERS else 'releases left off need a new order'})")
    rows, checks, summary = allocate(packing, erp)
    save(rows, output_path, checks, summary["unfilled"])
    bad = int((checks["Status"] != "Ties").sum()) if len(checks) else 0
    log(f"Matched {summary['matched']} packed line(s) ({summary['split']} split across releases), "
        f"not placed {summary['not_found']}, already filled {summary['skipped']}. "
        f"Part and batch totals that don't tie: {bad} (see the Check sheet)")
    reorder = sum(1 for u in summary["unfilled"] if u["Next step"].startswith("Not shipped"))
    later = sum(1 for u in summary["unfilled"] if u["Next step"].startswith("Stays open"))
    unknown = len(summary["unfilled"]) - reorder - later
    if summary["unfilled"]:
        log(f"Releases left off this shipment: {reorder} need a new order, {later} stay open for a later shipment"
            + (f", {unknown} with no customer in the export" if unknown else "") + " (see the Left off shipment sheet)")
    return {"output": output_path, **summary}
