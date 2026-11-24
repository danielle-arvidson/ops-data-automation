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
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

from . import config as _config
from .codes import codes_match, normalize_code

_P = _config.load().get("packout", {})
# Customers on rolling orders: a release left off a shipment stays open and ships later.
# For every other customer, a release left off the shipment has to be ordered again.
ROLLING_ORDER_CUSTOMERS = {c.strip().upper() for c in _P.get("rolling_order_customers", [])}
CUSTOMER_COLUMNS = _P.get("customer_columns", ["Customer", "CustID", "CustNum", "Customer Name"])

PACK = {"part": "Part Number", "batch": "Batch", "qty": "Quantity",
        "so": "SO", "line": "Line", "rel": "Rel"}
ERP = {"part": "PartNum", "lot": "LotNum", "qty": "Quantity", "ref": "Reference"}
BORDERED_COLS = 11
PINK = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")


def load_erp_exports(paths):
    """Combine ERP exports; list order is priority (oldest first)."""
    frames = []
    for priority, path in enumerate(paths, start=1):
        df = pd.read_csv(path, dtype=str, keep_default_na=False)
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


def run(erp_paths, packing_path, output_path, log=print):
    erp = load_erp_exports(erp_paths)
    packing = pd.read_excel(packing_path, dtype=str, keep_default_na=False)
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
