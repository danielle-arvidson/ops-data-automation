"""Pack-out SO/Line/Release filler.

The 3PL sends back a packing list with part, batch and quantity but no
sales order reference. The ERP shipping export has the reference as
"SO/Line/Release". This matches the two and fills SO, Line and Rel on the
packing list.

When one part/batch appears on several order lines, the oldest export
file wins, then the lowest SO number (first in, first out). An exact
quantity match overrides that order.
"""
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

from .codes import codes_match, normalize_code

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


def fill(packing, erp):
    """Fill SO/Line/Rel in place. Returns a summary dict."""
    for key in ("so", "line", "rel"):
        if PACK[key] not in packing.columns:
            packing[PACK[key]] = ""

    summary = {"total": len(packing), "matched": 0, "not_found": 0, "skipped": 0, "by_file": {}}
    for idx, row in packing.iterrows():
        if str(row.get(PACK["so"], "")).strip():
            summary["skipped"] += 1
            continue
        part = str(row.get(PACK["part"], "")).strip()
        batch = str(row.get(PACK["batch"], "")).strip()
        qty = str(row.get(PACK["qty"], "")).strip()
        if not part or not batch:
            continue

        hits = erp[(erp["_part_norm"] == normalize_code(part)) & (erp["_lot_norm"] == normalize_code(batch))]
        if hits.empty:
            packing.at[idx, PACK["so"]] = "NOT FOUND"
            summary["not_found"] += 1
            continue

        hits = hits.assign(_so=hits[ERP["ref"]].map(_so_number)).sort_values(["_priority", "_so"])
        chosen = hits.iloc[0]
        if qty:
            exact = hits[hits[ERP["qty"]].map(lambda q: codes_match(q, qty))]
            if not exact.empty:
                chosen = exact.iloc[0]

        parts = [p.strip() for p in str(chosen[ERP["ref"]]).split("/")]
        if len(parts) >= 3:
            packing.at[idx, PACK["so"]], packing.at[idx, PACK["line"]], packing.at[idx, PACK["rel"]] = parts[:3]
            summary["matched"] += 1
            src = chosen["_source"]
            summary["by_file"][src] = summary["by_file"].get(src, 0) + 1
        else:
            packing.at[idx, PACK["so"]] = "INVALID FORMAT"
            summary["not_found"] += 1
    return summary


def save(packing, output_path):
    """Write the packing list with borders, centred cells and failed rows highlighted."""
    packing.to_excel(output_path, index=False)
    wb = load_workbook(output_path)
    ws = wb.active
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
        if row[0].row > 1 and ws.cell(row[0].row, so_col).value in ("NOT FOUND", "INVALID FORMAT"):
            for c in range(1, BORDERED_COLS + 1):
                ws.cell(row[0].row, c).fill = PINK

    for column in ws.columns:
        width = max((len(str(c.value)) for c in column if c.value is not None), default=0)
        ws.column_dimensions[column[0].column_letter].width = min(width + 2, 50)
    wb.save(output_path)


def run(erp_paths, packing_path, output_path, log=print):
    erp = load_erp_exports(erp_paths)
    packing = pd.read_excel(packing_path, dtype=str, keep_default_na=False)
    summary = fill(packing, erp)
    save(packing, output_path)
    log(f"Matched {summary['matched']}, not found {summary['not_found']}, "
        f"already filled {summary['skipped']}")
    return {"output": output_path, **summary}
