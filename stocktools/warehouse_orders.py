"""Warehouse order checker.

Takes a warehouse order spreadsheet (part, batch, qty) and checks every
line against the 3PL's daily inventory report. Lines that can be covered
get a case count; lines that can't are marked and turned red. Blank
Date Required cells are filled with the next working day.
"""
import os
from datetime import datetime, timedelta

from openpyxl import load_workbook
from openpyxl.styles import PatternFill

from . import kpi, packconfig

REQUIRED = ["Item No", "Batch No", "Qty", "QtyBoxes/Bags", "Date Required"]
RED = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")
CUTOFF_HOUR = 14


def next_working_day(now=None):
    """Next working day, or the one after if it's past the 2pm cut-off. Skips weekends."""
    now = now or datetime.now()
    day = now + timedelta(days=2 if now.hour >= CUTOFF_HOUR else 1)
    while day.weekday() >= 5:
        day += timedelta(days=1)
    return day.replace(hour=0, minute=0, second=0, microsecond=0)


def find_header_row(ws, max_scan=10):
    for idx, row in enumerate(ws.iter_rows(min_row=1, max_row=max_scan, values_only=True), start=1):
        if row and "Item No" in [str(v).strip() if v is not None else "" for v in row]:
            return idx
    raise ValueError("Could not find a header row containing 'Item No'")


def _issue_type(msg):
    msg = str(msg)
    if msg == "NOT FOUND": return "Not found"
    if msg.startswith("SHORT"): return "Short"
    if "Pack Config" in msg: return "Invalid pack config"
    return "Missing data" if msg == "Missing data" else "Other"


RUNS_LOG = "warehouse_order_runs_log.csv"
ISSUES_LOG = "warehouse_order_issues_log.csv"
RUNS_HEADER = ["Run ID", "Run date", "Run by", "Order file", "3PL report age (hours)", "Lines", "OK", "Short", "Not found",
               "Invalid pack config", "Missing data", "Date Required filled"]
ISSUES_HEADER = ["Run ID", "Run date", "Order file", "Row", "Item", "Batch", "Qty", "Status", "Type", "Detail"]


def process(order_path, inventory, output_path, now=None, log=print, kpi_folder=None):
    """Check an order against inventory and save the marked-up copy. Returns a summary dict.
    kpi_folder: where the KPI logs go (see kpi.folder); False to skip logging."""
    wb = load_workbook(order_path)
    ws = wb.active
    header = find_header_row(ws)
    headers = [str(c.value).strip() if c.value is not None else "" for c in ws[header]]
    missing = [h for h in REQUIRED if h not in headers]
    if missing:
        raise ValueError(f"Missing columns: {', '.join(missing)}")
    col = {h: headers.index(h) + 1 for h in REQUIRED}

    due = next_working_day(now)
    log(f"Date Required default: {due:%d/%m/%Y}")
    counts = {"OK": 0, "SHORT": 0, "WARNING": 0}
    issues, dates_filled = [], 0

    for r in range(header + 1, ws.max_row + 1):
        item = ws.cell(r, col["Item No"]).value
        batch = ws.cell(r, col["Batch No"]).value
        qty = ws.cell(r, col["Qty"]).value
        if not item and not batch and not qty:
            continue
        if not ws.cell(r, col["Date Required"]).value:
            ws.cell(r, col["Date Required"]).value = due
            ws.cell(r, col["Date Required"]).number_format = "DD/MM/YYYY"
            dates_filled += 1

        out = ws.cell(r, col["QtyBoxes/Bags"])
        if not item or not batch or not qty:
            status, msg = "WARNING", "Missing data"
        else:
            qty = float(qty)
            free, rows = inventory.free_stock(item, batch)
            cfg = packconfig.parse_config(rows.iloc[0].get("Pack Config")) if not rows.empty else None
            if rows.empty:
                status, msg = "SHORT", "NOT FOUND"
            elif qty > free:
                status, msg = "SHORT", f"SHORT ({free:.0f} available)"
            elif not cfg or not cfg["units_per_case"]:
                status, msg = "WARNING", "ERROR: Invalid Pack Config"
            else:
                cases = qty / cfg["units_per_case"]
                msg = int(cases) if cases == int(cases) else round(cases, 2)
                status = "OK"

        out.value = msg
        counts[status] += 1
        log(f"Row {r}: {item} batch {batch}: {status} {msg}")
        if status != "OK":
            issues.append((r, item, batch, qty, status, msg))
            for cell in ws[r]:
                cell.fill = RED

    wb.save(output_path)
    log(f"Processed {counts['OK']}, shortages {counts['SHORT']}, warnings {counts['WARNING']}")
    run_id = None
    if kpi_folder is not False:
        run_id, (when, user) = kpi.new_run_id(), kpi.stamp()
        name = os.path.basename(order_path)
        detail = [str(i[5]) for i in issues]
        rt = getattr(inventory, "report_time", None)
        age = f"{(datetime.now() - rt).total_seconds() / 3600:.1f}" if rt else ""
        folder = kpi.folder("warehouse_orders", kpi_folder, output_path)
        # invalid pack configs are a master data measure: each one is a line someone works out by hand
        kpi.append_csv(folder, RUNS_LOG, RUNS_HEADER, [[
            run_id, when, user, name, age, sum(counts.values()), counts["OK"], sum(d.startswith("SHORT") for d in detail),
            detail.count("NOT FOUND"), sum("Pack Config" in d for d in detail), detail.count("Missing data"), dates_filled]], log)
        kpi.append_csv(folder, ISSUES_LOG, ISSUES_HEADER, [[run_id, when, name, r, item, batch, qty, status, _issue_type(msg), msg]
                                                           for r, item, batch, qty, status, msg in issues], log)
    return {"output": output_path, "run_id": run_id, **counts}
