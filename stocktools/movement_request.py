"""Movement request builder.

Takes an ERP order export (CSV) and turns it into the 3PL's movement
request spreadsheet: part, batch and quantity copied into the template,
case count worked out from the 3PL's own pack config, and any line the
3PL can't cover highlighted with a note on the cell.
"""
import os
import re
from datetime import datetime

import pandas as pd
from openpyxl import load_workbook
from openpyxl.comments import Comment
from openpyxl.styles import PatternFill

from . import packconfig

COLUMN_ALIASES = {
    "part": ["PartNum", "Part Num", "Part Number"],
    "lot": ["LotNum", "Lot Num", "Lot Number", "Batch"],
    "customer": ["Customer", "Cust Name", "CustomerName"],
    "qty": ["Quantity", "Qty", "Amount"],
}
SHEET = "Request Sheet"
FIRST_ROW = 4
RED = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")
YELLOW = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")


def find_column(df, names):
    """Match a column by any of its known names, ignoring case and spacing."""
    lookup = {c.strip().lower(): c for c in df.columns}
    for name in names:
        if name.strip().lower() in lookup:
            return lookup[name.strip().lower()]
    return None


def parse_qty(value):
    try:
        return float(str(value).strip().replace(",", ""))
    except ValueError:
        return 0.0


def stock_status(free, requested):
    if free >= requested:
        return "OK", "ok"
    if free > 0:
        return f"SHORT ({free:.0f} avail)", "short"
    return "NO STOCK", "none"


def build(erp_csv, template, output_dir, inventory=None, case_sku_pattern=None, log=print):
    """Build the movement request. Returns a summary dict."""
    df = pd.read_csv(erp_csv, dtype=str)
    cols = {key: find_column(df, names) for key, names in COLUMN_ALIASES.items()}
    missing = [k for k, v in cols.items() if v is None]
    if missing:
        raise ValueError(f"Missing columns: {', '.join(missing)}. Found: {', '.join(df.columns)}")

    customer = str(df[cols["customer"]].iloc[0]).split("\n")[0].strip()
    safe_customer = re.sub(r"[^\w\s-]", "", customer).replace(" ", "_")
    case_sku = re.compile(case_sku_pattern) if case_sku_pattern else None
    log(f"Customer: {customer}, {len(df)} lines")

    wb = load_workbook(template)
    if SHEET not in wb.sheetnames:
        raise ValueError(f"'{SHEET}' not found in template")
    ws = wb[SHEET]

    report_time = inventory.report_time.strftime("%d %b %Y %H:%M") if inventory else None
    issues, cases_written = [], 0

    for i, row in df.reset_index(drop=True).iterrows():
        r = FIRST_ROW + i
        part = str(row[cols["part"]]).strip()
        lot = str(row[cols["lot"]]).strip()
        qty = parse_qty(row[cols["qty"]])
        line = f"{part} | lot {lot} | qty {qty:,.0f}"

        cases = None
        if inventory is not None:
            free, rows = inventory.free_stock(part, lot)
            # Write the 3PL's version of the code if the ERP dropped a leading zero.
            if not rows.empty:
                part = rows.iloc[0]["Stock Code"]
                if lot.upper() != "N/A":
                    lot = rows.iloc[0]["Batch_id"]

            status, kind = stock_status(free, qty)
            if kind != "ok":
                issues.append({"part": part, "lot": lot, "requested": qty, "available": free})

            cfg_str = inventory.pack_config(part)
            cfg = packconfig.parse_config(cfg_str)
            if case_sku and case_sku.search(part.upper()) and qty > 0:
                cases = int(qty)  # already ordered in cases
                line += f" | case SKU: {cases} case(s)"
            elif cfg and cfg["units_per_case"] and qty > 0:
                cases = packconfig.cases_rounded_up(qty, cfg)
                line += f" | pack {cfg_str} ({cfg['units_per_case']:.0f}/case) | {cases} case(s)"
            line += f" | {status}"

            if kind != "ok":
                fill = RED if kind == "none" else YELLOW
                for c in range(1, 5):
                    ws.cell(r, c).fill = fill
                ws.cell(r, 1).comment = Comment(
                    f"{status} in 3PL report of {report_time}.\n"
                    "May have been booked in since. Check with the 3PL before chasing.",
                    "Movement Request Tool", width=250, height=80)

        ws.cell(r, 1, part).number_format = "@"
        ws.cell(r, 2, lot).number_format = "@"
        ws.cell(r, 3, int(qty)).number_format = "0"
        if cases is not None:
            ws.cell(r, 4, cases)
            cases_written += 1
        log(line)

    os.makedirs(output_dir, exist_ok=True)
    out = os.path.join(output_dir, f"{safe_customer}_{datetime.now():%d%b%Y_%H%M%S}.xlsx")
    wb.save(out)

    if issues:
        log(f"{len(issues)} stock issue(s). 3PL report time: {report_time}; "
            "anything booked in since then won't show.")
    return {"output": out, "customer": customer, "lines": len(df),
            "cases_written": cases_written, "issues": issues, "report_time": report_time}
