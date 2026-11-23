import datetime
import os
import sys

import pytest
from openpyxl import load_workbook

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
sys.path.insert(0, os.path.join(ROOT, "sample_data"))

import make_recon_samples  # noqa: E402
from stocktools import reconciliation as rec  # noqa: E402

S = os.path.join(ROOT, "sample_data")


@pytest.fixture(scope="module", autouse=True)
def samples():
    make_recon_samples.main()


def reasons(path, tab):
    ws = load_workbook(path)[tab]
    return {(r[1], r[3] or ""): r[9] for r in ws.iter_rows(min_row=2, values_only=True)}


def test_3pl_reasons_and_net(tmp_path):
    out = str(tmp_path / "r.xlsx")
    rec.run_3pl(os.path.join(S, "erp_stock_report_sample.xlsx"), os.path.join(S, "3pl_full_inventory_sample.xlsx"),
                out, os.path.join(S, "erp_transactions_sample.txt"),
                [os.path.join(S, "3pl_order_summary_sample.csv")], datetime.date(2026, 5, 15), log=lambda m: None)
    hi = reasons(out, "3PL higher")
    assert hi[("009201", "L201")] == rec.R_INVOICED
    assert hi[("009204", "L204")] == rec.R_INVOICED
    assert hi[("009205", "L205B")] == rec.R_RECEIVED_3PL
    assert hi[("009202", "L202")] == rec.R_OTHER_SITE
    assert hi[("009203", "L203")] == rec.R_NOT_AT_3PL
    assert hi[("009206", "L206")] == rec.R_TRANSIT
    assert hi[("009207", "L207A")] == rec.R_OFFSET
    assert hi[("P900001", "H1")] == rec.R_HOLD
    assert hi[("PALLET-01", "")] == rec.R_NONINV
    assert hi[("009208", "L208")] == rec.R_TBD
    lo = reasons(out, "ERP higher")
    assert lo[("009209", "L209")] == rec.R_DESPATCHED
    assert lo[("009210", "L210")] == rec.R_RECEIVED_EPI
    assert lo[("009211", "L211")] == rec.R_TRANSIT
    # leading zeros stripped by Excel still match; non-shop ecomm bins are left out
    assert ("009214", "045120") not in hi and ("009214", "045120") not in lo
    assert ("009213", "L213") not in lo


def test_3pl_totals_tie(tmp_path):
    out = str(tmp_path / "r.xlsx")
    rec.run_3pl(os.path.join(S, "erp_stock_report_sample.xlsx"), os.path.join(S, "3pl_full_inventory_sample.xlsx"),
                out, log=lambda m: None)
    wb = load_workbook(out)
    hi = sum(r[5] - r[4] for r in wb["3PL higher"].iter_rows(min_row=2, values_only=True))
    lo = sum(r[4] - r[5] for r in wb["ERP higher"].iter_rows(min_row=2, values_only=True))
    cmp_rows = list(wb["Comparison"].iter_rows(min_row=2, values_only=True))
    assert hi - lo == pytest.approx(sum(r[5] for r in cmp_rows) - sum(r[4] for r in cmp_rows))


def test_stocktake_reasons(tmp_path):
    out = str(tmp_path / "s.xlsx")
    rec.run_stocktake(os.path.join(S, "erp_stock_report_sample.xlsx"), os.path.join(S, "stock_take_count_sample.csv"),
                      out, {("Site A", "Stores")},
                      {"part": "Part", "lot": "Lot", "qty": "Counted Qty", "warehouse": None, "bin": "Bin"},
                      by_bin=True, log=lambda m: None)
    hi, lo = reasons(out, "Count higher"), reasons(out, "ERP higher")
    assert hi[("009303", "L303")] == rec.S_BIN and lo[("009303", "L303")] == rec.S_BIN
    assert hi[("009304", "L304B")] == rec.S_LOT and lo[("009304", "L304A")] == rec.S_LOT
    assert hi[("009307", "L307")] == rec.S_OTHER_WH
    assert hi[("FOUND-1", "X1")] == rec.S_FOUND
    assert lo[("009305", "L305")] == rec.S_RCV
    assert lo[("009306", "L306")] == rec.S_NOTFOUND


def test_kpi_logs(tmp_path):
    out = str(tmp_path / "r.xlsx")
    rec.run_3pl(os.path.join(S, "erp_stock_report_sample.xlsx"), os.path.join(S, "3pl_full_inventory_sample.xlsx"),
                out, os.path.join(S, "erp_transactions_sample.txt"),
                [os.path.join(S, "3pl_order_summary_sample.csv")], datetime.date(2026, 5, 15), log=lambda m: None,
                kpi_folder=str(tmp_path))
    import csv
    with open(tmp_path / "recon_runs_log.csv", newline="") as f:
        run_row = list(csv.DictReader(f))[0]
    wb = load_workbook(out)
    hi = sum(r[5] - r[4] for r in wb["3PL higher"].iter_rows(min_row=2, values_only=True))
    lo = sum(r[4] - r[5] for r in wb["ERP higher"].iter_rows(min_row=2, values_only=True))
    assert float(run_row["Net variance qty"]) == pytest.approx(hi - lo)
    assert float(run_row["Gross variance qty"]) == pytest.approx(hi + lo)
    assert (tmp_path / "recon_reasons_log.csv").exists() and (tmp_path / "recon_area_log.csv").exists()


def test_transactions_supply_cost_for_parts_missing_from_stock_report(tmp_path):
    from stocktools import reconciliation as rc
    hdr = ["Tran", "Date", "Part", "Description", "Rev", "Whse", "Warehouse", "Bin", "Type", "Type", "Invty",
           "Quantity", "UOM", "Mtl Unit Cost", "Lot Number", "Sys Date"]
    rows = [["1", "01/05/2026", "PX1", "d", "", "ZZ99", "Other site", "B1", "I", "STK-MTL", "", "5", "EA", "0", "L1", "01/05/2026"],
            ["2", "02/05/2026", "PX1", "d", "", "ZZ99", "Other site", "B1", "R", "PUR-STK", "", "5", "EA", "2.5", "L1", "02/05/2026"]]
    p = tmp_path / "txn.txt"
    p.write_text("\n".join("\t".join(r) for r in [hdr] + rows), encoding="utf-8")
    shipped, receipts, mtl = rc.load_transactions(str(p))
    assert mtl == {"PX1": 2.5}   # cost taken from any site, first positive value
    assert rc.load_transactions(None)[2] == {}
