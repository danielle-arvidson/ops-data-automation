import csv
import os
import sys

import pytest
from openpyxl import load_workbook

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
sys.path.insert(0, os.path.join(ROOT, "sample_data"))

import make_sku_samples as ms  # noqa: E402
from stocktools import sku_reconciliation as sku  # noqa: E402

S = os.path.join(ROOT, "sample_data")


@pytest.fixture(scope="module", autouse=True)
def samples():
    ms.main()


def run(tmp_path):
    out = str(tmp_path / "sku.xlsx")
    summary, checks = sku.run(ms.PART, [os.path.join(S, "sku_3pl_trace_sample.xlsx")],
                              os.path.join(S, "sku_erp_history_sample.txt"), out, ms.TPL_ONHAND, ms.ERP_ONHAND,
                              log=lambda m: None, kpi_folder=str(tmp_path))
    return {(r[0], r[1]): r for r in summary}, {(c[0], c[1]): c for c in checks}, out


def test_differences_and_causes(tmp_path):
    s, c, _ = run(tmp_path)
    assert s[("OWNER_B", "L401")][4] == -137 and "100234BK3" in s[("OWNER_B", "L401")][9]
    assert s[("OWNER_B", "L402")][4] == -163 and "100234BK3" in s[("OWNER_B", "L402")][9]
    assert s[("OWNER_B", "L403")][4] == -1 and "ADJ-QTY" in s[("OWNER_B", "L403")][9]
    assert s[("OWNER_A", "L404")][4] == 48 and s[("OWNER_B", "L404")][4] == -48
    assert c[("OWNER_A", "L404")][5].startswith("Attribute only")
    assert c[("OWNER_B", "L401")][5] == "Quantity"
    # every trace and history ties to its on-hand figure
    # owners with no typed on-hand use their movement totals and say so
    assert all(r[5] in ("Yes", "Not typed: trace total used") and
               r[6] in ("Yes", "Not typed: history total used") for r in s.values())
    assert s[("OWNER_A", "L404")][6] == "Not typed: history total used"


def test_kpi_log(tmp_path):
    run(tmp_path)
    with open(tmp_path / "sku_investigations_log.csv", newline="") as f:
        rows = list(csv.DictReader(f))
    assert len(rows) == 5
    cats = {(r["Owner"], r["Lot"]): r["Cause category"] for r in rows}
    assert cats[("OWNER_B", "L401")] == "3PL shipment not in ERP"
    assert cats[("OWNER_B", "L403")] == "ERP adjustment or misc"


def test_workbook_tabs(tmp_path):
    _, _, out = run(tmp_path)
    assert load_workbook(out).sheetnames == ["Summary", "Timeline", "Qty differs", "3PL movements", "ERP movements"]


def test_typed_onhand_default_owner():
    oh = sku.parse_onhand("L1 5\nL2, OWNER_B, 7", False, "OWNER_A")
    assert oh[("OWNER_A", "L1")] == 5 and oh[("OWNER_B", "L2")] == 7


def test_lot_tracker_export_and_reversals(tmp_path):
    """A lot tracker export has no Part or Lot column; a posting and its reversal cancel out."""
    from openpyxl import Workbook
    wb = Workbook(); ws = wb.active
    ws.append(["Tran Date", "Type", "Quantity", "UOM", "Warehouse", "Bin", "Order"])
    wh = next(iter(sku.WHSE_OWNER))
    ws.append([__import__("datetime").datetime(2026, 4, 10), "PUR-STK", 400, "EA", wh, wh, 0])
    ws.append([__import__("datetime").datetime(2026, 4, 10), "ADJ-CST", 400, "EA", wh, None, 0])
    for q in (5, -5, 5):
        ws.append([__import__("datetime").datetime(2026, 8, 20), "STK-CUS", q, "EA", wh, wh, 286938])
    path = str(tmp_path / "12940.xlsx"); wb.save(path)
    ev, sys_dated, ctx = sku.load_erp(path, "ANYPART", "12940")
    assert sys_dated == "lot tracker"
    assert [(e["type"], e["qty"]) for e in ev] == [("PUR-STK", 400), ("STK-CUS", -5)]
    assert all(e["lot"] == "12940" for e in ev)


def test_same_movement_different_qty_pairs():
    d = dict(sys="3PL", owner="A", lot="1", type="Receipt", qty=1, ref="R1",
             date=__import__("datetime").datetime(2026, 3, 23))
    e = dict(sys="ERP", owner="A", lot="1", type="STK-STK", qty=3001, ref="", refs=set(),
             date=__import__("datetime").datetime(2026, 3, 24))
    notes = sku.match_events([d], [e])
    assert "qty differs" in d["pair"] and notes[0]["dt"] == 1 and notes[0]["erp"] == 3001


def test_wrong_part_number_stops_with_clear_error(tmp_path):
    from openpyxl import Workbook
    wb = Workbook(); ws = wb.active
    ws.append(["Code", "Update Qty", "Batch ID", "Owner ID", "Date", "SKU"])
    ws.append(["Receipt", 10, "L1", "OWNER_A", __import__("datetime").datetime(2026, 1, 1), "REALPART"])
    p = str(tmp_path / "trace.xlsx"); wb.save(p)
    with pytest.raises(ValueError, match="It is for REALPART"):
        sku.load_3pl_traces([p], "TYPOPART")
