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
    assert all(r[5] == "Yes" and r[6] == "Yes" for r in s.values())


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
