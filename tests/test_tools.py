import os
import sys
from datetime import datetime

import pandas as pd
import pytest
from openpyxl import load_workbook

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
sys.path.insert(0, os.path.join(ROOT, "sample_data"))

import make_samples  # noqa: E402
from stocktools import movement_request, packconfig, packout, warehouse_orders  # noqa: E402
from stocktools.codes import codes_match, normalize_code  # noqa: E402
from stocktools.inventory import Inventory  # noqa: E402

SAMPLES = os.path.join(ROOT, "sample_data")


@pytest.fixture(scope="session", autouse=True)
def samples():
    make_samples.main()


@pytest.fixture
def inventory():
    return Inventory(os.path.join(SAMPLES, "3pl_inventory_sample.csv"), "OWNER_A")


@pytest.mark.parametrize("raw, per_case", [
    ("5000E25C1P", 200), ("2400B60C1P", 40), ("8160E60C1P", 136),
    ("100E1C", 100), ("20S1C", 20), ("1C", 1), ("5000E1P", 5000),
])
def test_pack_config_units_per_case(raw, per_case):
    assert packconfig.parse_config(raw)["units_per_case"] == per_case


@pytest.mark.parametrize("raw", [None, "", "N/A", "nan", "XYZ", float("nan")])
def test_pack_config_unreadable(raw):
    assert packconfig.parse_config(raw) is None


def test_units_per_pallet():
    assert packconfig.parse_config("5000E25C1P")["units_per_pallet"] == 5000


def test_case_rounding():
    cfg = packconfig.parse_config("5000E25C1P")
    assert packconfig.cases_needed(4500, cfg) == (22, 100)
    assert packconfig.cases_rounded_up(4500, cfg) == 23


def test_leading_zeros():
    assert normalize_code(" 0361 ") == "361"
    assert normalize_code("000") == "000"
    assert codes_match("97709", "097709")
    assert not codes_match("9770", "097709")


def test_inventory_matches_without_leading_zeros(inventory):
    free, rows = inventory.free_stock("361", "97709")
    assert free == 2400
    assert rows.iloc[0]["Stock Code"] == "0361"


def test_inventory_filters_owner(inventory):
    assert inventory.free_stock("009003")[0] == 0


def test_inventory_sums_all_batches(inventory):
    assert inventory.free_stock("009001")[0] == 12000


def test_movement_request(tmp_path, inventory):
    result = movement_request.build(
        os.path.join(SAMPLES, "erp_order_export_sample.csv"),
        os.path.join(SAMPLES, "movement_request_template.xlsx"),
        str(tmp_path), inventory=inventory, case_sku_pattern=".(CS).", log=lambda m: None)
    ws = load_workbook(result["output"])["Request Sheet"]
    rows = [[ws.cell(r, c).value for c in range(1, 5)] for r in range(4, 9)]
    assert rows[0] == ["009001", "045120", 4500, 23]
    assert rows[1] == ["0361", "097709", 200, 5]
    assert rows[2] == ["AB10CS02", "B7731", 12, 12]
    assert rows[3][3] == 3
    assert [i["part"] for i in result["issues"]] == ["009002", "009004"]
    assert ws.cell(7, 1).fill.start_color.rgb.endswith("FFFF99")
    assert ws.cell(8, 1).fill.start_color.rgb.endswith("FFCCCC")
    assert ws.cell(8, 1).comment is not None


def test_next_working_day():
    friday_morning = datetime(2026, 11, 20, 9)
    friday_afternoon = datetime(2026, 11, 20, 15)
    thursday_afternoon = datetime(2026, 11, 19, 15)
    assert warehouse_orders.next_working_day(friday_morning).date().isoformat() == "2026-11-23"
    assert warehouse_orders.next_working_day(friday_afternoon).date().isoformat() == "2026-11-23"
    assert warehouse_orders.next_working_day(thursday_afternoon).date().isoformat() == "2026-11-23"


def test_warehouse_orders(tmp_path, inventory):
    out = str(tmp_path / "out.xlsx")
    result = warehouse_orders.process(
        os.path.join(SAMPLES, "warehouse_order_sample.xlsx"), inventory, out,
        now=datetime(2026, 11, 19, 9), log=lambda m: None)
    ws = load_workbook(out).active
    assert [ws.cell(r, 4).value for r in range(4, 9)] == [
        25, 60, "SHORT (100 available)", "NOT FOUND", "Missing data"]
    assert ws.cell(4, 5).value.date().isoformat() == "2026-11-20"
    assert (result["OK"], result["SHORT"], result["WARNING"]) == (2, 2, 1)


def test_packout(tmp_path):
    out = str(tmp_path / "packed.xlsx")
    paths = [os.path.join(SAMPLES, "erp_shipping_export_old.csv"),
             os.path.join(SAMPLES, "erp_shipping_export_new.csv")]
    result = packout.run(paths, os.path.join(SAMPLES, "packing_list_sample.xlsx"), out, log=lambda m: None)
    df = pd.read_excel(out, dtype=str, keep_default_na=False)
    assert df["SO"].tolist() == ["50240", "50101", "50102", "50241", "NOT FOUND", "49999"]
    assert df.loc[3, "Line"] == "3" and df.loc[3, "Rel"] == "2"
    assert (result["matched"], result["not_found"], result["skipped"]) == (4, 1, 1)


def test_packout_multi_mqr_allocation(tmp_path):
    """One shipment, three movement requests, the same part and batch on several releases."""
    out = str(tmp_path / "packed.xlsx")
    paths = [os.path.join(SAMPLES, f"erp_shipping_export_mqr{i}.csv") for i in (1, 2, 3)]
    result = packout.run(paths, os.path.join(SAMPLES, "packing_list_multi_mqr_sample.xlsx"), out, log=lambda m: None)
    df = pd.read_excel(out, sheet_name="Packing list", dtype=str, keep_default_na=False)
    got = [(r["Pack line"], r["Quantity"], f'{r["SO"]}/{r["Line"]}/{r["Rel"]}', r["Match"]) for _, r in df.iterrows()]
    assert got == [
        ("2", "300", "50101/1/1", "Combination"), ("2", "200", "50102/2/1", "Combination"),
        ("3", "150", "50240/1/1", "Exact"),
        ("4", "60", "50300/4/1", "Oldest first"),
        ("5", "20", "50300/4/1", "Oldest first"),
        ("6", "10", "NO OPEN RELEASE//", ""),
        ("7", "100", "50101/3/1", "Oldest first"), ("7", "20", "50240/2/1", "Oldest first"),
        ("8", "80", "50240/2/1", "Oldest first"),
    ]
    # every release is filled exactly once, and the over-pack is flagged on the Check sheet
    placed = df[df["Match"] != ""].groupby(df["SO"] + "/" + df["Line"] + "/" + df["Rel"])["Quantity"].apply(
        lambda s: sum(float(x) for x in s)).to_dict()
    assert placed == {"50101/1/1": 300, "50102/2/1": 200, "50240/1/1": 150, "50300/4/1": 80,
                      "50101/3/1": 100, "50240/2/1": 100}
    chk = pd.read_excel(out, sheet_name="Check", dtype=str)
    assert dict(zip(chk["Part"], chk["Status"])) == {"009301": "Packed more than released", "009302": "Ties"}
    assert result["split"] == 2


def test_packout_releases_left_off_shipment(tmp_path, monkeypatch):
    """Rolling-order customer: release stays open. Any other customer: needs a new order."""
    monkeypatch.setattr(packout, "ROLLING_ORDER_CUSTOMERS", {"INTERCO-US"})
    erp = tmp_path / "mqr.csv"
    erp.write_text("PartNum,LotNum,Quantity,Reference,Customer\n"
                   "009401,B1,100,60001/1/1,INTERCO-US\n"
                   "009401,B1,50,60002/1/1,CUST-A\n"
                   "009402,B2,40,60003/2/1,CUST-B\n"
                   "009403,B3,10,60004/1/1,\n")
    from openpyxl import Workbook
    wb = Workbook(); ws = wb.active
    ws.append(["Part Number", "Batch", "Quantity", "SO", "Line", "Rel"])
    ws.append(["009401", "B1", "120", "", "", ""])      # 100 + 20 of 50: CUST-A short by 30
    wb.save(tmp_path / "pack.xlsx")
    out = str(tmp_path / "out.xlsx")
    result = packout.run([str(erp)], str(tmp_path / "pack.xlsx"), out, log=lambda m: None)
    left = pd.read_excel(out, sheet_name="Left off shipment", dtype=str)
    got = {r["SO/Line/Rel"]: (r["Left off shipment"], r["Next step"]) for _, r in left.iterrows()}
    assert got == {
        "60002/1/1": ("30", "Not shipped: needs a new order"),
        "60003/2/1": ("40", "Not shipped: needs a new order"),
        "60004/1/1": ("10", "Customer unknown: check whether it ships later or needs a new order"),
    }
    erp.write_text("PartNum,LotNum,Quantity,Reference,Customer\n009401,B1,200,60001/1/1,INTERCO-US\n")
    packout.run([str(erp)], str(tmp_path / "pack.xlsx"), out, log=lambda m: None)
    left = pd.read_excel(out, sheet_name="Left off shipment", dtype=str)
    assert left.iloc[0]["Next step"] == "Stays open: ships on a later shipment (rolling orders)"
    assert left.iloc[0]["Left off shipment"] == "80"


from stocktools import provision  # noqa: E402


def test_provision_rerun():
    stock = pd.read_csv(os.path.join(SAMPLES, "aged_stock_sample.csv"), dtype=str)
    df, summary = provision.rerun(stock, "2026-09-30", provision_sites=["Site C"],
                                  open_disposal_lots=[("009109", "L009")])
    got = dict(zip(df["part"], df["criterion"]))
    assert got["009101"] == "4 - Site provision"
    assert got["009102"] == got["009103"] == "1 - Expires within 12m"
    assert got["009104"] == got["009107"] == "3 - Over 2yr no use"
    assert got["009105"] == got["009106"] == "2 - SLOB"
    assert got["009110"] == got["009111"] == "None"
    excl = dict(zip(df["part"], df["excluded_by"]))
    assert excl["009106"] == "Planned use"
    assert excl["009107"] == "Disposal bin"
    assert excl["009108"] == "Quarantine warehouse"
    assert excl["009109"] == "Open disposal form"
    assert summary["total_value"] == 1050.00
    assert (summary["lines_included"], summary["lines_excluded"]) == (5, 4)
    assert summary["criteria_dates"] == {"expiry_by": "2027-09-30", "slob_before": "2025-09-30",
                                         "no_use_2yr_before": "2024-09-30"}


def test_provision_movement():
    assert provision.movement(1050.0, 1200.0) == {
        "prior": 1200.0, "current": 1050.0, "change": -150.0, "change_pct": -12.5}
