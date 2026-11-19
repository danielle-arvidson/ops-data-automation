"""Generate the synthetic sample files. All part numbers, batches, owners
and customers are made up.

    python sample_data/make_samples.py
"""
import csv
import os

from openpyxl import Workbook

HERE = os.path.dirname(os.path.abspath(__file__))

INVENTORY = [
    # Stock Code, Description, Batch_id, Owner ID, Total, Allocated, Free, Receipt, In Date, Pack Config
    ("009001", "Sample lancet 30G", "045120", "OWNER_A", 10000, 0, 10000, "R1001", "02/09/2026", "5000E25C1P"),
    ("009001", "Sample lancet 30G", "045121", "OWNER_A", 3000, 1000, 2000, "R1002", "14/09/2026", "5000E25C1P"),
    ("0361", "Sample pen needle 4mm", "097709", "OWNER_A", 2400, 0, 2400, "R1003", "20/09/2026", "2400B60C1P"),
    ("AB10CS02", "Sample cased kit", "B7731", "OWNER_A", 40, 0, 40, "R1004", "01/10/2026", "1C"),
    ("009002", "Sample sharps bin", "045300", "OWNER_A", 150, 50, 100, "R1005", "05/10/2026", "100E1C"),
    ("009003", "Sample swab pack", "045400", "OWNER_B", 500, 0, 500, "R1006", "06/10/2026", "20S1C"),
]

ERP_ORDER = [
    # PartNum, LotNum, Customer, Quantity
    ("9001", "45120", "Example Healthcare Ltd", "4,500"),
    ("361", "97709", "Example Healthcare Ltd", "200"),
    ("AB10CS02", "B7731", "Example Healthcare Ltd", "12"),
    ("009002", "045300", "Example Healthcare Ltd", "250"),
    ("009004", "N/A", "Example Healthcare Ltd", "10"),
]

ERP_SHIP_OLD = [("009001", "045120", "1000", "50101/1/1"), ("0361", "097709", "200", "50102/2/1")]
ERP_SHIP_NEW = [("009001", "045120", "600", "50240/1/1"), ("009002", "045300", "100", "50241/3/2")]


def main():
    with open(os.path.join(HERE, "3pl_inventory_sample.csv"), "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["Inventory Summary - Stored Goods (synthetic sample)"])
        w.writerow(["Stock Code", "Description", "Batch_id", "Owner ID", "Total",
                    "Allocated", "Free", "Receipt", "In Date", "Pack Config"])
        w.writerows(INVENTORY)

    with open(os.path.join(HERE, "erp_order_export_sample.csv"), "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["PartNum", "LotNum", "Customer", "Quantity"])
        w.writerows(ERP_ORDER)

    for name, rows in (("erp_shipping_export_old.csv", ERP_SHIP_OLD),
                       ("erp_shipping_export_new.csv", ERP_SHIP_NEW)):
        with open(os.path.join(HERE, name), "w", newline="") as f:
            w = csv.writer(f)
            w.writerow(["PartNum", "LotNum", "Quantity", "Reference"])
            w.writerows(rows)

    wb = Workbook()
    ws = wb.active
    ws.title = "Request Sheet"
    ws["A1"] = "Movement Request from Off-Site Storage"
    for col, head in enumerate(["Item No", "Batch No", "Qty", "QtyBoxes/Bags"], start=1):
        ws.cell(3, col, head)
    wb.save(os.path.join(HERE, "movement_request_template.xlsx"))

    wb = Workbook()
    ws = wb.active
    ws["A1"] = "Warehouse Order (synthetic sample)"
    for col, head in enumerate(["Item No", "Batch No", "Qty", "QtyBoxes/Bags", "Date Required"], start=1):
        ws.cell(3, col, head)
    for r, line in enumerate([("009001", "045120", 5000), ("361", "097709", 2400),
                              ("009002", "045300", 500), ("009009", "000001", 10),
                              ("009001", None, 100)], start=4):
        for c, v in enumerate(line, start=1):
            ws.cell(r, c, v)
    wb.save(os.path.join(HERE, "warehouse_order_sample.xlsx"))

    wb = Workbook()
    ws = wb.active
    ws.append(["Part Number", "Batch", "Quantity", "SO", "Line", "Rel"])
    ws.append(["9001", "45120", "600", "", "", ""])
    ws.append(["009001", "045120", "1000", "", "", ""])
    ws.append(["0361", "097709", "200", "", "", ""])
    ws.append(["009002", "045300", "100", "", "", ""])
    ws.append(["009005", "045999", "10", "", "", ""])
    ws.append(["009002", "045300", "50", "49999", "1", "1"])
    for row in ws.iter_rows():
        for cell in row:
            cell.number_format = "@"
    wb.save(os.path.join(HERE, "packing_list_sample.xlsx"))
    print("Sample files written to", HERE)


if __name__ == "__main__":
    main()
