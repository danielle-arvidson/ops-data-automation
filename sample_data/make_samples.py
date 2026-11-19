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

AGED_STOCK = [
    # part, lot, site, warehouse_type, bin, expires, first_receipt, last_used, planned_use, value
    ("009101", "L001", "Site C", "Stores", "A1", "", "01/03/2026", "15/09/2026", 0, 100.00),
    ("009102", "L002", "Site A", "Stores", "A2", "01/03/2027", "01/01/2026", "01/09/2026", 0, 200.00),
    ("009103", "L003", "Site A", "Stores", "A3", "30/09/2027", "01/01/2026", "01/09/2026", 0, 50.00),
    ("009104", "L004", "Site A", "Stores", "A4", "", "01/05/2023", "01/01/2024", 0, 300.00),
    ("009105", "L005", "Site B", "Stores", "B1", "", "01/03/2025", "", 0, 400.00),
    ("009106", "L006", "Site B", "Stores", "B2", "", "01/06/2025", "01/08/2025", 10, 150.00),
    ("009107", "L007", "Site A", "Stores", "DISPOSAL", "", "01/01/2023", "", 0, 80.00),
    ("009108", "L008", "Site B", "Quarantine", "Q1", "", "01/01/2024", "", 0, 60.00),
    ("009109", "L009", "Site A", "Stores", "A5", "", "01/02/2025", "", 0, 70.00),
    ("009110", "L010", "Site A", "Stores", "A6", "", "01/01/2025", "01/08/2026", 0, 500.00),
    ("009111", "L011", "Site B", "Stores", "B3", "", "01/01/2026", "", 0, 90.00),
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

    with open(os.path.join(HERE, "aged_stock_sample.csv"), "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["part", "lot", "site", "warehouse_type", "bin", "expires",
                    "first_receipt", "last_used", "planned_use", "value"])
        w.writerows(AGED_STOCK)

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
