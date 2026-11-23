"""Generate synthetic files for the stock reconciliation tool. Every part,
lot, warehouse and order is made up. Each line is set up to land on one
reason, so the tests can check the rules.

    python sample_data/make_recon_samples.py
"""
import csv
import datetime
import os

from openpyxl import Workbook

HERE = os.path.dirname(os.path.abspath(__file__))
REPORT_DATE = datetime.date(2026, 5, 15)

ERP = [
    # Site, Warehouse, Part, Description, Bin, Lot Number, On Hand, UOM, Value On Hand
    ("Supply Chain", "3PL Main", "PALLET-01", "Pallet 1200x1000", "TP01", "", 10, "EA", 0),
    ("Supply Chain", "3PL Main", "009201", "Sample lancet 28G", "TP01", "L201", 400, "EA", 40.0),
    ("Supply Chain", "3PL Main", "009204", "Sample pen needle 5mm", "TP01", "L204", 1000, "EA", 50.0),
    ("Supply Chain", "3PL Main", "009205", "Sample syringe 1ml", "TP01", "L205A", 300, "EA", 30.0),
    ("Sales", "External Warehouse", "009206", "Sample swab pack", "SL01", "L206", 10, "EA", 2.0),
    ("Supply Chain", "3PL Main", "009207", "Sample lancet 30G", "TP01", "L207B", 60, "EA", 6.0),
    ("Supply Chain", "3PL Main", "009208", "Sample sharps bin", "TP01", "L208", 100, "EA", 100.0),
    ("Supply Chain", "3PL Main", "EDGE-01", "Edge protector", "TP01", "", 30, "EA", 3.0),
    ("Supply Chain", "3PL Main", "009209", "Sample pen needle 8mm", "TP01", "L209", 500, "EA", 25.0),
    ("Supply Chain", "3PL Overflow", "009210", "Sample lancet 21G", "TP02", "L210", 80, "EA", 8.0),
    ("Sales", "External Warehouse", "009211", "Sample needle 6mm", "SL01", "L211", 25, "EA", 5.0),
    ("Supply Chain", "3PL Main", "P900001", "Pilot device sub-assembly", "TP01", "H2", 40, "EA", 400.0),
    ("Supply Chain", "3PL Main", "009212", "Sample test strip", "TP01", "L212", 70, "EA", 7.0),
    ("Sales", "Ecomm Warehouse", "009213", "Sample shop kit", "SHOP", "L213", 5, "EA", 5.0),
    ("Sales", "Ecomm Warehouse", "009213", "Sample shop kit", "FBA", "L213", 10, "EA", 10.0),
    ("Supply Chain", "3PL Main", "009214", "Sample cap", "TP01", "045120", 900, "EA", 9.0),
    # onsite warehouse used by the stock take tests (and the "own site" 3PL rule)
    ("Site A", "Stores", "009202", "Sample lancet 23G", "A0", "L202", 50, "EA", 5.0),
    ("Site A", "Stores", "009301", "Sample label", "A1", "L301", 100, "EA", 1.0),
    ("Site A", "Stores", "009302", "Sample carton", "A2", "L302", 50, "EA", 5.0),
    ("Site A", "Stores", "009303", "Sample spring", "A3", "L303", 20, "EA", 2.0),
    ("Site A", "Stores", "009304", "Sample body", "A4", "L304A", 30, "EA", 3.0),
    ("Site A", "Stores", "009305", "Sample tray", "RCV-01", "L305", 15, "EA", 1.5),
    ("Site A", "Stores", "009306", "Sample insert", "A6", "L306", 40, "EA", 4.0),
    ("Site B", "Stores", "009307", "Sample leaflet", "B1", "L307", 12, "EA", 1.2),
]

TPL = [
    # Owner, Sku, Description, Qty, Receive date, Allocated, Batch
    ("OWNER_A", "P900001", "Pilot device sub-assembly", 100, "01/05/2026", 0, "H1"),
    ("OWNER_A", "PALLET-01", "Pallet 1200x1000", 15, "01/05/2026", 0, "NONE"),
    ("OWNER_A", "009201", "Sample lancet 28G", 500, "01/04/2026", 0, "L201"),
    ("OWNER_A", "009202", "Sample lancet 23G", 50, "01/04/2026", 0, "L202"),
    ("OWNER_A", "009203", "Sample lancet 18G", 20, "01/04/2026", 0, "L203"),
    ("OWNER_A", "009204", "Sample pen needle 5mm", 1200, "01/04/2026", 200, "L204"),
    ("OWNER_A", "009205", "Sample syringe 1ml", 300, "01/04/2026", 0, "L205A"),
    ("OWNER_A", "009205", "Sample syringe 1ml", 2000, "14/05/2026", 0, "L205B"),
    ("OWNER_B", "009206", "Sample swab pack", 50, "01/04/2026", 0, "L206"),
    ("OWNER_A", "009207", "Sample lancet 30G", 60, "01/04/2026", 0, "L207A"),
    ("OWNER_A", "009208", "Sample sharps bin", 105, "01/04/2026", 0, "L208"),
    ("OWNER_A", "EDGE-01", "Edge protector", 20, "01/04/2026", 0, "NONE"),
    ("OWNER_A", "009212", "Sample test strip", 50, "01/04/2026", 0, "L212"),
    ("OWNER_D", "009213", "Sample shop kit", 5, "01/04/2026", 0, "L213"),
    ("OWNER_A", 9214, "Sample cap", 900, "01/04/2026", 0, 45120),   # Excel stripped the zeros
]

TXNS = [
    # Whse, Warehouse, Bin, Part, Tran type, Lot, Quantity
    ("TP01", "3PL Main", "TP01", "009201", "STK-CUS", "L201", "100"),
    ("SL01", "External Warehouse", "SL01", "009206", "PLT-STK", "L206", "40"),
    ("TP02", "3PL Overflow", "TP02", "009210", "MFG-STK", "L210", "80"),
    ("SL01", "External Warehouse", "SL01", "009211", "PLT-STK", "L211", "25"),
]

ORDERS = [("TP-140526-1", "009209", "L209", "500")]

COUNT = [
    # Part, Lot, Bin, Counted Qty
    ("009202", "L202", "A0", 50),
    ("009301", "L301", "A1", 100),
    ("009302", "L302", "A2", 50),
    ("009303", "L303", "B9", 20),
    ("009304", "L304B", "A4", 30),
    ("009307", "L307", "A7", 12),
    ("FOUND-1", "X1", "A8", 5),
]

TXN_HEADER = ["Tran", "Date", "Part", "Description", "Rev", "Whse", "Warehouse", "Bin", "Type", "Type", "Invty",
              "Quantity", "UOM", "Mtl Unit Cost", "Entry Person", "Reference", "Job", "Lot Number", "Pack",
              "Pack Line", "PO", "Line", "Order", "Line", "Supplier Num", "PO Receipt Qty", "Packing Slip",
              "Invoice", "Line", "Reason Code", "DMR Number", "RMA", "ID", "User", "Employee", "Sys Date",
              "Sys Time"]


def main():
    wb = Workbook(); ws = wb.active; ws.title = "Stock"
    ws.append(["Site", "Warehouse", "Part", "Description", "Bin", "Lot Number", "On Hand", "UOM", "Value On Hand"])
    for r in ERP:
        ws.append(list(r))
    wb.save(os.path.join(HERE, "erp_stock_report_sample.xlsx"))

    wb = Workbook(); ws = wb.active; ws.title = "Full inventory"
    ws.append(["Tag", "Client", "Owner", "Sku", "Description", "Qty", "Pack config", "Receive date",
               "Location", "Lock code", "Allocated", "Batch"])
    for i, (owner, sku, desc, qty, rd, alloc, batch) in enumerate(TPL, 1):
        ws.append([f"T{i:05d}", "CLIENT", owner, sku, desc, qty, "1E", rd, "LOC1", "UnLocked", alloc, batch])
    wb.save(os.path.join(HERE, "3pl_full_inventory_sample.xlsx"))

    with open(os.path.join(HERE, "erp_transactions_sample.txt"), "w", newline="") as f:
        w = csv.writer(f, delimiter="\t")
        w.writerow(TXN_HEADER)
        for n, (wc, wn, bn, part, tt, lot, qty) in enumerate(TXNS, 1):
            row = [""] * len(TXN_HEADER)
            row[0], row[1], row[2], row[5], row[6], row[7] = str(n), "14/05/2026", part, wc, wn, bn
            row[8], row[9], row[11], row[17], row[35] = "", tt, qty, lot, "14/05/2026"
            w.writerow(row)

    with open(os.path.join(HERE, "3pl_order_summary_sample.csv"), "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["Client ID", "Owner ID", "Deliver By Date", "Order ID", "Pallet ID", "SKU ID", "Batch ID",
                    "Qty Shipped"])
        for order, sku, batch, qty in ORDERS:
            w.writerow(["CLIENT", "OWNER_A", "15/05/2026", order, "P1", sku, batch, qty])

    with open(os.path.join(HERE, "stock_take_count_sample.csv"), "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["Part", "Lot", "Bin", "Counted Qty"])
        w.writerows(COUNT)


if __name__ == "__main__":
    main()
