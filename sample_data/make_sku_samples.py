"""Synthetic files for the single SKU reconciliation tool. Part 009401 has
four lots, each set up with a different kind of difference:

- L401: the 3PL shipped a second release of an order that the ERP never shipped
- L402: the same order's second release, on another lot
- L403: an ERP quantity adjustment with no matching 3PL movement
- L404: the 3PL holds stock under one owner, the ERP under another

    python sample_data/make_sku_samples.py
"""
import csv
import os

from openpyxl import Workbook

HERE = os.path.dirname(os.path.abspath(__file__))
PART = "009401"

TRACE = [
    # Code, Batch ID, Owner ID, Update Qty, Reference, Date, Hours
    ("Receipt", "L404", "OWNER_A", 48, "RCPT-0001", "01-Jun-2026", "09:00"),
    ("Receipt", "L401", "OWNER_A", 1260, "RCPT-0002", "09-Sep-2026", "08:00"),
    ("Owner Update", "L401", "OWNER_A", -1260, "TP-0001", "10-Sep-2026", "10:00"),
    ("Owner Update", "L401", "OWNER_B", 1260, "TP-0001", "10-Sep-2026", "10:00"),
    ("Shipment", "L401", "OWNER_B", 300, "100234BK1", "23-Sep-2026", "14:00"),
    ("Shipment", "L401", "OWNER_B", 823, "100300", "25-Sep-2026", "11:00"),
    ("Allocate", "L401", "OWNER_B", 137, "100234BK3", "28-Sep-2026", "09:00"),
    ("Shipment", "L401", "OWNER_B", 137, "100234BK3", "01-Oct-2026", "16:18"),
    ("Receipt", "L402", "OWNER_A", 3384, "RCPT-0003", "12-Sep-2026", "08:00"),
    ("Owner Update", "L402", "OWNER_A", -3384, "TP-0002", "12-Sep-2026", "12:00"),
    ("Owner Update", "L402", "OWNER_B", 3384, "TP-0002", "12-Sep-2026", "12:00"),
    ("Shipment", "L402", "OWNER_B", 3170, "100400", "20-Sep-2026", "10:00"),
    ("Shipment", "L402", "OWNER_B", 163, "100234BK3", "01-Oct-2026", "16:18"),
    ("Return", "L403", "OWNER_B", 1, "100500", "09-Oct-2026", "10:00"),
    ("Shipment", "L403", "OWNER_B", 1, "100600", "22-Oct-2026", "10:00"),
]

ERP = [
    # tran, date, whse, warehouse, bin, class, type, qty, lot, order, reason, rma, sys seconds
    (1, "09/09/2026", "TP01", "3PL Main", "TP01", "R", "PUR-STK", 1260, "L401", "", "", "", 30000),
    (2, "10/09/2026", "TP01", "3PL Main", "TP01", "S", "STK-PLT", 1260, "L401", "", "", "", 36000),
    (3, "10/09/2026", "SL01", "External Warehouse", "SL01", "R", "PLT-STK", 1260, "L401", "", "", "", 36100),
    (4, "24/09/2026", "SL01", "External Warehouse", "SL01", "S", "STK-CUS", 300, "L401", "100234", "", "", 40000),
    (5, "25/09/2026", "SL01", "External Warehouse", "SL01", "S", "STK-CUS", 823, "L401", "100300", "", "", 40000),
    (6, "12/09/2026", "TP01", "3PL Main", "TP01", "R", "MFG-STK", 3384, "L402", "", "", "", 30000),
    (7, "12/09/2026", "TP01", "3PL Main", "TP01", "S", "STK-PLT", 3384, "L402", "", "", "", 43000),
    (8, "12/09/2026", "SL01", "External Warehouse", "SL01", "R", "PLT-STK", 3384, "L402", "", "", "", 43100),
    (9, "20/09/2026", "SL01", "External Warehouse", "SL01", "S", "STK-CUS", 3170, "L402", "100400", "", "", 40000),
    (10, "12/10/2026", "SL01", "External Warehouse", "SL01", "R", "INS-STK", 1, "L403", "", "", "RMA-17", 40000),
    (11, "22/10/2026", "SL01", "External Warehouse", "SL01", "S", "STK-CUS", 1, "L403", "100600", "", "", 40000),
    (12, "22/10/2026", "SL01", "External Warehouse", "SL01", "A", "ADJ-QTY", 1, "L403", "", "COUNT", "", 40100),
    (13, "22/10/2026", "SL01", "External Warehouse", "SL01", "S", "STK-CUS", -1, "L403", "100600", "", "", 40200),
    (14, "22/10/2026", "SL01", "External Warehouse", "SL01", "S", "STK-CUS", 1, "L403", "100600", "", "", 40300),
    (15, "21/09/2026", "SL01", "External Warehouse", "SL01", "I", "STK-UKN", -48, "L404", "", "FOUND", "", 40000),
]

TPL_ONHAND = "Lot\tOwner\tQty\nL402\tOWNER_B\t51\nL404\tOWNER_A\t48\n"
ERP_ONHAND = "Lot\tWarehouse\tQty\nL402\tSL01\t214\nL401\tSL01\t137\nL403\tSL01\t1\nL404\tSL01\t48\n"

HEADER = ["Tran", "Date", "Part", "Description", "Rev", "Whse", "Warehouse", "Bin", "Type", "Type", "Invty",
          "Quantity", "UOM", "Mtl Unit Cost", "Entry Person", "Reference", "Job", "Lot Number", "Pack",
          "Pack Line", "PO", "Line", "Order", "Line", "Supplier Num", "PO Receipt Qty", "Packing Slip",
          "Invoice", "Line", "Reason Code", "DMR Number", "RMA", "ID", "User", "Employee", "Sys Date", "Sys Time"]


def main():
    wb = Workbook(); ws = wb.active
    ws.append(["Code", "Batch ID", "Owner ID", "Update Qty", "Reference", "Date", "Hours", "Sku ID", "Notes"])
    for r in TRACE:
        ws.append(list(r) + [PART, ""])
    wb.save(os.path.join(HERE, "sku_3pl_trace_sample.xlsx"))
    with open(os.path.join(HERE, "sku_erp_history_sample.txt"), "w", newline="") as f:
        w = csv.writer(f, delimiter="\t")
        w.writerow(HEADER)
        for t, d, wc, wn, b, cl, tt, q, lot, order, reason, rma, st in ERP:
            r = [""] * len(HEADER)
            r[0], r[1], r[2], r[5], r[6], r[7], r[8], r[9], r[10] = str(t), d, PART, wc, wn, b, cl, tt, "true"
            r[11], r[17], r[22], r[29], r[31], r[33], r[35], r[36] = str(q), lot, order, reason, rma, "user1", d, str(st)
            w.writerow(r)
    with open(os.path.join(HERE, "sku_onhand_sample.txt"), "w") as f:
        f.write("3PL on hand (paste into the first box)\n" + TPL_ONHAND +
                "\nERP on hand (paste into the second box)\n" + ERP_ONHAND)


if __name__ == "__main__":
    main()
