"""Synthetic data for the stock take reconciliation report.

    python powerbi/stock-take/make_synthetic_data.py

Same sheets and columns as the real report's source workbook, apart from
two product-line columns that are renamed, so the .pbix can be pointed at
this file instead (Transform data > Data source settings > Change source).
Every site, part, lot, user and value is made up.
"""
import os
import random
from datetime import date, timedelta

import pandas as pd

random.seed(2026)
HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "stock_take_synthetic.xlsx")

SITES = {"Site A": ["Site A"], "Site B": ["Site B"], "3PL": ["3PL Ops", "3PL Sales", "3PL Export", "3PL Ecomm"]}
TXN_TYPES = ["Qty adjustment (ADJ-QTY)", "Inventory transfer (STK-STK)", "Misc issue / return (STK-UKN)",
             "TO shipped (STK-PLT)", "TO received (PLT-STK)"]
CAUSES = ["Qty Wrong - Batch OK", "No LOT in ERP", "Found - LOT format mismatch", "Qty & Batch Wrong",
          "Zero Stock", "Part not on Input Sheet", "LOT NUMBER ERROR - qty matches", "Different batch - part known"]
ACTION_FOR_CAUSE = {
    "Qty Wrong - Batch OK": ("Adjust Quantity", "Qty Adjustment (ADJ-QTY)", "Requires written approval before posting"),
    "LOT NUMBER ERROR - qty matches": ("Correct LOT in ERP - do not adjust", "Lot reclass (A-10)", "Not required: net-zero lot reclass"),
    "Found - LOT format mismatch": ("Correct LOT in ERP - do not adjust", "Lot reclass (A-10)", "Not required: net-zero lot reclass"),
}
DEFAULT_ACTION = ("Manual Check", "Investigate first", "Any posting requires written approval")
USERS = [f"User {i}" for i in range(1, 7)]


def part():
    return f"P{random.randint(100000, 999999)}"


def lot():
    return f"L{random.randint(10000, 99999)}"


def discrepancies():
    rows = []
    for site, scopes in SITES.items():
        for scope in scopes:
            checked = random.randint(40, 700)
            disc = random.randint(0, max(1, checked // 8))
            posted = random.randint(0, disc)
            rows.append([site, f"{site} - {scope}" if site != scope else f"{site} - counted lines", checked, disc,
                         posted, disc - posted, random.randint(-5000, 5000), round(random.uniform(0, 20000), 2),
                         "Yes" if disc else "No"])
    return pd.DataFrame(rows, columns=["Site", "Scope", "LinesChecked", "DiscrepancyLines", "PostedResolved",
                                       "Outstanding", "NetUnits", "AbsErrorValueGBP", "InDiscrepancyCount"])


def transactions():
    rows = [[site, scope, t, random.randint(0, 40)] for site, scopes in SITES.items() for scope in scopes for t in TXN_TYPES]
    return pd.DataFrame(rows, columns=["Site", "Scope", "Transaction", "Count"])


def outstanding_actions():
    actions = [("Lot reclass (A-10)", "Not required (net zero)"), ("Root cause, then ADJ-QTY", "Covered by blanket approval"),
               ("Confirm samples, then misc issue", "Confirm they are samples first"), ("Chase 3PL for recount", "Chase 3PL"),
               ("Investigate", "No posting until the check is back")]
    rows = [[site, scope, a, random.randint(1, 15), appr] for site, scopes in SITES.items() for scope in scopes
            for a, appr in random.sample(actions, 3)]
    return pd.DataFrame(rows, columns=["Site", "Scope", "Action", "Lines", "Approval"])


def accuracy_history():
    rows = []
    for site, start_err in (("Site B", 0.35), ("Site A", 0.08), ("3PL", 0.12)):
        for i, year in enumerate(range(2019, 2027)):
            if year == 2025:
                continue  # no count that year, as in the real history
            stock = round(random.uniform(1.2e6, 2.5e6), 2)
            err = max(0.02, start_err * (0.75 ** i))
            txn = round(stock * err * random.uniform(0.15, 0.3), 2)
            lines = random.randint(200, 2400)
            err_lines = int(lines * err)
            rows.append([site, year, stock, txn, txn / stock, 1 - txn / stock, lines, err_lines, err_lines / lines])
    return pd.DataFrame(rows, columns=["Site", "Year", "StockValue", "TransactionValue", "ValuePct", "Accuracy",
                                       "LinesCounted", "ErrorLines", "ErrorLinePct"])


def site_lines(n=40):
    rows = []
    for i in range(n):
        cause = random.choice(CAUSES)
        action, txn, approval = ACTION_FOR_CAUSE.get(cause, DEFAULT_ACTION)
        sys_q = random.randint(0, 3000)
        cnt = max(0, sys_q + random.randint(-200, 200))
        rows.append([random.choice(["Input Sheet", "Found Stock"]), i + 2, f"{random.choice('ABCDEFGH')}{random.randint(1, 20):02d}A",
                     part(), lot(), sys_q, cnt, cnt - sys_q, round((cnt - sys_q) * random.uniform(0.05, 3), 2), cause,
                     action, txn, approval, None, "Outstanding", random.choice(["No", "No", "Yes"]), None, None])
    return pd.DataFrame(rows, columns=["Source tab", "Row", "Bin", "Part", "Lot", "System Qty", "Counted Qty", "Qty Diff",
                                       "Value Diff (£)", "Error / likely cause", "Workbook action", "Potential transaction",
                                       "Approval", "Note", "Status", "Product line flag", "In product recon", "Recon comparison"])


def pl_discrepancies(n=120):
    rows = []
    for _ in range(n):
        scope = random.choice(SITES["3PL"])
        erp_q = random.randint(0, 5000)
        pl_q = max(0, erp_q + random.choice([-erp_q, random.randint(-300, 300)]))
        diff = pl_q - erp_q
        status = random.choices(["Outstanding", "Posted", "Closed - no action", "Resolved"], [5, 2, 2, 1])[0]
        rows.append([scope, random.choice(["Inventory Transfer (STK-STK)", "Misc return (STK-UKN)", "Investigate - batch not at 3PL",
                                           "Lot reclass (A-10)"]),
                     random.choice(["Verify first", "Investigate first", "No action", "Ready now"]), part(), lot(),
                     lot() if pl_q else "not at 3PL", erp_q, pl_q, diff, abs(diff),
                     random.choice(["ACTIVE", "LARGE", "NOBATCH", "SMALL"]), "Requires written approval before posting",
                     status, f"U1{random.randint(15, 21)}092026" if status == "Posted" else None,
                     random.choice(["Not yet walked (3PL assumed)", "Not yet determined", "Both (timing)", "3PL"]),
                     round(diff * random.uniform(0.05, 1.5), 2)])
    return pd.DataFrame(rows, columns=["Scope", "Potential transaction", "Ready to post?", "Part", "Epicor lot", "DT batch",
                                       "Epicor qty", "DT qty", "Diff (DT - Epicor)", "Qty to post", "Check code", "Approval",
                                       "Status", "Posted by / date", "Correct side", "Value diff (£)"])


def pl_transactions(n=150):
    start = date(2026, 9, 15)
    rows = []
    for i in range(n):
        cat = random.choice(TXN_TYPES)
        rows.append([4100000 + i, start + timedelta(days=random.randint(0, 6)), random.choice(SITES["3PL"]), "3PL Warehouse",
                     random.choice(["BIN01", "BIN02", "ECOMM"]), part(), lot(), cat.split("(")[1].rstrip(")"),
                     random.randint(-500, 500), random.choice([None, "STAKE", "CORRE", "ISSERR", "RETERR"]),
                     random.choice(USERS), cat, f"REF{random.randint(1000, 9999)}", 0, 0])
    return pd.DataFrame(rows, columns=["Tran", "Tran Date", "Scope", "Warehouse", "Bin", "Part", "Lot", "Type", "Quantity",
                                       "Reason", "User", "Category", "Reference", "TO pack", "TO line"])


def posting_check(disc, txns):
    posted = disc[disc["Status"] == "Posted"]
    rows = []
    for _, r in posted.iterrows():
        found = random.random() > 0.3
        t = txns.sample(1).iloc[0] if found else None
        rows.append([r["Scope"], r["Potential transaction"], r["Part"], r["Epicor lot"], r["Qty to post"], "Posted",
                     r["Posted by / date"], "Found, check matches" if found else "Not found in export",
                     "Part Transactions export" if found else None,
                     t["Tran"] if found else None, t["Tran Date"] if found else None, t["Warehouse"] if found else None,
                     t["Type"] if found else None, t["Quantity"] if found else None, t["Reason"] if found else None])
    return pd.DataFrame(rows, columns=["Scope", "Potential transaction", "Part", "Lot", "Qty to post", "Status",
                                       "Posted by / date", "Check result", "Export", "Tran", "Tran date", "Warehouse",
                                       "Type", "Qty", "Reason"])


def unmatched(txns):
    df = txns.sample(5)[["Tran Date", "Warehouse", "Part", "Lot", "Type", "Quantity", "Reason", "User", "Tran"]]
    df.columns = ["Tran date", "Warehouse", "Part", "Lot", "Type", "Qty", "Reason", "User", "Tran"]
    return df


def main():
    disc = pl_discrepancies()
    txns = pl_transactions()
    sheets = {
        "Discrepancies": discrepancies(), "Transactions": transactions(), "OutstandingActions": outstanding_actions(),
        "AccuracyHistory": accuracy_history(), "WTY_Lines": site_lines(), "DT_Discrepancies": disc,
        "DT_Transactions": txns, "DT_PostingCheck": posting_check(disc, txns), "DT_UnmatchedPostings": unmatched(txns),
    }
    with pd.ExcelWriter(OUT) as xl:
        for name, df in sheets.items():
            df.to_excel(xl, sheet_name=name, index=False)
    print("Saved", OUT)


if __name__ == "__main__":
    main()
