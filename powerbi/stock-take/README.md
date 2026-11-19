# Year-end stock take reconciliation (Power BI)

The annual stock take covers two manufacturing sites and several accounts at the third-party warehouse, ahead of the external audit. This report tracks the reconciliation after the count: what was checked, what differs, what has been corrected, and what is still open and why.

## What the data model holds

| Table | What it answers |
|---|---|
| `Discrepancies` | Per site and account: lines checked, lines with a discrepancy, how many are posted or resolved, how many are outstanding, net units and absolute error value. |
| `Transactions` | The mix of ERP transactions posted per site during the reconciliation (adjustments, transfers, misc issues and returns, transfer orders). |
| `OutstandingActions` | Open actions grouped by type, with the approval route each one needs before anything is posted. |
| `AccuracyHistory` | Stock accuracy by site and year: value accuracy (adjustment value against stock value) and line accuracy (error lines against lines counted). |
| `WTY_Lines` | Line-level site count results: system against counted quantity, likely cause, and the correcting action. |
| `DT_Discrepancies` | Line-level ERP against 3PL differences, with the proposed transaction, whether it's ready to post, and which side is believed to be right. |
| `DT_Transactions` | The ERP transaction export for the reconciliation window. |
| `DT_PostingCheck` | Every correction marked as posted, checked against the ERP transaction export, to confirm it actually went through. |
| `DT_UnmatchedPostings` | ERP postings in the window that don't match any planned correction. |

The last two tables close the loop in both directions: nothing marked posted is missing from the system, and nothing posted in the system is unexplained.

## Files

- `make_synthetic_data.py` builds `stock_take_synthetic.xlsx`. It has the same sheets and columns as the real source workbook (two product-line columns are renamed), with made-up sites, parts and values.
- `theme.json` is the report theme.
