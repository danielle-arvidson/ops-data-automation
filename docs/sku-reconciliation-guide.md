# Single SKU reconciliation tool: user guide

The tool finds where the 3PL and the ERP stopped agreeing on one part, and which movements caused it.

For one part, it rebuilds the 3PL's and the ERP's balance for each lot and owner from their own histories. It then pairs the 3PL's movements with the ERP's and lists any that don't pair up. These are usually the cause.

Use it when a 3PL stock check, order or query shows a different quantity from the ERP, or an order can't be shipped in the ERP because the lot isn't there. The result is a lead to investigate, not a correction. The tool never changes anything in the ERP.

## Running it

```
python apps/sku_recon_app.py
```

| Input | Required | Notes |
| --- | --- | --- |
| Part number | Yes | |
| 3PL inventory trace(s) | Yes | .xls, .xlsx or .csv. Several files allowed, for example one per batch. The tool finds the header row itself |
| ERP transaction history | Yes | Part transaction export (.xlsx or tab-delimited .txt). Include Sys Date and Sys Time where possible |
| 3PL on hand | Recommended | Paste as Lot, Owner, Qty. Anything not listed counts as zero |
| ERP on hand | Recommended | Paste as Lot, Warehouse, Qty |

Both histories should go back to when the lot was first received. A movement dated before the other system's history starts is marked "may be outside the export" rather than treated as missing.

Try it on the samples: run `python sample_data/make_sku_samples.py`, then use part `009401`, `sample_data/sku_3pl_trace_sample.xlsx`, `sample_data/sku_erp_history_sample.txt`, and the on-hand figures in `sample_data/sku_onhand_sample.txt`.

## How it works

- **3PL balances** use Receipt, Shipment, Owner Update, Return and Receipt Reverse rows only. Allocate, Pick, Replenish, Vehicle Load, PreAdv Line, Stock Check and Repack net to zero at owner level.
- **ERP balances** use the transaction class (R, S, I, A) when the export has it. Otherwise they work from the transaction type, because some exports show every quantity as positive. Moves between bins within one owner net out.
- **Owners:** ERP warehouses are mapped to 3PL owners using the `reconciliation` section of `config.json`.
- **Pairing:** first by order or pack reference, ignoring 3PL suffixes, so 100234BK1 and 100234BK3 both count as order 100234. Where the 3PL has more against a reference than the ERP, the extra movement stays unpaired. Then by the same lot, owner and quantity within `match_days` (default 5).
- **Sequencing:** ERP movements are ordered on Sys Date and Sys Time. If the export only has a transaction date, the Summary says the order isn't verified.

## Reading the results

| Tab | What it shows |
| --- | --- |
| Summary | Per owner and lot: both on-hand figures, the difference, whether each history ties to its on hand, the last time both agreed, unmatched movements and the likely cause. Lines that differ are shaded |
| Before naming any correction (on the Summary) | For each difference: does the stock physically exist and where, which system holds the error, and is it a quantity or only an attribute (lot or owner) |
| Timeline | Every movement from both systems in date order, with both running balances. Unmatched movements are shaded |
| Qty differs | References found in both systems with different quantities |
| 3PL movements, ERP movements | The rows used, with signs applied |

Fix a "No" in either tie check first: it means a trace or part of the history is missing.

Any quantity correction it points to still needs root cause confirmed, physical evidence, and written approval on the adjustment approval form before posting. A net-zero lot reclass does not need approval.

## KPI log

Every run adds rows to `sku_investigations_log.csv` in the KPI log folder (blank means the results folder): one row per lot and owner that differs, or one "No difference" row. Each row has the cause category, error side, quantity or attribute, and who ran it. Reviewed by, Date resolved, Approval form no and Notes are left blank for the reviewer, and give time to resolve and corrections raised.

## Limits

- Units aren't converted. If the 3PL counts eaches and the ERP counts boxes, every movement looks different.
- Owner Update rows are expected as two signed rows, minus from the old owner and plus to the new one.
- 3PL and ERP dates don't always line up. A posting more than `match_days` apart shows as unmatched.
