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
| ERP transaction history | Yes | Part transaction export (.xlsx or tab-delimited .txt), or a single-lot lot tracker export with no Part or Lot column. Include Sys Date and Sys Time where possible. If the file has several tabs, the first with Type and Quantity columns is used |
| Lots to check | Optional | Lots to include, separated by commas or spaces. Blank means every lot. For a lot tracker export, type the one lot (otherwise it is taken from the file name) |
| Default owner | Optional | Owner used for on-hand lines typed without one. Set `default_owner` in `config.json` to change the starting value |
| 3PL on hand | Recommended | Type one line per lot: `Lot Qty` (default owner) or `Lot Owner Qty`. Spaces, tabs or commas between values |
| ERP on hand | Recommended | Type `Lot Qty` or `Lot Warehouse Qty` |

Every owner found in either history is checked. For an owner you typed figures for, a lot you didn't list counts as zero. For an owner you typed nothing for, that system's movement total is used, and the tie check says "Not typed".

Both histories should go back to when the lot was first received. A movement dated before the other system's history starts is marked "may be outside the export" rather than treated as missing.

Try it on the samples: run `python sample_data/make_sku_samples.py`, then use part `009401`, `sample_data/sku_3pl_trace_sample.xlsx`, `sample_data/sku_erp_history_sample.txt`, and the on-hand figures in `sample_data/sku_onhand_sample.txt`.

## How it works

- **3PL balances** use Receipt, Shipment, Owner Update, Return and Receipt Reverse rows only. Allocate, Pick, Replenish, Vehicle Load, PreAdv Line, Stock Check and Repack net to zero at owner level.
- **ERP balances** use the transaction class (R, S, I, A) when the export has it. Otherwise they work from the transaction type, because some exports show every quantity as positive. Moves between bins within one owner net out, as do transfers between two warehouses of the same owner, and a posting cancelled by its reversal (for example a shipment of -5, +5, -5 on one order leaves one -5). Movements at ERP warehouses outside the 3PL are shown on the Timeline as context and don't affect the balances.
- **Owners:** ERP warehouses are mapped to 3PL owners using the `reconciliation` section of `config.json`.
- **Pairing:** first by order or pack reference, ignoring 3PL suffixes, so 100234BK1 and 100234BK3 both count as order 100234. Where the 3PL has more against a reference than the ERP, the extra movement stays unpaired. Then by the same lot, owner and quantity within `match_days` (default 5). Finally, a 3PL and an ERP movement for the same lot and owner, in the same direction and within `match_days`, but with different quantities, are paired as "likely same movement, qty differs" (for example the ERP moves 3,001 in and the 3PL receives 1). This is a lead, not a confirmed match.
- **Sequencing:** ERP movements are ordered on Sys Date and Sys Time. If the export only has a transaction date, the Summary says the order isn't verified.

## Reading the results

| Tab | What it shows |
| --- | --- |
| Summary | Per owner and lot: both on-hand figures, the difference, whether each history ties to its on hand, the last time both agreed, unmatched movements and the likely cause. Lines that differ are shaded. A By owner table totals each owner. If the balances only agree because one system moved the stock on its own, "Last time both agreed" says so and gives the earlier genuine agreement |
| Before naming any correction (on the Summary) | For each difference: does the stock physically exist and where, which system holds the error, and is it a quantity or only an attribute (lot or owner) |
| Timeline | Every movement from both systems in date order, with both running balances. Unmatched and qty-differs movements are shaded. ERP movements outside the 3PL are marked CONTEXT |
| Qty differs | Movements found in both systems with different quantities, and whether they were matched on a shared reference or on date and direction |
| 3PL movements, ERP movements | The rows used, with signs applied |

Fix a "No" in either tie check first: it means a trace or part of the history is missing.

Any quantity correction it points to still needs root cause confirmed, physical evidence, and written approval on the adjustment approval form before posting. A net-zero lot reclass does not need approval.

## KPI log

Every run adds rows to `sku_investigations_log.csv` in the KPI log folder (blank means the results folder): one row per lot and owner that differs, or one "No difference" row. Each row has the cause category, error side, quantity or attribute, and who ran it. Reviewed by, Date resolved, Approval form no and Notes are left blank for the reviewer, and give time to resolve and corrections raised.

## Limits

- Units aren't converted. If the 3PL counts eaches and the ERP counts boxes, every movement looks different.
- Owner Update rows are expected as two signed rows, minus from the old owner and plus to the new one.
- 3PL and ERP dates don't always line up. A posting more than `match_days` apart shows as unmatched.
