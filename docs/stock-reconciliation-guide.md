# Stock reconciliation tool: user guide

The tool compares ERP stock with a second stock list and suggests a reason for every difference, in both directions. It then shows the net difference, which ties back to the two report totals.

It has two tabs:

- **3PL vs ERP**: the ERP stock report against the 3PL's full inventory report, by owner, part and lot.
- **Stock take vs ERP**: the ERP stock report against a physical count sheet, by part and lot, and by warehouse and bin if the count has them.

Each run saves one Excel workbook. Quantities are added as each system records them, with no unit conversion. The reasons are suggestions from rules, so every line still needs a review before it goes to anyone.

## Setting up

Run it from the repo:

```
pip install -r requirements.txt
python apps/stock_recon_app.py
```

To build a single `.exe` for colleagues:

```
pip install pyinstaller
pyinstaller --onefile --windowed --name Stock_Recon --paths . --add-data "config.json;." apps/stock_recon_app.py
```

The exe is in `dist/`. The settings are built into it, so rebuild after changing `config.json`. On macOS or Linux, use `:` instead of `;` in `--add-data`.

## 3PL vs ERP

Use this tab for 3PL audits, month-end checks or any time the 3PL and the ERP disagree.

| Input | Required | Where it comes from |
| --- | --- | --- |
| ERP stock report (.xlsx) | Yes | Full stock report, run at the same point as the 3PL's report |
| 3PL full inventory (.xlsx) | Yes | The 3PL's full inventory. The tool reads the "Full inventory" tab |
| ERP transactions (.txt) | No | Tab-delimited part transaction export covering the days before the report date |
| 3PL order summaries (.csv) | No | The 3PL's daily order summaries for the same days |
| Report date | Yes | The date both reports were run (dd/mm/yyyy) |

The optional files drive these reasons: Invoiced awaiting customer collection, In transit, Despatched by the 3PL, and Received in ERP. Without them, those lines fall to To Be Determined.

1. Open the **3PL vs ERP** tab.
2. Browse to the ERP report and the 3PL report.
3. Add the transactions export and the 3PL order summaries if you have them. You can select several CSVs at once.
4. Check the report date.
5. Choose where to save the results and click **Run**.
6. When it finishes, choose Yes to open the workbook.

The log at the bottom shows the totals for each side and the net. If anything fails, the error appears there.

## Stock take vs ERP

Use this tab after an onsite count, before any adjustments are posted.

| Input | Required | Notes |
| --- | --- | --- |
| ERP stock report (.xlsx) | Yes | Run after all transactions for the count day are posted |
| Warehouses counted | Yes | Ticked from the list the tool builds from the ERP report |
| Count sheet (.xlsx or .csv) | Yes | Any layout, as long as it has a Part column and a Qty column |
| Column choices | Part and Qty | Lot, Warehouse and Bin are optional |
| ERP transactions (.txt) | No | Same export as the 3PL tab, covering the count period |

1. Open the **Stock take vs ERP** tab.
2. Browse to the ERP report. The list fills with every site and warehouse in it.
3. Tick the warehouses that were counted. Hold Ctrl to tick several.
4. Browse to the count sheet. If it's an Excel file, pick the tab that holds the counts.
5. Check the column dropdowns. The tool guesses Part, Lot, Qty, Warehouse and Bin from the headings, so correct any it got wrong. Leave a dropdown blank if the count doesn't have that column.
6. Tick **Compare by bin** only if the count has a Bin column and you want bin-level differences.
7. Add the transactions export if you have it.
8. Choose where to save and click **Run**.

How the columns change the comparison:

- No Lot column: compared at part level.
- Warehouse column: compared per warehouse. Without it, all ticked warehouses are pooled.
- Compare by bin ticked: compared per bin. A part counted in the wrong bin shows as a bin difference on both sides.

## Reading the results

| Tab | What it shows |
| --- | --- |
| Summary | Quantity and value by reason for each direction, each direction's total, the net difference, the report totals and the files used |
| 3PL higher, or Count higher | Every line where the other list holds more than the ERP, with a reason and notes |
| ERP higher | Every line where the ERP holds more, with a reason and notes |
| Comparison | Every part and lot from both sides, including matches, and how each lot was matched |

**Check the net first.** The net difference on the Summary must equal the other list's total minus the ERP total, which is printed under it. If they don't match, don't use the results.

**Value** is quantity times the ERP unit cost from the stock report. Lines with no ERP cost are left blank.

**The Lot match column** on the Comparison tab says how each lot was paired:

- **exact:** same lot in both.
- **ignoring leading zeros:** for example, 045120 and 45120.
- **ignoring suffix:** for example, 20230711 and 20230711-99152.
- **part level:** the ERP part isn't lot-tracked, or the count has no lot.
- **part not in ERP:** the part has no ERP stock in the compared locations.

## Reason codes

Each line gets the first reason whose rule it meets, in the order below. To Be Determined means no rule matched.

**3PL higher than ERP**

| Reason | Rule |
| --- | --- |
| Product on hold (configurable name) | Description contains the hold keyword, or the part is on the hold list |
| Non-inventory | On the non-inventory list, or a part code with no digits |
| Invoiced awaiting customer collection | Shipped in the ERP transactions, or the 3PL has the exact difference allocated to orders |
| ERP shows it at an onsite warehouse | Not at the 3PL in the ERP, but the ERP holds it onsite |
| Part not held at the 3PL in ERP | Not at the 3PL in the ERP and not held anywhere else |
| Received at the 3PL not processed in ERP | The 3PL received it within the recent-days window before the report date |
| In transit | Transferred to another owner in the ERP, not yet moved to that owner at the 3PL |
| Booked to a different lot or area | The ERP holds exactly the same quantity of the part under another lot or owner |

**ERP higher than 3PL**

| Reason | Rule |
| --- | --- |
| Product on hold | As above |
| Non-inventory | As above |
| Booked to a different lot or area | The 3PL holds exactly the same quantity of the part under another lot or owner |
| Despatched by the 3PL, not yet moved in ERP | The 3PL's order summaries show it shipped, but the ERP still has it at the 3PL |
| Received in ERP, not yet at the 3PL | Received or manufactured into a 3PL warehouse in the ERP, not yet on the 3PL's report |
| In transit | Transferred in the ERP, not yet moved at the 3PL |

**Stock take: count higher than ERP**

| Reason | Rule |
| --- | --- |
| Non-inventory | As above |
| Booked to a different lot | The ERP holds the same quantity under another lot in the same place |
| In a different bin in ERP | The ERP holds the same quantity and lot in another bin |
| Issued or shipped in ERP, still on the shelf | Issued, shipped or moved out in the transactions |
| ERP shows it in another warehouse | The ERP has none here, but holds it in a warehouse that wasn't counted |
| Found stock (not in ERP anywhere) | The ERP has no stock of that part and lot anywhere |

**Stock take: ERP higher than count**

| Reason | Rule |
| --- | --- |
| Non-inventory | As above |
| Booked to a different lot | The count has the same quantity under another lot |
| In a different bin in ERP | The count has the same quantity and lot in another bin |
| Stranded in a receiving bin | The ERP holds it in the receiving bin. That isn't a physical location, so relocate it rather than adjusting |
| Received in ERP, not found in count | Received into the warehouse in the transactions |
| Not found in count | Nothing was counted for that part and lot |

"Booked to a different lot or area" matches on quantity alone, so treat it as a lead to check, not a finding.

## Settings

The `reconciliation` section of `config.json`:

| Setting | What it controls |
| --- | --- |
| `erp_warehouse_owner` | Which ERP warehouse names are at the 3PL, and which 3PL owner each maps to |
| `erp_site_warehouse_owner` | The same, where a warehouse name is only at the 3PL for one site |
| `erp_bin_owner` | Warehouses where only one bin is at the 3PL |
| `warehouse_code_owner`, `warehouse_code_bin_owner` | The same mapping by warehouse code, for the transactions export |
| `main_owner` | The 3PL owner that transfers go out from |
| `non_inventory_parts` | Pallets, edge protectors, cartons and other packaging |
| `hold_keyword`, `hold_parts`, `hold_reason` | Parts grouped under one hold reason, such as a new product awaiting release |
| `recent_days` | How many days before the report date count as recently received at the 3PL |
| `pending_receipt_bin` | The ERP's non-physical receiving bin |

## Before you share the results

1. Check the net difference on the Summary equals the report totals check under it.
2. Review every line under To Be Determined and Booked to a different lot or area.
3. Confirm a sample of the other reasons in the ERP before relying on them.
4. Add what you find to the Notes column. The tool doesn't remember findings between runs.

The tool never posts anything. Any correction it points to still needs root cause confirmed, physical evidence, and written approval on the adjustment approval form before posting.

## Limits

- Quantities in different units (eaches, boxes, cases) are added together as recorded. A shelf pack in the ERP against boxes at the 3PL shows as a difference on both sides.
- Stock in excluded bins (for example marketplace or despatch bins) is left out of the 3PL comparison.
- Without the transactions export and 3PL order summaries, timing differences fall to To Be Determined.
- Excel can strip leading zeros or turn part codes into dates. The tool repairs the common cases, so check the Lot match column if a line looks wrong.

## Troubleshooting

| Message or problem | What to do |
| --- | --- |
| Column not found | A report's headings have changed. Check the heading named in the message |
| The warehouse list is empty | The ERP report wasn't selected first, or it isn't the full stock report layout |
| Column dropdowns are empty | Pick the count sheet again, or choose the correct tab |
| Net doesn't match the totals check | Don't use the results. Check both reports were run at the same point and no rows were filtered out |
| Windows blocks the exe | Ask IT to allow it, or run the Python script instead |
