# KPIs: what every tool measures and why

Every tool in this repo logs its own runs. The aim is that each one answers a business question from its first day in use, not only saves time: where stock accuracy is lost, whether the 3PL is meeting its side of the process, where master data costs the team time, and what the automation is actually worth.

## The standard every tool meets

1. **A decision it informs, and who makes it.** A KPI nobody acts on is not logged.
2. **One row per run, one row per exception.** Written automatically to one KPI folder, append-only, as CSV. A log that is open in Excel is never overwritten; the rows go to a dated side file.
3. **A baseline.** The manual time and error rate are measured before the switch (the pack-out tool has fields for this during its parallel run), so the saving is measured, not estimated.
4. **A value.** Time saved × hourly rate, errors avoided, and stock value at risk, as set out in the [business case](business-case.md).
5. **System-neutral fields.** Part, batch, owner, customer, reason and quantity, not ERP-specific codes, so the history carries over an ERP migration and supports data cleansing before it.

The shared helper is [`stocktools/kpi.py`](../stocktools/kpi.py). The folder is set by `kpi_folder` in `config.json`, per tool or once at the top level. If none is set, logs go to a `kpi_logs` folder next to the output, so nothing goes unlogged.

## Coverage

| Tool | Logs | KPIs | Decision it supports |
|---|---|---|---|
| [Movement request](../apps/movement_request_app.py) | `movement_request_runs_log.csv`, `movement_request_issues_log.csv` | Lines and quantity requested, cases written, lines short or with no stock, age of the 3PL stock report at the time of request | Where stock positioning at the 3PL lets orders down; whether the daily stock feed is fresh enough to rely on |
| [Warehouse orders](../apps/warehouse_order_app.py) | `warehouse_order_runs_log.csv`, `warehouse_order_issues_log.csv` | Lines OK, short, not found, invalid pack config, missing data | Replenishment to the UK sites; invalid pack configs as a master data measure, each one a line someone works out by hand |
| [Pack-out](../apps/packout_app.py) | `packout_runs_log.csv`, `packout_exceptions_log.csv` | Lines placed automatically by method (exact, combination, oldest first), how often the 3PL combined lines across orders, lines not placed, totals not tying, releases left off by next step (new order needed, stays open, unknown), manual minutes and match during the parallel run | Customer service risk from releases left off; evidence for asking the 3PL to stop combining lines across sales orders; resourcing of the role |
| [Stock reconciliation](../apps/stock_recon_app.py) | `recon_runs_log.csv`, `recon_reasons_log.csv`, `recon_area_log.csv` | Line and quantity accuracy %, gross and net variance qty and value, % of variance explained, variance by reason and by area | Where count effort goes; whether differences are a 3PL process or an internal one; provision and audit evidence |
| [Single SKU reconciliation](../apps/sku_recon_app.py) | `sku_investigations_log.csv` | Last date the systems agreed, unmatched movements, cause category, which system holds the error, quantity or attribute, days to resolve, approval form reference | 3PL performance on investigations; recurring parts and causes to fix at source; a correction trail auditors can follow |
| Optional AI summary | `ai_usage_log.csv` | Tokens and cost per summary | Whether the AI step is worth its cost against the template summary |
| [Stock provision re-run](../stocktools/provision.py) | None yet | Provision total and movement against the prior quarter (calculated, not logged) | Run quarterly through the finance procedure, so the result lives in that procedure's sign-off. A quarter-by-quarter log is the next addition |

## Monthly scorecard

The logs feed one page for supply chain leadership, in Excel pivots now and Power BI later. It reads from the same CSVs, so there's no extra collection.

| Measure | Source |
|---|---|
| Stock accuracy % by value and by line, trend | `recon_runs_log` |
| Value at risk (gross variance value) and % explained | `recon_runs_log` |
| Top 5 variance causes | `recon_reasons_log`, `sku_investigations_log` |
| 3PL performance: investigations where the 3PL held the error, days to resolve | `sku_investigations_log` |
| Order fulfilment risk: lines short at request, releases left off needing a new order | `movement_request_issues_log`, `warehouse_order_issues_log`, `packout_exceptions_log` |
| Master data: invalid pack configs, recurring problem parts | `warehouse_order_issues_log`, `sku_investigations_log` |
| Automation: runs per tool, hours saved against the measured baseline | All runs logs |

## Power Query

Each log loads into Excel or Power BI with the same Power Query pattern: read every file in the KPI folder whose name starts with the log's name (so side files are included), promote headers, type the columns with the en-GB locale. Run IDs start with a letter so Excel keeps them as text.
