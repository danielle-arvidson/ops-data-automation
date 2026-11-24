# Ops data automation

## Summary

I'm a Stock and Data Integrity Manager at a regulated UK medical device manufacturer (ISO 13485, FDA 21 CFR 820). I'm responsible for inventory accuracy across three sites and an outsourced warehouse, and for supply chain master data in Epicor Kinetic. I'm also one of three leads on the company's move from Epicor to SAP.

This repo is the automation I built for that work:

- Three Python desktop tools that cut measured time per request from between 25 minutes and 1 hour 45 minutes down to 2 minutes or less, about 1,985 hours a year at minimum volumes.
- Power Automate flows that file 3PL reports automatically and give the master data team a single tracked request queue.
- A Power BI stock take reconciliation report, and an Excel approval form that flags write-offs already covered by finance provision.
- A [business case for every project](docs/business-case.md): why it exists, what it saves, and how each figure was worked out, about £43,700 a year in released capacity from the three Python tools alone.
- [Case studies](case-studies/) of analysis and control design: a cycle count coverage review, requirements extraction for the SAP migration, and rules for using AI assistants on controlled data.

Tools: Python (pandas, openpyxl, tkinter, pytest), Power Automate cloud and Desktop, Power BI, Excel, SharePoint, Microsoft Forms, Epicor Kinetic.

I build with AI assistants. The process design, requirements, testing against real data and corrections are mine; the AI drafts code and documents to that spec. The tools themselves don't need AI to run, to avoid per-use cost, keep company data off outside services, and avoid the energy use of a model where a script will do. The reasoning is in the [business case](docs/business-case.md#why-these-tools-run-without-ai).

## About this repo

The business runs on Epicor Kinetic, and finished goods are stored at a third-party logistics (3PL) warehouse that runs its own warehouse system. Most of the daily work sat in the gap between the two: exporting from the ERP, cross-checking against the 3PL's stock report, and retyping the result into the 3PL's spreadsheets.

The Python tools remove most of that retyping. They are written for non-technical colleagues: each one is a small desktop window, compiled to a single `.exe` with PyInstaller and run from a shared drive.

This repo is a sanitised copy. Company names, file paths, account codes and real part numbers have been replaced, and everything in `sample_data/` is synthetic.

## Python tools

| Tool | Problem | What it does |
|------|---------|--------------|
| [Movement request](apps/movement_request_app.py) | Every movement request and transfer order to the 3PL meant looking up stock line by line in the 3PL's web system, then typing part, batch, quantity and case count into their request template. | Reads the ERP order export, checks stock against the 3PL's daily stock CSV, fills the template, calculates cases from the 3PL's own pack config, and flags any line the 3PL can't cover. Getting the 3PL to email that daily CSV, in place of web lookups, is what made the tool possible. |
| [Warehouse orders](apps/warehouse_order_app.py) | Orders from the three UK sites were checked by hand against the 3PL stock report, with case counts worked out line by line. | Checks every line against the 3PL's daily stock report, adds case counts, fills in the required date (next working day, or the day after if it's past 2pm), and turns short lines red. |
| [Pack-out](apps/packout_app.py) | Packing lists came back from the 3PL with no order reference. One shipment can cover several movement requests, and the same part and batch can sit on several sales order lines and releases across them, so matching by hand was slow and error-prone. | Treats it as a quantity allocation, not a line match: within one part and batch the units are interchangeable, so it shares the packed quantity across the open releases so every quantity ties. Exact matches first, then exact combinations, then oldest request first, splitting lines where needed. A Check sheet compares packed and released totals per part and batch and flags anything that doesn't tie. A "Left off shipment" sheet lists releases not on this shipment: for a customer on rolling orders they stay open and ship later, for any other customer they need a new order, so they're caught on the day. Each shipment is for one customer, so it stops if the exports loaded belong to more than one. The packing list can be pasted straight from the 3PL's email. |
| [Stock reconciliation](apps/stock_recon_app.py) | Reconciling the ERP against the 3PL's audit report, or against an onsite stock take, meant lining up thousands of lines by part and lot in Excel and then investigating each difference one at a time. | Matches both lists by owner, part and lot, repairing codes Excel has damaged. Suggests a reason for every difference in both directions and shows the net, which ties back to the report totals. One window covers 3PL reports and onsite count sheets, and every run logs accuracy KPIs for Power BI. [User guide](docs/stock-reconciliation-guide.md). |
| [Single SKU reconciliation](apps/sku_recon_app.py) | Day-to-day queries on one part meant walking the 3PL's batch trace and the ERP's transaction history side by side to find where they split. | Rebuilds both balances per lot and owner, pairs the movements by order reference and then by quantity and date, and shows the last date they agreed and the movements that don't pair up. Answers the three questions needed before any correction is named. [User guide](docs/sku-reconciliation-guide.md). |

The tools are used by the supply chain team.

**Time per request, before and after** (deployed versions, measured April 2026):

| Tool | Manual | With tool | Volume |
|---|---|---|---|
| Movement requests and transfer orders | 43.5 min | 1 min | 28+ a week |
| Warehouse orders | 25 min | 30 sec | 10+ a week |
| Pack-out | 1 hr 45 min | 2 min | 11+ a week |

Across 46.4 working weeks at the minimum volumes, that's about 1,985 hours a year, about £43,700 of released capacity. See the [business case](docs/business-case.md) for the basis and the other projects.

### Optional AI step: investigation summaries

[`stocktools/ai_summary.py`](stocktools/ai_summary.py) adds one AI step to the single SKU reconciliation: for each line that differs, a model drafts a plain-English explanation for the stock team and a short email to the 3PL. It's off by default (`sku_reconciliation.ai.enabled` in `config.json`).

The step is deliberately small. Matching, balances and causes stay deterministic; the model only writes words about facts the code has already worked out, because those figures end up in audited stock records and AI adds cost and risk to arithmetic.

- **Only lines that differ** are sent, never matching lines.
- **Identifiers are masked** (part, lot and references become P1, L1, R1) before anything is sent, and restored afterwards.
- **Every figure in the answer is checked** against the facts it was given. An invented or altered number is flagged on the AI summary tab, as is an ERP code in the 3PL email or any suggestion to post or adjust stock.
- **Token use is logged** per run to `ai_usage_log.csv`, with cost if prices are configured, so cost is measured rather than guessed.

[`evals/run_sku_summary_eval.py`](evals/run_sku_summary_eval.py) scores the step against hand-written expectations for five kinds of difference (a move the 3PL never received, the same receipt booked with different quantities, shipments in one system only, and an owner move that is attribute only). A deterministic template baseline is included as the floor a model has to beat: it scores 0 of 5, because its 3PL email is generic rather than asking about the specific movement. Run with `--model <model id>` and an `ANTHROPIC_API_KEY` to score a model.

### Stock provision re-run

[`stocktools/provision.py`](stocktools/provision.py) is the quarter-end stock provision calculation I run for Finance. Every stock line is re-classified against the criteria as at quarter-end, rather than rolling the previous balance forward:

| Order | Criterion | Rule (rolling from quarter-end) |
|---|---|---|
| 1st | Site provision | Line sits at a site flagged for full provision, for example one being closed |
| 2nd | Expiry | Expires within 12 months |
| 3rd | Over 2 years no use | Received over 24 months ago, not used in the last 24 months |
| 4th | Slow-moving | Received over 12 months ago, not used in the last 12 months |

Qualifying lines are then excluded if they're in a disposal bin, a quarantine, shop-floor or maintenance warehouse, already on an open stock disposal form, or have any planned use. Planned use takes out the whole line, not just the planned quantity, which is Finance's standard basis. The output gives the total, a breakdown by criterion and by exclusion rule, and movement against the prior re-run. The breakdowns are checked to tie back to the total before anything is reported.

In production this runs as an Excel formula over the ERP's aged stock export. The Python version shows the same logic in testable form.

Shared logic is in `stocktools/`, kept separate from the GUI so it can be tested.

### Things the code handles that the manual process kept getting wrong

- **Leading zeros.** The ERP's CSV export drops them, and the 3PL keeps them, so `0361` and `361` are the same part. Matching ignores leading zeros, but refuses to guess if two different 3PL codes collapse to the same number. The output uses the 3PL's version of the code, so they recognise it.
- **Stale stock data.** The 3PL report is a snapshot. Every shortage carries the report's timestamp, and a cell note says to check with the 3PL before chasing, because stock may have been booked in since.
- **Pack configurations.** Case counts come only from the 3PL's pack config string (for example `5000E25C1P`), never from dividing order quantities. If there's no readable config, the tool says so rather than guessing.
- **Case-sold SKUs.** Some products are ordered in cases already. A configurable pattern identifies them so they aren't divided again.

## Power BI

- [Year-end stock take reconciliation](powerbi/stock-take/): tracks every count discrepancy across sites and the 3PL from count to correction, checks posted corrections against the ERP transaction export, and shows stock accuracy by site over time. Includes a synthetic dataset with the report's schema.

## Excel controls

[Adjustment approval form](templates/make_approval_form.py) is the control document for every manual stock correction: one form per correction, signed by requester, administrator and approver in order. Each line checks itself against the current quarter's finance provision snapshot, kept on a hidden sheet, and shows "In Finance Provision?" before anyone signs. A write-off on a line that's already provided for would count the same loss twice, so the check flags it before the form reaches the approver. The template ([`adjustment_approval_form.xlsx`](templates/adjustment_approval_form.xlsx)) is built by the script so the formulas can be read, and its provision rows are synthetic.

## Power Automate

These work in Outlook, SharePoint and the shared drive, and the data request queue also picks up notification emails from the quality management system. Neither one connects to the ERP.

- [3PL report filing](power-automate/3pl-report-filing/): a Power Automate Desktop flow that files 15 daily and weekly 3PL report types from a shared inbox into dated folders.
- [Data request intake](power-automate/data-request-intake/): Microsoft Form, SharePoint List and four cloud flows (intake, quality-system follow-ups, assignment and completion emails, lead-time and on-time tracking) that replaced email and Teams requests to the master data team with a single tracked queue.

## Case studies

[Write-ups](case-studies/) of work that can't be published as code: a cycle count coverage review, requirements extraction for the SAP migration, a packing spec review checklist, an audit gap analysis, rules for AI use on controlled data, and training and onboarding.

## Running it

```
pip install -r requirements.txt
python sample_data/make_samples.py
python sample_data/make_recon_samples.py
python sample_data/make_sku_samples.py
python apps/movement_request_app.py
python apps/stock_recon_app.py
python apps/sku_recon_app.py
python evals/run_sku_summary_eval.py
python -m pytest
```

Settings (folders, owner account codes, case-SKU pattern, and the reconciliation warehouse mapping) are in `config.example.json`. Copy it to `config.json` to change them. The defaults point at the sample data.
