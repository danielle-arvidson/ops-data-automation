# Business case and savings

Why each tool exists, what it saves, and how the figure was worked out. Every figure is labelled **measured**, **estimated** or **not yet measured**.

## Basis

- **Time** comes from the cost-benefit analysis (April 2026): manual times observed and averaged before deployment, tool times measured after it, at the minimum weekly volumes, over 46.4 working weeks a year.
- **Money** is hours saved at a loaded labour cost of £18.28 an hour: a £30,000 supply chain coordinator salary plus employer National Insurance and pension, over a 37.5-hour week. Coordinators are the people whose time the tools save, so their rate is the one used. It is capacity released for other work, not cash saved.
- **Development cost** was £0 beyond my own time. Everything uses open-source Python or the Microsoft 365 licences the business already had.

## Why these tools run without AI

I build with AI assistants, but I decided that none of these tools should need an AI model to run. Each one is ordinary code that runs the same way on a laptop with no internet connection. The reasons:

- **Cost.** A model charges for every request it handles. These tools run dozens of times a week, and matching, arithmetic and allocation need no model at all, so paying per run would add a cost every week for no benefit.
- **Data security.** The data is stock levels, batch numbers, customers and costs from a regulated manufacturer. Keeping the tools offline means none of it is sent to an outside service, and there is no data-sharing agreement or approval to manage.
- **Environmental impact.** Running a large model uses far more energy than running a short script. Using one for work a few lines of code can do adds energy use with nothing gained.
- **Same answer every time.** These figures end up in shipments, stock records and audit evidence. Code gives the same result from the same input and can be tested and explained line by line; a model's output can vary and has to be checked.

Where AI does help, with unstructured text that rules can't handle, it's added as a small optional step that is off by default. The one example in this repo, [investigation summaries](../README.md#optional-ai-step-investigation-summaries), is built around the same concerns: only lines that differ are sent, part numbers and references are masked before sending, every figure in the answer is checked, and token use is logged so the cost is measured.

## Summary

| Project | What it saves | Annual value | Status |
|---|---|---|---|
| Movement request tool | 42.5 min per request, about 920 hours a year | about £16,800 | Measured |
| Warehouse order tool | 24.5 min per order, about 190 hours a year | about £3,500 | Measured |
| Pack-out tool | 103 min per pack-out, about 876 hours a year | about £16,000 | Measured |
| **Three Python tools together** | **about 1,985 hours a year (97 to 98% less time per request)** | **about £36,300**, plus an estimated £5,000 of avoided errors | Measured (errors estimated) |
| Stock reconciliation tool | 3 to 3.5 days per full reconciliation | about £800 to £1,900 at 2 to 4 runs a year | Estimated |
| Single SKU reconciliation | Investigation time per query | Logged per run, not yet totalled | Not yet measured |
| Stock provision re-run | Accuracy of the quarter-end provision | Not a time saving | Not measured |
| Adjustment approval form | Write-offs counted twice | Not a time saving | Not measured |
| Stock take Power BI report | Chasing count corrections by hand | Not yet measured | Not yet measured |
| 3PL report filing flow | Daily manual filing of 15 report types | Not yet measured | Not yet measured |
| Data request intake | Requests lost or untracked across email and chat | Not yet measured | Not yet measured |

## Python tools

### Movement request tool

**Problem.** Each movement request or transfer order to the 3PL meant looking up stock line by line in the 3PL's web system, then typing part, batch, quantity and case count into their template.

**Why it matters.** It was the highest-volume manual task on the team (28 or more a week), and every retyped line was a chance to send the wrong batch or quantity to a regulated customer.

**Result (measured).** 43.5 minutes to 1 minute per request. At 1,299 requests a year that is 920 hours, about £16,800.

**AI at runtime.** None. See [why these tools run without AI](#why-these-tools-run-without-ai).

### Warehouse order tool

**Problem.** Orders from three UK sites were checked by hand against the 3PL stock report, with case counts worked out line by line.

**Why it matters.** Short lines found late hold up the sites; case counts worked out by dividing quantities were often wrong.

**Result (measured).** 25 minutes to 30 seconds per order. At 464 orders a year that is 190 hours, about £3,500.

**AI at runtime.** None. See [why these tools run without AI](#why-these-tools-run-without-ai).

### Pack-out tool

**Problem.** Packing lists came back from the 3PL with no order reference. One shipment can cover several movement requests, with the same part and batch on several order lines and releases.

**Why it matters.** Shipping against the wrong release mis-states what each customer has received. A release left off a shipment also has a cost: for a customer on rolling orders it simply ships later, but for any other customer it has to be ordered again, so it has to be caught on the day. The tool lists these on a "Left off shipment" sheet with the next step for each.

**Result (measured).** 1 hour 45 minutes to 2 minutes per pack-out. At 510 pack-outs a year that is 876 hours, about £16,000.

**Errors avoided (estimated).** Across the three tools, an estimated £5,000 a year: transcription errors causing wrong shipments (£500 to £2,000 an incident), leading zeros lost from part and batch codes, wrong case counts, and missed lines.

**AI at runtime.** None. See [why these tools run without AI](#why-these-tools-run-without-ai).

### Stock reconciliation tool

**Problem.** Reconciling the ERP against the 3PL's audit report, or against an onsite stock take, meant lining up thousands of lines in Excel and investigating each difference.

**Why it matters.** The reconciliation supports audit, year-end and the finance stock valuation. A difference with no explained reason is a finding.

**Result (estimated).** A manual reconciliation takes an estimated 4.5 to 6 days; the tool saves 3 to 3.5 of them a run (22.5 to 26 hours). At 2 to 4 runs a year that is 45 to 105 hours, about £800 to £1,900. Each run also logs accuracy KPIs, so the trend is measured from now on.

**AI at runtime.** None. See [why these tools run without AI](#why-these-tools-run-without-ai).

### Single SKU reconciliation tool

**Problem.** Day-to-day queries on one part meant reading the 3PL batch trace and the ERP history side by side to find where they split.

**Why it matters.** Coordinators can work through a query themselves, and a correction is only raised once the cause and the system at fault are shown.

**Result (not yet measured).** Every run logs the investigation, its cause and, once closed, the date resolved. Time to resolve will come from that log.

**AI at runtime.** None in the reconciliation itself. An optional AI summary step can be switched on; it is off by default.

## Controls and reporting

### Stock provision re-run

**Why it matters.** Rolling the previous quarter's provision forward carries old decisions into the new balance. Re-classifying every line against the criteria at quarter-end gives Finance a provision that can be traced line by line, with breakdowns that tie to the total.

**Savings.** Not a time saving; the value is an accurate, auditable figure.

**AI at runtime.** None. See [why these tools run without AI](#why-these-tools-run-without-ai).

### Adjustment approval form

**Why it matters.** A write-off on stock already covered by the finance provision counts the same loss twice. The form flags it before the approver signs.

**Savings.** Not measured. Each prevented double count is worth the value of the line.

**AI at runtime.** None. See [why these tools run without AI](#why-these-tools-run-without-ai).

### Stock take Power BI report

**Why it matters.** Tracks every count discrepancy from count to correction and checks posted corrections against the ERP, so nothing is corrected twice or missed.

**Savings.** Not yet measured.

**AI at runtime.** None. See [why these tools run without AI](#why-these-tools-run-without-ai).

## Power Automate

### 3PL report filing

**Why it matters.** 15 recurring report types arrive in a shared inbox and had to be saved by hand into dated folders. Unfiled reports mean decisions made on stale data.

**Savings.** Not yet measured. Time per day saved is the figure to add.

**AI at runtime.** None. See [why these tools run without AI](#why-these-tools-run-without-ai).

### Data request intake

**Why it matters.** Requests to the master data team came by email, chat and phone, with no record of what was asked, when, or whether it was done. One queue gives lead time and on-time figures for the first time.

**Savings.** Not yet measured. Requests a month and lead time will come from the list once it has run for a while.

**AI at runtime.** None. See [why these tools run without AI](#why-these-tools-run-without-ai).
