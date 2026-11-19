# Ops data automation

Automation I built as Stock and Data Integrity Manager at a UK medical device manufacturer. There are two groups: Python tools for stock work between the ERP and the 3PL warehouse, and Power Automate flows for report filing and data request handling.

The business runs on Epicor Kinetic, and finished goods are stored at a third-party logistics (3PL) warehouse that runs its own warehouse system. Most of the daily work sat in the gap between the two: exporting from the ERP, cross-checking against the 3PL's stock report, and retyping the result into the 3PL's spreadsheets.

The Python tools remove most of that retyping. They are written for non-technical colleagues: each one is a small desktop window, compiled to a single `.exe` with PyInstaller and run from a shared drive.

This repo is a sanitised copy. Company names, file paths, account codes and real part numbers have been replaced, and everything in `sample_data/` is synthetic.

## Python tools

| Tool | Problem | What it does |
|------|---------|--------------|
| [Movement request](apps/movement_request_app.py) | Every movement request and transfer order to the 3PL meant looking up stock line by line in the 3PL's web system, then typing part, batch, quantity and case count into their request template. | Reads the ERP order export, checks stock against the 3PL's daily stock CSV, fills the template, calculates cases from the 3PL's own pack config, and flags any line the 3PL can't cover. Getting the 3PL to email that daily CSV, in place of web lookups, is what made the tool possible. |
| [Warehouse orders](apps/warehouse_order_app.py) | Orders from the three UK sites were checked by hand against the 3PL stock report, with case counts worked out line by line. | Checks every line against the 3PL's daily stock report, adds case counts, fills in the required date (next working day, or the day after if it's past 2pm), and turns short lines red. |
| [Pack-out](apps/packout_app.py) | Packing lists came back from the 3PL with no order reference, so each line had to be matched by hand against several movement requests before it could be shipped in the ERP. | Matches part and batch against the movement request exports for that pack-out and fills SO, Line and Release, oldest request first. |

The tools are used by the supply chain team.

### Stock provision re-run

[`stocktools/provision.py`](stocktools/provision.py) is the quarter-end stock provision calculation I run for Finance. Every stock line is re-classified against the criteria as at quarter-end, rather than rolling the previous balance forward:

| Order | Criterion | Rule (rolling from quarter-end) |
|---|---|---|
| 1st | Site provision | Line sits at a site flagged for full provision, for example one being closed |
| 2nd | Expiry | Expires within 12 months |
| 3rd | Over 2 years no use | Received over 24 months ago, not used in the last 24 months |
| 4th | Slow-moving | Received over 12 months ago, not used in the last 12 months |

Qualifying lines are then excluded if they're in a disposal bin, a quarantine, shop-floor or maintenance warehouse, already on an open stock disposal form, or have any planned use. Planned use takes out the whole line, not just the planned quantity, because a partly used line isn't obsolete. The output gives the total, a breakdown by criterion and by exclusion rule, and movement against the prior re-run. The breakdowns are checked to tie back to the total before anything is reported.

In production this runs as an Excel formula over the ERP's aged stock export. The Python version shows the same logic in testable form.

Shared logic is in `stocktools/`, kept separate from the GUI so it can be tested.

### Things the code handles that the manual process kept getting wrong

- **Leading zeros.** The ERP's CSV export drops them, and the 3PL keeps them, so `0361` and `361` are the same part. Matching ignores leading zeros, but refuses to guess if two different 3PL codes collapse to the same number. The output uses the 3PL's version of the code, so they recognise it.
- **Stale stock data.** The 3PL report is a snapshot. Every shortage carries the report's timestamp, and a cell note says to check with the 3PL before chasing, because stock may have been booked in since.
- **Pack configurations.** Case counts come only from the 3PL's pack config string (for example `5000E25C1P`), never from dividing order quantities. If there's no readable config, the tool says so rather than guessing.
- **Case-sold SKUs.** Some products are ordered in cases already. A configurable pattern identifies them so they aren't divided again.

## Power Automate

These work in Outlook, SharePoint and the shared drive, and the data request queue also picks up notification emails from the quality management system. Neither one connects to the ERP.

- [3PL report filing](power-automate/3pl-report-filing/): a Power Automate Desktop flow that files a dozen daily and weekly 3PL report types from a shared inbox into dated folders.
- [Data request intake](power-automate/data-request-intake/): Microsoft Form, SharePoint List and four cloud flows (intake, quality-system follow-ups, assignment and completion emails, lead-time and on-time tracking) that replaced email and Teams requests to the master data team with a single tracked queue.

## Running it

```
pip install -r requirements.txt
python sample_data/make_samples.py
python apps/movement_request_app.py
python -m pytest
```

Settings (folders, owner account codes, case-SKU pattern) are in `config.example.json`. Copy it to `config.json` to change them. The defaults point at the sample data.
