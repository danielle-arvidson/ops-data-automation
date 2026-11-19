# Ops data automation

Automation I built as Stock and Data Integrity Manager at a UK medical device manufacturer. There are two groups: Python tools for stock work between the ERP and the 3PL warehouse, and Power Automate flows for report filing and data request handling.

The business runs on Epicor Kinetic, and finished goods are stored at a third-party logistics (3PL) warehouse that runs its own warehouse system. Most of the daily work sat in the gap between the two: exporting from the ERP, cross-checking against the 3PL's stock report, and retyping the result into the 3PL's spreadsheets.

The Python tools remove most of that retyping. They are written for non-technical colleagues: each one is a small desktop window, compiled to a single `.exe` with PyInstaller and run from a shared drive.

This repo is a sanitised copy. Company names, file paths, account codes and real part numbers have been replaced, and everything in `sample_data/` is synthetic.

## Python tools

| Tool | Problem | What it does |
|------|---------|--------------|
| [Movement request](apps/movement_request_app.py) | Customer orders picked from the 3PL had to be copied into the 3PL's request template by hand, with case counts worked out on a calculator. | Reads the ERP order export, fills the template, calculates cases from the 3PL's own pack config, and flags any line the 3PL can't cover. |
| [Warehouse orders](apps/warehouse_order_app.py) | Warehouse orders went to the 3PL without a stock check, and short lines came back days later. | Checks every line against the 3PL's daily stock report before sending, adds case counts, fills in the next working day as the required date, and turns short lines red. |
| [Pack-out](apps/packout_app.py) | Packing lists came back from the 3PL with no sales order reference, so each line had to be matched to an order by hand before it could be shipped in the ERP. | Matches part and batch against ERP shipping exports and fills SO, Line and Release, oldest order first. |

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
