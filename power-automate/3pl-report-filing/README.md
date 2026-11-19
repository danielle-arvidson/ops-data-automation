# 3PL report filing (Power Automate Desktop)

The 3PL emails around a dozen report types to a shared inbox: daily stock, receipts, orders, shipments, returns and weekly adjustments. Each one was being saved by hand into the right folder for its year and month. This flow does it in one run.

It only touches Outlook and a shared network drive. It doesn't read from or write to the ERP.

## What it does

1. Pulls every email from the 3PL's sender address and saves the attachments to a staging folder.
2. For each attachment, decides the destination from the file name. Subjects repeat across report types, so the file name is the only reliable key.
3. Creates `<report folder>\<yyyy>\<MM. MMMM>` if it doesn't exist and moves the file there.
4. Renames shipment summaries with a date stamp (for example `Shipment Summary 19NOV26`), because the 3PL sends them with the same name every day.
5. Copies the daily inventory report into a second folder, which is where the Python tools look for the newest stock file.
6. Moves the processed emails to a `Filed` subfolder so the next run doesn't pick them up again.

Unrecognised attachments stay in staging, so a new report type isn't lost.

## Notes

- Built in Power Automate Desktop (the free version), run by hand.
- The `SWITCH` takes the first match, so `Daily Warehouse Order Summary` has to come before the broader `Daily Order Summary`.
- `IfFileExists: DoNothing` means a re-run never overwrites a file that has already been filed.
- PAD's Outlook actions drive the classic desktop client, so the flow broke when the company moved to the new Outlook. Moving it to a cloud flow with the Office 365 Outlook connector would remove that dependency.

`flow.txt` is the flow as PAD exports it when you copy the actions. Paste it into an empty PAD flow to load it. The mailbox, sender, paths and report names are placeholders.
