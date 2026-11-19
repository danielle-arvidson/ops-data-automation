# Data request intake (Microsoft Forms, SharePoint, Power Automate)

Requests for master data work (new parts, part changes, lot fixes, reports) used to arrive by email, Teams message, phone call and in passing. Nothing was logged, so there was no queue, no view of workload and no way to show turnaround. This replaces all of those with one form and one list.

## How it fits together

```
Microsoft Form ──> Power Automate ──> SharePoint List (the queue)
                                         │
QMS follow-ups ──> Power Automate ───────┤
                                         ├──> assignment email to the analyst
                                         └──> completion email to the requester
```

- **Intake flow:** each form response becomes a list item, with any uploaded evidence attached.
- **QMS follow-up flow:** follow-up work raised from the quality system lands in the same list, so there is one queue rather than two.
- **Assignment email:** when an item is assigned, the analyst is told what it is and where to find it.
- **Completion email:** when an item is marked done, the requester is told.

Requesters have view-only access to the list, so they can check status without asking.

## The form

| # | Field | Type | Required | Guidance shown |
|---|-------|------|----------|----------------|
| 1 | Request summary | Text | Yes | One line, e.g. "Part 12345 wrong lot on receipt" |
| 2 | Request type | Choice | Yes | New part / Part change / Lot or stock data fix / Report or query / QMS follow-up / Other |
| 3 | QMS reference | Text | No | Document number, if it relates to a controlled document |
| 4 | Part numbers | Text | No | Separate multiple parts with commas |
| 5 | Lot or batch | Text | No | |
| 6 | Details | Long text | Yes | What's wrong or what's needed, and anything you've already checked |
| 7 | Business impact | Choice | Yes | Customer order blocked / Production blocked / Compliance or audit / No immediate impact |
| 8 | Needed by | Date | No | |
| 9 | Evidence | File upload | No | Screenshots, report exports, emails saved as PDF (up to 10 files, 10 MB each) |

Two design choices did most of the work:

- **Business impact instead of priority.** Everyone marks their own request high priority. Asking what is actually blocked gives a triage order people can't inflate.
- **"Anything you've already checked".** This cuts down on back-and-forth before work can start.

## What it gives the team

Every request now has a timestamp, a type, an impact and an owner, so workload, turnaround and request mix by type can be reported on rather than estimated.
