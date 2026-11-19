# Data request intake (Microsoft Forms, SharePoint, Power Automate)

Requests for master data work (new parts, part changes, lot fixes, reports) came in by email, Teams and phone. Some request types had their own logs, but there was no single queue, so there was no full view of workload and no way to show turnaround. This replaces those channels with one form and one list.

## How it fits together

```
Microsoft Form ------> intake flow -------+
                                          v
QMS notification ----> QMS flow -----> SharePoint List (the queue)
email                                     |
                                          +--> assignment flow --> email to analyst
                                          +--> done flow --------> email to requester,
                                                                   closed date, lead time,
                                                                   on-time flag
```

| Flow | Trigger | Definition |
|------|---------|------------|
| Intake | New form response | [`intake-flow.json`](intake-flow.json) |
| QMS follow-up | Notification email from the quality system | [`qms-followup-flow.json`](qms-followup-flow.json) |
| Assignment email | List item changed | [`assignment-email-flow.json`](assignment-email-flow.json) |
| Done email | List item changed | [`done-email-flow.json`](done-email-flow.json) |

The definitions are the flows as exported from Power Automate. Site, list, form and folder IDs, and links, are replaced with placeholders such as `<SITE_URL>`.

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

## Intake flow, step by step

1. **Trigger:** a new form response.
2. **Get the response details.**
3. **Create the list item.** Each answer maps to a column. Status is set to `New` and `Received via` to `Form`, so form requests can be told apart from QMS follow-ups in the same list. The requester is set from the form's signed-in responder, not typed in.
4. **Attach evidence.** If anything was uploaded, the flow loops through the files, reads each one from the folder where Forms stores uploads, and attaches it to the list item, so the evidence stays with the request.
5. **Email the requester** with a reference number, their summary, the business impact, a link to a "My requests" view, and a link to the item itself. The email asks them to add questions as comments on the request rather than by email or Teams, so the whole history stays in one place.

Details that stop it breaking:

- A blank "Needed by" date is sent as `null`. SharePoint rejects an empty string in a date column, so without this the whole item fails.
- The confirmation email runs whether the attachment step succeeds or fails. A file that won't copy doesn't stop the requester getting their reference number.

## QMS follow-up flow

Spec changes in the quality management system often need matching ERP changes, and the QMS sends a notification email for each one. This flow turns those emails into requests in the same queue as everything else.

1. **Trigger:** an email from the QMS notification address whose subject contains a document reference (`DOC-`).
2. **Pull out the reference** (for example `DOC-00316`) from the subject.
3. **Check for an open item** with that reference. The QMS sends several emails per document as it moves through approval, so without this check one change would become several requests.
4. **If none is open, create one:** titled `<reference> ERP review`, with the email subject as the details, Received via `QMS`. If the email is an assignment notice, the due date is parsed from the end of the subject (`d MMM yyyy`, UK format) and set as both the needed-by and target date.

## Assignment and done flows

Both trigger on any change to a list item. Firing on every edit would spam people, so each one first asks SharePoint which columns actually changed since the last run, and only acts if its column did.

**Assignment:** if `Assigned to` changed and isn't blank, email the analyst with the request, business impact, target date and a link.

**Done:** if `Status` changed to `Done`:
1. Email the requester with the resolution notes and a link, asking them to comment on the request if anything isn't right.
2. Stamp the item with:
   - **Closed date:** now.
   - **Lead time (days):** whole days from created to closed, worked out from date ticks.
   - **On time:** true if closed on or before the target date. Items with no target date are marked false.

The update in step 2 is itself a change to the item, but it doesn't touch `Status`, so the flow doesn't trigger itself again.

## What it gives the team

Every request now has a timestamp, a type, an impact and an owner, so workload, turnaround and request mix by type can be reported on rather than estimated.
