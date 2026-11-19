# Case studies

Analysis, control design and documentation work from the same role. None of these are code in this repo. The source material holds real part numbers, people and internal paths, so each one is written up here instead, with the status stated plainly: delivered, or proposed only.

## Cycle count coverage review

Status: analysis delivered; fix proposed, not implemented.

The cycle count programme looked healthy on paper, so I tested what it was actually counting. I pulled the stock bin export, the cycle count parts-selected report, the ABC code settings and the adjustment history from the ERP and compared them.

Two gaps came out of it:

- Parts with no recent usage were given a classification the cycle count schedule never selects. Those parts held about 16% of the stock value in the export analysed, and none of it was ever going to be counted.
- Counting ran at a fraction of its configured schedule. From January to June 2026, 82 of 378 expected counts happened (22%), and 5 of 144 for the highest-value A parts (3.5%).

The proposed fix was to give every no-usage part a countable code and add a monthly step to keep it that way, with a root-cause diagram behind it. The method uses standard ERP reports and settings, so it carries over to other sites.

## Requirements extraction for an ERP migration

Status: step 1 (extraction) complete.

I'm one of three leads for the company's move from Epicor Kinetic to SAP. Before anyone writes data quality rules, the existing requirements have to be pulled out of the documents that already govern part creation.

I built a field-by-field register from 30 work instructions and procedures and 5 workbooks. Each row gives the requirement, the lifecycle state it applies to (manufactured, purchased, development), its source document, and whether a query can test it. The register also logs 15 places where documents contradict each other and lists what isn't documented at all. It deliberately proposes no rules yet, so the requirements can be agreed before the rules are written.

Alongside it I proposed a design for running data quality rules as a single ERP query with a defect count, controls at the point of part creation, and a part creation log generated from the system change log instead of kept by hand. That design wasn't taken forward because it was judged too broad for the programme it was put to.

## Packing spec review checklist

Status: delivered and in use.

Every new or changed packing specification has to be checked against the bill of materials in the ERP before it's released. I covered this review during a colleague's leave and found the same errors recurring: header revisions that didn't match the system, and component lines that didn't match the spec.

I wrote a 9-step review procedure from the errors found on real specs. It's now the reference used for later reviews. A small Python checker that reads a spec document and compares it against pasted ERP grids came out of the same work; it caught 3 of 3 in-scope errors on 4 test specs and is still to be added here with synthetic data.

## Movement request process: audit gap analysis

Status: current-state analysis complete; future state not yet designed.

After the acquisition, the stock processes faced public-company audit standards. I mapped the current state of one movement request process end to end and identified four control gaps: a completeness tracker kept by hand, reference numbers assigned by hand, a shared warehouse system login, and no second check before submission. The Python tools in this repo remove some of the manual steps but don't close these gaps on their own.

## Rules for using AI assistants on controlled data

Status: delivered and in use.

I use AI assistants for checks on stock, bills of materials and packing specs, where a wrong number can end up in an audited record. I went back through past sessions, found six confirmed errors and one inferred one (for example, comparing quantities before establishing the unit basis, and treating one document as the full population), and turned each into a rule the assistant has to follow. The rules cover stating sources and arithmetic for every figure, putting assumptions above the output, never deriving case sizes by dividing quantities, and stopping when an input is missing. They now govern my day-to-day AI use, and I wrote a reusable prompt so the same rules can be set up for other work.

## Training and onboarding

Status: delivered, except where noted.

I learned the supply chain coordinator role from trainee notes, then turned those notes into the team's working documentation: SOPs for receipting, pack-out, ordering stock to sites, material transfers and movement requests, plus a single flow chart of how goods move inbound, between sites, outbound and through non-conformance.

When a new coordinator started in April 2026, I built and ran the onboarding programme: a training plan, weekly pre-reading packs and a first-day system orientation. For the live sessions I used a paired format, with a trainer script beside a learner desk reference, using real inbox emails as the cases. When the learner kept mixing up factory orders and transfer orders, I wrote a three-question decision framework for choosing the right stock movement, and cut that session down to fit.

I also rewrote training for the ERP data analyst role (part creation, method changes, packing specs, monthly reports, housekeeping) and wrote a standalone guide to method changes in Epicor's Engineering Workbench. Separately, I designed a six-document coordinator curriculum with a glossary, quick reference cards and a scenario question bank; that one was proposed and not issued.

Because training material can look like procedure to an auditor, I raised the risk with Quality myself. My materials now carry an "information only, not a controlled document" label, so they can't be mistaken for controlled work instructions.
