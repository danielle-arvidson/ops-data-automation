"""Optional AI step for the single SKU reconciliation: drafts a plain-English
explanation of one difference and a short email to the 3PL.

Design rules (why this step is small and fenced in):
- Off by default. Matching, balances and causes stay deterministic; the model
  only writes words about facts the code has already worked out.
- Only lines that differ are sent, never matching lines.
- Part numbers, lots and references are masked before sending (P1, L1, R1...)
  and put back afterwards, so no identifiers leave the machine.
- Every number in the model's answer is checked against the facts it was given.
  An invented or altered number is flagged, not silently accepted.
- The email is checked for ERP transaction codes and for any instruction to post
  or adjust stock: the model drafts, a person decides.
- Token use is logged per run, so cost is a measured number.
"""
import csv
import datetime
import json
import os
import re

SYSTEM = (
    "You explain stock differences between a third-party warehouse (the 3PL) and the company's ERP. "
    "Use only the facts given. Do not add, round or change any number, date or reference. "
    "If the facts do not establish which system is wrong, say so. Never recommend posting, adjusting "
    "or writing off stock: a person decides that after the root cause is confirmed. "
    "Reply with JSON only: {\"summary\": \"...\", \"email\": \"...\"}. "
    "summary: at most 4 sentences for the stock team. "
    "email: a short plain-language request to the 3PL asking them to check what the facts need checked. "
    "In the email do not use ERP transaction codes, reason codes, warehouse codes or bin codes."
)

ERP_CODE = re.compile(r"\b[A-Z]{3}-[A-Z]{3}\b")            # STK-CUS, ADJ-QTY, PUR-STK ...
POSTING = re.compile(r"\b(post|posting|adjust|adjustment|write[- ]?off|write it off)\b", re.I)
NUMBER = re.compile(r"(?<![\w/])[-+]?\d[\d,]*(?:\.\d+)?")


def build_case(part, summary_row, check_row, unmatched, qty_notes):
    """Facts for one owner and lot that differ. Plain dicts and strings only."""
    owner, lot, tpl, erp, diff, ties_tpl, ties_erp, agreed, n_unmatched, cause, category = summary_row
    fmt = lambda d: d.strftime("%d/%m/%Y") if isinstance(d, datetime.datetime) else (d or "")
    return {
        "part": part, "owner": owner, "lot": str(lot),
        "on_hand": {"3PL": tpl, "ERP": erp, "3PL minus ERP": diff},
        "history_ties_to_on_hand": {"3PL": ties_tpl, "ERP": ties_erp},
        "last_agreed": fmt(agreed) if isinstance(agreed, datetime.datetime) else str(agreed),
        "cause_category": category, "likely_cause": cause,
        "which_system": check_row[4] if check_row else "Not established",
        "quantity_or_attribute": check_row[5] if check_row else "",
        "unmatched_movements": [{"system": t[3], "date": fmt(t[2]), "type": t[4], "qty": t[5], "ref": t[6]}
                                for t in unmatched],
        "qty_differs": [{"ref": n["ref"], "3PL": n["dt"], "ERP": n["erp"], "date": fmt(n.get("date"))}
                        for n in qty_notes],
    }


def mask(case):
    """Replace the part, lot and every reference with tokens. Returns (masked case, token -> real)."""
    mapping, rev = {}, {}

    def tok(value, prefix):
        v = str(value)
        if not v:
            return v
        if v not in rev:
            rev[v] = f"{prefix}{sum(1 for k in mapping if k.startswith(prefix)) + 1}"
            mapping[rev[v]] = v
        return rev[v]

    c = json.loads(json.dumps(case, default=str))
    c["part"], c["lot"] = tok(c["part"], "P"), tok(c["lot"], "L")
    for m in c["unmatched_movements"] + c["qty_differs"]:
        m["ref"] = tok(m["ref"], "R")
    for real, t in sorted(rev.items(), key=lambda x: -len(x[0])):   # also inside the cause text
        c["likely_cause"] = c["likely_cause"].replace(real, t)
    return c, mapping


def unmask(text, mapping):
    for t, real in sorted(mapping.items(), key=lambda x: -len(x[0])):
        text = re.sub(rf"\b{re.escape(t)}\b", real, text)
    return text


def _numbers(text):
    out = set()
    for m in NUMBER.findall(str(text)):
        v = m.replace(",", "").lstrip("+")
        try:
            f = abs(float(v))
        except ValueError:
            continue
        out.add(f"{f:g}")
    return out


def check_output(result, case):
    """Deterministic checks on the model's answer. Returns a list of issues (empty = passed)."""
    issues = []
    allowed = _numbers(json.dumps(case, default=str))
    for field in ("summary", "email"):
        text = result.get(field, "")
        if not text:
            issues.append(f"{field} missing")
            continue
        made_up = sorted(_numbers(text) - allowed)
        if made_up:
            issues.append(f"{field} has figures not in the facts: {', '.join(made_up)}")
        for sentence in re.split(r"(?<=[.!?])\s+", text):
            if POSTING.search(sentence) and not re.search(r"\b(not|no|never|until|before|approval|approve|decides?)\b",
                                                          sentence, re.I):
                issues.append(f"{field} suggests posting or adjusting stock: \"{sentence.strip()[:80]}\"")
    if ERP_CODE.search(result.get("email", "")):
        issues.append("email contains ERP transaction codes")
    return issues


def parse(text):
    try:
        start, end = text.index("{"), text.rindex("}") + 1
        data = json.loads(text[start:end])
        return {"summary": str(data.get("summary", "")), "email": str(data.get("email", ""))}
    except (ValueError, json.JSONDecodeError):
        return {"summary": text.strip(), "email": ""}


class AnthropicClient:
    """Thin wrapper so tests can pass a fake client with the same complete() method."""

    def __init__(self, model, max_tokens=600):
        import anthropic   # optional dependency: pip install anthropic
        self.client, self.model, self.max_tokens = anthropic.Anthropic(), model, max_tokens

    def complete(self, system, user):
        r = self.client.messages.create(model=self.model, max_tokens=self.max_tokens, system=system,
                                        messages=[{"role": "user", "content": user}])
        text = "".join(b.text for b in r.content if getattr(b, "type", "") == "text")
        return text, r.usage.input_tokens, r.usage.output_tokens


def summarise(cases, client, mask_ids=True, log=print):
    """Run the model on each case. Returns rows: case, result, issues, tokens in, tokens out."""
    rows = []
    for case in cases:
        sent, mapping = mask(case) if mask_ids else (case, {})
        user = "Facts:\n" + json.dumps(sent, indent=1, default=str)
        text, tin, tout = client.complete(SYSTEM, user)
        result = parse(text)
        issues = check_output(result, sent)
        result = {k: unmask(v, mapping) for k, v in result.items()}
        rows.append({"case": case, "result": result, "issues": issues, "in": tin, "out": tout})
        log(f"  AI summary for {case['owner']} lot {case['lot']}: {tin} in / {tout} out tokens"
            + (f"; CHECK: {'; '.join(issues)}" if issues else ""))
    return rows


def log_usage(folder, run_id, part, model, rows, price_in=None, price_out=None, log=print):
    """Append token use (and cost, if prices per million tokens are configured) to ai_usage_log.csv."""
    tin, tout = sum(r["in"] for r in rows), sum(r["out"] for r in rows)
    cost = round(tin / 1e6 * price_in + tout / 1e6 * price_out, 6) if price_in is not None and price_out is not None else ""
    path = os.path.join(folder, "ai_usage_log.csv")
    try:
        os.makedirs(folder, exist_ok=True)
        new = not os.path.exists(path)
        with open(path, "a", newline="", encoding="utf-8") as f:
            w = csv.writer(f)
            if new:
                w.writerow(["Run ID", "Run date", "Part", "Model", "Lines summarised", "Tokens in", "Tokens out",
                            "Cost (USD)", "Lines with check issues"])
            w.writerow([run_id, f"{datetime.datetime.now():%d/%m/%Y %H:%M}", part, model, len(rows), tin, tout, cost,
                        sum(1 for r in rows if r["issues"])])
    except OSError:
        log(f"AI usage log folder {folder} could not be reached. Usage not logged.")
