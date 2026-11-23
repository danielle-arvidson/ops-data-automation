"""Evaluate the optional AI summary step against hand-written expectations.

Each case in sku_summary_cases.json holds the facts the reconciliation code
produces for one difference, plus what a correct answer must contain
(expect_any: every group needs at least one match; email_expect_any: the same
for the 3PL email, which must ask about the specific movement) and must not contain.
On top of that, ai_summary.check_output() fails any answer that invents a
figure, uses ERP codes in the email, or suggests posting stock.

    python evals/run_sku_summary_eval.py              # template baseline, no API key needed
    python evals/run_sku_summary_eval.py --model claude-haiku-5-5   # needs ANTHROPIC_API_KEY

The baseline shows the harness working and gives a floor to beat: a model is
only worth its tokens if it passes at least as many cases with clearer text.
"""
import argparse
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from stocktools import ai_summary  # noqa: E402


class TemplateClient:
    """Deterministic baseline: fills a fixed sentence pattern from the facts."""
    model = "template-baseline"

    def complete(self, system, user):
        f = json.loads(user.split("\n", 1)[1])
        oh = f["on_hand"]
        summary = (f"For {f['part']} lot {f['lot']} ({f['owner']}), the 3PL shows {oh['3PL']:,} and the ERP shows "
                   f"{oh['ERP']:,}. Likely cause: {f['likely_cause']}. {f['which_system']}.")
        email = f"Please check lot {f['lot']} for item {f['part']} and confirm the quantity you physically hold."
        return json.dumps({"summary": summary, "email": email}), 0, 0


def score(case, row):
    text = (row["result"]["summary"] + " " + row["result"]["email"]).lower()
    email = row["result"]["email"].lower()
    misses = [g for g in case["expect_any"] if not any(x.lower() in text for x in g)]
    misses += [g for g in case.get("email_expect_any", []) if not any(x.lower() in email for x in g)]
    bad = [x for x in case["expect_none"] if x.lower() in text]
    problems = row["issues"] + [f"missing one of {g}" for g in misses] + [f"contains {x}" for x in bad]
    return problems


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--model", help="Anthropic model id; omit for the template baseline")
    ap.add_argument("--no-mask", action="store_true")
    a = ap.parse_args()
    cases = json.load(open(os.path.join(os.path.dirname(__file__), "sku_summary_cases.json")))
    client = ai_summary.AnthropicClient(a.model) if a.model else TemplateClient()
    rows = ai_summary.summarise([c["facts"] for c in cases], client, mask_ids=not a.no_mask, log=lambda m: None)
    passed = 0
    for c, r in zip(cases, rows):
        problems = score(c, r)
        passed += not problems
        print(f"{'PASS' if not problems else 'FAIL'}  {c['id']}" + ("" if not problems else "  -> " + "; ".join(problems)))
    tin, tout = sum(r["in"] for r in rows), sum(r["out"] for r in rows)
    print(f"\n{client.model}: {passed}/{len(cases)} passed; tokens in {tin}, out {tout}")
    return passed, len(cases)


if __name__ == "__main__":
    main()
