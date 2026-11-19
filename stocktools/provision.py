"""Quarter-end stock provision re-run.

Every stock line is re-classified against the provision criteria as at
quarter-end. It is a full re-evaluation, not a roll-forward of the
previous balance, so a line that has started moving again drops out on
its own.

Criteria, all rolling from the quarter-end date, checked in this order
(first match wins):

    4  Site flagged for full provision (for example a site being closed)
    1  Expires within 12 months
    3  Received over 24 months ago and not used in the last 24 months
    2  Received over 12 months ago and not used in the last 12 months (SLOB)

Lines that qualify are then excluded if they sit in a disposal bin, a
quarantine, shop-floor or maintenance warehouse, are already on an open
stock disposal form, or have any planned use. Planned use excludes the
whole line, not just the planned quantity.

Input is an aged stock export with one row per part/lot/bin.
"""
import pandas as pd

CRITERIA = {
    "4": "4 - Site provision",
    "1": "1 - Expires within 12m",
    "3": "3 - Over 2yr no use",
    "2": "2 - SLOB",
}
NONE = "None"
EXCLUDED_WAREHOUSE_TYPES = ("Quarantine", "Shop Floor", "Maintenance")


def criteria_dates(quarter_end):
    q = pd.Timestamp(quarter_end)
    return {
        "expiry_by": q + pd.DateOffset(months=12),
        "slob_before": q - pd.DateOffset(months=12),
        "no_use_2yr_before": q - pd.DateOffset(months=24),
    }


def _not_used_since(row, cutoff):
    return pd.isna(row["last_used"]) or row["last_used"] <= cutoff


def classify_line(row, dates, provision_sites=()):
    if row["site"] in provision_sites:
        return CRITERIA["4"]
    if pd.notna(row["expires"]) and row["expires"] <= dates["expiry_by"]:
        return CRITERIA["1"]
    received = row["first_receipt"]
    if pd.notna(received) and received <= dates["no_use_2yr_before"] and _not_used_since(row, dates["no_use_2yr_before"]):
        return CRITERIA["3"]
    if pd.notna(received) and received <= dates["slob_before"] and _not_used_since(row, dates["slob_before"]):
        return CRITERIA["2"]
    return NONE


def exclusion_reason(row, disposal_bins, open_disposal_lots):
    """First exclusion rule that applies, or None. Order is fixed so the breakdown adds up."""
    if row["bin"] in disposal_bins:
        return "Disposal bin"
    if row["warehouse_type"] in EXCLUDED_WAREHOUSE_TYPES:
        return f"{row['warehouse_type']} warehouse"
    if (row["part"], row["lot"]) in open_disposal_lots:
        return "Open disposal form"
    if row["planned_use"] > 0:
        return "Planned use"
    return None


def rerun(stock, quarter_end, provision_sites=(), disposal_bins=("DISPOSAL",), open_disposal_lots=()):
    """Classify every line and return (line-level frame, summary dict)."""
    df = stock.copy()
    for col in ("expires", "first_receipt", "last_used"):
        df[col] = pd.to_datetime(df[col], dayfirst=True, errors="coerce")
    df["planned_use"] = pd.to_numeric(df["planned_use"], errors="coerce").fillna(0)
    df["value"] = pd.to_numeric(df["value"], errors="coerce").fillna(0)

    dates = criteria_dates(quarter_end)
    open_lots = set(open_disposal_lots)
    df["criterion"] = df.apply(classify_line, axis=1, dates=dates, provision_sites=set(provision_sites))
    df["excluded_by"] = df.apply(exclusion_reason, axis=1, disposal_bins=set(disposal_bins),
                                 open_disposal_lots=open_lots)
    df.loc[df["criterion"] == NONE, "excluded_by"] = None  # only qualifying lines can be excluded
    df["in_provision"] = (df["criterion"] != NONE) & df["excluded_by"].isna()

    included = df[df["in_provision"]]
    excluded = df[(df["criterion"] != NONE) & df["excluded_by"].notna()]
    summary = {
        "quarter_end": pd.Timestamp(quarter_end).date().isoformat(),
        "criteria_dates": {k: v.date().isoformat() for k, v in dates.items()},
        "total_value": round(included["value"].sum(), 2),
        "lines_included": len(included),
        "lines_excluded": len(excluded),
        "by_criterion": included.groupby("criterion")["value"].agg(["count", "sum"]).round(2).to_dict("index"),
        "by_exclusion": excluded.groupby("excluded_by")["value"].agg(["count", "sum"]).round(2).to_dict("index"),
    }
    # The breakdowns must tie back to the total, or something has been double counted.
    assert round(sum(v["sum"] for v in summary["by_criterion"].values()), 2) == summary["total_value"]
    return df, summary


def movement(current_total, prior_total):
    """Change against the prior re-run, for the sign-off note."""
    change = round(current_total - prior_total, 2)
    pct = round(change / prior_total * 100, 1) if prior_total else None
    return {"prior": prior_total, "current": current_total, "change": change, "change_pct": pct}
