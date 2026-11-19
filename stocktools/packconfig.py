"""Parse 3PL pack configuration strings.

The 3PL stock report carries a pack config per line, written as a run of
quantity + unit-of-measure pairs, outermost level last:

    5000E25C1P   5000 each across 25 cases on 1 pallet (200 each per case)
    2400B60C1P   2400 boxes across 60 cases on 1 pallet (40 boxes per case)
    100E1C       100 each in 1 case
    20S1C        20 shrink-wraps in 1 case

UOM codes: E each, B box, S shrink, C case, P pallet.

The string itself is the source of truth. Nothing here guesses a pack
size from order quantities.
"""
import math
import re

import pandas as pd

UOM_NAMES = {"E": "Each", "B": "Box", "C": "Case", "P": "Pallet", "S": "Shrink"}
_PAIR = re.compile(r"(\d+)([EBCPS])")


def parse_config(config_str):
    """Return a dict describing the pack structure, or None if unreadable."""
    if config_str is None or (not isinstance(config_str, str) and pd.isna(config_str)):
        return None
    raw = str(config_str).strip().upper()
    if raw in ("", "N/A", "NAN"):
        return None

    pairs = _PAIR.findall(raw)
    if not pairs:
        return None

    config = {
        "raw": raw,
        "levels": [{"qty": int(q), "uom": u, "uom_full": UOM_NAMES[u]} for q, u in pairs],
        "units_per_case": None,
        "cases_per_pallet": None,
        "units_per_pallet": None,
        "units_per_shrink": None,
    }

    unit_qty = shrink_qty = case_qty = pallet_qty = None
    for level in config["levels"]:
        if level["uom"] in ("E", "B"):
            unit_qty = level["qty"]
        elif level["uom"] == "S":
            shrink_qty = level["qty"]
        elif level["uom"] == "C":
            case_qty = level["qty"]
        elif level["uom"] == "P":
            pallet_qty = level["qty"]

    if unit_qty and case_qty:
        config["units_per_case"] = unit_qty / case_qty
        config["cases_per_pallet"] = case_qty
    elif shrink_qty and case_qty:
        config["units_per_shrink"] = shrink_qty
        config["units_per_case"] = shrink_qty / case_qty
    elif unit_qty:
        config["units_per_case"] = unit_qty
    elif case_qty and len(config["levels"]) == 1:
        config["units_per_case"] = case_qty

    if config["units_per_case"] and case_qty and pallet_qty:
        config["units_per_pallet"] = config["units_per_case"] * case_qty * pallet_qty

    return config


def cases_needed(qty, config):
    """Return (full_cases, leftover_units) for a quantity, or (None, None)."""
    if not config or not config.get("units_per_case"):
        return None, None
    per_case = config["units_per_case"]
    return int(qty // per_case), int(qty % per_case)


def cases_rounded_up(qty, config):
    """Cases to pick when a part case has to go as a whole case."""
    if not config or not config.get("units_per_case"):
        return None
    return math.ceil(qty / config["units_per_case"])
