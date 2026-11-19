"""Part and lot number helpers.

The ERP export and the 3PL warehouse system disagree on leading zeros:
the ERP CSV export (and Excel, if anyone opens it) drops them, so part
0361 or lot 097709 arrive as 361 and 97709. Every comparison between the
two systems goes through these helpers.
"""


def normalize_code(code):
    """Strip whitespace and leading zeros so codes from both systems compare equal."""
    code = str(code).strip()
    return code.lstrip("0") or code


def codes_match(a, b):
    """True if two part or lot numbers are the same once leading zeros are ignored."""
    a, b = str(a).strip(), str(b).strip()
    return a == b or normalize_code(a) == normalize_code(b)
