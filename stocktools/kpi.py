"""Shared KPI logging.

Every tool appends to CSV logs in one folder, one row per run and one row per
exception, so performance can be trended in Excel or Power BI from the first
run. Logs are append-only: a log that is open in Excel is never overwritten,
the rows go to a dated side file instead.
"""
import csv
import datetime
import getpass
import os

from . import config as _config


def folder(section, override=None, output_path=None):
    """Where the KPI logs go: the folder passed in, else the tool's kpi_folder in config, else the
    top-level kpi_folder, else a kpi_logs folder next to the output so nothing goes unlogged."""
    cfg = _config.load()
    chosen = override or cfg.get(section, {}).get("kpi_folder") or cfg.get("kpi_folder")
    if chosen:
        return chosen
    base = os.path.dirname(os.path.abspath(output_path)) if output_path else os.getcwd()
    return os.path.join(base, "kpi_logs")


def new_run_id():
    """Run ID with a letter first so Excel doesn't turn it into 2.02611E+16."""
    return "R" + datetime.datetime.now().strftime("%Y%m%d-%H%M%S-%f")[:-3]


def stamp():
    return datetime.datetime.now().strftime("%d/%m/%Y %H:%M"), getpass.getuser()


def append_csv(folder, name, header, rows, log=print):
    """Append rows to a CSV log, header first if the file is new."""
    if not rows:
        return
    try:
        os.makedirs(folder, exist_ok=True)
    except OSError:
        log(f"KPI log folder {folder} could not be reached. Nothing logged for {name}.")
        return
    path = os.path.join(folder, name)
    try:
        new = not os.path.exists(path)
        with open(path, "a", newline="", encoding="utf-8") as f:
            w = csv.writer(f)
            if new:
                w.writerow(header)
            w.writerows(rows)
    except PermissionError:
        alt = os.path.join(folder, name.replace(".csv", f"_{datetime.datetime.now():%Y%m%d_%H%M%S}.csv"))
        with open(alt, "w", newline="", encoding="utf-8") as f:
            w = csv.writer(f)
            w.writerow(header)
            w.writerows(rows)
        log(f"KPI log {name} is open or locked. Rows saved to {os.path.basename(alt)}: add them to the main log later.")
