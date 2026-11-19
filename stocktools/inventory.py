"""Load and query the 3PL daily inventory report."""
import os
from datetime import datetime

import pandas as pd

from .codes import normalize_code

NUMERIC_COLUMNS = ("Free", "Total", "Allocated")


class Inventory:
    """Stock on hand at the 3PL for one owner account.

    The report is a CSV with a title on row 1 and headers on row 2.
    Everything is read as text so part and batch codes keep their
    leading zeros.
    """

    def __init__(self, csv_path, owner_id):
        df = pd.read_csv(csv_path, skiprows=1, dtype=str)
        df["Stock Code"] = df["Stock Code"].astype(str).str.strip()
        df["Batch_id"] = df["Batch_id"].astype(str).str.strip()
        for col in NUMERIC_COLUMNS:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

        self.df = df[df["Owner ID"] == owner_id].copy()
        self.df["_part_norm"] = self.df["Stock Code"].map(normalize_code)
        self.df["_batch_norm"] = self.df["Batch_id"].map(normalize_code)
        self.owner_id = owner_id
        self.path = csv_path
        # File time, not run time: the report can be hours old by the time it's used.
        self.report_time = datetime.fromtimestamp(os.path.getmtime(csv_path))

        self._pack_exact, self._pack_norm = {}, {}
        if "Pack Config" in self.df.columns:
            for code, cfg in zip(self.df["Stock Code"], self.df["Pack Config"]):
                if isinstance(cfg, str) and cfg.strip():
                    self._pack_exact.setdefault(code, cfg.strip())
                    self._pack_norm.setdefault(normalize_code(code), cfg.strip())

    def __len__(self):
        return len(self.df)

    @property
    def pack_config_count(self):
        return len(self._pack_exact)

    def pack_config(self, part):
        part = str(part).strip()
        return self._pack_exact.get(part) or self._pack_norm.get(normalize_code(part))

    def lines(self, part, lot=None):
        """Inventory rows for a part (and lot, if given).

        Exact match first. If that fails, match ignoring leading zeros, but
        refuse to guess when two different 3PL codes collapse to the same
        number.
        """
        part = str(part).strip()
        rows = self.df[self.df["Stock Code"] == part]
        if rows.empty:
            rows = self.df[self.df["_part_norm"] == normalize_code(part)]
            if rows["Stock Code"].nunique() > 1:
                return rows.iloc[0:0]
        if lot and str(lot).strip().upper() != "N/A":
            rows = rows[rows["_batch_norm"] == normalize_code(lot)]
        return rows

    def free_stock(self, part, lot=None):
        rows = self.lines(part, lot)
        return float(rows["Free"].sum()), rows

    @staticmethod
    def find_newest(folder, name_contains="Inventory"):
        """Newest matching CSV anywhere under folder, or None."""
        newest, newest_time = None, -1
        for dirpath, _, files in os.walk(folder):
            for name in files:
                if name.lower().endswith(".csv") and name_contains in name:
                    path = os.path.join(dirpath, name)
                    mtime = os.path.getmtime(path)
                    if mtime > newest_time:
                        newest, newest_time = path, mtime
        return newest
