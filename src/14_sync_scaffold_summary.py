"""
Re-synchronise the ``scaffolds`` column of dataset_summary.csv with the
curated per-target tables, after fixing the counting rule in 02_curate.py.

Background: MurckoScaffoldSmiles() returns an empty string for acyclic
compounds.  The original 02_curate.py counted that empty string as one unique
scaffold (FAP 162, DPP4 1597), while 05_analysis.py read the same tables back
from CSV (where "" becomes NaN) and counted 161 / 1596.  Both scripts now use
``replace("", nan).nunique()``; this small script refreshes the already-built
summary without re-running the whole curation or touching any modelled value.

Usage: python src/14_sync_scaffold_summary.py
"""
from __future__ import annotations

from pathlib import Path

import numpy as np
import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
TAB = ROOT / "results" / "tables"
PROC = ROOT / "data" / "processed"
SUMMARY = TAB / "dataset_summary.csv"


def main() -> None:
    s = pd.read_csv(SUMMARY)
    changed = []
    for i, row in s.iterrows():
        name = str(row["target"]).lower()
        d = pd.read_csv(PROC / f"{name}_curated.csv")
        n = int(d.scaffold.replace("", np.nan).nunique())
        old = int(row["scaffolds"])
        if old != n:
            changed.append((row["target"], old, n))
        s.loc[i, "scaffolds"] = n
    s.to_csv(SUMMARY, index=False)
    print("scaffold counts (target, before, after):", changed or "already consistent")
    print(s[["target", "compounds", "scaffolds"]].to_string(index=False))
    print("->", SUMMARY)


if __name__ == "__main__":
    main()
