"""
Step 1 -- Data acquisition from ChEMBL.

Targets (all human):

    FAP   CHEMBL4683  Prolyl endopeptidase FAP (Q12884)   -- theranostic target
    PREP  CHEMBL3202  Prolyl endopeptidase     (P48147)   -- closest paralogue
    DPP4  CHEMBL284   Dipeptidyl peptidase 4   (P27487)   -- anti-target / toxicity

PREP and DPP4 are the two prolyl-endopeptidase-family off-targets that dominate
the selectivity problem of the cyano-pyrrolidine FAP warhead. They are pulled as
(i) selectivity labels for screening and (ii) an auxiliary task for transfer
learning.

Output: data/raw/chembl_<target>.json
"""
from __future__ import annotations

import json
import sys
import time
import urllib.parse
import urllib.request
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
RAW = ROOT / "data" / "raw"
RAW.mkdir(parents=True, exist_ok=True)

TARGETS = {
    "FAP": "CHEMBL4683",
    "PREP": "CHEMBL3202",
    "DPP4": "CHEMBL284",
}
ACTIVITY_TYPES = ("Ki", "Kd", "IC50")
BASE = "https://www.ebi.ac.uk/chembl/api/data/activity.json"
PAGE = 1000


def fetch_page(url: str, retries: int = 4) -> dict:
    for attempt in range(retries):
        try:
            with urllib.request.urlopen(url, timeout=120) as fh:
                return json.loads(fh.read())
        except Exception as exc:
            if attempt == retries - 1:
                raise
            print(f"  retry {attempt + 1} ({exc})", file=sys.stderr)
            time.sleep(3 * (attempt + 1))
    raise RuntimeError("unreachable")


def fetch_activities(target_id: str, std_type: str) -> list[dict]:
    records: list[dict] = []
    offset = 0
    while True:
        query = urllib.parse.urlencode(
            {
                "target_chembl_id": target_id,
                "standard_type": std_type,
                "limit": PAGE,
                "offset": offset,
            }
        )
        payload = fetch_page(f"{BASE}?{query}")
        activities = payload.get("activities", [])
        records.extend(activities)
        total = payload["page_meta"]["total_count"]
        offset += PAGE
        print(f"    {std_type}: {len(records)}/{total}")
        if offset >= total or not activities:
            break
    return records


def main() -> None:
    for name, chembl_id in TARGETS.items():
        print(f"[{name}] {chembl_id}")
        all_records = []
        for std_type in ACTIVITY_TYPES:
            all_records.extend(fetch_activities(chembl_id, std_type))
        out = RAW / f"chembl_{name}.json"
        out.write_text(json.dumps(all_records, indent=1), encoding="utf-8")
        print(f"  -> {out}  ({len(all_records)} records)")


if __name__ == "__main__":
    main()
