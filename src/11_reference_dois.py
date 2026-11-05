"""
Resolve a DOI for every reference in the AIDD8122 report and re-verify the
bibliographic details against Crossref.

The reference strings live in ``09_build_report.py`` (``refs = [...]`` inside
``build_body``).  This script parses that literal with ``ast`` so the two can
never drift apart, queries Crossref with the full citation as a bibliographic
query, and keeps the best match only when the returned title is contained in
the original citation (or is a close fuzzy match).  The report builder then
reads ``results/tables/reference_dois.csv`` and appends the verified DOI.

Usage:  python src/11_reference_dois.py
Output: results/tables/reference_dois.csv
"""
from __future__ import annotations

import ast
import difflib
import json
import re
import time
import urllib.parse
import urllib.request
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
REPORT_SRC = ROOT / "src" / "09_build_report.py"
OUT = ROOT / "results" / "tables" / "reference_dois.csv"

# Citations that are legitimately not journal articles: give a persistent
# identifier where one exists (arXiv registers DataCite DOIs) and otherwise
# keep the official URL, flagged as such.  These are NOT invented DOIs.
MANUAL = {
    18: ("10.48550/arXiv.2107.07511", "arXiv preprint (DataCite DOI)"),
    # RDKit has no official publication.  Its documentation explicitly asks
    # authors to cite https://www.rdkit.org AND the Zenodo DOI of the version
    # used (https://doi.org/10.5281/zenodo.591637 is the concept DOI, which
    # currently resolves to Release 2026_09_1 -- exactly the version pinned in
    # this project).  Verified against DataCite on 2026-11-05.
    27: ("10.5281/zenodo.591637",
         "software; Zenodo concept DOI for the RDKit release used (2026.09.1), "
         "as recommended by rdkit.org/docs/Overview.html"),
}

UA = "AIDD8122-report-builder/1.0 (mailto:p2600222@mpu.edu.mo)"


def extract_refs(path: Path) -> list[str]:
    tree = ast.parse(path.read_text(encoding="utf-8"))
    for node in ast.walk(tree):
        if isinstance(node, ast.Assign):
            for t in node.targets:
                if isinstance(t, ast.Name) and t.id == "refs":
                    return [ast.literal_eval(e) for e in node.value.elts]
    raise RuntimeError("refs = [...] not found in 09_build_report.py")


def norm(s: str) -> str:
    return re.sub(r"[^a-z0-9]+", " ", s.lower()).strip()


def crossref_query(citation: str, retries: int = 3) -> list[dict]:
    url = ("https://api.crossref.org/works?rows=4&query.bibliographic="
           + urllib.parse.quote(citation[:400]))
    req = urllib.request.Request(url, headers={"User-Agent": UA})
    for attempt in range(retries):
        try:
            with urllib.request.urlopen(req, timeout=25) as r:
                return json.load(r)["message"]["items"]
        except Exception as e:                       # noqa: BLE001
            if attempt == retries - 1:
                print(f"    query failed: {type(e).__name__}: {e}")
                return []
            time.sleep(1.5 * (attempt + 1))
    return []


def best_match(citation: str, items: list[dict]) -> tuple[dict | None, float]:
    nc = norm(citation)
    best, score = None, 0.0
    for it in items:
        title = (it.get("title") or [""])[0]
        nt = norm(title)
        if not nt:
            continue
        s = 1.0 if nt in nc else difflib.SequenceMatcher(None, nt, nc[:len(nt) * 3]).ratio()
        # reward agreement on year
        yr = ""
        dp = (it.get("issued") or {}).get("date-parts") or [[]]
        if dp and dp[0]:
            yr = str(dp[0][0])
        if yr and yr in citation:
            s += 0.08
        if s > score:
            best, score = it, s
    return best, score


def main() -> None:
    import pandas as pd

    refs = extract_refs(REPORT_SRC)
    print(f"parsed {len(refs)} references from 09_build_report.py\n")
    rows = []
    for i, ref in enumerate(refs, 1):
        if i in MANUAL:
            doi, note = MANUAL[i]
            rows.append({"ref": i, "doi": doi, "status": "manual",
                         "crossref_title": "", "score": "", "note": note,
                         "citation": ref})
            print(f"[{i:2d}] {note}")
            continue
        items = crossref_query(ref)
        best, score = best_match(ref, items)
        if best is None or score < 0.55:
            rows.append({"ref": i, "doi": "", "status": "UNRESOLVED",
                         "crossref_title": "", "score": round(score, 3),
                         "note": "Crossref returned no confident title match",
                         "citation": ref})
            print(f"[{i:2d}] !! UNRESOLVED (score {score:.2f}) {ref[:70]}")
            continue
        doi = best.get("DOI", "")
        title = (best.get("title") or [""])[0]
        meta = (f"{best.get('container-title', [''])[0]} "
                f"{(best.get('issued') or {}).get('date-parts', [['']])[0][0]} "
                f"vol {best.get('volume','')} issue {best.get('issue','')} "
                f"p {best.get('page','')}")
        rows.append({"ref": i, "doi": doi, "status": "verified",
                     "crossref_title": title, "score": round(score, 3),
                     "note": meta.strip(), "citation": ref})
        print(f"[{i:2d}] {doi:38s} score {score:.2f}  {title[:58]}")
        time.sleep(0.35)                            # be polite to Crossref

    OUT.parent.mkdir(parents=True, exist_ok=True)
    df = pd.DataFrame(rows)
    df.to_csv(OUT, index=False)
    ok = (df.doi != "").sum()
    print(f"\n{ok}/{len(df)} references carry a persistent identifier -> {OUT}")
    bad = df[df.doi == ""]
    if len(bad):
        print("without DOI:", bad.ref.tolist())


if __name__ == "__main__":
    main()
