# -*- coding: utf-8 -*-
"""Keep Appendix A and the package indices in sync with the real layout.

Appendix A originally described the *repository* view (``03_Code/results/``),
while the submission package actually ships those files in ``04_Results/`` and
its ``src/`` now contains scripts 01-24 rather than 01-20.  This idempotent
pass rewrites those two paragraphs, aligns the two index documents, and can
re-export the report PDF (``--pdf``).  It changes wording only: no number,
figure, caption, reference or formatting attribute is touched.

    python src/25_fix_appendix_paths.py --check   # read-only verification
    python src/25_fix_appendix_paths.py           # patch docx + markdown
    python src/25_fix_appendix_paths.py --pdf     # patch + re-export PDF
"""
from __future__ import annotations

import shutil
import sys
import tempfile
import time
from pathlib import Path

from docx import Document

ROOT = Path(__file__).resolve().parents[1]
REPORT = ROOT / "report"
DOCX = REPORT / "AIDD8122_Final_Report_HuangYong_P2600222_v2.docx"
PDF = REPORT / "AIDD8122_Final_Report_HuangYong_P2600222_v2.pdf"
START = ROOT / "docs" / "00_START_HERE.md"
README = ROOT / "README.md"

OLD_INTRO = (
    "The complete project is submitted as a single package, "
    "Submission_HuangYong_P2600222/. Its 03_Code/ directory is the evidence "
    "bundle referenced by this report:"
)
NEW_INTRO = (
    "The complete project is submitted as a single package; 03_Code/ inside "
    "it holds the source code and process evidence, and 04_Results/ holds "
    "every table and figure cited here. The package root is "
    "Submission_HuangYong_P2600222/."
)

OLD_TREE = (
    "03_Code/README.md — one-command reproduction, table-to-script map and "
    "environment fingerprint; 03_Code/requirements.txt — pinned dependency "
    "versions; 03_Code/src/ — scripts 01–20 in execution order; "
    "03_Code/data/raw/ — the three ChEMBL release-37 JSON exports together "
    "with their retrieval date and API query strings; 03_Code/data/processed/ "
    "— the 972 → 438 → 384 curation chain; 03_Code/results/tables/ — every "
    "number quoted in Tables 1–10 and Appendices B–C; "
    "03_Code/results/figures/ — all eight figures at ≥2,000 px; "
    "03_Code/logs/ — random seeds, package versions, runtime and the "
    "reconciliation log. Running the scripts in numerical order from "
    "README.md regenerates Tables 1, 2, 8, 9 and 10 from the public API "
    "without any manual step."
)
NEW_TREE = (
    "03_Code/README.md — one-command reproduction, table-to-script map and "
    "environment fingerprint; 03_Code/requirements.txt — pinned dependency "
    "versions; 03_Code/src/ — scripts 01–24 in execution order (01–13 "
    "analysis, 14–24 delivery); 03_Code/data/raw/ — the three ChEMBL "
    "release-37 JSON exports together with their retrieval date and API "
    "query strings; 03_Code/data/processed/ — the 972 → 438 → 384 curation "
    "chain; 04_Results/tables/ — every number quoted in Tables 1–10 and "
    "Appendices B–C; 04_Results/figures/ — all eight figures at ≥2,000 px; "
    "03_Code/logs/ — random seeds, package versions, runtime and the "
    "reconciliation log. Running scripts 01–13 in numerical order from "
    "README.md regenerates Tables 1, 2, 8, 9 and 10 from the public API; "
    "scripts 14–24 are delivery-format passes that rebuild the report and "
    "the deck."
)

# (path, old, new) -- every edit is exact and idempotent
DOC_FIXES: list[tuple[Path, str, str]] = [
    (START,
     "report, 30 pages, live TOC (submission copy)",
     "report, 32 pages, live TOC (submission copy)"),
    (START,
     "    figures/                  fig0–fig8, as they appear in the report",
     "    figures/                  the eight report figures (fig0/1/2/3/4/5/7/8)"),
    (START,
     "`AIDD8122_Final_Presentation_HuangYong/`; `README.md` there documents that view.",
     "`AIDD8122_Final_Presentation_HuangYong/`; `README.md` there documents that view.\n"
     "`03_Code/README.md` is written from that repository view: its `results/tables/`\n"
     "and `results/figures/` correspond to `04_Results/tables/` and\n"
     "`04_Results/figures/` in this package. Appendix A of the report uses the\n"
     "package paths."),
    (START,
     "steps 21–23 must be re-applied afterwards.",
     "steps 21–23 (layout) and 25 (Appendix A paths) must be re-applied afterwards."),
    (START,
     "`23_final_layout.py`\n"
     "is the **final report** idempotent pass",
     "`23_final_layout.py`\n"
     "is the last report-layout idempotent pass"),
    (START,
     "The crest/wordmark and the two deck-only\nschematics it needs live in `assets/`.",
     "The crest/wordmark and the two deck-only schematics it needs live in\n"
     "`assets/`. `25_fix_appendix_paths.py` keeps Appendix A's directory tree\n"
     "aligned with `04_Results/` and re-exports the report PDF."),
    (README,
     "  24_restyle_mpu_deck.py rebuilds the 14-slide deck in the MPU presentation style (--pdf exports)\n",
     "  24_restyle_mpu_deck.py rebuilds the 14-slide deck in the MPU presentation style (--pdf exports)\n"
     "  25_fix_appendix_paths.py  syncs Appendix A paths with 04_Results/ (--pdf re-exports)\n"),
    (README,
     "repo_url.txt             the public repository URL quoted in Appendix A\n```",
     "repo_url.txt             the public repository URL quoted in Appendix A\n```\n\n"
     "**Submission-package view.** When this README is copied to\n"
     "`Submission_HuangYong_P2600222/03_Code/README.md`, the repository paths\n"
     "above map as follows: `results/tables/` → `../04_Results/tables/`,\n"
     "`results/figures/` → `../04_Results/figures/`, `report/` → `../01_Report/`,\n"
     "`data/processed/` → `data/processed/` and `../05_Data/`. Appendix A of the\n"
     "report quotes the package paths directly."),
]


def _set_paragraph_text(par, text: str) -> None:
    """Replace paragraph text while preserving the first run's formatting."""
    if not par.runs:
        par.add_run(text)
        return
    par.runs[0].text = text
    for run in par.runs[1:]:
        run.text = ""


def patch_docx() -> list[str]:
    doc = Document(DOCX)
    intro = tree = None
    for par in doc.paragraphs:
        t = par.text.strip()
        if t.startswith("The complete project is submitted as a single package"):
            intro = par
        elif t.startswith("03_Code/README.md —"):
            tree = par
    if intro is None or tree is None:
        raise SystemExit("ERROR: could not locate the two Appendix A paragraphs")

    changes = []
    for label, par, old, new in (("intro", intro, OLD_INTRO, NEW_INTRO),
                                 ("tree", tree, OLD_TREE, NEW_TREE)):
        cur = par.text.strip()
        if cur == new:
            changes.append(f"  {label:6s}: already current")
        elif cur == old:
            _set_paragraph_text(par, new)
            changes.append(f"  {label:6s}: rewritten")
        else:
            raise SystemExit(f"ERROR: {label} text is neither old nor expected new form")
    if any("rewritten" in c for c in changes):
        doc.save(DOCX)
    return changes


def patch_docs() -> list[str]:
    out = []
    for path, old, new in DOC_FIXES:
        text = path.read_text(encoding="utf-8")
        if old in text:
            path.write_text(text.replace(old, new, 1), encoding="utf-8")
            out.append(f"  {path.relative_to(ROOT)}: edited  <- {old[:52]!r}")
        elif new in text:
            out.append(f"  {path.relative_to(ROOT)}: already current")
        else:
            out.append(f"  {path.relative_to(ROOT)}: WARNING anchor not found: {old[:52]!r}")
    return out


def export_pdf() -> None:
    import win32com.client as win32

    tmp = Path(tempfile.gettempdir()) / f"aidd_report_{int(time.time())}.pdf"
    word = win32.DispatchEx("Word.Application")
    word.Visible = False
    try:
        word.DisplayAlerts = 0
        doc = word.Documents.Open(str(DOCX), False, True, False)
        doc.Fields.Update()
        for sec in doc.Sections:
            for i in (1, 2, 3):
                for coll in (sec.Footers, sec.Headers):
                    try:
                        coll(i).Range.Fields.Update()
                    except Exception:
                        pass
        doc.Repaginate()
        print(f"  Word pages        : {doc.ComputeStatistics(2)}")
        doc.ExportAsFixedFormat(str(tmp), 17, False, 0, 0, 0, 0, 0,
                                True, True, 0, True, True, False)
        doc.Close(False)
    finally:
        word.Quit()
    shutil.copy2(tmp, PDF)
    tmp.unlink(missing_ok=True)
    print(f"  report PDF        : {PDF.stat().st_size:,} bytes")


def check() -> None:
    doc = Document(DOCX)
    texts = [p.text.strip() for p in doc.paragraphs]
    for label, old, new in (("Appendix A intro", OLD_INTRO, NEW_INTRO),
                            ("Appendix A tree ", OLD_TREE, NEW_TREE)):
        state = "OK (current)" if new in texts else ("OLD (needs patch)" if old in texts
                                                     else "MISSING (unexpected)")
        print(f"  {label}: {state}")
    for p in (START, README):
        t = p.read_text(encoding="utf-8")
        # an append-style edit keeps `old` as a prefix of `new`; only an `old`
        # whose replacement is absent counts as stale
        stale = [old for path, old, new in DOC_FIXES
                 if path == p and old in t and new not in t]
        print(f"  {str(p.relative_to(ROOT)):28s}: {len(stale)} stale anchor(s)")
    print("  old path in docx   :", any("03_Code/results/" in t for t in texts))
    print("  report PDF exists  :", PDF.exists(), PDF.stat().st_size if PDF.exists() else "")


def main() -> None:
    if "--check" in sys.argv:
        print("=== CHECK ===")
        check()
        return
    print("=== DOCX ===")
    for line in patch_docx():
        print(line)
    print("=== MARKDOWN ===")
    for line in patch_docs():
        print(line)
    if "--pdf" in sys.argv:
        print("=== PDF ===")
        export_pdf()
    print("=== DONE ===")


if __name__ == "__main__":
    main()
