# -*- coding: utf-8 -*-
"""Final copy-edit of the report: terminology and one redundant paragraph.

Two things are fixed here, both small and checkable:

1. The deck now says "four regressors plus a mean baseline"; the report still
   called the same benchmark "five model families" in three places. A mean
   predictor is a baseline, not a model family, so the report is made precise.
2. Section 5.8 repeated every number of Table 9 a second time in the following
   paragraph. The argument is kept; the repeated figures are replaced by a
   pointer to the table.

Usage:
    python src/33_report_polish.py --check
    python src/33_report_polish.py          # rewrites the docx and exports the PDF
    python src/33_report_polish.py --no-pdf
"""
from __future__ import annotations

import argparse
import importlib.util
import shutil
import sys
import tempfile
from pathlib import Path

from docx import Document

ROOT = Path(__file__).resolve().parents[1]
DOCX = ROOT / "report" / "AIDD8122_Final_Report_HuangYong_P2600222_v2.docx"

# (paragraph index, run index, old fragment, new fragment)
RUN_EDITS = [
    (62, 0,
     "and benchmarked five models under random, scaffold and time splits.",
     "and benchmarked four regressors against a mean baseline under random, "
     "scaffold and time splits."),
    (77, 0,
     "of five model families with identical features and hyperparameters",
     "of four regressors and a mean baseline, with identical features and hyperparameters"),
    (146, 1,
     "Benchmark of five model families under three split protocols",
     "Benchmark of four regressors and a mean baseline under three split protocols"),
]

OLD_197 = ("Two features of Table 9 need care. First, the affinity-only 100% pass rate follows "
           "from the metric, not better chemistry: selection drives all 369 novel molecules above "
           "pFAP 8, and since margin is pFAP minus the largest paralogue prediction, a high pFAP "
           "almost automatically meets margin \u2265 1. The same circularity explains why adding "
           "selectivity and developability pressure lowers the pass rate to 61%: the optimizer can "
           "no longer spend its whole budget on pFAP. What that budget bought matters more than the "
           "pass rate \u2014 median QED 0.61, MW 413 Da, 115 scaffolds at pFAP 8.24, versus QED 0.25, "
           "MW 548 Da, 83 scaffolds for affinity-only. The top-ranked designs are "
           "quinoline-carboxamides with modified amide linkers (Table 10); Section 5.11 and Appendix "
           "B explain why their warhead is not the clinical FAPI warhead.")

NEW_197 = ("Two features of Table 9 need care. First, the affinity-only 100% pass rate follows "
           "from the metric, not from better chemistry: selection drives all 369 novel molecules "
           "above pFAP 8, and because margin is pFAP minus the largest paralogue prediction, a high "
           "pFAP almost automatically satisfies margin \u2265 1. The same circularity explains why "
           "adding selectivity and developability pressure lowers the pass rate to 61% \u2014 the "
           "optimizer can no longer spend its whole budget on pFAP. What that budget bought is "
           "visible in the other columns of Table 9 and matters more than the pass rate itself. The "
           "top-ranked designs are quinoline-carboxamides with modified amide linkers (Table 10); "
           "Section 5.11 and Appendix B explain why their warhead is not the clinical FAPI warhead.")


def word_count(doc) -> int:
    return sum(len(p.text.split()) for p in doc.paragraphs)


def apply(doc) -> list[str]:
    log = []
    for idx, run_idx, old, new in RUN_EDITS:
        run = doc.paragraphs[idx].runs[run_idx]
        if new in run.text:
            log.append(f"  [{idx:3d}] already applied")
            continue
        assert old in run.text, f"paragraph {idx} run {run_idx}: fragment not found"
        run.text = run.text.replace(old, new, 1)
        log.append(f"  [{idx:3d}] run {run_idx}: {old[:44]!r} -> {new[:44]!r}")

    para = doc.paragraphs[197]
    if para.text.strip() != NEW_197:
        assert para.text.strip() == OLD_197, "paragraph 197 differs from the expected original"
        para.runs[0].text = NEW_197
        for extra in para.runs[1:]:
            extra._r.getparent().remove(extra._r)
        log.append("  [197] section 5.8: dropped the repeated Table 9 figures")
    else:
        log.append("  [197] already applied")
    return log


def check(doc) -> None:
    joined = "\n".join(p.text for p in doc.paragraphs)
    assert "five model families" not in joined, "old terminology still present"
    assert "four regressors against a mean baseline" in joined
    assert doc.paragraphs[197].text.strip() == NEW_197
    print("  check          : terminology and paragraph 5.8 are in final form")


def export_pdf() -> None:
    spec = importlib.util.spec_from_file_location(
        "c27", Path(__file__).with_name("27_condense_report.py"))
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    module.export_pdf()


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--check", action="store_true")
    ap.add_argument("--no-pdf", action="store_true")
    args = ap.parse_args()

    doc = Document(str(DOCX))
    if args.check:
        check(doc)
        return

    before = word_count(doc)
    backup = Path(tempfile.gettempdir()) / "report_before_polish.docx"
    shutil.copy2(DOCX, backup)

    for line in apply(doc):
        print(line)
    doc.save(str(DOCX))

    after = word_count(Document(str(DOCX)))
    print(f"  words          : {before:,} -> {after:,} ({after - before:+,})")
    print(f"  backup         : {backup}")

    if not args.no_pdf:
        export_pdf()
    check(Document(str(DOCX)))


if __name__ == "__main__":
    main()
