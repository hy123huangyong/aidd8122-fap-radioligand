"""Polish the layout of the final report (post-processing step 21).

The report generator (``09_build_report.py``) builds the document with
justified body text at 12 pt Times New Roman, 1.5-line spacing and a 6 pt
paragraph gap, but leaves every paragraph flush left.  This step applies the
typesetting conventions used for the submitted version:

* body paragraphs get a **first-line indent of two characters** (24 pt);
* the paragraph gap is reduced from 6 pt to 3 pt, so the indent and the gap do
  not double-separate consecutive paragraphs;
* everything else is deliberately left untouched: headings, the table of
  contents, figure/table captions, the cover page, the boxed repository URL,
  bulleted contribution/limitation lists, the RQ1-RQ4 list and the numbered
  references (the latter three already carry hanging indents).

Two long-standing layout defects are corrected in the same pass:

* the table of contents used to spill its final entry ("Appendix C") alone
  onto page 3.  The last blank paragraph between the "Table of Contents"
  heading and the first entry is collapsed to a 2 pt leader, which lifts that
  entry back onto page 2 without changing any entry or its page number;
* the file-path inventory in Appendix A is a list, not prose.  Justifying it
  after adding an indent stretched the word spacing badly, so it is set flush
  left.

The script is idempotent: it writes absolute values, never increments.

Usage (from the project root)::

    <venv>/Scripts/python.exe src/21_polish_layout.py           # patch docx
    <venv>/Scripts/python.exe src/21_polish_layout.py --pdf     # also export PDF
    <venv>/Scripts/python.exe src/21_polish_layout.py --check   # verify only
"""

from __future__ import annotations

import shutil
import sys
import tempfile
import time
from collections import Counter
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Pt

ROOT = Path(__file__).resolve().parents[1]
DOCX = ROOT / "report" / "AIDD8122_Final_Report_HuangYong_P2600222_v2.docx"
PDF = ROOT / "report" / "AIDD8122_Final_Report_HuangYong_P2600222_v2.pdf"

FIRST_LINE_PT = 24          # 2 characters at the 12 pt body size
FIRST_LINE_CHARS = "200"    # Word's native "2 characters" attribute
BODY_SPACE_AFTER_PT = 3     # was 6 pt; keeps the page count stable
TOC_LEADING_GAP_PT = 2      # collapsed blank line before the first TOC entry


def classify(par) -> str:
    """Return ``body`` for prose paragraphs, otherwise the reason to skip."""
    if par.style.name != "Normal":
        return "style:" + par.style.name
    text = par.text.strip()
    if not text:
        return "empty"
    ppr = par._p.find(qn("w:pPr"))
    if ppr is not None:
        if ppr.find(qn("w:pBdr")) is not None:
            return "boxed"
        if ppr.find(qn("w:numPr")) is not None:
            return "numbered"
        ind = ppr.find(qn("w:ind"))
        if ind is not None and (
            ind.get(qn("w:hanging")) or ind.get(qn("w:left"))
        ):
            return "hanging-indent"
    if par._p.findall(".//" + qn("w:drawing")):
        return "image"
    if par._p.findall(qn("w:hyperlink")):
        return "hyperlink"
    if par.alignment != WD_ALIGN_PARAGRAPH.JUSTIFY:
        return "not-justified"
    for prefix in ("Figure ", "Table ", "Keywords:", "\u2022"):
        if text.startswith(prefix):
            return "caption-or-list"
    return "body"


def set_indent(par) -> None:
    pf = par.paragraph_format
    pf.first_line_indent = Pt(FIRST_LINE_PT)
    ind = par._p.get_or_add_pPr().find(qn("w:ind"))
    # Word honours firstLineChars over firstLine; both express 2 characters.
    ind.set(qn("w:firstLineChars"), FIRST_LINE_CHARS)
    pf.space_after = Pt(BODY_SPACE_AFTER_PT)


def polish_toc_leading(doc) -> int:
    """Collapse the last blank line before the first TOC entry."""
    paras = doc.paragraphs
    heads = [i for i, p in enumerate(paras)
             if p.text.strip() == "Table of Contents"]
    if not heads:
        return 0
    blanks = []
    for par in paras[heads[0] + 1:]:
        if par.text.strip():
            break
        blanks.append(par)
    if not blanks:
        return 0
    pf = blanks[-1].paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing = Pt(TOC_LEADING_GAP_PT)
    return 1


def relax_path_inventory(doc) -> int:
    """Left-align the Appendix A path inventory (a list, not prose)."""
    n = 0
    for par in doc.paragraphs:
        if par.style.name != "Normal":
            continue
        text = par.text.strip()
        if not text or text.count(" \u2014 ") < 3 or "/" not in text:
            continue
        ppr = par._p.find(qn("w:pPr"))
        ind = ppr.find(qn("w:ind")) if ppr is not None else None
        if ind is not None and (ind.get(qn("w:hanging")) or ind.get(qn("w:left"))):
            continue                      # already a hanging list
        par.alignment = WD_ALIGN_PARAGRAPH.LEFT
        n += 1
    return n


def export_pdf() -> None:
    """Refresh fields and export the PDF through Word (Windows only)."""
    import win32com.client as win32

    tmp = Path(tempfile.gettempdir()) / f"aidd_layout_{int(time.time())}.pdf"
    word = win32.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    try:
        doc = word.Documents.Open(str(DOCX), False, True, False)
        doc.Fields.Update()
        for sec in doc.Sections:
            for i in (1, 2, 3):
                try:
                    sec.Footers(i).Range.Fields.Update()
                except Exception:
                    pass
                try:
                    sec.Headers(i).Range.Fields.Update()
                except Exception:
                    pass
        doc.Repaginate()
        print(f"  Word page count : {doc.ComputeStatistics(2)}")
        doc.ExportAsFixedFormat(
            str(tmp), 17, False, 0, 0, 0, 0, 0, True, True, 0, True, True, False
        )
        doc.Close(False)
    finally:
        word.Quit()
    shutil.copy2(tmp, PDF)
    tmp.unlink(missing_ok=True)
    print(f"  PDF              : {PDF.name} ({PDF.stat().st_size:,} bytes)")


def main() -> None:
    check_only = "--check" in sys.argv
    want_pdf = "--pdf" in sys.argv
    doc = Document(DOCX)
    counts: Counter = Counter()
    bodies = []
    for par in doc.paragraphs:
        verdict = classify(par)
        counts[verdict] += 1
        if verdict == "body":
            bodies.append(par)

    print(f"  paragraphs       : {len(doc.paragraphs)}")
    for reason, n in counts.most_common():
        print(f"    {reason:<22} {n}")

    if check_only:
        ok = sum(
            1
            for p in bodies
            if p.paragraph_format.first_line_indent
            and abs(p.paragraph_format.first_line_indent.pt - FIRST_LINE_PT) < 0.1
        )
        print(f"  first-line indent: {ok}/{len(bodies)} body paragraphs")
        return

    for par in bodies:
        set_indent(par)
    print(f"  patched          : {len(bodies)} body paragraphs")
    print(f"    first-line     : {FIRST_LINE_PT} pt / 2 characters")
    print(f"    space after    : {BODY_SPACE_AFTER_PT} pt")
    print(f"  TOC leading      : {polish_toc_leading(doc)} blank line collapsed")
    print(f"  path inventory   : {relax_path_inventory(doc)} paragraph left-aligned")
    doc.save(DOCX)
    if want_pdf:
        export_pdf()


if __name__ == "__main__":
    main()
