# -*- coding: utf-8 -*-
"""Put the live public GitHub URL into the final report and refresh deliverables.

Run from the project root, after the repository has been created and pushed::

    <venv>/Scripts/python.exe src/20_set_repo_url.py https://github.com/USER/REPO

IMPORTANT
---------
Do **not** change the link by re-running ``src/09_build_report.py``.  That
script rebuilds the .docx from the MPU template and would discard the
post-build corrections (recomposed figures, corrected figure numbers,
caption spacing, TOC, page fields).  This script patches the delivered
document in place instead.

What it does
------------
1. Writes the URL to ``repo_url.txt`` so any later full rebuild keeps it.
2. Replaces the Appendix A "Version-control location" line with a clickable
   hyperlink inside a light grey box (the prominent closing statement).
3. Re-exports ``report/<stem>.pdf`` through Word COM with fields refreshed.
4. Updates the two docs that describe code-availability status.
5. Rebuilds ``Submission_HuangYong_P2600222/`` and its .zip.
"""
from __future__ import annotations

import re
import shutil
import subprocess
import sys
import tempfile
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
STEM = "AIDD8122_Final_Report_HuangYong_P2600222_v2"
DOCX = ROOT / "report" / f"{STEM}.docx"
PDF = ROOT / "report" / f"{STEM}.pdf"

ANCHOR = "Version-control location:"
PLACEHOLDER = ("the submitted project repository "
               "(folder 03_Code/ of this submission package)")


def check(url: str) -> str:
    url = url.strip().rstrip("/")
    if not re.fullmatch(r"https://github\.com/[A-Za-z0-9_.-]+/[A-Za-z0-9_.-]+", url):
        raise SystemExit(
            "URL must look like https://github.com/USER/REPO "
            "(https, no trailing slash, no query string).")
    return url


# ----------------------------------------------------------------- docx patch
def _box(par) -> None:
    """Light grey four-sided paragraph border, valid OOXML child order."""
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    ppr = par._p.get_or_add_pPr()
    old = ppr.find(qn("w:pBdr"))
    if old is not None:
        ppr.remove(old)
    border = OxmlElement("w:pBdr")
    for edge in ("top", "left", "bottom", "right"):
        e = OxmlElement("w:" + edge)
        e.set(qn("w:val"), "single")
        e.set(qn("w:sz"), "6")
        e.set(qn("w:space"), "6")
        e.set(qn("w:color"), "808080")
        border.append(e)
    # pBdr must precede spacing / ind / jc in the schema
    before = None
    for child in ppr:
        if child.tag in (qn("w:spacing"), qn("w:ind"), qn("w:jc")):
            before = child
            break
    if before is None:
        ppr.append(border)
    else:
        before.addprevious(border)


def _hyperlink(par, url: str, rid: str) -> None:
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    hl = OxmlElement("w:hyperlink")
    hl.set(qn("r:id"), rid)
    run = OxmlElement("w:r")
    rpr = OxmlElement("w:rPr")
    fonts = OxmlElement("w:rFonts")
    for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        fonts.set(qn(attr), "Times New Roman")
    rpr.append(fonts)
    colour = OxmlElement("w:color")
    colour.set(qn("w:val"), "0563C1")
    rpr.append(colour)
    under = OxmlElement("w:u")
    under.set(qn("w:val"), "single")
    rpr.append(under)
    sz = OxmlElement("w:sz")
    sz.set(qn("w:val"), "22")           # 11 pt
    rpr.append(sz)
    run.append(rpr)
    text = OxmlElement("w:t")
    text.text = url
    run.append(text)
    hl.append(run)
    par._p.append(hl)


def patch_docx(url: str) -> None:
    from docx import Document
    from docx.opc.constants import RELATIONSHIP_TYPE as RT
    from docx.shared import Pt

    doc = Document(DOCX)
    target = None
    for par in doc.paragraphs:
        if par.text.strip().startswith(ANCHOR):
            target = par
            break
    if target is None:
        raise SystemExit("Appendix A 'Version-control location' paragraph not found")

    for run in list(target.runs):
        run._element.getparent().remove(run._element)
    for hl in target._p.findall("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}hyperlink"):
        target._p.remove(hl)

    label = target.add_run(ANCHOR + "  ")
    label.bold = True
    label.font.name = "Times New Roman"
    label.font.size = Pt(11)
    rid = doc.part.relate_to(url, RT.HYPERLINK, is_external=True)
    _hyperlink(target, url, rid)
    _box(target)
    target.paragraph_format.space_before = Pt(10)
    target.paragraph_format.space_after = Pt(6)
    doc.save(DOCX)
    print(f"  + docx  : {DOCX.name} -> {url}")


# ------------------------------------------------------------------ PDF export
def export_pdf() -> None:
    import win32com.client as win32

    tmp = Path(tempfile.gettempdir()) / f"aidd_report_{int(time.time())}.pdf"
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
        pages = doc.ComputeStatistics(2)
        doc.ExportAsFixedFormat(str(tmp), 17, False, 0, 0, 0, 0, 0,
                                True, True, 0, True, True, False)
        doc.Close(False)
    finally:
        word.Quit()
    shutil.copy2(tmp, PDF)
    tmp.unlink(missing_ok=True)
    print(f"  + pdf   : {PDF.name} ({pages} pages)")


# ----------------------------------------------------------------- docs + pack
def patch_docs(url: str) -> None:
    edits = {
        ROOT / "docs" / "00_START_HERE.md": [
            ("Appendix A of the report points to **the submitted project repository (folder\n"
             "`03_Code/` of this submission package)**, not an external GitHub URL. If the code\n"
             "is later published to GitHub, set the URL once in `src/09_build_report.py`\n"
             "(`REPO_PLACEHOLDER`) or via `repo_url.txt` / the `AIDD_REPO_URL` environment\n"
             "variable, then rebuild with `python src/09_build_report.py` and re-export the PDF\n"
             "with `src/build_report_pdf.ps1`.",
             f"Appendix A of the report carries a clickable link to the public repository\n"
             f"**{url}**.  The URL is stored in `repo_url.txt`; to change it, run\n"
             f"`python src/20_set_repo_url.py <new-url>` (do not rebuild with\n"
             f"`09_build_report.py`, which would discard post-build corrections)."),
            ("- [ ] Appendix A points to the submitted project repository (no dangling external link)",
             f"- [x] Appendix A links to the live public GitHub repository: {url}"),
        ],
        ROOT / "docs" / "ROUTE_DECISION.md": [
            ("| Submission | \"final report (PDF) and presentation slides (PDF) … link to an accessible GitHub repository\" | **One open item — see below.** |",
             f"| Submission | \"final report (PDF) and presentation slides (PDF) … link to an accessible GitHub repository\" | **Satisfied — {url}** |"),
            ("### One open item before submission\n\n"
             "The brief requires a link to an accessible GitHub repository. Appendix A currently\n"
             "points to the submitted `03_Code/` evidence bundle (no dangling placeholder URL).\n"
             "Once the repository exists, write its URL to `repo_url.txt` in the project root\n"
             "(or set `AIDD_REPO_URL`), then rebuild with `python src/09_build_report.py` and\n"
             "re-export the PDF. No other text needs editing.",
             f"### Repository link — resolved\n\n"
             f"The public repository is {url} . Appendix A carries the clickable URL;\n"
             f"`repo_url.txt` stores it for reproducibility. The `03_Code/` evidence bundle in the\n"
             f"submission package mirrors the same source tree."),
        ],
    }
    for path, pairs in edits.items():
        if not path.exists():
            continue
        text = path.read_text(encoding="utf-8")
        for old, new in pairs:
            if old in text:
                text = text.replace(old, new)
            else:
                print(f"  ! doc pattern not found in {path.name}: {old[:48]}...")
        path.write_text(text, encoding="utf-8")
        print(f"  + docs  : {path.name}")


def repack() -> None:
    subprocess.run([sys.executable, str(ROOT / "src" / "pack_submission.py"), "--clean"],
                   cwd=ROOT, check=True)


def main() -> None:
    if len(sys.argv) != 2:
        raise SystemExit(__doc__)
    url = check(sys.argv[1])
    (ROOT / "repo_url.txt").write_text(url + "\n", encoding="utf-8")
    print(f"  + config: repo_url.txt -> {url}")
    patch_docx(url)
    export_pdf()
    patch_docs(url)
    repack()
    print("\nDone. Verify the Appendix A link opens the public repository before submitting.")


if __name__ == "__main__":
    main()
