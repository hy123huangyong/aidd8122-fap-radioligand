"""Step 23 -- final typesetting pass for the submitted report (and deck).

This is the LAST step of the report chain.  It is idempotent and never touches
the scientific content:

1. embeds the legible figures produced by ``22_figures_clear.py`` in the docx
   (and in the 10-minute deck), resizing each picture so the new aspect ratio
   is preserved;
2. turns all 16 data tables into **unshaded three-line tables** -- a 1.5 pt
   rule above the header, a 0.75 pt rule below it, a 1.5 pt rule at the foot,
   no vertical rules, no header fill;
3. makes paragraph spacing uniform (one value per structural tier);
4. removes the frame around the cover title;
5. changes the cover label "Submission Date:" to "Date:";
6. refreshes fields and exports the report PDF through Word; the deck PDF is
   re-exported through PowerPoint.

Usage (from the project root)::

    <venv>/Scripts/python.exe src/23_final_layout.py --check   # report only
    <venv>/Scripts/python.exe src/23_final_layout.py           # patch docx+deck
    <venv>/Scripts/python.exe src/23_final_layout.py --pdf     # also export PDFs
"""
from __future__ import annotations

import hashlib
import io
import shutil
import sys
import tempfile
import time
from pathlib import Path

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt
from PIL import Image

ROOT = Path(__file__).resolve().parents[1]
FIG = ROOT / "results" / "figures"
OLD_FIG = ROOT / "results" / "figures_bak_20261110"
OLDER_FIG = ROOT / "results" / "figures_bak"
DOCX = ROOT / "report" / "AIDD8122_Final_Report_HuangYong_P2600222_v2.docx"
PDF = ROOT / "report" / "AIDD8122_Final_Report_HuangYong_P2600222_v2.pdf"
DECK_DIR = ROOT / "AIDD8122_Final_Presentation_HuangYong"
DECK_STEM = "AIDD8122_Final_Presentation_HuangYong_10min"
PPTX = DECK_DIR / f"{DECK_STEM}.pptx"
DECK_PDF = DECK_DIR / f"{DECK_STEM}.pdf"

# one spacing value per structural tier, in points
SPACE = {
    "heading": (12, 6, 1.2),
    "toc": (0, 2, 1.15),
    "caption": (6, 6, 1.15),
    "image": (6, 2, 1.0),
    "body": (0, 6, 1.5),
    "cell": (2, 2, 1.15),
}
FIG_WIDTH_IN = 6.35          # full text-column width
LOGO_WIDTH_IN = 4.19         # native width of the university mark on the cover


def md5(data: bytes) -> str:
    return hashlib.md5(data).hexdigest()


def old_figure_index() -> dict[str, str]:
    """md5 -> current figure filename, for both generations of backup."""
    idx: dict[str, str] = {}
    for folder in (OLD_FIG, OLDER_FIG):
        for p in sorted(folder.glob("*.png")):
            if (FIG / p.name).exists():
                idx.setdefault(md5(p.read_bytes()), p.name)
    return idx


def replace_docx_images(doc, idx: dict[str, str]) -> int:
    """Swap the eight data figures; leave the university logo at native size."""
    current = {md5(p.read_bytes()) for p in FIG.glob("*.png")}
    figure_rids = set()
    replaced = 0
    for rid, part in doc.part.related_parts.items():
        if "image" not in part.content_type:
            continue
        name = idx.get(md5(part.blob))
        if name and hasattr(part, "_blob"):
            part._blob = (FIG / name).read_bytes()
            replaced += 1
        if md5(part.blob) in current:
            figure_rids.add(rid)

    for shape in doc.inline_shapes:
        blip = shape._inline.graphic.graphicData.pic.blipFill.blip
        rid = blip.get(qn("r:embed"))
        part = doc.part.related_parts.get(rid)
        if part is None:
            continue
        with Image.open(io.BytesIO(part.blob)) as im:
            ratio = im.width / im.height
        target_in = FIG_WIDTH_IN if rid in figure_rids else LOGO_WIDTH_IN
        cx = int(target_in * 914400)
        cy = int(cx / ratio)
        shape._inline.extent.cx = cx
        shape._inline.extent.cy = cy
        xfrm = shape._inline.graphic.graphicData.pic.spPr.xfrm
        if xfrm is not None and xfrm.ext is not None:
            xfrm.ext.set("cx", str(cx))
            xfrm.ext.set("cy", str(cy))
    return replaced


def restore_cover_spacing(doc) -> int:
    """Keep the template's cover header compact (it is not body text)."""
    n = 0
    for child in doc.element.body:
        if child.tag == qn("w:tbl"):
            break
        if child.tag != qn("w:p"):
            continue
        if child.findall(".//" + qn("w:drawing")):
            continue                                   # logo / green rule
        from docx.text.paragraph import Paragraph

        par = Paragraph(child, doc)
        text = par.text.strip()
        pf = par.paragraph_format
        pf.space_before = Pt(0)
        pf.space_after = Pt(2 if text.startswith("DOCTOR OF PHILOSOPHY") else 0)
        pf.line_spacing = 1.15 if text.startswith("FACULTY OF APPLIED SCIENCES") else 1.0
        n += 1
    return n


def replace_deck_images(idx: dict[str, str]) -> tuple[int, int]:
    """Refresh the seven pictures in the 10-minute deck without distorting them."""
    try:
        from pptx import Presentation
        from pptx.enum.shapes import MSO_SHAPE_TYPE
    except ImportError:
        print("  ! python-pptx unavailable -- deck images left untouched")
        return 0, 0
    if not PPTX.exists():
        return 0, 0
    prs = Presentation(str(PPTX))
    n = 0
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.shape_type != MSO_SHAPE_TYPE.PICTURE:
                continue
            name = idx.get(md5(shape.image.blob))
            if not name:
                continue
            new_bytes = (FIG / name).read_bytes()
            rid = shape._element.blipFill.blip.rEmbed
            part = shape.part.related_part(rid)
            if not hasattr(part, "_blob"):
                continue
            part._blob = new_bytes
            with Image.open(io.BytesIO(new_bytes)) as im:
                ratio = im.width / im.height
            box_w, box_h = shape.width, shape.height
            w = min(box_w, int(box_h * ratio))
            h = int(w / ratio)
            if w <= 0 or h <= 0:
                continue
            old_left, old_top = shape.left, shape.top
            shape.width, shape.height = w, h
            shape.left = old_left + (box_w - w) // 2
            shape.top = old_top                    # captions sit below the box
            n += 1
    if n:
        prs.save(str(PPTX))
    return n, len(prs.slides)


# ------------------------------------------------------------------ tables
def _edge(tag: str, val: str, sz: str) -> OxmlElement:
    el = OxmlElement("w:" + tag)
    el.set(qn("w:val"), val)
    el.set(qn("w:sz"), sz)
    el.set(qn("w:space"), "0")
    el.set(qn("w:color"), "000000")
    return el


def three_line_table(t, header: bool = True) -> None:
    """Replace grid borders/fills with three horizontal rules only."""
    try:
        t.style = "Normal Table"        # drops the Table Grid border set
    except KeyError:
        pass
    tblpr = t._tbl.tblPr
    borders = tblpr.find(qn("w:tblBorders"))
    if borders is not None:
        tblpr.remove(borders)
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        borders.append(_edge(edge, "single" if edge in ("top", "bottom") else "nil",
                             "12" if edge in ("top", "bottom") else "0"))
    # insert before tblLayout / tblCellMar to keep the CT_TblPr child order
    anchor = None
    for tag in ("w:shd", "w:tblLayout", "w:tblCellMar", "w:tblLook"):
        anchor = tblpr.find(qn(tag))
        if anchor is not None:
            break
    if anchor is not None:
        anchor.addprevious(borders)
    else:
        tblpr.append(borders)

    # WPS writes a full grid into a w:tblPrEx on almost every row.  Those
    # row-level exceptions override tblPr, so leaving them in place would keep
    # the very grid we just removed.  Drop only their border definitions.
    for ex in t._tbl.iter(qn("w:tblPrEx")):
        for b in ex.findall(qn("w:tblBorders")):
            ex.remove(b)

    for r, row in enumerate(t.rows):
        for cell in row.cells:
            tcpr = cell._tc.get_or_add_tcPr()
            for shd in tcpr.findall(qn("w:shd")):
                tcpr.remove(shd)                 # no header fill, no banding
            for old in tcpr.findall(qn("w:tcBorders")):
                tcpr.remove(old)
            if header and r == 0:
                b = OxmlElement("w:tcBorders")
                b.append(_edge("bottom", "single", "6"))
                tcpr.append(b)


def strip_cover_borders(t) -> int:
    """Remove the rectangle drawn around the cover title."""
    n = 0
    for row in t.rows:
        for cell in row.cells:
            tcpr = cell._tc.tcPr
            if tcpr is None:
                continue
            for b in tcpr.findall(qn("w:tcBorders")):
                tcpr.remove(b)
                n += 1
            for shd in tcpr.findall(qn("w:shd")):
                tcpr.remove(shd)
    return n


# ------------------------------------------------------------------ spacing
def _set_spacing(par, tier: str) -> None:
    before, after, line = SPACE[tier]
    pf = par.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    pf.line_spacing = line


def classify(par, cover_ids: set, tbl_ids: set) -> str | None:
    # lxml element proxies are recreated on access, so compare the elements
    # themselves, never id(proxy).
    if par._p in cover_ids or par._p in tbl_ids:
        return None
    p = par._p
    if p.findall(".//" + qn("w:fldChar")) or p.findall(".//" + qn("w:instrText")):
        return None                                   # TOC field itself
    ppr = p.pPr
    if ppr is not None and ppr.find(qn("w:pBdr")) is not None:
        return None                                   # boxed repository URL
    if p.findall(".//" + qn("w:drawing")):
        return "image"                                # picture-only paragraph
    text = par.text.strip()
    if not text:
        return None                                   # structural blank lines
    st = par.style.name
    if st.startswith("Heading"):
        return "heading"
    if st.lower().startswith("toc"):
        return "toc"
    if text.startswith(("Figure ", "Table ")):
        return "caption"
    return "body"


def uniform_spacing(doc) -> dict[str, int]:
    cover_ids: set = set()
    for child in doc.element.body:
        if child.tag == qn("w:tbl"):
            break
        if child.tag == qn("w:p"):
            cover_ids.add(child)
    tbl_ids = {p._p for t in doc.tables
               for row in t.rows for cell in row.cells for p in cell.paragraphs}

    counts: dict[str, int] = {}
    for par in doc.paragraphs:
        tier = classify(par, cover_ids, tbl_ids)
        if tier is None:
            continue
        _set_spacing(par, tier)
        counts[tier] = counts.get(tier, 0) + 1

    # table cells: a single uniform value, cover table excluded
    cells = 0
    for t in doc.tables[1:]:
        for row in t.rows:
            for cell in row.cells:
                for par in cell.paragraphs:
                    _set_spacing(par, "cell")
                    cells += 1
    counts["cell"] = cells
    return counts


def relabel_date(doc) -> int:
    n = 0
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                for par in cell.paragraphs:
                    for run in par.runs:
                        if "Submission Date" in run.text:
                            run.text = run.text.replace("Submission Date", "Date")
                            n += 1
    for par in doc.paragraphs:
        for run in par.runs:
            if "Submission Date" in run.text:
                run.text = run.text.replace("Submission Date", "Date")
                n += 1
    return n


# ------------------------------------------------------------------ export
def export_report_pdf() -> None:
    import win32com.client as win32

    tmp = Path(tempfile.gettempdir()) / f"aidd_final_{int(time.time())}.pdf"
    word = win32.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    try:
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


def export_deck_pdf() -> None:
    if not PPTX.exists():
        return
    import win32com.client as win32

    tmp = Path(tempfile.gettempdir()) / f"aidd_deck_{int(time.time())}.pdf"
    app = win32.DispatchEx("PowerPoint.Application")
    try:
        pres = app.Presentations.Open(str(PPTX), WithWindow=False)
        pres.SaveCopyAs(str(tmp), 32)            # 32 = ppSaveAsPDF
        pres.Close()
    finally:
        app.Quit()
    shutil.copy2(tmp, DECK_PDF)
    tmp.unlink(missing_ok=True)
    print(f"  deck PDF          : {DECK_PDF.stat().st_size:,} bytes")


def main() -> None:
    check_only = "--check" in sys.argv
    want_pdf = "--pdf" in sys.argv
    idx = old_figure_index()
    doc = Document(DOCX)

    if check_only:
        from collections import Counter
        c = Counter()
        cover_ids = set()
        for child in doc.element.body:
            if child.tag == qn("w:tbl"):
                break
            if child.tag == qn("w:p"):
                cover_ids.add(child)
        tbl_ids = {p._p for t in doc.tables
                   for row in t.rows for cell in row.cells for p in cell.paragraphs}
        for par in doc.paragraphs:
            tier = classify(par, cover_ids, tbl_ids)
            if tier is None:
                continue
            pf = par.paragraph_format
            c[(tier, pf.space_before.pt if pf.space_before else None,
               pf.space_after.pt if pf.space_after else None,
               pf.line_spacing)] += 1
        print("  spacing tiers:")
        for k, v in sorted(c.items(), key=lambda x: str(x[0])):
            print("   ", k, v)
        shaded = sum(1 for t in doc.tables[1:] for row in t.rows for cell in row.cells
                     if cell._tc.tcPr is not None
                     and cell._tc.tcPr.find(qn("w:shd")) is not None)
        grid_ex = sum(len(ex.findall(qn("w:tblBorders"))) for t in doc.tables[1:]
                      for ex in t._tbl.iter(qn("w:tblPrEx")))
        print(f"  tables            : {len(doc.tables) - 1} data tables, {shaded} shaded cells, "
              f"{grid_ex} row-level grid overrides")
        print("  cover borders     :", sum(1 for c0 in doc.tables[0].rows[0].cells
                                          if c0._tc.tcPr is not None
                                          and c0._tc.tcPr.find(qn("w:tcBorders")) is not None))
        print("  cover date labels :", [r.cells[0].text for r in doc.tables[0].rows
                                        if "Date" in r.cells[0].text])
        return

    print("  images in docx    :", replace_docx_images(doc, idx))
    print("  cover paras reset :", restore_cover_spacing(doc))
    for t in doc.tables[1:]:
        three_line_table(t)
    print(f"  three-line tables : {len(doc.tables) - 1}")
    print("  cover borders off :", strip_cover_borders(doc.tables[0]))
    counts = uniform_spacing(doc)
    print("  uniform spacing   :", ", ".join(f"{k}={v}" for k, v in sorted(counts.items())))
    print("  date relabelled   :", relabel_date(doc))
    doc.save(DOCX)

    n, slides = replace_deck_images(idx)
    print(f"  images in deck    : {n} (across {slides} slides)")

    if want_pdf:
        export_report_pdf()
        try:
            export_deck_pdf()
        except Exception as exc:                     # PowerPoint may not be installed
            print(f"  ! deck PDF not re-exported: {exc}")


if __name__ == "__main__":
    main()
