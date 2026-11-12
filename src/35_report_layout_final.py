"""Final typography pass on the v2 report.

What this pass changes, and nothing else:

  cover      the ID / name / date block becomes three tight rows on the same
             left edge, the four filler rows between name and date are dropped,
             the gap under the course title is reduced, and the date is the day
             the file was last edited (12/11/2026)
  TOC        no rule under the heading, no dot leaders; page numbers stay flush
             right on the same tab stop
  headings   H1 15.5 pt and H2 13 pt (body stays 12 pt), more air above each
             heading than below it, keep-with-next so no heading ends a page
  body       keeps 1.5 line spacing but drops the extra 6 pt after every
             paragraph, so paragraphs are separated by the indent alone
  isotopes   68Ga, 18F, 90Y and 177Lu get a true superscript mass number

Word regenerates the table of contents when the fields are refreshed, and that
regeneration puts the dot leaders back. The export is therefore a two-stage
hand-off: Word refreshes fields and saves, the leaders are stripped from the
saved file, then Word exports the PDF without refreshing again.

Usage
    python src/35_report_layout_final.py            # edit docx, refresh, export
    python src/35_report_layout_final.py --check    # report state, write nothing
    python src/35_report_layout_final.py --no-pdf   # edit and refresh only
"""

from __future__ import annotations

import argparse
import re
import shutil
import tempfile
from copy import deepcopy
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.shared import Pt

ROOT = Path(__file__).resolve().parent.parent
DOCX = ROOT / "report" / "AIDD8122_Final_Report_HuangYong_P2600222_v2.docx"
PDF = ROOT / "report" / "AIDD8122_Final_Report_HuangYong_P2600222_v2.pdf"

COVER_DATE = "12/11/2026"
TOC_TITLE = "Table of Contents"

# heading geometry: (size pt, space before pt, space after pt)
H1 = (15.5, 24, 10)
H2 = (13.0, 16, 6)
BODY_SIZE = 12.0
BULLET_SIZE = 12.0

# mass number immediately in front of an element symbol, not part of a word or
# a decimal; C/N/O are left out on purpose so SMILES strings stay untouched
ISOTOPE = re.compile(
    r"(?<![0-9A-Za-z.])(\d{1,3}m?)(Ga|Lu|F|Tc|I|Cu|Zr|Ac|In|Y)(?![A-Za-z0-9])"
)


# ── small helpers ───────────────────────────────────────────────────────────
def inch(value) -> float:
    return 0.0 if value is None else value / 914400


def set_spacing(par, before=None, after=None, line=None, rule=None) -> None:
    pf = par.paragraph_format
    if before is not None:
        pf.space_before = Pt(before)
    if after is not None:
        pf.space_after = Pt(after)
    if line is not None:
        pf.line_spacing = line
        if rule is not None:
            pf.line_spacing_rule = rule


def set_runs_size(par, size: float) -> int:
    n = 0
    for run in par.runs:
        run.font.size = Pt(size)
        if run.text.strip():
            n += 1
    return n


def drop_row(table, index: int) -> None:
    tr = table.rows[index]._tr
    tr.getparent().remove(tr)


# ── cover ───────────────────────────────────────────────────────────────────
def cover(doc) -> dict:
    tbl = doc.tables[0]
    rows = list(tbl.rows)
    label = lambda r: r.cells[0].text.strip().rstrip(":")

    info = {}
    for ri, row in enumerate(rows):
        key = label(row)
        if key in ("Student ID", "Student Name", "Date"):
            info[key] = ri
    assert set(info) == {"Student ID", "Student Name", "Date"}, info

    # date = the day this file was last edited
    date_cell = rows[info["Date"]].cells[1]
    par = date_cell.paragraphs[0]
    if par.runs:
        par.runs[0].text = COVER_DATE
        for extra in par.runs[1:]:
            extra.text = ""
    else:  # pragma: no cover - the template always ships a run here
        par.add_run(COVER_DATE)

    # the four filler rows between name and date
    filler = [ri for ri in range(info["Student Name"] + 1, info["Date"])
              if not "".join(c.text for c in rows[ri].cells).strip()]
    for ri in reversed(filler):
        drop_row(tbl, ri)

    # spacer under the course title: seven blank lines is too much air
    spacer = rows[1].cells[0]
    blanks = [p for p in spacer.paragraphs if not p.text.strip()]
    keep = 5
    for p in blanks[keep:]:
        p._p.getparent().remove(p._p)

    return {"date": COVER_DATE, "rows_removed": len(filler),
            "spacer_lines": len(blanks[:keep])}


# ── table of contents ───────────────────────────────────────────────────────
def strip_toc_rule(doc) -> int:
    """Remove the underline from the 'Table of Contents' heading only."""
    n = 0
    for par in doc.paragraphs:
        if par.text.strip() != TOC_TITLE:
            continue
        for u in par._p.findall(".//" + qn("w:u")):
            u.getparent().remove(u)
            n += 1
    return n


def strip_leaders(doc) -> int:
    """Turn every dot-leader tab stop inside the TOC into a plain tab stop."""
    n = 0
    for par in doc.paragraphs:
        if not par.style.name.lower().startswith("toc"):
            continue
        for tab in par._p.findall(".//" + qn("w:tab")):
            if tab.get(qn("w:leader")) in ("dot", "middleDot", "heavy"):
                tab.set(qn("w:leader"), "none")
                n += 1
    return n


def strip_leaders_in_styles(doc) -> int:
    """Word rebuilds the TOC from the TOC1/TOC2 styles, which carry their own
    dot leader; clearing it there is what survives a field refresh."""
    n = 0
    for style in doc.styles:
        name = (style.name or "").lower()
        if not name.startswith("toc"):
            continue
        for tab in style.element.findall(".//" + qn("w:tab")):
            if tab.get(qn("w:leader")) in ("dot", "middleDot", "heavy"):
                tab.set(qn("w:leader"), "none")
                n += 1
    return n


TOC_RUN = {"toc 1": (12.0, True), "toc 2": (10.0, False), "toc 3": (10.0, False)}


def toc_type(doc) -> int:
    """Word rewrites every TOC entry at 11 pt and copies the heading's bold into
    the level-2 entries. Put the hierarchy back: level 1 bold, level 2 light.

    The entry text sits inside <w:hyperlink>, so paragraph.runs sees almost none
    of it; every w:r descendant has to be wrapped instead.
    """
    from docx.text.run import Run

    n = 0
    for par in doc.paragraphs:
        spec = TOC_RUN.get(par.style.name.lower())
        if spec is None:
            continue
        size, bold = spec
        for el in par._p.findall(".//" + qn("w:r")):
            run = Run(el, par)
            run.font.size = Pt(size)
            run.font.bold = bold
        n += 1
    return n


# ── headings and body ───────────────────────────────────────────────────────
def headings(doc) -> dict:
    counts = {"h1": 0, "h2": 0, "h3": 0}
    for par in doc.paragraphs:
        name = par.style.name
        if name == "Heading 1":
            size, before, after = H1
            counts["h1"] += 1
        elif name == "Heading 2":
            size, before, after = H2
            counts["h2"] += 1
        elif name == "Heading 3":
            size, before, after = 12.0, 12, 4
            counts["h3"] += 1
        else:
            continue
        set_spacing(par, before=before, after=after, line=1.15,
                    rule=WD_LINE_SPACING.MULTIPLE)
        set_runs_size(par, size)
    return counts


def toc_styles(doc) -> int:
    """Set the size on the TOC styles too, so a field update in Word keeps the
    hierarchy instead of falling back to Word's 11 pt for both levels."""
    n = 0
    for style in doc.styles:
        spec = TOC_RUN.get((style.name or "").lower())
        if spec is None:
            continue
        size, bold = spec
        style.font.size = Pt(size)
        style.font.bold = bold
        n += 1
    return n


def keep_headings_with_next(doc) -> int:
    """Applied after Word has saved: Word's own save drops the flag."""
    n = 0
    for par in doc.paragraphs:
        if par.style.name in ("Heading 1", "Heading 2", "Heading 3"):
            par.paragraph_format.keep_with_next = True
            par.paragraph_format.keep_together = True
            n += 1
    return n


def captions(doc) -> int:
    """A three-line table caption was being cut by a page break while its table
    moved on alone; keep captions whole and glued to the float they label."""
    n = 0
    for par in doc.paragraphs:
        text = par.text.strip()
        if not text or text == TOC_TITLE:
            continue
        if text.startswith(("Table ", "Figure ")) and len(text) < 600:
            par.paragraph_format.keep_with_next = True
            par.paragraph_format.keep_together = True
            n += 1
    return n


def path_lists(doc) -> int:
    """Appendix A lists file paths in running prose. Justified, every line with a
    long unbreakable path gets stretched, so that block is set flush left."""
    n = 0
    for par in doc.paragraphs:
        text = par.text
        if text.count("03_Code/") + text.count("04_Results/") >= 3:
            par.alignment = WD_ALIGN_PARAGRAPH.LEFT
            n += 1
    return n


def body(doc) -> dict:
    n_body = n_bullet = 0
    for par in doc.paragraphs:
        if par.style.name != "Normal":
            continue
        text = par.text.strip()
        if not text:
            continue
        ind = par._p.pPr.find(qn("w:ind")) if par._p.pPr is not None else None
        first = ind.get(qn("w:firstLineChars")) if ind is not None else None
        bullet = text.startswith(("•", "–", "-"))
        if text.startswith("Keywords"):
            # sits directly under the abstract, which no longer carries a gap
            set_spacing(par, before=6, after=0, line=1.5,
                        rule=WD_LINE_SPACING.MULTIPLE)
            set_runs_size(par, BODY_SIZE)
            n_body += 1
        elif bullet:
            set_spacing(par, before=0, after=3, line=1.5,
                        rule=WD_LINE_SPACING.MULTIPLE)
            set_runs_size(par, BULLET_SIZE)
            n_bullet += 1
        elif first == "200":
            set_spacing(par, before=0, after=0, line=1.5,
                        rule=WD_LINE_SPACING.MULTIPLE)
            set_runs_size(par, BODY_SIZE)
            n_body += 1
    return {"prose": n_body, "bullets": n_bullet}


# ── isotopes ────────────────────────────────────────────────────────────────
def _split_run(run) -> int:
    """Rewrite one run so each mass number becomes a superscript fragment."""
    text = run.text
    if not ISOTOPE.search(text):
        return 0
    pieces: list[tuple[str, bool]] = []
    pos = 0
    for m in ISOTOPE.finditer(text):
        if m.start() > pos:
            pieces.append((text[pos:m.start()], False))
        pieces.append((m.group(1), True))
        pieces.append((m.group(2), False))
        pos = m.end()
    if pos < len(text):
        pieces.append((text[pos:], False))

    parent = run._r.getparent()
    anchor = run._r
    made = 0
    for seg, sup in pieces:
        el = deepcopy(run._r)
        for child in list(el):
            if child.tag != qn("w:rPr"):
                el.remove(child)
        if sup:
            rPr = el.find(qn("w:rPr"))
            if rPr is None:
                rPr = el.makeelement(qn("w:rPr"), {})
                el.insert(0, rPr)
            for old in rPr.findall(qn("w:vertAlign")):
                rPr.remove(old)
            va = rPr.makeelement(qn("w:vertAlign"), {qn("w:val"): "superscript"})
            rPr.append(va)
        t = el.makeelement(qn("w:t"), {})
        t.text = seg
        t.set(qn("xml:space"), "preserve")
        el.append(t)
        anchor.addnext(el)
        anchor = el
        made += 1
    parent.remove(run._r)
    return made


def isotopes(doc) -> int:
    n = 0
    for par in doc.paragraphs:
        for run in list(par.runs):
            n += _split_run(run)
    return n


# ── Word hand-off ───────────────────────────────────────────────────────────
def refresh_fields() -> int:
    """Let Word update the TOC and page-number fields, then save."""
    import win32com.client as win32

    word = win32.DispatchEx("Word.Application")
    word.Visible = False
    try:
        word.DisplayAlerts = 0
        doc = word.Documents.Open(str(DOCX), False, False, False)
        doc.Fields.Update()
        try:
            doc.TablesOfContents(1).Update()
            print("  TOC             : regenerated by Word")
        except Exception as exc:  # pragma: no cover - depends on the document
            print(f"  TOC             : Word could not refresh it ({type(exc).__name__})")
        for sec in doc.Sections:
            for idx in (1, 2, 3):
                for coll in (sec.Footers, sec.Headers):
                    try:
                        coll(idx).Range.Fields.Update()
                    except Exception:
                        pass
        doc.Repaginate()
        pages = doc.ComputeStatistics(2)
        doc.Save()
        doc.Close(False)
    finally:
        word.Quit()
    return int(pages)


def export_pdf() -> None:
    """Export without touching fields: the saved TOC already carries the numbers."""
    import win32com.client as win32

    tmp = Path(tempfile.gettempdir()) / f"aidd_layout_{int(__import__('time').time())}.pdf"
    word = win32.DispatchEx("Word.Application")
    word.Visible = False
    try:
        word.DisplayAlerts = 0
        doc = word.Documents.Open(str(DOCX), False, True, False)
        doc.Repaginate()
        print(f"  Word pages      : {doc.ComputeStatistics(2)}")
        doc.ExportAsFixedFormat(str(tmp), 17, False, 0, 0, 0, 0, 0,
                                True, True, 0, True, True, False)
        doc.Close(False)
    finally:
        word.Quit()
    shutil.copy2(tmp, PDF)
    tmp.unlink(missing_ok=True)
    print(f"  PDF exported    : {PDF.name} ({PDF.stat().st_size:,} bytes)")


# ── verification ────────────────────────────────────────────────────────────
def verify(doc) -> None:
    tbl = doc.tables[0]
    rows = list(tbl.rows)
    labels = [r.cells[0].text.strip().rstrip(":") for r in rows]
    for need in ("Student ID", "Student Name", "Date"):
        assert need in labels, f"cover label {need!r} missing"
    i_id, i_name, i_date = (labels.index(k) for k in ("Student ID", "Student Name", "Date"))
    assert i_name == i_id + 1, "ID and name are not adjacent"
    assert i_date == i_name + 1, "a filler row still separates name and date"
    assert rows[i_date].cells[1].text.strip() == COVER_DATE

    toc_heading = next(p for p in doc.paragraphs if p.text.strip() == TOC_TITLE)
    assert not toc_heading._p.findall(".//" + qn("w:u")), "TOC heading still underlined"

    leaders = [t for p in doc.paragraphs if p.style.name.lower().startswith("toc")
               for t in p._p.findall(".//" + qn("w:tab"))
               if t.get(qn("w:leader")) not in (None, "none")]
    assert not leaders, f"{len(leaders)} TOC tab stops still carry a leader"
    style_leaders = [t for st in doc.styles if (st.name or "").lower().startswith("toc")
                     for t in st.element.findall(".//" + qn("w:tab"))
                     if t.get(qn("w:leader")) not in (None, "none")]
    assert not style_leaders, f"{len(style_leaders)} TOC style tab stops still carry a leader"

    for par in doc.paragraphs:
        if par.style.name in ("Heading 1", "Heading 2"):
            assert par.paragraph_format.keep_with_next, f"heading without keep-next: {par.text[:40]}"

    loose = [p.text[:40] for p in doc.paragraphs
             if p.text.strip().startswith(("Table ", "Figure "))
             and p.text.strip() != TOC_TITLE
             and not (p.paragraph_format.keep_with_next and p.paragraph_format.keep_together)]
    assert not loose, f"{len(loose)} captions can still split or drift: {loose[:3]}"

    flat = [r.text for p in doc.paragraphs for r in p.runs
            if ISOTOPE.search(r.text) and not r.font.superscript]
    assert not flat, f"{len(flat)} runs still hold a flat isotope label: {flat[:4]}"
    sup = sum(1 for p in doc.paragraphs for r in p.runs if r.font.superscript)
    assert sup >= 29, f"only {sup} superscript runs"


def audit_toc(doc=None) -> None:
    """Every TOC page number must point at the page the heading actually sits on."""
    import fitz

    doc = doc or Document(str(DOCX))
    pdf = fitz.open(PDF)
    body = [re.sub(r"\s+", " ", p.get_text()).lower() for p in pdf]

    rows = []
    for par in doc.paragraphs:
        if not par.style.name.lower().startswith("toc"):
            continue
        text = par.text.strip()
        if "\t" not in text:
            continue
        title, _, page = text.rpartition("\t")
        title, page = title.strip(), page.strip()
        if len(title) < 5 or not page.isdigit():
            continue
        rows.append((title, int(page)))

    bad = []
    for title, page in rows:
        key = re.sub(r"\s+", " ", title)[:28].lower()
        hits = [i + 1 for i, txt in enumerate(body) if key in txt]
        if not hits:
            bad.append((title, page, "not found"))
        elif page not in hits and min(hits, key=lambda x: abs(x - page)) != page:
            bad.append((title, page, hits[:3]))
    print(f"  TOC audit       : {len(rows)} entries, {len(bad)} mismatched")
    for b in bad[:6]:
        print("     ", b)


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--check", action="store_true")
    ap.add_argument("--audit", action="store_true", help="only re-check the TOC pages of the PDF")
    ap.add_argument("--no-pdf", action="store_true")
    args = ap.parse_args()

    doc = Document(str(DOCX))
    if args.audit:
        audit_toc(doc)
        return
    if args.check:
        verify(doc)
        print("  layout state OK")
        return

    backup = Path(tempfile.gettempdir()) / "aidd_report_before_layout.docx"
    shutil.copy2(DOCX, backup)

    c = cover(doc)
    rule = strip_toc_rule(doc)
    lead = strip_leaders(doc)
    h = headings(doc)
    b = body(doc)
    cap = captions(doc)
    paths = path_lists(doc)
    iso = isotopes(doc)
    doc.save(str(DOCX))

    print(f"  cover           : date {c['date']}, {c['rows_removed']} filler rows dropped, "
          f"{c['spacer_lines']} spacer lines")
    print(f"  TOC             : {rule} underline runs removed, {lead} leaders -> plain tabs")
    print(f"  headings        : H1 {h['h1']} @ {H1[0]} pt, H2 {h['h2']} @ {H2[0]} pt, H3 {h['h3']}")
    print(f"  body            : {b['prose']} paragraphs @ 1.5 / no gap, {b['bullets']} bullets")
    print(f"  captions        : {cap} kept whole and with their float; {paths} path list left-aligned")
    print(f"  isotopes        : {iso} runs rewritten, "
          f"{sum(1 for p in doc.paragraphs for r in p.runs if r.font.superscript)} superscripts")

    pages = refresh_fields()
    print(f"  fields updated  : Word reports {pages} pages")

    doc2 = Document(str(DOCX))
    again = strip_leaders(doc2) + strip_leaders_in_styles(doc2)
    strip_toc_rule(doc2)
    typed = toc_type(doc2)
    styled = toc_styles(doc2)
    kept = keep_headings_with_next(doc2)
    doc2.save(str(DOCX))
    print(f"  post-refresh    : {again} leaders stripped (paragraph + style), "
          f"{typed} TOC entries retyped, {styled} TOC styles fixed, {kept} headings set keep-next")

    if not args.no_pdf:
        export_pdf()
        audit_toc()

    doc3 = Document(str(DOCX))
    verify(doc3)
    print(f"  backup          : {backup}")


if __name__ == "__main__":
    main()
