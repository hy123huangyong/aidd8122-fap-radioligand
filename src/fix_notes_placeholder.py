# -*- coding: utf-8 -*-
"""Rebuild speaker notes as one proper body placeholder per slide.

The earlier pass left plain text boxes, and a second pass left duplicates with
the same shape id. PowerPoint's Notes pane only reads the notes body
placeholder, so each notes slide is now rebuilt from scratch: every shape is
removed and a single ``p:ph type="body" idx="1"`` shape is written with the
note text from 10_add_notes.py.
"""
import importlib.util
import sys
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.oxml.ns import qn
from pptx.util import Emu, Pt

HERE = Path(__file__).resolve().parent
spec = importlib.util.spec_from_file_location("notes_src", HERE / "10_add_notes.py")
notes_src = importlib.util.module_from_spec(spec)
spec.loader.exec_module(notes_src)
NOTES = notes_src.NOTES

deck = Path(sys.argv[1])
prs = Presentation(str(deck))

A = "http://schemas.openxmlformats.org/drawingml/2006/main"


def build_notes_shape(shape_id: int, text: str):
    sp = etree.Element(qn("p:sp"))
    nv = etree.SubElement(sp, qn("p:nvSpPr"))
    cnv = etree.SubElement(nv, qn("p:cNvPr"))
    cnv.set("id", str(shape_id))
    cnv.set("name", "Notes Placeholder %d" % shape_id)
    cnv_sp = etree.SubElement(nv, qn("p:cNvSpPr"))
    etree.SubElement(cnv_sp, qn("a:spLocks")).set("noGrp", "1")
    nvpr = etree.SubElement(nv, qn("p:nvPr"))
    ph = etree.SubElement(nvpr, qn("p:ph"))
    ph.set("type", "body")
    ph.set("idx", "1")

    sppr = etree.SubElement(sp, qn("p:spPr"))
    xfrm = etree.SubElement(sppr, qn("a:xfrm"))
    off = etree.SubElement(xfrm, qn("a:off"))
    off.set("x", "685800")
    off.set("y", "4343400")
    ext = etree.SubElement(xfrm, qn("a:ext"))
    ext.set("cx", "5486400")
    ext.set("cy", "4800600")
    prst = etree.SubElement(sppr, qn("a:prstGeom"))
    prst.set("prst", "rect")
    etree.SubElement(prst, qn("a:avLst"))

    body = etree.SubElement(sp, qn("p:txBody"))
    bp = etree.SubElement(body, qn("a:bodyPr"))
    bp.set("wrap", "square")
    etree.SubElement(body, qn("a:lstStyle"))
    for line in text.splitlines():
        p = etree.SubElement(body, qn("a:p"))
        r = etree.SubElement(p, qn("a:r"))
        rpr = etree.SubElement(r, qn("a:rPr"))
        rpr.set("lang", "en-GB")
        rpr.set("sz", "1100")
        fill = etree.SubElement(rpr, qn("a:solidFill"))
        clr = etree.SubElement(fill, qn("a:srgbClr"))
        clr.set("val", "222222")
        etree.SubElement(rpr, qn("a:latin")).set("typeface", "Consolas")
        t = etree.SubElement(r, qn("a:t"))
        t.text = line
    return sp


rebuilt = 0
for i, slide in enumerate(prs.slides):
    if i >= len(NOTES):
        break
    ns = slide.notes_slide
    tree = ns._element.find(qn("p:cSld")).find(qn("p:spTree"))
    for child in list(tree):
        if child.tag == qn("p:sp"):
            tree.remove(child)
    tree.append(build_notes_shape(i + 901, NOTES[i].strip()))
    rebuilt += 1

prs.save(str(deck))
print(f"rebuilt notes on {rebuilt} slides -> {deck}")
