# -*- coding: utf-8 -*-
"""Rebuild the FAP background page in place from `24_restyle_mpu_deck.s03b`.

The deck is hand-edited outside this repository, so nothing else is touched: the
new page is appended at the end with python-pptx, PowerPoint then deletes the
existing page 3 and moves the new one into its place. Running it again simply
replaces page 3 again, which is what makes layout tweaks repeatable.

Usage:
    python src/31_fap_slide_layout.py            # rebuild in place
    python src/31_fap_slide_layout.py --check
    python src/31_fap_slide_layout.py --out X.pptx
"""
from __future__ import annotations

import argparse
import importlib.util
import shutil
import sys
import tempfile
from copy import deepcopy
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(Path(__file__).resolve().parent))

from deck_notes import NOTES  # noqa: E402
from pptx import Presentation  # noqa: E402
from pptx.oxml.ns import qn  # noqa: E402

PPTX = (ROOT / "AIDD8122_Final_Presentation_HuangYong"
        / "AIDD8122_Final_Presentation_HuangYong_10min.pptx")
TITLE = "FAP sits on the stroma, not on the tumour cell"
OLD_TITLE = "FAP is the fastest-moving small-molecule target in nuclear medicine"
POSITION = 2          # 0-based: third slide


def load_deck_module():
    path = Path(__file__).with_name("24_restyle_mpu_deck.py")
    spec = importlib.util.spec_from_file_location("d24", path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def inch(v) -> float:
    return 0.0 if v is None else v / 914400


def title_of(slide) -> str:
    for shape in slide.shapes:
        if shape.has_text_frame and inch(shape.top) < 0.70 and inch(shape.width) > 8:
            text = shape.text_frame.text.strip()
            if len(text) > 8:
                return text
    return ""


def build_stage1() -> Path:
    """Add the new page at the end with python-pptx (no deletion, no reordering)."""
    prs = Presentation(str(PPTX))
    module = load_deck_module()
    module.s03b(prs, prs.slide_layouts[6])
    out = Path(tempfile.gettempdir()) / "aidd_fap_stage1.pptx"
    prs.save(str(out))
    return out


def com_reorder(stage1: Path) -> Path:
    """Delete the two superseded pages and move the new one to position 3.

    python-pptx cannot drop a slide without disturbing the r:id map, so the
    destructive part is handed to PowerPoint itself.
    """
    import win32com.client as win32

    out = Path(tempfile.gettempdir()) / "aidd_fap_stage2.pptx"
    app = win32.DispatchEx("PowerPoint.Application")
    try:
        pres = app.Presentations.Open(str(stage1), WithWindow=False)
        pres.Slides(3).Delete()      # the page being replaced
        n = pres.Slides.Count
        pres.Slides(n).MoveTo(3)     # new page to position 3
        pres.SaveAs(str(out), 24)    # ppSaveAsOpenXMLPresentation
        pres.Close()
    finally:
        app.Quit()
    return out


def ensure_notes_frame(prs, slide):
    notes = slide.notes_slide
    if notes.notes_text_frame is not None:
        return notes.notes_text_frame
    for other in prs.slides:
        if other is slide:
            continue
        for ph in other.notes_slide.placeholders:
            if ph.placeholder_format.idx == 3:
                notes.shapes._spTree.append(deepcopy(ph._element))
                return notes.notes_text_frame
    raise RuntimeError("no notes body placeholder to clone")


def renumber(prs) -> int:
    fixed = 0
    for idx, slide in enumerate(prs.slides, 1):
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            text = shape.text_frame.text.strip()
            if (inch(shape.left) > 11.4 and inch(shape.top) > 6.9
                    and text.isdigit() and int(text) != idx):
                shape.text_frame.text = str(idx)
                fixed += 1
    return fixed


def write_notes(prs) -> None:
    assert len(prs.slides) == len(NOTES), f"{len(prs.slides)} slides vs {len(NOTES)} notes"
    for i, slide in enumerate(prs.slides, 1):
        ensure_notes_frame(prs, slide).text = NOTES[i].strip()


def rebuild() -> Path:
    stage1 = build_stage1()
    stage2 = com_reorder(stage1)
    print(f'  stage1         : {stage1.name} (page appended)')
    print(f'  stage2         : {stage2.name} (old pages removed, new one moved)')
    return stage2


def check(prs) -> None:
    assert len(prs.slides) == 13, f"expected 13 slides, got {len(prs.slides)}"
    titles = [title_of(s) for s in prs.slides]
    assert titles[2] == TITLE, f"slide 3 title is {titles[2]!r}"
    assert OLD_TITLE not in titles, "old FAP page still present"
    pics = sum(1 for sh in prs.slides[2].shapes if sh.shape_type == 13)
    assert pics == 3, f"slide 3 should hold 3 pictures (figure, MIP, crest), got {pics}"
    bad = []
    for idx, slide in enumerate(prs.slides, 1):
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            t = shape.text_frame.text.strip()
            if inch(shape.left) > 11.4 and inch(shape.top) > 6.9 and t.isdigit():
                if int(t) != idx:
                    bad.append((idx, t))
    assert not bad, f"page numbers out of step: {bad}"
    for i, slide in enumerate(prs.slides, 1):
        note = slide.notes_slide.notes_text_frame.text.strip() if slide.has_notes_slide else ""
        assert note == NOTES[i].strip(), f"notes mismatch on slide {i}"
        for token in ("［", "］", "分钟", "小时"):
            assert token not in note, f"slide {i} notes contain {token!r}"
    print(f"  slides         : 13, FAP page at position 3")
    print(f"  pictures on p3 : {pics}")
    print(f"  notes          : 13/13, no timestamps")


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--check", action="store_true")
    ap.add_argument("--out", type=Path)
    args = ap.parse_args()

    prs = Presentation(str(PPTX))
    if args.check:
        check(prs)
        return

    backup = Path(tempfile.gettempdir()) / "aidd_deck_before_fap_rebuild.pptx"
    shutil.copy2(PPTX, backup)

    prs = Presentation(str(rebuild()))
    fixed = renumber(prs)
    write_notes(prs)

    out = args.out or PPTX
    try:
        prs.save(str(out))
    except PermissionError:
        raise SystemExit(f"ERROR: {out} is locked; close PowerPoint first "
                         f"(backup at {backup})")
    print(f"  page numbers fixed: {fixed}")
    print(f"  saved          : {out}")
    print(f"  backup         : {backup}")
    check(Presentation(str(out)))


if __name__ == "__main__":
    main()
