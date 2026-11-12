# -*- coding: utf-8 -*-
"""Insert the FAP background slide into the hand-edited 10-minute deck.

Adds one page after slide 3 (what FAP is, mechanism schematic + a real FAPI
PET image), renumbers the page footers, and rewrites all speaker notes from
``deck_notes.py`` (now 15 entries). Nothing else on the deck is touched.

Usage:
    python src/30_add_fap_slide.py            # insert in place (backs up first)
    python src/30_add_fap_slide.py --check
    python src/30_add_fap_slide.py --out X.pptx
"""
from __future__ import annotations

import argparse
import importlib.util
import shutil
import sys
from copy import deepcopy
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(Path(__file__).resolve().parent))

from deck_notes import NOTES  # noqa: E402
from pptx import Presentation  # noqa: E402

PPTX = (ROOT / "AIDD8122_Final_Presentation_HuangYong"
        / "AIDD8122_Final_Presentation_HuangYong_10min.pptx")
TITLE = "FAP sits on the stroma, not on the tumour cell"
INSET = 3          # 0-based position of the new page (after slide 3)


def load_deck_module():
    path = Path(__file__).with_name("24_restyle_mpu_deck.py")
    spec = importlib.util.spec_from_file_location("d24", path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def inch(v) -> float:
    return 0.0 if v is None else v / 914400


def slide_title(slide) -> str:
    best, size = "", -1.0
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        text = shape.text_frame.text.strip()
        if len(text) > 8 and inch(shape.top) < 0.70 and inch(shape.width) > 8:
            if shape.width > size:
                best, size = text, shape.width
    return best


def renumber(prs) -> int:
    fixed = 0
    for idx, slide in enumerate(prs.slides, 1):
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            text = shape.text_frame.text.strip()
            if (inch(shape.left) > 11.4 and inch(shape.top) > 6.9 and text.isdigit()
                    and int(text) != idx):
                tf = shape.text_frame
                tf.paragraphs[0].runs[0].text = str(idx)
                for run in tf.paragraphs[0].runs[1:]:
                    run.text = ""
                fixed += 1
    return fixed


def ensure_notes_frame(prs, slide):
    """A freshly inserted slide gets an empty notes page on this file.

    Its notes master exposes no body placeholder, so python-pptx cannot create
    one. Clone the body placeholder from a slide that already has notes; the
    text is overwritten immediately afterwards.
    """
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
    raise RuntimeError("no notes body placeholder available to clone")


def write_notes(prs) -> None:
    assert len(prs.slides) == len(NOTES), f"{len(prs.slides)} slides vs {len(NOTES)} notes"
    for i, slide in enumerate(prs.slides, 1):
        ensure_notes_frame(prs, slide).text = NOTES[i].strip()


def add_slide(prs, module) -> bool:
    if any(slide_title(s) == TITLE for s in prs.slides):
        print("  already present; skipping insertion")
        return False
    module.s03b(prs, prs.slide_layouts[6])
    lst = prs.slides._sldIdLst
    ids = list(lst)
    new = ids[-1]
    lst.remove(new)
    lst.insert(INSET, new)
    return True


def check(prs) -> None:
    assert len(prs.slides) == 15, f"expected 15 slides, got {len(prs.slides)}"
    titles = [slide_title(s) for s in prs.slides]
    assert titles[3] == TITLE, f"slide 4 title is {titles[3]!r}"
    s4 = prs.slides[3]
    pics = sum(1 for sh in s4.shapes if sh.shape_type == 13)
    assert pics == 3, f"slide 4 should hold 3 pictures (crest + 2 figures), got {pics}"
    # page numbers must match physical order
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
    print(f"  slides         : 15 (new page at position 4)")
    print(f"  pictures on p4 : {pics}")
    print(f"  page numbers   : 1-14 in order")
    print(f"  notes          : 15/15, no timestamps")


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--check", action="store_true")
    ap.add_argument("--out", type=Path)
    args = ap.parse_args()

    prs = Presentation(str(PPTX))
    if args.check:
        check(prs)
        return

    backup = Path(tempfile.gettempdir()) / "aidd_deck_before_fap_slide.pptx"
    shutil.copy2(PPTX, backup)

    module = load_deck_module()
    inserted = add_slide(prs, module)
    fixed = renumber(prs)
    write_notes(prs)

    out = args.out or PPTX
    try:
        prs.save(str(out))
    except PermissionError:
        raise SystemExit(f"ERROR: {out} is locked; close PowerPoint first "
                         f"(backup at {backup})")
    print(f"  inserted       : {inserted} | page numbers fixed: {fixed}")
    print(f"  saved          : {out}")
    print(f"  backup         : {backup}")
    check(Presentation(str(out)))


if __name__ == "__main__":
    main()
