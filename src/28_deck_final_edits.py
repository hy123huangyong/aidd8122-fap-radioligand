# -*- coding: utf-8 -*-
"""Final content edits applied in place to the hand-edited 10-minute deck.

Two changes only, on the file as it stands:

1. Slide 13 keeps the limitations and drops the "next steps" column. The four
   limitation cards are widened to the full text width, the heading loses
   "— and what I would do next", and the summary bar stops promising a fix.
2. Every speaker note is rewritten from ``deck_notes.py``, which no longer
   carries timestamps or any mention of duration.

Nothing else is touched: no slide is rebuilt, no layout, colour or picture is
re-created, so manual edits already in the file survive.

Usage:
    python src/28_deck_final_edits.py            # edit in place (backs up first)
    python src/28_deck_final_edits.py --check    # verify only
    python src/28_deck_final_edits.py --out X.pptx
"""
from __future__ import annotations

import argparse
import shutil
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(Path(__file__).resolve().parent))

from deck_notes import NOTES  # noqa: E402
from pptx import Presentation  # noqa: E402
from pptx.enum.shapes import MSO_SHAPE_TYPE  # noqa: E402
from pptx.util import Inches  # noqa: E402

PPTX = (ROOT / "AIDD8122_Final_Presentation_HuangYong"
        / "AIDD8122_Final_Presentation_HuangYong_10min.pptx")

TITLE_OLD = "Limitations — and what I would do next"
TITLE_NEW = "Limitations"
SUB_OLD = "Where the evidence stops, and the cheapest experiment that would move it"
SUB_NEW = "Where the evidence stops, and what this study cannot claim"
SUMMARY_OLD = "The fix is a consistent assay campaign, not a better architecture."
SUMMARY_NEW = ("The evidence stops at the applicability domain; "
               "beyond it, nothing in this study is tested.")

# geometry of the old two-column layout (inches)
LEFT_MAX = 6.5          # anything left of this belongs to the limitations column
CARD_W, TEXT_W = 5.79, 5.25
CARD_NEW, TEXT_NEW = 12.40, 11.86


def set_text(shape, text: str) -> None:
    """Replace the paragraph text but keep the first run's formatting."""
    tf = shape.text_frame
    p = tf.paragraphs[0]
    if p.runs:
        p.runs[0].text = text
        for run in p.runs[1:]:
            run.text = ""
    else:
        p.text = text
    for extra in list(tf.paragraphs[1:]):
        extra._p.getparent().remove(extra._p)


def _geom(shape):
    """Position/size in inches (shape.left is already an EMU Length)."""
    def inch(v):
        return 0.0 if v is None else v / 914400
    return inch(shape.left), inch(shape.top), inch(shape.width)


def edit_slide13(slide) -> list[str]:
    log: list[str] = []
    seen = {"title": False, "sub": False, "summary": False}

    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        text = shape.text_frame.text.strip()
        if text == TITLE_OLD:
            set_text(shape, TITLE_NEW); seen["title"] = True; log.append("title")
        elif text == TITLE_NEW:
            seen["title"] = True
        elif text == SUB_OLD:
            set_text(shape, SUB_NEW); seen["sub"] = True; log.append("subtitle")
        elif text == SUB_NEW:
            seen["sub"] = True
        elif text == SUMMARY_OLD:
            set_text(shape, SUMMARY_NEW); seen["summary"] = True; log.append("summary")
        elif text == SUMMARY_NEW:
            seen["summary"] = True

    # widen the four limitation cards to the full width
    widened = 0
    for shape in slide.shapes:
        left, top, width = _geom(shape)
        if left >= LEFT_MAX or top <= 1.7:
            continue
        if CARD_W - 0.35 < width < CARD_W + 0.35:
            shape.width = Inches(CARD_NEW); widened += 1
        elif TEXT_W - 0.35 < width < TEXT_W + 0.35:
            shape.width = Inches(TEXT_NEW); widened += 1
    if widened:
        log.append(f"widened {widened}")

    # re-space the four cards evenly: 0.72 high, 0.393 apart, bottom at 5.94
    cards, bars, texts = [], [], []
    for shape in slide.shapes:
        left, top, width = _geom(shape)
        if left >= LEFT_MAX or not (1.7 < top < 6.0):   # card band only, not the summary bar
            continue
        if shape.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE:
            (cards if width > 1.0 else bars).append((top, shape))
        elif shape.shape_type == MSO_SHAPE_TYPE.TEXT_BOX:
            texts.append((top, shape))
    if len(cards) == len(bars) == len(texts) == 4:
        for i in range(4):
            y = 1.88 + i * 1.113
            card = cards[i][1]
            card.top, card.height = Inches(y), Inches(0.72)
            bar = bars[i][1]
            bar.top, bar.height = Inches(y + 0.13), Inches(0.46)
            body = texts[i][1]
            body.top, body.height = Inches(y + 0.19), Inches(0.34)
        log.append("respace 4")

    # drop the right-hand column (heading, cards, accent bars, body text)
    dropped = 0
    for shape in list(slide.shapes):
        left, top, _ = _geom(shape)
        if left > LEFT_MAX and 1.4 < top < 6.2:
            shape._element.getparent().remove(shape._element)
            dropped += 1
    if dropped:
        log.append(f"removed {dropped}")

    missing = [k for k, v in seen.items() if not v]
    assert not missing, f"slide 13 anchors not found: {missing}"
    return log


def write_notes(prs) -> None:
    assert len(prs.slides) == len(NOTES) == 14, "slide/notes count mismatch"
    for i, slide in enumerate(prs.slides, 1):
        slide.notes_slide.notes_text_frame.text = NOTES[i].strip()


def check(prs) -> None:
    s13 = prs.slides[12]
    texts = [sh.text_frame.text.strip() for sh in s13.shapes if sh.has_text_frame]
    assert TITLE_NEW in texts, "slide 13 title not updated"
    assert not any(t == TITLE_OLD for t in texts), "old title still present"
    assert not any(t == "Next steps" for t in texts), "Next steps heading still present"
    right = [sh for sh in s13.shapes
             if _geom(sh)[0] > LEFT_MAX and 1.4 < _geom(sh)[1] < 6.2]
    assert not right, f"{len(right)} right-column shapes remain"
    for i, slide in enumerate(prs.slides, 1):
        got = slide.notes_slide.notes_text_frame.text.strip() if slide.has_notes_slide else ""
        assert got == NOTES[i].strip(), f"notes mismatch on slide {i}"
        for bad in ("［", "］", "分钟", "小时"):
            assert bad not in got, f"slide {i} notes still contain {bad!r}"
    print(f"  slide 13       : limitations only ({len(s13.shapes)} shapes)")
    print(f"  notes          : 14/14 rewritten, 0 timestamps")
    print(f"  words          : {sum(len(n) for n in NOTES.values())} chars")


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--check", action="store_true")
    ap.add_argument("--out", type=Path, help="write here instead of overwriting")
    args = ap.parse_args()

    prs = Presentation(str(PPTX))
    if args.check:
        check(prs)
        return

    backup = Path(tempfile.gettempdir()) / "aidd_deck_before_s13_edit.pptx"
    shutil.copy2(PPTX, backup)

    log = edit_slide13(prs.slides[12])
    write_notes(prs)

    out = args.out or PPTX
    try:
        prs.save(str(out))
    except PermissionError:
        raise SystemExit(
            f"ERROR: {out} is locked (close PowerPoint, then re-run). "
            f"A backup of the input is at {backup}")
    print(f"  edited         : {', '.join(log)}")
    print(f"  saved          : {out}")
    print(f"  backup         : {backup}")
    check(Presentation(str(out)))


if __name__ == "__main__":
    main()
