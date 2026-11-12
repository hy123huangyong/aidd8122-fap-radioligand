# -*- coding: utf-8 -*-
"""Align the hand-edited 13-slide deck with its speaker notes.

The deck on the Desktop is the version the speaker actually owns: 13 slides and
a rewritten plain-Chinese script. This pass does not restyle it. It fixes only
where the printed text contradicts what the notes say:

  * cover date 24 -> 16 November (the submission date on the last slide)
  * outline 05 no longer promises a "next steps" part that the deck dropped
  * outline 01 describes the FAP page that is actually shown
  * slide 3 gives FAP its full name (the notes open with it)
  * slide 7 counts four regressors plus the mean baseline, not "five models"
  * slide 10 carries the 61% / 100% gate numbers the notes quote
  * footers 10-12 were left behind by the deleted slide, now renumbered
  * slide 12 drops the "Limitations" label that repeats the slide title

Usage:
    python src/32_desktop_deck_polish.py --check
    python src/32_desktop_deck_polish.py            # writes in place
"""
from __future__ import annotations

import argparse
import shutil
import tempfile
from copy import deepcopy
from pathlib import Path

from pptx import Presentation
from pptx.oxml.ns import qn

DECK = Path("C:/Users/13260/Desktop/AIDD8122_Final_Presentation_HuangYong_10min.pptx")

# (slide number, text that must be present, replacement) - one paragraph each
REPLACE = [
    (2, "FAP radiopharmaceuticals and what limits the current ligands",
     "FAP and FAPI PET: where the target sits, why the contrast is clean"),
    (2, "Take-home messages, limitations, what I would do next",
     "Take-home messages and where the evidence stops"),
    (2, "Benchmark, selectivity, uncertainty, temporal drift, design campaigns",
     "Benchmark, selectivity, uncertainty, design campaigns"),
    (3, "Where the target is expressed, what the tracer binds, and why the contrast is so clean",
     "Fibroblast activation protein (FAP): where it is expressed, what the tracer binds, "
     "why the contrast is clean"),
    (7, "Five models × three protocols: interpolation, scaffold hopping and a prospective time split",
     "Four regressors plus a mean baseline, under three protocols: interpolation, "
     "scaffold hopping, prospective prediction"),
]

# extra line appended to the generative cards, so the slide carries the numbers
# the notes quote (61% of novel molecules pass; affinity-only passes 100%)
EXTRA_LINES = [
    (10, "median MW 413 Da, SA 2.93 — the most usable chemistry",
     "61% pass the pFAP ≥ 8 / margin ≥ 1 gate"),
    (10, "median MW 548 Da — the model is exploited, chemistry degenerates",
     "100% pass the same gate"),
]

DROP_LABEL = (12, "Limitations")   # repeats the slide title one inch below it


def inch(v) -> float:
    return 0.0 if v is None else v / 914400


def shapes_with(slide, needle):
    for sh in slide.shapes:
        if sh.has_text_frame and needle in sh.text_frame.text:
            yield sh


def set_para(para, text: str) -> None:
    """Replace a paragraph's text without losing its run formatting."""
    runs = para.runs
    if not runs:                      # empty paragraph: just write into it
        para.text = text
        return
    runs[0].text = text
    for extra in runs[1:]:
        extra._r.getparent().remove(extra._r)


def append_para(shape, text: str) -> None:
    """Add a paragraph that inherits the first one's paragraph and run formatting."""
    first = shape.text_frame.paragraphs[0]._p
    new = deepcopy(first)
    runs = new.findall(qn("a:r"))
    for extra in runs[1:]:
        new.remove(extra)
    if runs:
        runs[0].find(qn("a:t")).text = text
    shape.text_frame._txBody.append(new)


def renumber(prs) -> list[tuple[int, str, int]]:
    fixed = []
    for idx, slide in enumerate(prs.slides, 1):
        for sh in slide.shapes:
            if not sh.has_text_frame:
                continue
            text = sh.text_frame.text.strip()
            if (inch(sh.left) > 11.8 and inch(sh.top) > 6.9
                    and text.isdigit() and int(text) != idx):
                old = text
                set_para(sh.text_frame.paragraphs[0], str(idx))
                fixed.append((idx, old, idx))
    return fixed


def apply(prs) -> list[str]:
    log = []
    for slide_no, old, new in REPLACE:
        slide = prs.slides[slide_no - 1]
        hits = list(shapes_with(slide, old))
        if not hits:
            if list(shapes_with(slide, new)):
                log.append(f"slide {slide_no:2d}: already applied")
                continue
        assert len(hits) == 1, f"slide {slide_no}: {len(hits)} matches for {old[:40]!r}"
        set_para(hits[0].text_frame.paragraphs[0], new)
        log.append(f"slide {slide_no:2d}: replaced {old[:44]!r}")

    for slide_no, anchor, extra in EXTRA_LINES:
        slide = prs.slides[slide_no - 1]
        hits = list(shapes_with(slide, anchor))
        assert len(hits) == 1, f"slide {slide_no}: {len(hits)} matches for {anchor[:40]!r}"
        if extra in hits[0].text_frame.text:
            log.append(f"slide {slide_no:2d}: already applied {extra!r}")
            continue
        append_para(hits[0], extra)
        log.append(f"slide {slide_no:2d}: appended {extra!r}")

    slide = prs.slides[DROP_LABEL[0] - 1]
    hits = [sh for sh in shapes_with(slide, DROP_LABEL[1])
            if inch(sh.top) > 1.2 and len(sh.text_frame.text.strip()) < 40]
    if hits:
        assert len(hits) == 1, f"slide {DROP_LABEL[0]}: label not unique"
        hits[0]._element.getparent().remove(hits[0]._element)
        log.append(f"slide {DROP_LABEL[0]:2d}: removed duplicate label 'Limitations'")
    else:
        log.append(f"slide {DROP_LABEL[0]:2d}: label already removed")

    for idx, old, new in renumber(prs):
        log.append(f"slide {idx:2d}: footer {old} -> {new}")
    return log


def cover_date(prs) -> str:
    slide = prs.slides[0]
    for sh in slide.shapes:
        if sh.has_text_frame and "November 2026" in sh.text_frame.text:
            para = sh.text_frame.paragraphs[0]
            before = para.text
            if para.runs:
                para.runs[0].text = "16"
            return f"cover date: {before!r} -> {para.text!r}"
    raise AssertionError("cover date not found")


def check(prs) -> None:
    texts = ["\n".join(sh.text_frame.text for sh in s.shapes if sh.has_text_frame)
             for s in prs.slides]
    assert "24 November 2026" not in texts[0], "cover still shows 24 November"
    assert "16 November 2026" in texts[0], "cover date not updated"
    assert "what I would do next" not in texts[1], "outline still promises next steps"
    assert "temporal drift" not in texts[1], "outline still promises a drift slide"
    assert "Fibroblast activation protein (FAP)" in texts[2], "FAP full name missing"
    assert "Five models" not in texts[6], "slide 7 still says five models"
    assert "mean baseline" in texts[6], "slide 7 lacks the baseline wording"
    assert "61%" in texts[9] and "100%" in texts[9], "gate numbers missing on slide 10"
    labels = [sh.text_frame.text.strip() for sh in prs.slides[11].shapes
              if sh.has_text_frame and sh.text_frame.text.strip() == "Limitations"]
    assert len(labels) == 1, f"slide 12 still carries {len(labels)} 'Limitations' labels"
    for idx, slide in enumerate(prs.slides, 1):
        for sh in slide.shapes:
            if not sh.has_text_frame:
                continue
            t = sh.text_frame.text.strip()
            if inch(sh.left) > 11.8 and inch(sh.top) > 6.9 and t.isdigit():
                assert int(t) == idx, f"slide {idx} footer says {t}"
    print("  check          : all text and footer assertions passed")


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--check", action="store_true")
    args = ap.parse_args()

    prs = Presentation(str(DECK))
    if args.check:
        check(prs)
        return

    backup = Path(tempfile.gettempdir()) / "deck_desktop_before_polish.pptx"
    shutil.copy2(DECK, backup)

    print(" ", cover_date(prs))
    for line in apply(prs):
        print(" ", line)
    prs.save(str(DECK))
    print(f"  saved          : {DECK}")
    print(f"  backup         : {backup}")
    check(Presentation(str(DECK)))


if __name__ == "__main__":
    main()
