"""Restyle the 10-minute deck to the visual language of the group-meeting deck.

Reference: PSMA-FAP_组会汇报20261029.pptx
  - white pages, no full-width colour band
  - navy "▶" + bold navy title at top left, grey section label at right
  - 1 pt navy hairline under the header
  - tinted cards with a thin border
  - Arial throughout
"""
from __future__ import annotations

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Inches, Pt

SRC = "Submission_HuangYong_P2600222/02_Presentation/AIDD8122_Final_Presentation_HuangYong_10min.pptx"
DST = "Submission_HuangYong_P2600222/02_Presentation/AIDD8122_Final_Presentation_HuangYong_10min.pptx"

NAVY = RGBColor(0x00, 0x20, 0x60)
GREY = RGBColor(0x8B, 0x97, 0xA8)
TINT = RGBColor(0xEA, 0xF1, 0xF9)
LINE = RGBColor(0xD6, 0xDC, 0xE5)
OLD_NAVY = RGBColor(0x1E, 0x4F, 0xA8)


def is_band(shape) -> bool:
    try:
        return (shape.top == 0 and shape.width > Inches(13) and shape.height < Inches(1)
                and str(shape.fill.fore_color.rgb) == "1E4FA8")
    except Exception:
        return False


def restyle_run(run) -> None:
    run.font.name = "Arial"
    try:
        if run.font.color and run.font.color.type is not None:
            if str(run.font.color.rgb) == "1E4FA8":
                run.font.color.rgb = NAVY
            elif str(run.font.color.rgb) == "C9D9F2":
                run.font.color.rgb = GREY
    except Exception:
        pass


def main() -> None:
    prs = Presentation(SRC)

    for idx, slide in enumerate(prs.slides, start=1):
        band = None
        for shape in slide.shapes:
            if is_band(shape):
                band = shape
                break

        header_page = band is not None
        if band is not None:
            band._element.getparent().remove(band._element)

        for shape in slide.shapes:
            if shape.has_text_frame:
                for par in shape.text_frame.paragraphs:
                    for run in par.runs:
                        restyle_run(run)

            # tinted cards -> reference tint + hairline border
            try:
                if str(shape.fill.fore_color.rgb) == "E8EFF8":
                    shape.fill.solid()
                    shape.fill.fore_color.rgb = TINT
                    shape.line.color.rgb = LINE
                    shape.line.width = Pt(0.75)
            except Exception:
                pass

        if not header_page:
            continue

        # title: navy, prefixed with the reference deck's marker
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            if shape.top is not None and shape.top < Inches(0.62) and shape.left < Inches(1):
                frame = shape.text_frame
                if frame.text.strip() and not frame.text.startswith("▶"):
                    run = frame.paragraphs[0].runs[0]
                    run.text = "▶  " + run.text
                    run.font.size = Pt(18.5)
                    run.font.bold = True
                    run.font.color.rgb = NAVY
                    frame.word_wrap = True
                    shape.width = Inches(10.9)

        # hairline under the header
        line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, Inches(0.70),
                                      prs.slide_width, Pt(1.1))
        line.fill.solid()
        line.fill.fore_color.rgb = NAVY
        line.line.fill.background()
        line.shadow.inherit = False

    # outline copy: the slot is now ten minutes
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame and "Ten to fifteen minutes" in shape.text_frame.text:
                for par in shape.text_frame.paragraphs:
                    for run in par.runs:
                        run.text = run.text.replace("Ten to fifteen minutes", "Ten minutes")

    prs.save(DST)
    print("saved", DST)


if __name__ == "__main__":
    main()
