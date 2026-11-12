"""Rebuild the 10-minute AIDD8122 deck in the MPU presentation style.

Style reference: ``AIDD8124_Presentation_MPUstyle_HUANG_Yong_CHEN_Wan.pptx``
(course 5).  Its design system is reproduced here as literal constants so the
deck can be regenerated from source: navy #002060 + teal #029BAB, a two-stripe
top rule, the MPU crest at the top right, a play-triangle title with subtitle
and hairline divider, pastel numbered cards, a dark "Summary" bar and a small
page number.  Cover and closing slides use the same navy band / full-bleed
treatment.

All 14 slides, all seven figures and every published number are preserved.
Speaker notes come from ``src/deck_notes.py``.

Usage:
    python src/24_restyle_mpu_deck.py            # rebuild pptx
    python src/24_restyle_mpu_deck.py --pdf      # rebuild pptx + export PDF
    python src/24_restyle_mpu_deck.py --check    # report structure only
"""
from __future__ import annotations

import argparse
import shutil
import sys
import tempfile
import time
from pathlib import Path

from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

ROOT = Path(__file__).resolve().parents[1]
ASSETS = ROOT / "assets"
FIGS = ROOT / "results" / "figures"
DECK_DIR = ROOT / "AIDD8122_Final_Presentation_HuangYong"
PPTX = DECK_DIR / "AIDD8122_Final_Presentation_HuangYong_10min.pptx"
PDF = DECK_DIR / "AIDD8122_Final_Presentation_HuangYong_10min.pdf"

sys.path.insert(0, str(Path(__file__).resolve().parent))
from deck_notes import NOTES  # noqa: E402

W, H = 13.333, 7.5

NAVY = "002060"       # primary
TEAL = "029BAB"       # accent
INK = "1A2230"        # body ink / summary bar
MUTED = "4A5568"
GREY = "8B97A8"
RULE = "D6DCE5"
WHITE = "FFFFFF"
PALE = "F7F9FC"
BLUE = "EAF1F9"
CREAM = "FDF3E7"
MINT = "E8F5F7"
BADGE = "1B4B8F"
AMBER = "B45309"
LIGHT = "7FA6D9"
PALEBLUE = "C7D3E6"
CREDIT = "AFC6EA"
RED = "B03A2E"

FONT = "Arial"

CREST = ASSETS / "mpu_round_logo.png"
WORDMARK = ASSETS / "mpu_wordmark.png"


def I(v: float):
    return Inches(v)


def C(hexstr: str) -> RGBColor:
    return RGBColor.from_string(hexstr)


def rect(slide, x, y, w, h, color):
    sh = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, I(x), I(y), I(w), I(h))
    sh.fill.solid()
    sh.fill.fore_color.rgb = C(color)
    sh.line.fill.background()
    try:
        sh.shadow.inherit = False
    except Exception:
        pass
    return sh


def txt(slide, x, y, w, h, text, size=11, color=INK, bold=False,
        align="left", anchor="top", leading=1.0, italic=False):
    sh = slide.shapes.add_textbox(I(x), I(y), I(w), I(h))
    tf = sh.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = {"top": MSO_ANCHOR.TOP,
                          "middle": MSO_ANCHOR.MIDDLE,
                          "bottom": MSO_ANCHOR.BOTTOM}[anchor]
    al = {"left": PP_ALIGN.LEFT, "right": PP_ALIGN.RIGHT,
          "center": PP_ALIGN.CENTER}[align]
    for i, line in enumerate(str(text).split("\n")):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = al
        p.line_spacing = leading
        p.space_after = Pt(0)
        r = p.add_run()
        r.text = line
        r.font.name = FONT
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.italic = italic
        r.font.color.rgb = C(color)
    return sh


def lines(slide, x, y, w, h, items, size=10.5, color=MUTED, gap=5,
          leading=1.12):
    sh = slide.shapes.add_textbox(I(x), I(y), I(w), I(h))
    tf = sh.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    for i, line in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.line_spacing = leading
        p.space_after = Pt(gap)
        r = p.add_run()
        r.text = line
        r.font.name = FONT
        r.font.size = Pt(size)
        r.font.color.rgb = C(color)
    return sh


def pic(slide, path, x, y, w):
    path = Path(path)
    with Image.open(path) as im:
        h = w * im.height / im.width
    slide.shapes.add_picture(str(path), I(x), I(y), width=I(w), height=I(h))
    return h


def new_slide(prs, blank):
    s = prs.slides.add_slide(blank)
    rect(s, 0, 0, W, H, WHITE)
    return s


def stripes(slide):
    rect(slide, 0, 0, W, 0.05, NAVY)
    rect(slide, 0, 0.05, W, 0.02, TEAL)


def crest(slide):
    slide.shapes.add_picture(str(CREST), I(12.04), I(0.17),
                             width=I(0.88), height=I(0.88))


def content_slide(prs, blank, title, subtitle, page):
    s = new_slide(prs, blank)
    stripes(s)
    crest(s)
    txt(s, 0.42, 0.55, 0.20, 0.24, "▶", 12, TEAL)
    txt(s, 0.70, 0.47, 11.15, 0.44, title, 22.5, NAVY, True)
    txt(s, 0.42, 0.94, 12.50, 0.26, subtitle, 12.5, MUTED)
    rect(s, 0.42, 1.25, 12.50, 0.012, RULE)
    txt(s, 11.85, 7.02, 1.05, 0.22, str(page), 9.5, GREY, align="right")
    return s


def summary(slide, text, label="Summary"):
    rect(slide, 0.42, 6.34, 12.50, 0.48, INK)
    txt(slide, 0.60, 6.47, 1.18, 0.22, label, 10.5, LIGHT, True)
    txt(slide, 1.87, 6.40, 10.88, 0.36, text, 10.8, WHITE, leading=1.02)


def card(slide, x, y, w, h, fill, badge, label, heading, items,
         accent=BADGE, line_size=10.5, head_size=13.5):
    rect(slide, x, y, w, h, fill)
    bw = 0.28 if len(badge) <= 2 else 0.50
    rect(slide, x + 0.15, y + 0.15, bw, 0.24, accent)
    txt(slide, x + 0.15, y + 0.15, bw, 0.24, badge, 8.5, WHITE, True,
        align="center", anchor="middle")
    txt(slide, x + 0.15 + bw + 0.12, y + 0.16, w - bw - 0.40, 0.22,
        label, 9.0, accent, True)
    txt(slide, x + 0.17, y + 0.46, w - 0.34, 0.50, heading, head_size,
        NAVY, True, leading=1.04)
    lines(slide, x + 0.17, y + 1.02, w - 0.34, h - 1.10, items,
          size=line_size, gap=5)


def stat(slide, x, y, w, h, value, caption, fill=PALE, accent=NAVY,
         value_size=20, cap_size=10.5):
    rect(slide, x, y, w, h, fill)
    txt(slide, x + 0.18, y + 0.11, w - 0.36, 0.44, value,
        value_size, accent, True)
    txt(slide, x + 0.18, y + 0.59, w - 0.36, h - 0.66, caption,
        cap_size, MUTED, leading=1.10)


def outline_row(slide, x, y, w, h, number, heading, detail, timing=None,
                fill=PALE):
    rect(slide, x, y, w, h, fill)
    txt(slide, x + 0.17, y + 0.15, 0.62, 0.40, number, 18, TEAL, True)
    txt(slide, x + 0.80, y + 0.12, w - 2.65, 0.30, heading, 14, NAVY, True)
    txt(slide, x + 0.80, y + 0.41, w - 2.65, 0.26, detail, 12, MUTED)
    if timing:
        txt(slide, x + w - 1.85, y + 0.20, 1.65, 0.30, timing, 12, GREY,
            True, align="right")


# ---------------------------------------------------------------- slides ----

def s01(prs, blank):
    s = new_slide(prs, blank)
    stripes(s)
    s.shapes.add_picture(str(WORDMARK), I(0.67), I(0.42),
                         width=I(3.82), height=I(1.04))
    txt(s, 6.90, 0.62, 5.76, 0.26,
        "AIDD8122  ·  Advanced Topics in AI & Drug Discovery  ·  Final Project",
        10.5, TEAL, True, align="right")
    rect(s, 0, 2.21, W, 0.012, NAVY)
    rect(s, 0, 2.24, W, 1.95, NAVY)
    rect(s, 0, 4.19, W, 0.012, NAVY)
    txt(s, 0.62, 2.48, 12.10, 0.62, "Trust Before Throughput",
        30, WHITE, True)
    txt(s, 0.62, 3.18, 12.10, 0.82,
        "Benchmarking, uncertainty-aware screening and multi-objective generative "
        "design of FAP-selective radioligand warheads",
        17, PALEBLUE, leading=1.16)
    txt(s, 0.62, 4.55, 12.10, 0.30,
        "972 records   →   384 FAP inhibitors   →   3 split protocols   →   one calibrated decision rule",
        14.5, TEAL, True)
    txt(s, 0.62, 5.20, 5.60, 0.32, "HUANG Yong   (P2600222)", 15, INK)
    txt(s, 0.62, 5.58, 5.60, 0.32, "Group 2  ·  AIDD8122 Final Project", 13, MUTED)
    rect(s, 0.64, 6.06, 4.60, 0.012, RULE)
    txt(s, 0.62, 6.20, 6.40, 0.30, "12 November 2026", 12.5, GREY)
    txt(s, 0.62, 6.62, 8.40, 0.26,
        "Prepared for AIDD8122 · Advanced Topics in Artificial Intelligence & Drug Discovery",
        10.5, GREY)
    return s


def s02(prs, blank):
    s = content_slide(prs, blank, "Outline",
                      "Five parts; limitations are dealt with where they arise", 2)
    rows = [
        ("01", "Background", "FAP and FAPI PET: where the target sits, why the contrast is clean"),
        ("02", "Research questions", "Accuracy, transfer, trust, and generative design"),
        ("03", "Methods", "Curation, three split protocols, conformal intervals, graph GA"),
        ("04", "Results", "Benchmark, selectivity, uncertainty, design campaigns"),
        ("05", "Conclusions", "Take-home messages and where the evidence stops"),
    ]
    for i, (n, h, d) in enumerate(rows):
        outline_row(s, 0.45, 1.66 + i * 0.93, 12.40, 0.72, n, h, d)
    return s


def s03(prs, blank):
    s = content_slide(prs, blank,
                      "FAP is the fastest-moving small-molecule target in nuclear medicine",
                      "A validated stromal address, one warhead, and a therapeutic gap that has not closed", 3)
    card(s, 0.45, 1.45, 3.95, 2.85, BLUE, "1", "CONTEXT",
         "A stromal address with a clean selectivity profile",
         ["FAP sits on cancer-associated fibroblasts",
          "Largely absent from healthy tissue",
          "Accepts drug-like small molecules",
          "Selectivity over PREP and DPP4 is the hard constraint"])
    card(s, 4.62, 1.45, 3.95, 2.85, CREAM, "2", "CHEMISTRY",
         "One warhead, four radionuclides",
         ["Quinoline-carboxamide FAPIs",
          "68Ga / 18F-AlF for imaging",
          "90Y / 177Lu for therapy",
          "Only the chelator changes"], accent=AMBER)
    card(s, 8.79, 1.45, 4.06, 2.85, MINT, "3", "LIMIT",
         "Imaging contrast has not become therapeutic benefit",
         ["Tumour retention is short",
          "Responses are rare; myelosuppression is dose-limiting",
          "Each design cycle costs weeks of chemistry plus an animal study"],
         accent=TEAL)
    rect(s, 0.45, 4.52, 12.40, 1.58, PALE)
    txt(s, 0.65, 4.62, 12.00, 0.28,
        "Three numbers that frame the design problem", 12.5, NAVY, True)
    triples = [
        ("0.65", "28", "tumour entities imaged", "68Ga-FAPI PET/CT · Kratochwil 2019"),
        ("4.48", "4", "radionuclides, one warhead", "68Ga / 18F-AlF imaging · 90Y / 177Lu therapy"),
        ("8.31", "14%", "objective response (4/28)", "DCR 46% · 21% grade 3/4 haematological tox"),
    ]
    for xs, big, label, sub in triples:
        x = float(xs)
        txt(s, x, 4.95, 3.54, 0.34, big, 20, NAVY, True)
        txt(s, x, 5.34, 3.54, 0.24, label, 11, INK, True)
        txt(s, x, 5.62, 3.54, 0.38, sub, 9.5, GREY, leading=1.05)
    summary(s, "The design levers are warhead, linker, chelator and albumin binder — with selectivity over PREP and DPP4 as the hard constraint.")
    return s


def s03b(prs, blank):
    """Background page: what FAP is, and what a FAPI PET image actually shows."""
    s = content_slide(prs, blank, "FAP sits on the stroma, not on the tumour cell",
                      "Fibroblast activation protein (FAP): where it is expressed, what the tracer binds, "
                      "why the contrast is clean", 3)
    mech_h = pic(s, ASSETS / "fap_mechanism.png", 0.42, 1.52, 5.30)
    txt(s, 0.42, 1.52 + mech_h + 0.16, 5.30, 0.50,
        "Schematic of the FAP–FAPI interaction at the plasma membrane. The quinoline "
        "warhead occupies the catalytic cleft; DOTA carries the 68Ga label. Not to scale.",
        8.4, MUTED, leading=1.20)
    pet_h = pic(s, ASSETS / "fap_refs" / "krat_fig3.jpg", 5.92, 1.44, 3.24)
    txt(s, 5.92, 1.44 + pet_h + 0.12, 3.24, 0.64,
        "68Ga-FAPI PET/CT maximum-intensity projections in 15 of the 28 tumour entities studied. "
        "Kratochwil et al., JNM 2019;60:801–805, Fig. 3 (© SNMMI; academic use).",
        7.8, MUTED, leading=1.24)
    txt(s, 9.48, 1.50, 3.42, 0.20, "KEY POINTS", 8.8, TEAL, True)
    bullets = [
        "FAP sits on cancer-associated fibroblasts, not on tumour cells.",
        "68Ga-FAPI binds the warhead pocket; uptake reports stromal volume.",
        "Blood-pool and muscle background stays low (SUVmax < 2).",
        "Brightest in sarcoma, breast, lung and oesophageal cancer.",
    ]
    y = 1.94
    for body in bullets:
        rect(s, 9.48, y + 0.068, 0.10, 0.030, TEAL)
        txt(s, 9.71, y, 3.19, 0.62, body, 10.6, INK, leading=1.16)
        y += 0.95
    summary(s, "A stromal target, not a tumour-cell target: the image reports the tumour's supporting tissue, which is why it works across entities.")
    return s


def s04(prs, blank):
    s = content_slide(prs, blank, "Four questions a radiochemist would actually ask",
                      "Published FAP QSAR models report random-split R² ≈ 0.75 — an interpolation result, not a design decision", 4)
    rq = [
        (0.45, 1.45, "RQ1", "ACCURACY", BLUE, BADGE,
         "How accurate is a FAP affinity model under real use?",
         ["Random vs scaffold vs time split: interpolation,",
          "scaffold hopping, prospective prediction."]),
        (6.78, 1.45, "RQ2", "TRANSFER", CREAM, AMBER,
         "Does paralogue bioactivity transfer to FAP?",
         ["PREP and DPP4 have 1.3× and 11× more data;",
          "can selectivity be predicted at all?"]),
        (0.45, 3.65, "RQ3", "TRUST", MINT, TEAL,
         "When are the uncertainty estimates honest?",
         ["Applicability-domain gating plus split-conformal",
          "intervals at the 80% and 90% levels."]),
        (6.78, 3.65, "RQ4", "DESIGN", BLUE, BADGE,
         "What does multi-objective design add?",
         ["Graph GA with NSGA-II over affinity, margin,",
          "QED and SA, against two controls."]),
    ]
    for x, y, tag, label, fill, accent, head, body in rq:
        rect(s, x, y, 6.10, 2.05, fill)
        rect(s, x + 0.15, y + 0.15, 0.57, 0.24, accent)
        txt(s, x + 0.15, y + 0.15, 0.57, 0.24, tag, 8.5, WHITE, True,
            align="center", anchor="middle")
        txt(s, x + 0.84, y + 0.16, 4.90, 0.22, label, 9.0, accent, True)
        txt(s, x + 0.17, y + 0.50, 5.76, 0.52, head, 14, NAVY, True,
            leading=1.05)
        lines(s, x + 0.17, y + 1.14, 5.76, 0.72, body, 11, MUTED, gap=4)
    txt(s, 0.62, 5.84, 12.10, 0.42,
        "Success criteria fixed in advance: scaffold-split R² > 0.5, screening AUROC > 0.85, and conformal coverage predictable from the applicability domain.",
        11.25, MUTED, leading=1.10)
    summary(s, "A claim about FAP prediction that does not name its split protocol cannot be falsified.")
    return s


def s05(prs, blank):
    s = content_slide(prs, blank, "Workflow: from ChEMBL records to designed warheads",
                      "Curation, benchmarking, calibrated error bars and a four-objective design loop", 5)
    # taller cards after the 2026-11 redraw: 10.2 in wide keeps the bottom at 5.31,
    # clear of the chip row at 5.42 (the old 12.10 in width ran into it)
    pic(s, ASSETS / "deck_workflow.png", 1.57, 1.40, 10.20)
    chips = [
        ("DATA", "972 FAP records, curated with an auditable filter log"),
        ("MODELS", "Morgan + MACCS + descriptors; ridge / SVR / RF / XGBoost"),
        ("TRUST", "Applicability domain + conformal intervals"),
        ("DESIGN", "Graph GA, NSGA-II, four objectives"),
    ]
    for i, (tag, body) in enumerate(chips):
        x = 0.45 + i * 3.11
        rect(s, x, 5.42, 2.98, 0.72, PALE)
        txt(s, x + 0.14, 5.50, 2.70, 0.20, tag, 9.5, TEAL, True)
        txt(s, x + 0.14, 5.72, 2.70, 0.38, body, 9.5, MUTED, leading=1.05)
    summary(s, "Random, scaffold and time splits are applied identically to every model, so failures can be attributed to the protocol rather than the algorithm.")
    return s


def s06(prs, blank):
    s = content_slide(prs, blank, "The public FAP record is small, noisy and non-stationary",
                      "Potency coverage, two publication eras and scaffold diversity across the three paralogues", 6)
    pic(s, ASSETS / "deck_dataset.png", 0.62, 1.45, 12.10)
    stat(s, 0.45, 5.06, 4.02, 1.10, "972 → 384",
         "unique FAP inhibitors after curation, 161 Murcko scaffolds", BLUE)
    stat(s, 4.66, 5.06, 4.02, 1.10, "0.67 log",
         "median replicate SD — the error floor no model can beat (best RMSE 0.70)", CREAM, AMBER)
    stat(s, 8.87, 5.06, 4.02, 1.10, "+2.0 log",
         "potency jump after 2018, median MW +80 Da — a new chemical series", MINT, TEAL)
    summary(s, "Assay noise, not model capacity, sets the ceiling; the post-2018 level shift is what breaks prospective prediction.")
    return s


def s07(prs, blank):
    s = content_slide(prs, blank, "The split, not the algorithm, decides the conclusion",
                      "Four regressors plus a mean baseline, under three protocols: interpolation, scaffold hopping, "
                      "prospective prediction", 7)
    pic(s, FIGS / "fig2_benchmark.png", 0.55, 1.50, 7.55)
    rect(s, 8.30, 1.50, 4.45, 3.18, BLUE)
    txt(s, 8.52, 1.66, 4.00, 0.28, "XGBoost R² across protocols", 13, NAVY, True)
    txt(s, 8.52, 2.10, 4.00, 0.46, "0.75  →  0.68  →  −0.79", 21, NAVY, True)
    lines(s, 8.52, 2.74, 4.00, 1.80,
          ["Random split: interpolation, R² 0.75, RMSE 0.71.",
           "Scaffold split: R² 0.68 — still usable.",
           "Time split: worse than the mean predictor;",
           "EF10 = 0.96, no better than random picks."], 10.8, MUTED, gap=6)
    txt(s, 0.62, 4.86, 12.10, 1.20,
        "RF ≈ XGBoost, both clearly ahead of SVR and ridge on the interpolation protocols. "
        "Ridge degrades least under time splitting because it relies on additive physicochemical "
        "trends rather than scaffold-specific fingerprints. A model trained on pre-2019 chemistry "
        "cannot rank the quinoline FAPI series at all.",
        11.5, MUTED, leading=1.18)
    summary(s, "The same model is excellent or worse than the mean predictor depending only on the split protocol.")
    return s


def s08(prs, blank):
    s = content_slide(prs, blank, "Selectivity is a cliff, not a slope",
                      "Twenty seeds, two modelling routes and one dominant off-target liability", 8)
    pic(s, FIGS / "fig3_transfer_selectivity.png", 0.55, 1.50, 8.10)
    rect(s, 8.85, 1.50, 3.90, 2.78, CREAM)
    txt(s, 9.05, 1.66, 3.50, 0.28, "The DPP4 liability", 13, AMBER, True)
    txt(s, 9.05, 2.06, 3.50, 0.50, "86%", 27, NAVY, True)
    lines(s, 9.05, 2.70, 3.50, 1.44,
          ["of dual-measured compounds are non-selective vs DPP4",
           "mean ΔpActivity = −1.75 log",
           "reversal only in the mixed-assay stratum (n = 30)"],
          10.5, MUTED, gap=6)
    txt(s, 0.62, 4.52, 12.10, 1.60,
        "Selectivity AUROC 0.84 within known scaffolds — but 0.38 (below chance) on new scaffolds, over 20 seeds. "
        "Used as a filter, the model would enrich the non-selective series. Direct modelling of ΔpActivity and "
        "subtraction of two single-target models are indistinguishable; the effect reverses only in a small "
        "mixed-assay stratum (n = 30), which is not used as evidence.",
        11.5, MUTED, leading=1.18)
    summary(s, "AUROC 0.84 inside known scaffolds, 0.38 across scaffolds — a filter built on it would invert the shortlist.")
    return s


def s09(prs, blank):
    s = content_slide(prs, blank, "Error bars you can defend — and where you cannot",
                      "Applicability-domain gating and split-conformal coverage at nominal 80% / 90%", 9)
    pic(s, FIGS / "fig4_applicability_domain.png", 0.55, 1.55, 6.10)
    txt(s, 0.62, 4.62, 6.40, 1.48,
        "Error rises monotonically with distance from the training set: RMSE 0.70 at similarity above 0.7 "
        "versus ≈1.0 below 0.5. Exchangeability holds under random splitting and fails for "
        "out-of-domain compounds — the interval is optimistic exactly where it matters most.",
        11.3, MUTED, leading=1.18)

    # three-line table, no shading and no vertical rules
    x0, x1, x2 = 7.90, 10.60, 11.60
    txt(s, x0, 1.45, 4.85, 0.28,
        "Conformal coverage (nominal 80% / 90%)", 12.5, NAVY, True)
    rect(s, x0, 1.80, 4.85, 0.016, NAVY)
    txt(s, x0, 1.88, 2.60, 0.24, "Protocol", 10.5, NAVY, True)
    txt(s, x1, 1.88, 0.90, 0.24, "80%", 10.5, NAVY, True, align="center")
    txt(s, x2, 1.88, 0.90, 0.24, "90%", 10.5, NAVY, True, align="center")
    rect(s, x0, 2.16, 4.85, 0.010, NAVY)
    rows = [
        ("Random split (control)", "0.79", "0.89", False),
        ("Scaffold split, in domain", "0.77", "0.88", False),
        ("Scaffold split, all", "0.73", "0.87", False),
        ("Scaffold split, out of domain", "0.54", "0.77", True),
    ]
    yy = 2.24
    for label, a, b, warn in rows:
        col = RED if warn else INK
        txt(s, x0, yy, 2.60, 0.34, label, 10.5, col, warn)
        txt(s, x1, yy, 0.90, 0.34, a, 10.5, col, warn, align="center")
        txt(s, x2, yy, 0.90, 0.34, b, 10.5, col, warn, align="center")
        yy += 0.38
    rect(s, x0, 3.80, 4.85, 0.016, NAVY)
    txt(s, x0, 3.94, 4.85, 1.10,
        "Quote a conformal interval only when nearest-neighbour similarity exceeds 0.6. "
        "Mondrian calibration restores out-of-domain coverage but doubles the width — "
        "a diagnostic, not a fix.", 10.5, MUTED, leading=1.15)
    summary(s, "Below similarity 0.6, rank a series internally but keep the interval out of the synthesis decision.")
    return s


def s10(prs, blank):
    s = content_slide(prs, blank, "Prospective failure is a level shift, not a ranking signal",
                      "Train on compounds published up to 2018, test on 2019–2024", 10)
    pic(s, FIGS / "fig5_temporal_drift.png", 0.55, 1.50, 6.28)
    txt(s, 0.62, 4.62, 6.20, 1.48,
        "The model under-predicts every new compound by roughly an order of magnitude — it has simply "
        "never seen the quinoline FAPI series. Recalibrating one intercept on 25% of the test period "
        "repairs the level; the ordering of compounds is unchanged.",
        11.3, MUTED, leading=1.18)
    stat(s, 7.10, 1.50, 5.65, 1.42, "+1.18 log",
         "systematic offset of the raw model on post-2018 compounds",
         fill=BLUE, value_size=24)
    stat(s, 7.10, 3.02, 5.65, 1.42, "−0.82 → −0.15",
         "R² after recalibrating one intercept on 25% of the test set",
         fill=CREAM, accent=AMBER, value_size=21)
    stat(s, 7.10, 4.54, 5.65, 1.42, "−0.12",
         "Spearman correlation — unchanged by recalibration: ranking information is absent",
         fill=MINT, accent=TEAL, value_size=24)
    summary(s, "Recalibration restores the level, never the ranking — if the downstream use is ‘pick the top ten’, it buys nothing.")
    return s


def s11(prs, blank):
    s = content_slide(prs, blank, "Generative design: the objectives shape the chemistry",
                      "Three campaigns, one 384-member archive, four objectives under NSGA-II", 10)
    pic(s, FIGS / "fig7_generative.png", 0.55, 1.45, 7.20)
    arms = [
        (1.45, BLUE, BADGE, "Multi-objective", "115 scaffolds · QED 0.61",
         "median MW 413 Da, SA 2.93 — the most usable chemistry\n"
         "61% pass the pFAP ≥ 8 / margin ≥ 1 gate"),
        (2.97, CREAM, AMBER, "Affinity only", "83 scaffolds · QED 0.25",
         "median MW 548 Da — the model is exploited, chemistry degenerates\n"
         "100% pass the same gate"),
        (4.49, PALE, GREY, "Random", "80 scaffolds · QED 0.52",
         "no selection pressure; population fitness drifts down"),
    ]
    for y, fill, accent, tag, value, detail in arms:
        rect(s, 7.95, y, 4.80, 1.42, fill)
        txt(s, 8.15, y + 0.13, 4.40, 0.22, tag.upper(), 9.5, accent, True)
        txt(s, 8.15, y + 0.40, 4.40, 0.30, value, 15, INK, True)
        txt(s, 8.15, y + 0.78, 4.40, 0.56, detail, 10.5, MUTED, leading=1.10)
    txt(s, 0.62, 5.05, 7.10, 1.10,
        "All campaigns hit the affinity-model ceiling (pFAP 9.6–9.7) within a few generations; "
        "361 of 384 archive molecules are novel, but the top hits are warheads, not radioligands — "
        "no chelator, no leaving group.",
        10.5, MUTED, leading=1.15)
    summary(s, "The objective function is the chemistry — single-objective greed returns potent, heavy, unusable matter.")
    return s


def s12(prs, blank):
    s = content_slide(prs, blank, "Take-home messages",
                      "Four results, one decision rule, and a limit on each claim", 11)
    rows = [
        ("01", "The evaluation protocol outweighs the algorithm",
         "The same XGBoost is excellent (R² 0.75) or worse than the mean predictor (R² −0.79) depending only on the split."),
        ("02", "Assay noise sets the ceiling, not model architecture",
         "Median replicate SD 0.67 log units vs best RMSE 0.70. The highest-value intervention is a consistent assay campaign."),
        ("03", "Selectivity must be verified, not optimised in silico",
         "AUROC 0.84 inside known scaffolds, 0.38 across scaffolds; conformal coverage 0.79 vs 0.54 at nominal 80%."),
        ("04", "The deliverable is a calibrated decision rule",
         "Use the model above similarity 0.6, distrust it across series and eras; multi-objective design keeps chemistry usable."),
    ]
    for i, (n, head, body) in enumerate(rows):
        y = 1.45 + i * 1.17
        rect(s, 0.45, y, 12.40, 1.02, PALE)
        txt(s, 0.62, y + 0.16, 0.62, 0.40, n, 18, TEAL, True)
        txt(s, 1.25, y + 0.12, 11.20, 0.30, head, 14, NAVY, True)
        txt(s, 1.25, y + 0.46, 11.20, 0.46, body, 11.5, MUTED, leading=1.10)
    summary(s, "None of this is specific to FAP — the protocol transfers; the numbers do not.")
    return s


def s13(prs, blank):
    s = content_slide(prs, blank, "Limitations",
                      "Where the evidence stops, and what this study cannot claim", 12)
    limits = [
        "ChEMBL snapshot: publication bias towards potent compounds; assay conditions inconsistent across ~60 papers.",
        "Ligand-based only — no docking or protein-structure information; physicochemical coincidence cannot be excluded.",
        "The generative scorer is the same model whose failure modes are documented; pFAP > 9 means ‘inside the confident region’, not ‘10 nM’.",
        "Similarity gate at 0.40 constrains novelty by construction.",
    ]
    for i, body in enumerate(limits):
        y = 1.88 + i * 1.113
        rect(s, 0.62, y, 12.40, 0.72, CREAM)
        rect(s, 0.78, y + 0.13, 0.06, 0.46, AMBER)
        txt(s, 0.98, y + 0.19, 11.86, 0.34, body, 10.5, INK, leading=1.12)
    summary(s, "The evidence stops at the applicability domain; beyond it, nothing in this study is tested.")
    return s


def s14(prs, blank):
    s = new_slide(prs, blank)
    rect(s, 0, 0, W, H, NAVY)
    crest(s)
    txt(s, 0.90, 2.90, 11.50, 0.78, "Thank you — questions welcome",
        34, WHITE, True)
    rect(s, 0.92, 3.88, 1.60, 0.03, TEAL)
    txt(s, 0.92, 4.14, 11.50, 0.30,
        "HUANG Yong (P2600222)   ·   AIDD8122 Final Project   ·   12 November 2026",
        12, PALEBLUE)
    txt(s, 0.92, 4.60, 11.50, 0.55,
        "Trust before throughput: the useful output of a model on a small, drifting target is a calibrated\n"
        "decision rule — not a single predicted number. Code, curated data and filter logs are in the repository.",
        12, CREDIT, leading=1.20)
    return s


# s10 (temporal drift) is kept for reference but is no longer in the talk:/n# the delivered deck is 13 pages and the time-split result is covered on s07.
SLIDES = [s01, s02, s03b, s04, s05, s06, s07, s08, s09, s11, s12, s13, s14]


def build() -> None:
    prs = Presentation()
    prs.slide_width = I(W)
    prs.slide_height = I(H)
    blank = prs.slide_layouts[6]
    for i, fn in enumerate(SLIDES, 1):
        slide = fn(prs, blank)
        slide.notes_slide.notes_text_frame.text = NOTES[i].strip()
    PPTX.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(PPTX))
    print(f"  deck rebuilt   : {PPTX}  ({len(prs.slides)} slides)")
    print(f"  notes written  : {sum(1 for s in prs.slides if s.has_notes_slide)} pages")


def export_pdf() -> None:
    import win32com.client as win32

    tmp = Path(tempfile.gettempdir()) / f"aidd_deck_{int(time.time())}.pdf"
    app = win32.DispatchEx("PowerPoint.Application")
    try:
        pres = app.Presentations.Open(str(PPTX), WithWindow=False)
        pres.SaveCopyAs(str(tmp), 32)          # 32 = ppSaveAsPDF
        pres.Close()
    finally:
        app.Quit()
    shutil.copy2(tmp, PDF)
    tmp.unlink(missing_ok=True)
    print(f"  PDF exported   : {PDF}  ({PDF.stat().st_size:,} bytes)")


def check() -> None:
    prs = Presentation(str(PPTX))
    print(f"  slides         : {len(prs.slides._sldIdLst)}")
    print(f"  size           : {prs.slide_width/914400:.3f} x {prs.slide_height/914400:.3f} in")
    print(f"  notes pages    : {sum(1 for s in prs.slides if s.has_notes_slide)}")
    for i, s in enumerate(prs.slides, 1):
        n = len(s.shapes)
        pics = sum(1 for sh in s.shapes if sh.shape_type == 13)
        nt = s.notes_slide.notes_text_frame.text if s.has_notes_slide else ""
        print(f"   {i:2d}: shapes={n:2d} pictures={pics} notes={len(nt):4d} chars")


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--pdf", action="store_true", help="also export the PDF")
    ap.add_argument("--check", action="store_true", help="inspect an existing deck")
    args = ap.parse_args()
    if args.check:
        check()
        return
    build()
    if args.pdf:
        export_pdf()
    check()


if __name__ == "__main__":
    main()
