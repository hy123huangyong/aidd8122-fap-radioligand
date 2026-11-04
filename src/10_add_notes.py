# -*- coding: utf-8 -*-
"""Add 15-minute speaker notes with per-slide timing to the rebuilt deck."""
import sys
from pptx import Presentation

path = sys.argv[1] if len(sys.argv) > 1 else "deck_v3.pptx"

NOTES = [
# 1 title (0:30)
"""TIME 0:00-0:30 | 30 s

Opening line:
"Good morning. I'm Huang Yong. This project asks a narrow question: if I only
trust public FAP affinity data, how far can a model actually carry a
radiochemist -- and where does it quietly stop being useful?"

Do not read the subtitle. State the framing in one breath: four questions,
three splits, one honest error bar.""",

# 2 outline (0:30)
"""TIME 0:30-1:00 | 30 s

Signpost only. Name the four questions and say the talk is organised around
them, not around methods.
"Target, data, decision, design -- in that order."

Explicitly say: methods detail is in the report; here I show only what changes
a decision.""",

# 3 target background (1:00)
"""TIME 1:00-2:00 | 60 s

Drive the three cards, then the lever list.
- 28 tumour entities imaged with 68Ga-FAPI PET/CT (Kratochwil 2019) -- the
  reason the field moved so fast.
- One warhead, four radionuclides: 68Ga / 18F-AlF for imaging, 90Y / 177Lu for
  therapy. Only the chelator changes, so a single chemistry serves both.
- ~21% objective response in the largest phase II cohort (177Lu-LNC1004,
  Fu 2025). Contrast has not become benefit.

Then the four levers: warhead (nitrile vs boronate), linker, chelator, albumin
binder -- with selectivity over PREP and DPP4 as the hard constraint.

Bridge: "Each design cycle costs weeks of chemistry plus an animal study.
That is the cost the model is supposed to cut." """,

# 4 four questions (1:00)
"""TIME 2:00-3:00 | 60 s

Read the four questions as questions, then give the one-line answer each:
1 Can we predict affinity? Yes within a series (R2 ~0.68), no across time.
2 Is selectivity learnable? Barely -- it is a cliff, not a slope.
3 Can we trust the error bar? Only inside the applicability domain.
4 Can we design new warheads? Yes, but only if the objective is multi-objective.

Say this: "Everything that follows is evidence for those four one-liners." """,

# 5 workflow (1:15)
"""TIME 3:00-4:15 | 75 s

Walk left to right, do not linger on any box.
- 972 raw FAP records from ChEMBL_37 (target CHEMBL4683); every record keeps
  its assay ID so the audit trail survives all the way to the figure.
- Curation chain: 972 raw -> 510 exact-relation -> 438 in range -> 384 unique
  structures, 162 Murcko scaffolds. The filter log is a deliverable, not
  housekeeping.
- Descriptors: Morgan-2048/r2, MACCS-167, 15 physicochemical.
- Three splits (random / scaffold / time), four model families, conformal
  intervals, then the graph GA with NSGA-II.

Emphasise the order: "The split is fixed before any model is fitted. That is
the whole methodological bet of this project." """,

# 6 data quality (1:30)
"""TIME 4:15-5:45 | 90 s

Three numbers, in this order:
- 972 -> 384 unique inhibitors, 162 scaffolds. For scale: DPP4 in the same
  release has 4,338 compounds and 1,597 scaffolds. FAP is roughly one-eleventh
  the size.
- 0.67 log median replicate SD. That is the floor no model can beat -- and our
  best RMSE is 0.70. "We are already at the noise floor. More architecture buys
  nothing; a better assay buys everything."
- +2.0 log median potency jump after 2018, with median MW up 80 Da. That is not
  a trend, it is a new chemical series arriving (quinoline FAPIs).

Caveat to say out loud: assay conditions differ across the ~60 contributing
papers and cannot be corrected from public metadata alone. """,

# 7 split protocol (1:30)
"""TIME 5:45-7:15 | 90 s

This is the core slide. Give the three numbers slowly.
- Random: XGB RMSE 0.71, R2 0.75, Spearman 0.85.
- Scaffold: RMSE 0.83, R2 0.68, Spearman 0.81. Costs about 0.07 R2 -- modest.
- Time (train <=2018, test 2019-2024): R2 goes negative. Ridge 1.58 / -0.07;
  XGB 2.05 / -0.79; RF 2.18 / -1.03. Spearman collapses to -0.12 for XGB.

The sentence to say: "The ranking signal inverts. A model trained on the
pre-2019 literature is not merely weaker on 2019-2024 compounds -- it is
anti-correlated with the truth."

Then why: the field changed chemotype (quinoline/urea -> heteroaryl
boronic-acid-like and covalent warheads), and the assay standard moved with it.""",

# 8 selectivity cliff (1:20)
"""TIME 7:15-8:35 | 80 s

Two halves: the model, then the biology.
Model: selectivity AUROC 0.84 on a random split, 0.38 on a scaffold split --
below chance -- averaged over 20 seeds. "It is not weak on new scaffolds, it is
inverted: used as a filter it would actively enrich the non-selective series."

Direct modelling of delta-pActivity and subtraction of two single-target models
are indistinguishable (0.84 vs 0.82 random, 0.38 vs 0.36 scaffold). So the
problem is the data, not the formulation.

Biology: DPP4 is the dominant liability -- 86% of the 232 dual-measured
compounds are non-selective, mean delta -1.75 log. Against PREP, 100 pairs,
mean delta only +0.59.

The sentence to say: "Selectivity is a cliff. One atom flips it, and 384
compounds do not map a cliff." """,

# 9 uncertainty (1:15)
"""TIME 8:35-9:50 | 75 s

Point at the table, not the paragraph.
- Random split (the control): 0.79 / 0.89 at nominal 80 / 90. Theory holds.
- Scaffold, in domain: 0.77 / 0.88 -- still fine.
- Scaffold, all: 0.73 / 0.87.
- Scaffold, out of domain: 0.54 / 0.77. "At the 80% level the interval covers
  half the time. That is not a conservative estimate, it is a wrong one."

Error vs distance: RMSE 0.70 above similarity 0.7, about 1.0 below 0.5.
Interval width 1.73 and 2.38 log units -- a 50x to 240x affinity range.

The rule to state: quote a conformal interval only above similarity 0.6.
Below that, rank a series internally, but keep the interval out of the
synthesis decision. """,

# 10 time drift (1:20)
"""TIME 9:50-11:10 | 80 s

The failure is a level shift, not a ranking loss. Show the two rows:
RMSE 2.07 -> 1.64 after recalibrating the intercept on 25% of the test period;
R2 -0.82 -> -0.15; mean offset +1.18 log units.
But Spearman stays at -0.12 before and after.

"So recalibration fixes the offset and nothing else. You can rescue the number,
you cannot rescue the ordering. If your downstream use is 'pick the top ten',
recalibration buys you nothing."

This is the single most transferable lesson of the project.""",

# 11 generative (1:20)
"""TIME 11:10-12:30 | 80 s

Three arms, same 384-member archive.
- Affinity only: median pFAP 9.36, 100% pass the pFAP>=8 & margin>=1 filter --
  but median QED 0.25, MW 548 Da, SA 3.51, and only 83 scaffolds. "It found
  potent, ugly, heavy molecules; the optimiser exploited the model."
- Multi-objective (pFAP, margin, QED, SA): median pFAP 8.24, 60.6% pass,
  QED 0.61, MW 413 Da, SA 2.93, 115 scaffolds -- more diversity, slightly less
  potency.
- Random baseline: QED 0.52, SA 3.59, 80 scaffolds -- no selection pressure,
  population fitness drifts down.

The point: "All three hit the ceiling of the affinity model (pFAP 9.6-9.7) in
a few generations. They differ only in what they sacrifice. The objective
function is the chemistry." """,

# 12 take-home (0:45)
"""TIME 12:30-13:15 | 45 s

Four lines, no elaboration -- they were all justified earlier.
1 Random-split numbers are a ceiling, not an estimate.
2 Selectivity needs its own model and its own data.
3 Conformal intervals are only valid inside the AD; publish the similarity.
4 Multi-objective design is the only version that yields synthesisable matter.

Close with: "None of this is specific to FAP. The protocol transfers; the
numbers do not." """,

# 13 limitations (1:00)
"""TIME 13:15-14:15 | 60 s

Be explicit and unhedged:
- 384 compounds contributed by ~60 papers across 78 assays, with no assay-level
  random effect in the model. Publication bias runs towards potent compounds.
- Selectivity pairs are few (100 for PREP) and heavily scaffold-clustered.
- The time split is one boundary (2018); I did not test 2016 or 2020.
- The generative scorer is the same model whose failure modes are documented --
  pFAP > 9 means "inside the confident region", not "10 nM".
- Similarity gate at 0.40 constrains novelty by construction.
- No synthesis, no wet-lab validation: this is a retrospective audit.

Next steps: consistent assay campaign; structure-aware features; freeze the
model and score the next 12 months prospectively; dual-objective active
learning where the assay, not the model, picks the scaffold.""",

# 14 references (0:20)
"""TIME 14:15-14:35 | 20 s

Do not read the list. Say: "Six primary sources carry the argument -- ChEMBL
for the data, the FAPI PET trials for the clinical frame, the conformal and
AD papers for the statistics, and the two FAP reviews for chemistry. Full list
in the report."

If asked about a specific citation, the report has all 29 with DOIs.""",

# 15 closing (0:15)
"""TIME 14:35-14:50 | 15 s

"Thank you. I'm happy to take questions -- particularly on the time-split
result, which is the part I would most like to be challenged on. Code, curated
data and the full filter log are in the submitted repository."

Then stop talking. Leave the slide up.""",
]


def main():
    prs = Presentation(path)
    n = len(prs.slides)
    print(f"slides in deck: {n}, notes prepared: {len(NOTES)}")
    from pptx.util import Inches, Pt
    from pptx.dml.color import RGBColor
    from pptx.oxml.ns import qn
    from lxml import etree

    NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
    NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"

    def make_textbox(ns, left, top, width, height, sid):
        sp = etree.SubElement(ns.shapes._spTree, qn("p:sp"))
        nv = etree.SubElement(sp, qn("p:nvSpPr"))
        cnv = etree.SubElement(nv, qn("p:cNvPr"))
        cnv.set("id", str(sid))
        cnv.set("name", "NotesBox %d" % sid)
        etree.SubElement(nv, qn("p:cNvSpPr"))
        etree.SubElement(nv, qn("p:nvPr"))
        sppr = etree.SubElement(sp, qn("p:spPr"))
        xfrm = etree.SubElement(sppr, qn("a:xfrm"))
        off = etree.SubElement(xfrm, qn("a:off"))
        off.set("x", str(int(left)))
        off.set("y", str(int(top)))
        ext = etree.SubElement(xfrm, qn("a:ext"))
        ext.set("cx", str(int(width)))
        ext.set("cy", str(int(height)))
        prst = etree.SubElement(sppr, qn("a:prstGeom"))
        prst.set("prst", "rect")
        etree.SubElement(prst, qn("a:avLst"))
        etree.SubElement(sp, qn("p:txBody"))
        body = sp.find(qn("p:txBody"))
        bodypr = etree.SubElement(body, qn("a:bodyPr"))
        bodypr.set("wrap", "square")
        etree.SubElement(body, qn("a:lstStyle"))
        p0 = etree.SubElement(body, qn("a:p"))
        r0 = etree.SubElement(p0, qn("a:r"))
        t0 = etree.SubElement(r0, qn("a:t"))
        t0.text = ""
        return sp

    for i, slide in enumerate(prs.slides):
        if i >= len(NOTES):
            break
        ns = slide.notes_slide
        tf = ns.notes_text_frame
        if tf is None:
            sp = make_textbox(ns, 548640, 457200, 7772400, 5943600,
                              i + 900)
            from pptx.text.text import TextFrame
            tf = TextFrame(sp.find(qn("p:txBody")), ns)
            tf.word_wrap = True
        tf.text = NOTES[i].strip()
        for p in tf.paragraphs:
            p.line_spacing = 1.15
            for r in p.runs:
                r.font.size = Pt(12)
                r.font.name = "Consolas"
                r.font.color.rgb = RGBColor(0x22, 0x22, 0x22)
    prs.save(path)
    print("notes written ->", path)


if __name__ == "__main__":
    main()
