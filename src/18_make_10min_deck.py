"""Rewrite the 15-minute deck into a 10-minute cut.

Deletes the reference-list slide, renumbers the page footers and replaces all
speaker notes with a 10:00 timeline.
"""
from __future__ import annotations

import re

from pptx import Presentation

SRC = "Submission_HuangYong_P2600222/02_Presentation/AIDD8122_Final_Presentation_HuangYong_15min.pptx"
DST = "Submission_HuangYong_P2600222/02_Presentation/AIDD8122_Final_Presentation_HuangYong_10min.pptx"

NOTES = [
    # 1 — title, 20 s
    """TIME 0:00-0:20 | 20 s

Walk on, one sentence, no warm-up.
"Good morning. One question: if I only trust the public FAP affinity
data, how far can a model carry a radiochemist -- and where does it stop
being useful?"

Do not read the subtitle. Do not preview the agenda here.""",

    # 2 — outline, 20 s
    """TIME 0:20-0:40 | 20 s

Signpost only, four words.
"Target, data, decision, design -- in that order."

Then: "Methods detail is in the report; here I show only what changes a
decision." Move on immediately.""",

    # 3 — background, 85 s
    """TIME 0:40-2:05 | 85 s

Three cards, then the levers. Keep it to 85 s.
- 28 tumour entities imaged with 68Ga-FAPI PET/CT (Kratochwil 2019) -- the
  reason the field moved fast.
- One warhead, four radionuclides: 68Ga / 18F-AlF for imaging, 90Y / 177Lu
  for therapy; only the chelator changes.
- Objective response 14% (4 of 28), disease control 46%, grade 3/4
  haematological toxicity 21% (Fu 2025). Contrast has not become benefit.

Then the four levers in one breath: warhead, linker, chelator, albumin
binder -- with selectivity over PREP and DPP4 as the hard constraint.

Bridge: "Each cycle costs weeks of chemistry and an animal study. That is
the cost the model is supposed to cut.\"""",

    # 4 — questions, 40 s
    """TIME 2:05-2:45 | 40 s

Read the four questions as a list, ten seconds each, no commentary.
1 Can the public record support a potency model at all?
2 Does paralogue data transfer to FAP?
3 Can selectivity over PREP be predicted, or only measured?
4 Does multi-objective design return synthesisable matter?

One line to close: "All four are answerable with data that already exist.\"""",

    # 5 — workflow, 40 s
    """TIME 2:45-3:25 | 40 s

Left to right, do not linger.
- 972 raw FAP records (ChEMBL_37, CHEMBL4683) -> 384 unique structures and
  161 Murcko scaffolds; the filter log ships with the code.
- Morgan-2048/r2 + MACCS-167 + 15 physicochemical descriptors.
- Three splits, four model families, conformal intervals, then a graph GA
  with NSGA-II selection.

The sentence that matters: "The split is fixed before any model is fitted.
That is the whole methodological bet.\"""",

    # 6 — data, 50 s
    """TIME 3:25-4:15 | 50 s

Three numbers, in this order.
- 972 -> 384 compounds, 161 scaffolds. DPP4 in the same release has 4,338
  compounds: FAP is one-eleventh the size.
- 0.67 log median replicate SD. Our best RMSE is 0.70. "We are at the noise
  floor -- more architecture buys nothing, a better assay buys everything."
- +2.0 log median potency jump after 2018, MW up 80 Da: a new series, not a
  trend.

Say the caveat: assay conditions differ across ~60 papers and cannot be
corrected from public metadata.""",

    # 7 — split, 50 s
    """TIME 4:15-5:05 | 50 s

Core slide. Three numbers, slowly.
- Random: R2 0.75, RMSE 0.71, Spearman 0.85.
- Scaffold: R2 0.68, RMSE 0.83 -- costs 0.07; still usable.
- Time (train <=2018, test 2019-2024): R2 -0.79, RMSE 2.05, Spearman -0.12.

Then: "The ranking inverts. A model trained on the pre-2019 literature is
not merely weaker on new compounds -- it is anti-correlated with the truth."

Why: the chemotype changed and the assay standard moved with it.""",

    # 8 — selectivity, 50 s
    """TIME 5:05-5:55 | 50 s

Model, then biology.
Model: selectivity AUROC 0.84 random, 0.38 scaffold -- below chance, over
20 seeds. "Used as a filter it would enrich the non-selective series."
Direct delta modelling and subtraction of two models are indistinguishable
(0.84 vs 0.82; 0.38 vs 0.36): the data are the limit, not the formulation.

Biology: DPP4 dominates -- 86% of 232 dual-measured compounds non-selective,
mean delta -1.75 log. Against PREP, mean delta only +0.59.

Close: "Selectivity is a cliff. One atom flips it, and 384 compounds do not
map a cliff.\"""",

    # 9 — AD / conformal, 50 s
    """TIME 5:55-6:45 | 50 s

Two gates, two numbers each.
- Error rises with distance: RMSE 0.70 above 0.7 similarity (n=263), 0.94 to
  1.12 below 0.5.
- Coverage: 0.79 / 0.89 random at nominal 80 / 90; 0.73 / 0.87 scaffold;
  0.54 at 80% outside the domain.

The rule: quote an interval only above 0.6 similarity. Below it, rank a
series internally and keep the number out of a synthesis decision.

One line on the repair: Mondrian calibration restores out-of-domain coverage
but doubles the width -- a diagnostic, not a fix.""",

    # 10 — temporal, 40 s
    """TIME 6:45-7:25 | 40 s

The most transferable result.
RMSE 2.07 -> 1.64 after recalibrating the intercept on 25% of the test
period; R2 -0.82 -> -0.15; offset +1.18 log. Spearman stays at -0.12.

"So recalibration fixes the level and nothing else. If the downstream use is
'pick the top ten', recalibration buys you nothing.\"""",

    # 11 — generative, 60 s
    """TIME 7:25-8:25 | 60 s

Three arms, same 384-member archive.
- Affinity only: 100% pass the potency filter, median pFAP 9.36 -- but QED
  0.25, MW 548 Da, 83 scaffolds. "Potent, ugly, heavy: it exploited the model."
- Multi-objective: 61% pass, pFAP 8.24, QED 0.61, MW 413 Da, 115 scaffolds.
- Random: QED 0.52, fitness drifts down -- no selection pressure.

The point: "All three hit the model ceiling (pFAP 9.6-9.7) within a few
generations. They differ only in what they sacrifice. The objective function
is the chemistry."

Guard: "These are warheads, not radioligands -- no chelator, no leaving
group; DOTA conjugation leaves the applicability domain (Appendix B).\"""",

    # 12 — take-home, 60 s
    """TIME 8:25-9:25 | 60 s

Four lines, one sentence each, then stop.
1 Random-split numbers are a ceiling, not an estimate.
2 Selectivity needs its own model and its own data.
3 A conformal interval is valid only inside the AD -- publish the similarity.
4 Multi-objective design is the only version that returns synthesisable matter.

Close with: "None of this is specific to FAP. The protocol transfers; the
numbers do not.\"""",

    # 13 — limitations, 27 s
    """TIME 9:25-9:52 | 27 s

Pick three, do not read the slide.
- 384 compounds from ~60 papers and 78 assays, with no assay-level random
  effect; publication bias runs towards potent compounds.
- 100 selectivity pairs, heavily scaffold-clustered.
- One time boundary (2018), and the generative scorer is the same model whose
  failures are documented.

Then one line of next step: "The fix is a consistent assay campaign, not a
better architecture.\"""",

    # 14 — thank you, 8 s
    """TIME 9:52-10:00 | 8 s

"Thank you. I would most like to be challenged on the time-split result.
Code, curated data and the filter log are in the repository."

Then stop talking.""",
]


def delete_slide(prs, index: int) -> None:
    sld_id_lst = prs.slides._sldIdLst
    slides = list(sld_id_lst)
    r_id = slides[index].rId
    prs.part.drop_rel(r_id)
    sld_id_lst.remove(slides[index])


def main() -> None:
    prs = Presentation(SRC)
    assert len(prs.slides) == 15, len(prs.slides)
    delete_slide(prs, 13)  # the reference list
    assert len(prs.slides) == 14, len(prs.slides)

    for i, slide in enumerate(prs.slides):
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            m = re.fullmatch(r"(\d{2})\s*/\s*15", shape.text_frame.text.strip())
            if m:
                shape.text_frame.paragraphs[0].runs[0].text = f"{m.group(1)} / 14"
        slide.notes_slide.notes_text_frame.text = NOTES[i]

    prs.save(DST)
    print("saved", DST)


if __name__ == "__main__":
    main()
