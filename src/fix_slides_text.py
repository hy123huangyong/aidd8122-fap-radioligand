# -*- coding: utf-8 -*-
"""P0/P1 text fixes for slides 3, 8, 11 of the 15-min deck (v2 evaluation)."""
import sys
from pptx import Presentation

path = sys.argv[1]

# (slide_index_0based, [(old_substring, new_substring), ...])
FIXES = [
    (2, [
        # P0-1: 21% is haematological toxicity, not objective response
        ("~21%", "14%"),
        ("best objective response in therapy",
         "objective response (4/28); DCR 46%"),
        # keep the source line inside the 420-px card (it overflowed the canvas)
        ("grade 3/4 haematological toxicity: 21% — phase II 177Lu-LNC1004, Fu 2025",
         "Fu 2025 · 21% grade 3/4 haematological tox"),
        # P1: compress slide 3 to a ~60-second read
        ("Fibroblast activation protein (FAP) sits on cancer-associated fibroblasts, "
         "is largely absent from healthy tissue, and accepts small drug-like inhibitors "
         "— the first stromal target with a small-molecule theranostic pair.",
         "FAP sits on cancer-associated fibroblasts, is largely absent from healthy "
         "tissue, and accepts drug-like small molecules."),
        ("Quinoline-carboxamide FAPIs (FAPI-04, -46, -74) are labelled with "
         "68Ga / 18F-AlF for imaging and 90Y / 177Lu for therapy — only the chelator changes.",
         "Quinoline-carboxamide FAPIs: 68Ga / 18F-AlF for imaging, 90Y / 177Lu for "
         "therapy — only the chelator changes."),
        ("Retention in tumour is short and response rates are modest, so the field "
         "needs better ligands — and each design cycle costs weeks of chemistry plus "
         "an animal study.",
         "Tumour retention is short, responses are rare and myelosuppression is "
         "dose-limiting — each design cycle costs weeks of chemistry plus an animal study."),
        ("Warhead (nitrile vs boronate) · linker · chelator · albumin binder",
         "Warhead · linker · chelator · albumin binder"),
    ]),
    (7, [
        # P1: label vs DPP4 and disclose the mixed-assay stratum reversal
        ("86% of dual-measured compounds are non-selective (mean Δ = −1.75 log).",
         "86% of dual-measured compounds are non-selective vs DPP4 (mean Δ = −1.75 log); "
         "the effect reverses only in the small mixed-assay stratum (n = 30, report §5.5)."),
    ]),
    (10, [
        # P1: warheads, not radioligands
        ("top hits are plausible quinoline-carboxamide analogues.",
         "top hits are plausible quinoline-carboxamide warheads — warheads only, "
         "not radioligands (no chelator, no leaving group; Appendix B)."),
        # the paragraph must stay two lines, above the footer band (y = 672)
        ("All campaigns hit the ceiling of the affinity model (pFAP 9.6–9.7) within a few "
         "generations; they differ in what they sacrifice. 361 of 384 multi-objective "
         "archive molecules are novel; top hits are plausible quinoline-carboxamide "
         "warheads — warheads only, not radioligands (no chelator, no leaving group; "
         "Appendix B).",
         "All campaigns hit the affinity-model ceiling (pFAP 9.6–9.7) in a few generations; "
         "361 of 384 archive molecules are novel, but the top hits are warheads, not "
         "radioligands — no chelator, no leaving group (Appendix B)."),
    ]),
]


def iter_text_paras(slide):
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        for para in shape.text_frame.paragraphs:
            yield para


def main():
    prs = Presentation(path)
    total = 0
    for idx, reps in FIXES:
        slide = prs.slides[idx]
        for old, new in reps:
            hit = False
            for para in iter_text_paras(slide):
                full = "".join(r.text for r in para.runs)
                if old in full:
                    new_full = full.replace(old, new)
                    para.runs[0].text = new_full
                    for r in para.runs[1:]:
                        r.text = ""
                    hit = True
                    break
            print(("OK  " if hit else "MISS") + f" slide {idx+1}: {old[:60]}...")
            total += hit
    prs.save(path)
    print(f"saved {path}; {total} replacements applied")


if __name__ == "__main__":
    main()
