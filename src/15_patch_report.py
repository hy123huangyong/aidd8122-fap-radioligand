"""
Final consistency patch for 09_build_report.py.

Applies the textual corrections as verified, single-match replacements.
Every patch asserts that its anchor occurs exactly once, so a later edit to the
report builder cannot silently drop a fix.  Re-running this script is a no-op
once the patches are in place (it detects the already-patched state).

Usage: python src/15_patch_report.py
"""
from __future__ import annotations

from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
TARGET = ROOT / "src" / "09_build_report.py"

# Each entry: (label, old, new).  `old` must match exactly once.
PATCHES: list[tuple[str, str, str]] = []

PATCHES.append((
    "P0.2 scaffold count in abstract (162 -> 161)",
    '"inhibitors (162 Murcko scaffolds) and benchmarked five models under random, scaffold and "',
    '"inhibitors (161 Murcko scaffolds) and benchmarked five models under random, scaffold and "',
))

PATCHES.append((
    "P1.2 bound the PK gap in 1.2 and point to Appendix B",
    '''         "an affinity problem and is not addressed here. This project concerns the first.")''',
    '''         "an affinity problem and is outside the ligand-based model built here — but its chemical "
         "preconditions are not unknowable. Appendix B works out, for the six generated structures, "
         "where a radionuclide could be attached, what a chelator would add to molecular weight "
         "and polar surface area, and why that modification would push a candidate outside the "
         "applicability domain established in Section 5.6. This project concerns the first problem; "
         "the second is recorded as a bounded, testable gap rather than left implicit.")
    para(doc,
         "The distinction matters because a generative model returns a warhead, not a "
         "radiopharmaceutical. A 18F label needs a displaceable leaving group on an aryl or alkyl "
         "carbon, while 68Ga and 177Lu need a macrocyclic chelator such as DOTA or NOTA. None of "
         "the six highest-ranked designs in Table 10 carries a chelator (Appendix B, Table B1), "
         "and a DOTA mono-amide conjugate would add about 358 Da, taking the median molecular "
         "weight from 499 to 857 Da and the topological polar surface area from 126 to about "
         "248 Å². Those are properties of a different chemical class from the 384 FAP inhibitors "
         "the model was trained on, so a predicted pFAP for the conjugated molecule would fall "
         "outside the applicability domain as well. Tumour retention and normal-tissue background "
         "cannot be inferred from this model; the testable next step is to synthesise a small "
         "chelator-bearing subset and measure it, rather than to extrapolate from the warhead.")''',
))

PATCHES.append((
    "P0.2 rewrite 3.3 success criteria as a three-row table",
    '''    para(doc,
         "The project is deliberately restricted to ligand-based modelling of the FAP warhead region; "
         "it does not model the chelator, the radiometal, or in vivo pharmacokinetics, and it does not "
         "perform docking or free-energy calculations. Success was defined in advance as: (i) a "
         "scaffold-split R² above 0.5 with a screening AUROC above 0.85 for the best model, which "
         "would justify use as a triage filter; (ii) a measurable and statistically defensible "
         "difference between modelling protocols; and (iii) conformal intervals whose empirical "
         "coverage can be predicted from the applicability domain. All three were met, although the "
         "third only conditionally.")''',
    '''    para(doc,
         "The project is deliberately restricted to ligand-based modelling of the FAP warhead region; "
         "it does not model the chelator, the radiometal, or in vivo pharmacokinetics, and it does not "
         "perform docking or free-energy calculations. Three success criteria were fixed before any "
         "model was fitted. Their disposition at the end of the project is reported explicitly below, "
         "because a pre-registered criterion is only useful if a partial or negative result is allowed "
         "to stand.")
    table(doc,
          ["Pre-registered criterion", "Threshold", "Disposition", "Evidence"],
          [["C1  Triage-grade accuracy", "scaffold-split R² > 0.5 and screening AUROC > 0.85",
            "Achieved", "Table 2: R² 0.68, AUROC 0.94 (XGBoost, five seeds)"],
           ["C2  Protocol effect", "a measurable, seed-replicated difference between protocols",
            "Achieved", "Table 2: R² 0.75 (random) to −0.79 (time); s.d. ≤ 0.06 per cell"],
           ["C3  Coverage predictable from the AD", "coverage near nominal in each AD class",
            "Partially achieved",
            "Table 7: 0.76 in-domain, 0.50 out-domain at nominal 0.80"]],
          caption="**Pre-registered success criteria and their disposition.** C3 is deliberately "
                  "recorded as partial: the applicability domain does predict where conformal "
                  "coverage fails, but the intervals themselves are not valid out of domain, which "
                  "is a weaker outcome than the criterion anticipated.",
          widths=[3.9, 4.6, 2.5, 5.0], size=8.0)
    para(doc,
         "Stating C3 as partial rather than met is the point of the exercise: the analysis that "
         "validated the first two criteria is the same analysis that refused the third.")''',
))

PATCHES.append((
    "P0.2 single 0.6 AD threshold declared in 4.7",
    '''         "which yields an operational rule (a similarity threshold) rather than a binary claim.")''',
    '''         "which yields an operational rule (a similarity threshold) rather than a binary claim. "
         "A single threshold of 0.60 separates in-domain from out-of-domain compounds throughout "
         "this report: it is the boundary at which RMSE begins to rise in Table 6 and at which "
         "out-of-domain conformal coverage falls to 0.54 in Table 7. The narrower 0.1-wide bins "
         "in Table 6 are descriptive partitions, not alternative decision rules; the 0.40 filter "
         "in Section 4.9 is a novelty constraint on the generator and is a different quantity.")''',
))

PATCHES.append((
    "P0.1 update 4.10 implementation and reproducibility",
    '''    para(doc,
         "All code is Python 3.13 with RDKit 2026.09.1 [27], scikit-learn 1.9.1, XGBoost 3.4.1 and "
         "NumPy 2.5.4. The pipeline is organised as numbered scripts (fetch, curate, benchmark, "
         "transfer, analysis, generative, figures, report), each writing a table or figure that the "
         "next stage reads. The curated datasets, the full set of predictions and the code are "
         "available in the repository listed in Appendix A.")''',
    '''    para(doc,
         "All code is Python 3.13 with RDKit 2026.09.1 [27], scikit-learn 1.9.1, XGBoost 3.4.1, "
         "NumPy 2.5.4, pandas 3.1.0 and Matplotlib 3.11.2. The pipeline is organised as numbered "
         "scripts: 01 fetch, 02 curate, 03 benchmark, 04 transfer/selectivity, 05 trust analysis, "
         "06 generative design, 07-08 figures, 09 this report, 10 presentation notes, 11 DOI "
         "verification, 12 retrosynthetic/radiolabelling triage, 13 robustness analyses, 14 "
         "summary reconciliation and 15 this optimisation pass. Each script writes a table or "
         "figure that the next stage reads; random seeds, package versions and run logs are shipped "
         "with the code. The curated datasets, the raw ChEMBL exports, the full prediction tables, "
         "the environment fingerprint and a one-command reproduction guide are listed in Appendix A.")''',
))

PATCHES.append((
    "P1.3 paired bootstrap and stratified re-analysis after Figure 4",
    '''           "twenty seeds. (c) Distribution of ΔpActivity for the dual-measured compounds against PREP "
           "and DPP4.")''',
    '''           "twenty seeds. (c) Distribution of ΔpActivity for the dual-measured compounds against PREP "
           "and DPP4. ChEMBL release 37; five benchmark seeds, twenty selectivity seeds; n = 100 "
           "FAP/PREP and 232 FAP/DPP4 dual-measured compounds.")
    para(doc,
         "Because the cliff is the strongest negative claim in this report, it was re-tested with a "
         "paired bootstrap over the twenty seeds (Appendix C, Table C1). Pairing the random- and "
         "scaffold-split AUROC obtained on the same seed gives a mean difference of 0.463 "
         "(95% CI 0.392–0.532) for the indirect protocol and 0.454 (0.382–0.525) for the direct "
         "one; 100% of 5,000 bootstrap replicates are positive. The collapse is therefore not an "
         "artefact of seed variation.")
    para(doc,
         "A stratified re-analysis is less comforting (Appendix C, Table C2). The cliff persists in "
         "the IC50-only stratum (Δ = 0.451, n = 70) and in the single-measurement stratum "
         "(Δ = 0.523, n = 93), but it reverses in the small mixed-assay stratum "
         "(Δ = −0.158, n = 30, only 13 evaluable random folds). The replicated stratum contains "
         "just 7 compounds and cannot be fitted at all. Assay-type confounding cannot therefore be "
         "excluded for that subgroup; the honest reading is that the cliff is robust for the "
         "IC50-dominated bulk of the data and unresolved for the heterogeneous remainder.")''',
))

PATCHES.append((
    "P0.2 unify 5.6 wording below-threshold to 0.6",
    '''         "is the basis for a usable rule: a prediction supported by a similarity above 0.6 can be "
         "quoted with an error comparable to assay noise, and below 0.5 it cannot.")''',
    '''         "is the basis for a usable rule: a prediction supported by a similarity above 0.6 can be "
         "quoted with an error comparable to assay noise, and below 0.6 it cannot.")''',
))

PATCHES.append((
    "P1.4 Mondrian conformal sensitivity analysis after 5.6 rule",
    '''         "may still be useful for ranking a series internally, but its numeric interval should not "
         "appear in a synthesis decision.")''',
    '''         "may still be useful for ranking a series internally, but its numeric interval should not "
         "appear in a synthesis decision.")
    para(doc,
         "A natural repair is class-conditional (Mondrian) conformal calibration: bin the "
         "out-of-fold calibration residuals by how well each training compound is represented in "
         "the training chemistry, and let each bin use its own quantile. Appendix C, Table C2 and "
         "Figure 8 report that sensitivity analysis. Out-of-domain coverage at the 80% level rises "
         "from 0.54 to 0.74 and at the 90% level from 0.77 to 1.00, while in-domain coverage is "
         "essentially unchanged (0.77 to 0.74 and 0.88 to 0.88).")
    para(doc,
         "This is a diagnostic, not a deployable fix. Only 27 of roughly 280 training compounds "
         "fall in the out-of-domain calibration bin, so its quantile is noisy: at the 90% level it "
         "produces complete over-coverage (1.00) and more than doubles the interval width from 2.40 "
         "to 5.01 log units. Conditional calibration restores nominal coverage by spending "
         "precision, which is exactly the trade-off the rest of this report argues should be "
         "reported rather than hidden behind a single reassuring number.")
    figure(doc, FIG / "fig8_mondrian.png",
           "**Figure 8.** Class-conditional conformal calibration under scaffold splitting "
           "(mean over five seeds). Dark bars: global split conformal; light bars: bin-wise "
           "quantiles. The dashed line is the nominal level. ChEMBL release 37; XGBoost, "
           "five seeds, 0.6 applicability-domain threshold; n ≈ 280 training / ≈ 100 test "
           "compounds per seed.")''',
))

# ---------------------------------------------------------------- 5.8 / Table 9
PATCHES.append((
    "P0.2 add median novel pFAP to Table 9",
    '''                     "mw": float(d.mw.median()),
                     "best": float(nov.pFAP.max()),
                     "pass": float(len(ok) / len(nov)) if len(nov) else float("nan")}
    table(doc,
          ["Campaign", "Archived", "Novel", "Scaffolds", "Median QED", "Median SA",
           "Median MW (Da)", "Best novel pFAP", "Pass rate"],
          [[stats[mo]["lab"], stats[mo]["n"], stats[mo]["novel"], stats[mo]["scaf"],
            f"{stats[mo]['qed']:.2f}", f"{stats[mo]['sa']:.2f}", f"{stats[mo]['mw']:.0f}",
            f"{stats[mo]['best']:.2f}", f"{stats[mo]['pass']*100:.0f}%"]
           for mo in ["ga", "affinity_only", "random"]],
          caption="**Table 9.** Outcome of the three design campaigns with an identical evaluation "
                  "budget (24 generations, population 96). Novelty = InChIKey absent from all three "
                  "curated datasets; pass rate = fraction of novel molecules with predicted pFAP ≥ 8 "
                  "and selectivity margin ≥ 1.",
          widths=[4.4, 1.7, 1.4, 1.8, 1.9, 1.7, 2.1, 2.2, 1.6], size=8.5)''',
    '''                     "mw": float(d.mw.median()),
                     "best": float(nov.pFAP.max()),
                     "mednov": float(nov.pFAP.median()),
                     "pass": float(len(ok) / len(nov)) if len(nov) else float("nan")}
    table(doc,
          ["Campaign", "Arch.", "Novel", "Scaf.", "QED", "SA", "MW (Da)",
           "Median novel pFAP", "Best novel pFAP", "Pass rate"],
          [[stats[mo]["lab"], stats[mo]["n"], stats[mo]["novel"], stats[mo]["scaf"],
            f"{stats[mo]['qed']:.2f}", f"{stats[mo]['sa']:.2f}", f"{stats[mo]['mw']:.0f}",
            f"{stats[mo]['mednov']:.2f}", f"{stats[mo]['best']:.2f}",
            f"{stats[mo]['pass']*100:.0f}%"]
           for mo in ["ga", "affinity_only", "random"]],
          caption="**Table 9.** Outcome of the three design campaigns with an identical evaluation "
                  "budget (24 generations, population 96). Novelty = InChIKey absent from all three "
                  "curated datasets; pass rate = fraction of novel molecules with predicted pFAP ≥ 8 "
                  "and selectivity margin ≥ 1. QED, SA and MW are medians over the whole archive; "
                  "the two pFAP columns are over novel molecules only. All values are recomputed "
                  "from generative_archive.csv by this script.",
          widths=[3.2, 1.2, 1.2, 1.1, 1.1, 1.0, 1.4, 2.2, 1.9, 1.3], size=8.0)''',
))

PATCHES.append((
    "P0.2 explain the 100% pass rate and the multi-objective trade-off",
    '''    para(doc,
         "The multi-objective campaign pays a small affinity price (median novel pFAP 8.24 versus "
         "9.36 for affinity-only) and in exchange returns the most usable chemistry of the three "
         "regimes: 115 distinct Murcko scaffolds against 80–83 for the controls, a median QED of "
         "0.61, a median synthetic-accessibility score of 2.93 (easier to make than either control) "
         "and a median molecular weight of 413 Da, squarely in the range of the clinical FAPIs. "
         "61% of its novel molecules clear the pre-registered design thresholds of pFAP ≥ 8 and "
         "selectivity margin ≥ 1, essentially the same pass rate as the random control (60%) — the "
         "value of the objectives shows up in diversity and developability rather than in the pass "
         "rate, which is exactly what a Pareto front should do. The top-ranked novel molecules are "
         "quinoline-carboxamide derivatives with modified amide linkers, i.e. plausible next analogues "
         "rather than exotic chemistry (Table 10).")''',
    '''    para(doc,
         "The multi-objective campaign pays a small affinity price (median novel pFAP 8.24 in "
         "Table 9 versus 9.36 for affinity-only) and in exchange returns the most usable chemistry "
         "of the three regimes: 115 distinct Murcko scaffolds against 80–83 for the controls, a "
         "median QED of 0.61, a median synthetic-accessibility score of 2.93 (easier to make than "
         "either control) and a median molecular weight of 413 Da, squarely in the range of the "
         "clinical FAPIs. Its 61% pass rate is essentially the same as the random control (60%) — "
         "the value of the objectives shows up in diversity and developability rather than in the "
         "pass rate, which is exactly what a Pareto front should do.")
    para(doc,
         "Two features of Table 9 need to be read with care. First, the affinity-only campaign "
         "shows a 100% pass rate, and this is a consequence of the metric's definition, not of "
         "superior chemistry: selection drives all 369 of its novel molecules above pFAP 8, and "
         "because the margin is pFAP minus the largest paralogue prediction, a high pFAP almost "
         "automatically satisfies margin ≥ 1 as well. The pass rate is therefore close to a "
         "restatement of the objective that campaign was given. Second, the same circularity "
         "explains why adding selectivity and developability pressure lowers the pass rate to "
         "61%: the optimiser is no longer allowed to spend its entire budget on pFAP. A post-hoc "
         "look at what the four-objective budget bought is more informative than the pass rate — "
         "the affinity-only archive has median QED 0.25, median MW 548 Da and 83 scaffolds, "
         "whereas the four-objective archive has median QED 0.61, median MW 413 Da and 115 "
         "scaffolds at a median pFAP of 8.24 that still clears the pre-registered threshold. The "
         "top-ranked novel molecules are quinoline-carboxamide derivatives with modified amide "
         "linkers, i.e. plausible next analogues rather than exotic chemistry (Table 10), but see "
         "Section 5.10 and Appendix B for why their warhead is not the clinical FAPI warhead.")''',
))

PATCHES.append((
    "P0.2 add warhead-mechanism and structure-method boundaries to 5.10",
    '''        "test would require an external, later dataset.",
    ]:
        bullet(doc, b)''',
    '''        "test would require an external, later dataset.",
        "Five of the six highest-ranked designs in Table 10 carry an alpha-diketone / ketoamide "
        "electrophile, whereas clinical FAPI tracers use a prolyl nitrile warhead that forms a "
        "reversible covalent adduct with the catalytic serine. The generated molecules are "
        "therefore not close analogues of the clinical series: pFAP values above 9.5 place them "
        "inside the model's confident region, not inside a validated SAR series, and imply "
        "neither potency nor synthetic tractability (Appendix B).",
        "No protein-structure, docking, pharmacophore or free-energy calculation was performed. "
        "The assignment brief lists molecular docking and structure prediction as possible "
        "directions, not as required deliverables, so this project stays with ligand-based "
        "methods. Adding a docking score would create a second, uncalibrated confidence scale "
        "and would contradict the audit framing of this report; if structure-based work is "
        "undertaken later it should be used only for qualitative binding-mode explanations and "
        "be reported with the same applicability-domain and failure checks as the models here.",
        "The selectivity cliff is robust in the IC50-dominated bulk of the paired data but "
        "reverses in a 30-compound mixed-assay stratum (Appendix C, Table C2); assay-type "
        "confounding in that subgroup cannot be excluded.",
    ]:
        bullet(doc, b)''',
))

PATCHES.append((
    "P0.4 append verified DOI to each reference",
    '''    for i, ref in enumerate(refs, 1):
        p = doc.add_paragraph()''',
    '''    for i, ref in enumerate(refs, 1):
        doi = D["doi_map"].get(i, "")
        if doi:
            ref = f"{ref.rstrip('.')} doi:{doi}."
        p = doc.add_paragraph()''',
))

PATCHES.append((
    "P0.1 rewrite Appendix A as a real evidence-package description",
    '''    para(doc,
         "The complete project — numbered analysis scripts, curated datasets, prediction tables and "
         "figure-generation code — is available at:")
    para(doc, repo_url(), align="center", italic=True)
    para(doc,
         "The repository is organised as src/01_fetch_chembl.py through src/09_build_report.py, with "
         "data/raw holding the retrieved ChEMBL records and results/tables holding every number "
         "reported in this document. Running the scripts in numerical order reproduces all figures "
         "and tables from the public API; a pinned requirements file and a short README are included.")''',
    '''    para(doc,
         "The complete project is submitted as a single package, "
         "Submission_HuangYong_P2600222/. Its 03_Code/ directory is the evidence bundle referenced "
         "by this report:")
    para(doc,
         "03_Code/README.md — one-command reproduction, table-to-script map and environment "
         "fingerprint; 03_Code/requirements.txt — pinned dependency versions; 03_Code/src/ — "
         "scripts 01–15 in execution order; 03_Code/data/raw/ — the three ChEMBL release-37 JSON "
         "exports together with their retrieval date and API query strings; 03_Code/data/processed/ "
         "— the 972 → 438 → 384 curation chain; 03_Code/results/tables/ — every number quoted in "
         "Tables 1–10 and Appendices B–C; 03_Code/results/figures/ — all eight figures at "
         "≥2,000 px; 03_Code/logs/ — random seeds, package versions, runtime and the reconciliation "
         "log. Running the scripts in numerical order from README.md regenerates Tables 1, 2, 8, "
         "9 and 10 from the public API without any manual step.")
    para(doc,
         "Code and data are additionally mirrored in the accessible version-control repository "
         "at " + repo_url() + ". The ChEMBL snapshot is release 37 (May 2026), targets "
         "CHEMBL4683 / CHEMBL3202 / CHEMBL284; the environment is Python 3.13, RDKit 2026.09.1 "
         "[27], scikit-learn 1.9.1, XGBoost 3.4.1, NumPy 2.5.4, pandas 3.1.0 and Matplotlib 3.11.2, "
         "on Windows 11. All five benchmark seeds and all twenty selectivity seeds are fixed in "
         "the scripts.")''',
))


PATCHES.append((
    "P0.2 fix 5.6 cross-reference to Table C3",
    "Appendix C, Table C2 and ",
    "Appendix C, Table C3 and ",
))

PATCHES.append((
    "P1.1/P1.2/P1.3/P1.4 appendices B and C",
    '''         "on Windows 11. All five benchmark seeds and all twenty selectivity seeds are fixed in "
         "the scripts.")''',
    '''         "on Windows 11. All five benchmark seeds and all twenty selectivity seeds are fixed in "
         "the scripts.")

    # -------------------------------------------------------- appendix B
    heading(doc, "Appendix B. Radioligand feasibility and retrosynthetic triage", 1)
    para(doc,
         "This appendix closes the gap left open in Section 1.2. It is a rule-based structural "
         "analysis, not a literature-verified synthesis: BRICS disconnections, reaction "
         "assignments and step counts are heuristics computed from the molecular graph by "
         "src/12_design_feasibility.py, and no reaction, yield or supplier check was performed. "
         "Descriptors are RDKit calculations; the chelator figures are stoichiometric estimates "
         "for a DOTA mono-amide (DOTA free acid 376.4 Da; conjugation adds 358.4 Da).")
    lab = D["radiolabel"]
    table(doc,
          ["ID", "Warhead", "MW (Da)", "TPSA (Å²)", "Chelator", "MW + DOTA (Da)"],
          [[r.id, r.warhead, f"{r.MW_Da:.1f}", f"{r.TPSA:.1f}",
            r.chelator_present, f"{r.MW_after_DOTA_est:.1f}"] for r in lab.itertuples()],
          caption="**Table B1.** Calculated properties of the six highest-ranked novel designs "
                  "(Table 10) and the estimated effect of conjugating one DOTA chelator. None of "
                  "the six carries a chelator; median MW rises from 498.6 to 856.9 Da and median "
                  "TPSA from 126.0 to 247.7 Å², which is outside the training chemistry of "
                  "Table 1. Source: radiolabel_feasibility.csv.",
          widths=[1.6, 3.2, 1.8, 1.9, 2.0, 2.7], size=8.0)
    para(doc,
         "The labelling handles lead to the same conclusion. A 18F label requires a displaceable "
         "leaving group on an aryl or alkyl carbon; the fluorine atoms present in these designs "
         "are stable aryl-F or gem-difluoro substituents, not leaving groups. 68Ga and 177Lu "
         "require a macrocyclic chelator, and every nitrogen and oxygen in the six structures is "
         "already engaged in an amide or in the electrophilic warhead, so there is no free amine "
         "or carboxylate for a chelator coupling. All three radionuclide routes therefore demand "
         "a structural modification first — and the modified molecule would no longer resemble "
         "the 384 compounds that define the model's applicability domain. This is why tumour "
         "retention and normal-tissue background are declared out of scope here rather than "
         "predicted: they are pharmacokinetic properties of a conjugate this model has never "
         "seen. The testable next step is narrow — synthesise three to five chelator-bearing "
         "analogues, measure affinity and logD, and check whether they still fall inside the "
         "0.6 similarity gate.")
    ret = D["retro"]
    table(doc,
          ["ID", "Warhead", "BRICS frags", "Est. steps", "SA", "Dominant disconnection"],
          [[r.id,
            "alpha-diketone / ketoamide" if "alpha" in r.warhead else "prolyl nitrile",
            int(r.brics_fragments), int(r.estimated_steps), f"{r.SA_score:.2f}",
            r.key_reactions.split(";")[0].strip()] for r in ret.itertuples()],
          caption="**Table B2.** Rule-based retrosynthetic triage of the six designs. BRICS "
                  "fragments are graph disconnections; the step count is a heuristic "
                  "(one coupling per junction plus one interconversion per specialised warhead), "
                  "not an experimental route. Source: retrosynthesis.csv.",
          widths=[1.5, 3.2, 1.7, 1.6, 1.1, 5.1], size=8.0)
    para(doc,
         "The medicinal-chemistry answer to the warhead question is therefore: yes, an "
         "alpha-diketone is accessible (oxidation of the corresponding alpha-hydroxy ketone, "
         "followed by amide couplings onto a quinoline-4-carboxamide and a proline-derived "
         "amine), but it is not the clinical FAPI warhead. Clinical tracers use a prolyl nitrile "
         "that forms a reversible covalent adduct with the catalytic serine; only the sixth "
         "design in Table 10 retains that motif. A high predicted pFAP for the first five "
         "consequently says that the optimizer found the model's confident region for a "
         "different chemotype, not that it rediscovered FAPI-04. This appendix supplies a "
         "triage and a set of checkable questions, not a claimed synthesis.")

    # -------------------------------------------------------- appendix C
    heading(doc, "Appendix C. Robustness analyses", 1)
    para(doc,
         "All three analyses below reuse the existing ChEMBL release-37 tables generated by "
         "src/13_robustness.py: no external dataset and no new model class was introduced. They "
         "were added to test, rather than to decorate, the two strongest claims in the report — "
         "the selectivity cliff and the out-of-domain conformal failure.")
    boot = D["boot"]
    table(doc,
          ["Protocol", "Paired seeds", "AUROC random [95% CI]",
           "AUROC scaffold [95% CI]", "Paired Δ [95% CI]"],
          [[r.protocol, int(r.n_paired_seeds),
            f"{r.AUROC_random:.3f} {r.random_CI95}",
            f"{r.AUROC_scaffold:.3f} {r.scaffold_CI95}",
            f"{r.paired_delta:.3f} [{r.delta_CI95_low:.3f}, {r.delta_CI95_high:.3f}]"]
           for r in boot.itertuples()],
          caption="**Table C1.** Paired bootstrap over twenty seeds (5,000 replicates, seed 0 "
                  "paired with seed 0). The random- and scaffold-split AUROCs are differenced "
                  "within the same seed, so the interval describes the cliff itself. Source: "
                  "selectivity_bootstrap.csv.",
          widths=[2.2, 2.0, 3.7, 3.7, 3.0], size=8.0)
    strat = D["strat"]
    table(doc,
          ["Stratum", "n", "AUROC random", "AUROC scaffold", "Δ", "Folds (r / s)"],
          [[r.stratum, int(r.n), fmt(r.AUROC_random, 3), fmt(r.AUROC_scaffold, 3),
            fmt(r.cliff_delta, 3), f"{fmt(r.n_folds_random, 0)} / {fmt(r.n_folds_scaffold, 0)}"]
           for r in strat.itertuples()],
          caption="**Table C2.** Stratified re-analysis of the selectivity experiment using the "
                  "identical twenty-seed protocol, restricted to each stratum. Strata with "
                  "n < 30 have unstable AUROCs; the mixed-assay stratum reverses the cliff and "
                  "the replicated stratum (n = 7) cannot be fitted. Source: "
                  "selectivity_stratified.csv.",
          widths=[4.4, 1.1, 2.3, 2.4, 1.4, 2.6], size=8.0)
    mc = (D["mondrian"].groupby(["level", "domain"], as_index=False)
          [["coverage_global", "coverage_mondrian", "width_global", "width_mondrian"]].mean())
    table(doc,
          ["Nominal", "Domain", "Global coverage", "Mondrian coverage",
           "Global width", "Mondrian width"],
          [[f"{r.level:.0%}", r.domain.replace("_", " "),
            f"{r.coverage_global:.3f}", f"{r.coverage_mondrian:.3f}",
            f"{r.width_global:.2f}", f"{r.width_mondrian:.2f}"] for r in mc.itertuples()],
          caption="**Table C3.** Class-conditional (Mondrian) conformal calibration, mean over "
                  "five scaffold-split seeds. Calibration residuals are binned by each training "
                  "compound's maximum Tanimoto similarity to the other training compounds, and "
                  "test compounds use the bin matching their own similarity to the training set "
                  "at the 0.6 threshold. Only about 27 training compounds fall in the "
                  "out-of-domain bin. Source: mondrian_conformal.csv.",
          widths=[1.8, 2.8, 2.5, 2.7, 2.2, 2.4], size=8.0)
    para(doc,
         "Read together, these tables strengthen the first claim and qualify the repair of the "
         "second. The paired interval excludes zero in every replicate, so the cliff is not seed "
         "noise — but it reverses in a 30-compound mixed-assay stratum, so assay-type "
         "confounding remains a live limitation rather than a dismissed one. Mondrian "
         "calibration lifts out-of-domain coverage without touching in-domain coverage, yet it "
         "does so with a 27-member calibration bin whose 90% quantile over-covers completely "
         "(1.00) at more than twice the interval width. Conditional calibration is thus a "
         "diagnostic of where exchangeability fails, not a new interval to quote in a synthesis "
         "meeting. That is the same conclusion Section 5.6 reaches by a different route.")''',
))


# ------------------------------------------------------- P0.3 figure captions
_CAPTIONS = [
    ("Figure 1",
     '''           "The dashed arrow indicates that error analysis feeds back into feature and model choice.")''',
     '''           "The dashed arrow indicates that error analysis feeds back into feature and model "
           "choice. ChEMBL release 37; RDKit 2026.09.1 / Matplotlib 3.11.2; schematic of the "
           "01–09 pipeline, no plotted observations.")'''),
    ("Figure 2",
     '''           "relative to its compound count.")''',
     '''           "relative to its compound count. ChEMBL release 37; RDKit 2026.09.1; n = 384 / 512 / "
           "4,338 compounds (FAP / PREP / DPP4); panel b n = 237 (2004–2018) and 147 (2019–2024).")'''),
    ("Figure 3",
     '''           "focused evaluation set; the dashed line marks the no-enrichment level.")''',
     '''           "focused evaluation set; the dashed line marks the no-enrichment level. ChEMBL "
           "release 37; five seeds × five models × three protocols, n = 384 FAP compounds; "
           "XGBoost 3.4.1 / scikit-learn 1.9.1.")'''),
    ("Figure 5",
     '''           "applicability class; the dashed line is the overall RMSE.")''',
     '''           "applicability class; the dashed line is the overall RMSE. ChEMBL release 37; "
           "scaffold split, XGBoost, five seeds pooled, n = 384 FAP compounds; 0.6 AD threshold.")'''),
    ("Figure 6",
     '''           "recalibrated points move onto the diagonal in level but remain uncorrelated in rank.")''',
     '''           "recalibrated points move onto the diagonal in level but remain uncorrelated in "
           "rank. ChEMBL release 37; train ≤ 2018 (n = 237), test 2019–2024 (n = 147); single "
           "fit at seed 0, XGBoost 3.4.1.")'''),
    ("Figure 7",
     '''           "archived molecules by campaign.")''',
     '''           "archived molecules by campaign. ChEMBL release 37; 24 generations × population "
           "96, 384 archived molecules per campaign; RDKit 2026.09.1, five-seed XGBoost "
           "ensemble scoring.")'''),
]
for _label, _old, _new in _CAPTIONS:
    PATCHES.append((f"P0.3 caption n/version: {_label}", _old, _new))

PATCHES.append((
    "P0.1 natural wording for the version-control location in Appendix A",
    '''    para(doc,
         "Code and data are additionally mirrored in the accessible version-control repository "
         "at " + repo_url() + ". The ChEMBL snapshot is release 37 (May 2026), targets "''',
    '''    para(doc,
         "The ChEMBL snapshot is release 37 (May 2026), targets "''',
))

PATCHES.append((
    "P0.1 separate version-control line in Appendix A",
    '''         "the scripts.")

    # -------------------------------------------------------- appendix B''',
    '''         "the scripts.")
    para(doc, "Version-control location: " + repo_url(), align="center", italic=True)

    # -------------------------------------------------------- appendix B''',
))


def main() -> None:
    s = TARGET.read_text(encoding="utf-8")
    applied, skipped, failed = 0, 0, []
    for label, old, new in PATCHES:
        # Check the post-patch text FIRST: several patches append content after
        # their anchor, so `old` remains a prefix of `new` and would otherwise
        # be re-applied on every run.
        if new in s:
            skipped += 1
            print(f"  SKIP {label} (already applied)")
            continue
        n = s.count(old)
        if n == 1:
            s = s.replace(old, new, 1)
            applied += 1
            print(f"  OK   {label}")
        else:
            failed.append(label)
            print(f"  FAIL {label}: {n} anchor matches")
    if failed:
        raise SystemExit(f"{len(failed)} patches could not be applied; source left unchanged")
    TARGET.write_text(s, encoding="utf-8")
    print(f"\n{applied} patches applied, {skipped} already present -> {TARGET}")


if __name__ == "__main__":
    main()
