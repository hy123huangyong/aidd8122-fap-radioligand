# -*- coding: utf-8 -*-
"""Condense the v2 report prose without touching its structure.

One idempotent, paragraph-level text pass: long expository paragraphs are
rewritten more tightly (shorter sentences, no restatement of table values,
no throat-clearing). Headings, figure/table captions, the reference list,
Appendix A paths and the live repository hyperlink are deliberately excluded,
so the TOC, cross-references, three-line tables and pagination stay intact.

Every replacement is anchored by paragraph index AND a prefix of the original
text; the pass aborts rather than guessing if the document has moved. Before
saving it asserts that the multiset of ``Table n`` / ``Figure n`` / ``[n]``
references is unchanged, so no citation can be lost during editing.

Usage:
    python src/27_condense_report.py            # rewrite docx and re-export PDF
    python src/27_condense_report.py --check    # verify the condensed state
    python src/27_condense_report.py --no-pdf   # rewrite docx only
"""
from __future__ import annotations

import argparse
import re
import shutil
import sys
import tempfile
import time
from collections import Counter
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DOCX = ROOT / "report" / "AIDD8122_Final_Report_HuangYong_P2600222_v2.docx"
PDF = ROOT / "report" / "AIDD8122_Final_Report_HuangYong_P2600222_v2.pdf"

# index -> (unique prefix of the ORIGINAL paragraph, condensed replacement)
PATCHES: dict[int, tuple[str, str]] = {
    # ---- Abstract -------------------------------------------------------
    62: ("Fibroblast activation protein (FAP) is one of", """Fibroblast activation protein (FAP) is a leading nuclear-medicine target, but its public structure–activity record is small, assay-heterogeneous and chemically non-stationary. We ask what machine learning can deliver for FAP radioligand design, and where a chemist should not trust it. We curated 972 ChEMBL bioactivities into 384 unique inhibitors (161 Murcko scaffolds) and benchmarked five models under random, scaffold and time splits. Random splitting looks excellent (XGBoost R² = 0.75); a prospective time split collapses to R² = −0.79 because post-2019 papers describe a different series (median MW +80 Da, potency +2.0 log units). Selectivity over PREP and DPP4 is predictable within known scaffolds (AUROC 0.84) but anti-correlated on new ones (0.38). Split-conformal intervals cover 79% and 89% at nominal 80% and 90% under random splitting, but only 54% of out-of-domain compounds at 80%. A graph genetic algorithm with NSGA-II selection designed warheads under four objectives; applicability-domain and selectivity gates changed the output more than the optimizer. For targets of this size, the usable deliverable is a calibrated decision rule stating when a prediction may be acted on, not a single predicted number."""),

    # ---- 1. Introduction ------------------------------------------------
    66: ("In desmoplastic tumours such as pancreatic", """In desmoplastic tumours such as pancreatic ductal adenocarcinoma, breast carcinoma and sarcoma, cancer-associated fibroblasts (CAFs) can dominate tumour mass and deposit the collagen-rich stroma that limits drug penetration. FAP is a type II transmembrane serine protease of the S9B prolyl-oligopeptidase family, displayed on activated fibroblasts and largely absent from healthy tissue [9]. Its catalytic site accepts small drug-like inhibitors, so FAP became the first stromal target with a purely small-molecule theranostic pair: one quinoline inhibitor with a prolyl warhead can be labelled with 68Ga or 18F for imaging and 90Y or 177Lu for therapy, changing only the chelator [1,6]."""),
    67: ("The clinical trajectory has been steep.", """The clinical trajectory has been steep. FAPI-04, optimised in 2018 [1], gave high-contrast PET images across tumour entities [3,4,10]; FAPI-46 and FAPI-74 improved retention and enabled 18F-AlF labelling [5,7]. Several hundred patients have been imaged, and 90Y-FAPI-46, 177Lu-FAP-2286 and albumin-binding 177Lu-LNC1004 have entered early-phase trials [6–8]. For nuclear medicine this is now a ligand-design, not a target-validation, problem."""),
    69: ("Imaging performance has not translated", """Imaging success has not translated into therapeutic benefit. In published FAPI radioligand-therapy series, disease control reaches 20–80% but objective responses are uncommon (14%, 4 of 28, versus 46% disease control in the largest phase II cohort); grade 3/4 haematological toxicity occurred in 21%, making myelosuppression the practical constraint [7,8,29]. Tumour retention is shorter than for PSMA ligands, and uptake in salivary glands, pancreas and marrow creates a background that dosimetry cannot remove. The levers are well defined — warhead (nitrile versus boronic acid), linker, chelator, and an Evans blue motif for albumin binding — but every combination must be synthesised, labelled, assayed and imaged in mice. One design–test cycle costs weeks plus a full animal experiment."""),
    70: ("Two selectivity problems dominate", """Two selectivity problems dominate. First, selectivity against the close paralogues prolyl endopeptidase (PREP) and dipeptidyl peptidase IV (DPP4/CD26), which share the catalytic fold and are widely expressed; DPP4 inhibition would add a systemic liability no dosimetry can offset. Second, selectivity against FAP in healthy tissue — a pharmacokinetic problem, outside this ligand-based model, though its chemical preconditions are knowable. Appendix B shows where a radionuclide could be attached to the six generated structures, what a chelator adds to molecular weight and polar surface area, and why that moves them outside the applicability domain of Section 5.6. This project addresses the first problem and records the second as a bounded, testable gap."""),
    71: ("The distinction matters because a generative", """The distinction matters: a generative model returns a warhead, not a radiopharmaceutical. 18F needs a displaceable leaving group; 68Ga and 177Lu need a macrocyclic chelator such as DOTA or NOTA. None of the six top-ranked designs in Table 10 carries one (Appendix B, Table B1). A DOTA mono-amide would add about 386 Da, moving median molecular weight from 499 to 885 Da and TPSA from 126 to about 248 Å² — a different chemical class from the 384 training inhibitors, so a predicted pFAP for the conjugate also falls outside the applicability domain. Tumour retention and background cannot be inferred from this model. The testable next step is to synthesise and measure a small chelator-bearing subset, not to extrapolate from the warhead."""),
    73: ("Public bioactivity data for FAP are", """Public FAP bioactivity data are an order of magnitude smaller than kinase datasets: ChEMBL holds a few hundred unique inhibitors from sixty-odd publications across 78 assays [11,28]. Published FAP QSAR models report random-split R² of 0.7–0.8 [12,14], inviting the conclusion that affinity prediction is solved. A chemist needs something else: will the model rank a new chemotype correctly, and will its error bars hold when the literature moves on? Random splitting answers neither, because near-duplicate analogues remain in the training set."""),
    74: ("This project therefore treats the evaluation", """We therefore treat the evaluation protocol as the object of study. The same models are tested under three protocols matching three real decisions — interpolation within a known series, scaffold hopping, and prospective use on later literature — and we ask under which of them the output is actionable."""),
    76: ("•\u2003A reproducible curation of the public", """•\u2003A reproducible curation of the public FAP, PREP and DPP4 bioactivity record, with an audit trail from 972 raw records to 384 compounds and an assay-noise estimate that sets the error floor."""),
    77: ("•\u2003A three-protocol benchmark", """•\u2003A three-protocol benchmark (random, scaffold, time) of five model families with identical features and hyperparameters, reporting screening as well as regression metrics."""),
    78: ("•\u2003A quantitative test of whether", """•\u2003A quantitative test of whether paralogue bioactivity transfers to FAP, and whether selectivity can be predicted outside the training scaffolds."""),
    79: ("•\u2003A trust analysis that combines", """•\u2003A trust analysis combining applicability domain, conformal intervals and temporal recalibration, localising when a predicted interval is honest."""),
    80: ("•\u2003A multi-objective generative design", """•\u2003A multi-objective generative design campaign (graph GA with NSGA-II) compared against single-objective and random controls, with novelty and applicability tracked."""),

    # ---- 2. Literature Review ------------------------------------------
    83: ("The modern FAP inhibitor series", """Modern FAP inhibitors derive from quinoline-2-carboxamides with a prolyl nitrile warhead, a linker and a chelator. Lindner et al. optimised fifteen candidates into FAPI-04, the first clinical tracer [1], with better serum stability and FAP/CD26 selectivity than FAPI-02. Loktev et al. showed high-contrast imaging of CAF-rich tumours [2], and Kratochwil et al. documented uptake across 28 tumour entities [3], establishing FAP as a pan-tumour stromal target. FAPI-46 improved retention [7]; FAPI-74 enabled one-step 18F-AlF labelling [5], removing the dependence on a 68Ga generator."""),
    84: ("The therapeutic branch followed quickly.", """The therapeutic branch followed: 90Y-FAPI-46 in advanced sarcoma [7], 177Lu-FAP-2286 as the first DOTA-bearing FAP ligand in a first-in-human dosimetry study [6], and 177Lu-LNC1004 with an Evans blue albumin-binding motif [8]. Across these series imaging contrast is excellent, objective response rates low, and the field attributes the gap to short intratumoural retention. For design, the consequence is clear: affinity alone is insufficient, and paralogue selectivity must be designed in rather than measured afterwards."""),
    86: ("On ChEMBL-scale benchmarks, gradient", """On ChEMBL-scale benchmarks, gradient-boosted trees over circular fingerprints remain competitive with graph neural networks when measurements per target number in the hundreds [26]. Mayr et al. showed on 1,300 targets that deep architectures only win when data are abundant [12], and van Tilborg et al. found much of the apparent advantage of learned representations disappears once activity cliffs and scaffold novelty are accounted for [13]. For FAP, with fewer than 500 unique compounds, a tuned tree ensemble over combined fingerprints and descriptors is therefore the rational baseline against which any deep model must justify itself."""),
    87: ("A second line of work uses transfer", """Transfer learning is the second line of work. PREP and DPP4 belong to the same protease family and contribute 1.3-fold and 11-fold more compounds than FAP, making their data an obvious auxiliary source. Reported gains are small and depend on whether the auxiliary task shares chemotypes with the target; where a different series dominates, transfer can be neutral or harmful."""),
    89: ("The gap between random and scaffold", """The gap between random and scaffold splits is a well-documented reproducibility problem. MoleculeNet adopted scaffold splitting because random splits leak near-analogues of test compounds into training [14]; Wallach and Heifets showed many ligand-based benchmarks reward memorisation [15]; Sheridan argued time-split validation best proxies prospective prediction [16]. For FAP, whose post-2018 literature is dominated by one chemotype, the time split is not a technicality but the only protocol reproducing how a chemist would actually use the model."""),
    91: ("Point predictions without an error", """Point predictions without error estimates are of limited use in a synthesis meeting. Applicability-domain analysis asks whether a training set represents the query, typically via a similarity threshold or descriptor-space bounding box [17]. Conformal prediction instead builds distribution-free intervals from held-out residuals, with a finite-sample coverage guarantee [18] under exchangeability of calibration and test data. Scaffold and temporal splits violate exactly that assumption, making conformal prediction a diagnostic rather than a universal remedy: if coverage collapses, the interval should not be quoted."""),
    93: ("Generative design for small molecules", """Small-molecule generative design has converged on graph-based genetic algorithms that mutate molecular graphs, and deep generative models trained on large corpora [19]. Jensen's graph GA remains a strong baseline: it needs no training data and optimises arbitrary scoring functions. When affinity, selectivity, drug-likeness and synthetic accessibility must be balanced, multi-objective algorithms such as NSGA-II [20] return a Pareto front rather than one scalarised optimum, exposing the trade-offs. Their weakness is that the scoring function is itself a model, so the optimizer will exploit its errors unless novelty and applicability are enforced explicitly."""),
    95: ("The literature supplies mature tools", """The literature supplies mature tools but few honest audits of them on small radiopharmaceutical targets. This project proposes no new architecture; it asks whether the standard toolkit, applied carefully to FAP, produces anything a radiochemist should act on, and reports where the answer is no."""),

    # ---- 3. Research Questions -----------------------------------------
    103: ("Four objectives follow directly.", """Four objectives follow. O1: curate and document deduplicated FAP, PREP and DPP4 datasets with an auditable filter log and an assay-noise estimate. O2: benchmark representative model families under three protocols, reporting regression and virtual-screening metrics. O3: quantify paralogue transfer and direct versus indirect selectivity modelling. O4: determine where predictions and intervals can be trusted, and use that region as a hard constraint in generative design."""),
    105: ("The project is deliberately restricted", """The project is restricted to ligand-based modelling of the FAP warhead; it does not model the chelator, radiometal or in vivo pharmacokinetics, and performs no docking or free-energy calculations. Three success criteria were fixed before any model was fitted, and their disposition is reported below — a pre-registered criterion is only useful if a partial or negative result may stand."""),
    108: ("Stating C3 as partial", """Stating C3 as partial rather than met is the point of the exercise: the analysis that validated the first two criteria also refused the third."""),
    110: ("The workflow has four stages", """The workflow has four stages — data, models, trust, design — each producing a table consumed by the next (Figure 1). Every experiment used five random seeds with identical features and hyperparameters; all figures and tables are regenerated by the accompanying code."""),

    # ---- 4. Methods -----------------------------------------------------
    114: ("Bioactivity records were retrieved", """Bioactivities came from the ChEMBL REST API (release 37, May 2026) for FAP (CHEMBL4683), PREP (CHEMBL3202) and DPP4 (CHEMBL284). Only records with an exact activity relation, standard type Ki, IC50 or Kd, nanomolar units and ChEMBL validity annotations were retained [28]: 972 FAP, 962 PREP and 6,553 DPP4 records, each with assay metadata, document year, validity flags and canonical SMILES."""),
    116: ("Curation followed a fixed sequence", """Curation followed a fixed sequence, with surviving counts logged (Table 1). Records were dropped for a data-validity comment, potential duplication, non-biochemical assay (cell-based or ADME), pActivity outside 3–12, RDKit parsing failure or mass outside 150–1,500 Da. Replicate InChIKeys were collapsed to median pActivity, retaining measurement count and standard deviation; Murcko scaffolds [25] were computed for grouping and splitting. FAP is the modelling target; PREP and DPP4 are auxiliary and off-target references."""),
    117: ("Assay noise was estimated directly", """Assay noise was estimated from replicate structure: for each compound measured more than once we took the standard deviation of its pActivity values and report the median across compounds. This is a floor on the RMSE attainable from these labels and the reference against which model error is judged throughout."""),
    119: ("Each molecule is represented by a", """Each molecule is a concatenation of 2,048-bit Morgan [26] fingerprints (radius 2), 167-bit MACCS keys and 15 physicochemical descriptors (molecular weight, cLogP, TPSA, hydrogen-bond donors/acceptors, rotatable bonds, ring and aromatic-ring counts, sp3 fraction, heavy-atom count, and nitrile, amide, carboxylic acid, N+O and halogen counts). Descriptors were included because the FAP warhead is defined by functional groups, and they extrapolate better than fingerprints when the scaffold changes."""),
    121: ("Five estimators were compared", """Five estimators were compared: a mean predictor (trivial baseline), ridge regression, RBF support-vector regression, a random forest [24] with 500 trees, and XGBoost [23] with 700 trees, learning rate 0.03, depth 6, subsample 0.8, column subsample 0.4, L2 1.0 and minimum child weight 3. Hyperparameters were fixed a priori, since tuning on the test protocol would erase the difference this benchmark measures. All models used the same feature matrix and training indices."""),
    123: ("Three protocols were used.", """Three protocols were used. Random splitting holds out 20% of compounds uniformly (interpolation, the usual published QSAR protocol). Scaffold splitting holds out 20% of Murcko scaffolds (ranking a new series, i.e. scaffold hopping). Time splitting trains on compounds last reported by 2018 and tests on 2019–2024 (prospective use, reproducing a model trained before the FAPI series matured). Each was repeated over five seeds; random and scaffold splits vary by seed, the time split is deterministic."""),
    125: ("Regression quality is reported as", """Regression quality is reported as RMSE, MAE, R², Pearson and Spearman correlation. Since the practical use is triage rather than quantitation, screening metrics were computed on a focused set of clear actives (pActivity ≥ 7, IC50 ≤ 100 nM) and inactives (≤ 5.5), excluding the ambiguous middle: AUROC, enrichment in the top 10% (EF10) and top-10% precision."""),
    127: ("The applicability domain is defined", """The applicability domain is defined by maximum Tanimoto similarity to any training compound over Morgan fingerprints. Test compounds are binned by similarity and bin-wise RMSE reported, yielding an operational threshold rather than a binary claim. A single threshold of 0.60 separates in-domain from out-of-domain throughout: it is where RMSE begins to rise in Table 6 and where out-of-domain conformal coverage falls to 0.54 in Table 7. The 0.1-wide bins in Table 6 are descriptive partitions, not alternative rules; the 0.40 filter in Section 4.9 is a novelty constraint on the generator and a different quantity."""),
    128: ("Uncertainty is quantified by split", """Uncertainty uses split conformal prediction. Out-of-fold residuals come from five-fold cross-validation inside the training set; the conformal half-width is the ⌈(n+1)(1−α)⌉-th smallest absolute residual. Intervals are evaluated on held-out compounds at nominal 80% and 90%, with coverage stratified by applicability-domain class. A random-split experiment is the positive control: with exchangeable data, coverage should match the nominal level."""),
    130: ("For the time split, the raw model", """For the time split, the raw model is compared with a version whose single intercept correction is estimated on the first quarter of the test set and applied to the rest, simulating a small prospective calibration campaign. This separates a shift in overall potency level, which an intercept can absorb, from a loss of ranking information, which it cannot."""),
    132: ("Design used a graph genetic algorithm", """Design used a graph genetic algorithm in the style of Jensen [19]: parents are recombined by breaking one acyclic bond and joining the largest fragments; offspring are mutated by atom substitution, bond-order change, addition or deletion, with sanitisation after each operation. NSGA-II [20] selection used fast non-dominated sorting and crowding distance. Four objectives were optimised: predicted FAP affinity (mean of five XGBoost models), selectivity margin (pFAP − max(pPREP, pDPP4)), QED and −SA/10 [21,22]. Hard filters removed invalid structures, compounds outside 250–650 Da, and those with maximum training-set similarity below 0.40. The initial population was the 130 measured FAP actives with pActivity ≥ 7."""),
    133: ("Two controls were run with the same", """Two controls used the same evaluation budget: affinity-only selection, where NSGA-II degenerates into a hill climber on predicted potency, and a random campaign where offspring replace parents without selection. All archived molecules were annotated for novelty (InChIKey absent from all three datasets) and Murcko scaffold diversity, so rediscovery of known inhibitors is not rewarded."""),
    135: ("All code is Python 3.13", """All code is Python 3.13 with RDKit 2026.09.1 [27], scikit-learn 1.9.1, XGBoost 3.4.1, NumPy 2.5.4, pandas 3.1.0 and Matplotlib 3.11.2. The pipeline is numbered: 01 fetch, 02 curate, 03 benchmark, 04 transfer/selectivity, 05 trust analysis, 06 generative design, 07–08 figures, 09 report, 10 presentation notes, 11 DOI verification, 12 retrosynthetic triage, 13 robustness, 14 reconciliation, 15 report assets, 16 evidence bundle, 17 asset-integrity checks. Each script writes a table or figure the next reads; seeds, versions and logs ship with the code. Curated data, raw exports, prediction tables, the environment fingerprint and a one-command reproduction guide are listed in Appendix A."""),

    # ---- 5. Results -----------------------------------------------------
    138: ("Curation reduced 972 raw FAP", """Curation reduced 972 raw FAP bioactivities to 438 records covering 384 compounds and 161 Murcko scaffolds (Table 1). Median potency is pActivity 6.40 (≈ 398 nM, s.d. 1.49); 34% reach the 100 nM threshold. PREP and DPP4 are larger but structurally similar; DPP4 is dominated by potent compounds (59% active), reflecting its history as a diabetes target with industrial screening behind it."""),
    141: ("The most consequential property", """The most consequential property is not size but noise. Among FAP compounds measured more than once, the median replicate standard deviation is 0.67 log units (PREP 0.53, DPP4 0.37) — roughly a five-fold disagreement in IC50 between two laboratories. This is the error floor: a model predicting within 0.67 log units already reaches the limit of what the labels resolve, so 0.05-unit differences between algorithms are not interpretable."""),
    142: ("The dataset is also chemically non", """The dataset is also chemically non-stationary (Figure 2). Compounds first reported in 2019 or later have median molecular weight 80 Da higher (452 vs 372 Da), median potency two orders of magnitude greater (pActivity 7.60 vs 5.57), and the active fraction rises from 15% to 64%. Before 2018 most FAP measurements came from DPP4 programmes using FAP as a counter-screen; afterwards the literature is dominated by purpose-built quinoline FAPI warheads. Two different chemical questions therefore sit under one target name, which matters for every result that follows."""),
    148: ("Random splitting gives the flattering", """Random splitting gives the flattering picture usual in the literature: random forest and XGBoost both reach R² 0.75 with RMSE near 0.70 and rank correlation 0.85, statistically indistinguishable and clearly ahead of ridge (0.56) and SVR (0.62). Scaffold splitting costs the tree ensembles about 0.07 in R² and 0.12 in RMSE (Figure 3a) — still usable at R² 0.68, Spearman 0.81, screening AUROC 0.94."""),
    149: ("The time split does not degrade", """The time split does not degrade the models, it destroys them (Figure 3b). XGBoost reaches R² −0.79 and rank correlation −0.11; random forest is worse (−1.03), and the mean predictor is as good as any learned model (≈ −1.06). Enrichment falls to 0.96, no better than random selection. Ridge degrades far less (−0.07): relying on additive physicochemical trends, it cannot rank the new series but does not confidently rank it wrongly."""),
    152: ("This is the central empirical result", """This is the central empirical result, and not an artefact of a small test set (n = 147). It follows directly from the era structure in Section 5.1: a model trained on pre-2019 DPP4-like chemistry is asked to rank quinoline FAPI warheads it has never seen. A FAP model trained on literature available before the FAPI era has no prospective value for FAPI design, despite excellent random-split statistics."""),
    156: ("The combined block is the most", """The combined block is most accurate in regression (RMSE 0.828 versus 0.866 Morgan alone, 0.922 MACCS, 0.987 descriptors alone), confirming that fingerprints and descriptors carry complementary information: substructure versus the functional-group composition defining a warhead. Ranking metrics differ — descriptor-only gives the highest AUROC (0.96) despite the worst RMSE. Regression accuracy and screening utility are therefore not the same quantity, which is why both are reported."""),
    160: ("Transfer produces a consistent but", """Transfer produces a consistent but tiny improvement: under random splitting RMSE falls from 0.710 to 0.697 and R² rises from 0.751 to 0.760; under scaffold splitting, 0.828 to 0.824. Paired comparisons give p = 0.19 and p = 0.62 (Wilcoxon signed-rank), so transfer is neutral at this scale. Removing similarity features changes nothing, indicating the small gain comes from the paralogue prediction itself rather than a neighbourhood signal."""),
    161: ("This null result is informative", """This null result is informative. DPP4 has eleven times more measurements than FAP, but its chemistry is dominated by gliptin-like scaffolds that FAP chemists do not want. Transfer helps when auxiliary and target share chemotypes; here they mostly share an enzyme fold. It also bounds what multi-task architectures could achieve with the same labels, since a shared encoder faces the same mismatch."""),
    165: ("Inside known scaffolds, selectivity", """Inside known scaffolds selectivity is learnable: both protocols reach AUROC 0.82–0.84 and rank correlation 0.58–0.63 for ΔpActivity. Direct modelling of the difference and indirect subtraction of two single-target models are statistically indistinguishable — useful, because the indirect route needs no dual-measured training set."""),
    166: ("Outside known scaffolds the same", """Outside known scaffolds the same models are worse than chance: AUROC falls to 0.36–0.38 (s.d. ≈ 0.12 over twenty seeds) and ΔR² becomes negative. A model that ranks selective and non-selective compounds in the wrong order would actively enrich the non-selective series if used as a filter. The likely explanation is a sign flip in the SAR: within a series the substituent filling the FAP S2 pocket also fills the PREP S2 pocket, whereas across series the linker shape dominates — variation the model has never seen."""),
    167: ("The descriptive statistics explain", """Descriptive statistics show why DPP4 matters more than PREP. Among 100 FAP/PREP dual-measured compounds the mean ΔpActivity is +0.59 (32% selective, 34% non-selective), a broad two-sided distribution (Figure 4c). Among 232 FAP/DPP4 pairs the mean is −1.75 and 86% are non-selective. DPP4 is the dominant off-target liability of this chemotype; the practical objective is to preserve FAP affinity while pushing DPP4 below the selectivity threshold."""),
    170: ("Because the cliff is the strongest", """Because this cliff is the strongest negative claim, it was re-tested with a paired bootstrap over twenty seeds (Appendix C, Table C1). Pairing random- and scaffold-split AUROC on the same seed gives a mean difference of 0.463 (95% CI 0.392–0.532) for the indirect protocol and 0.454 (0.382–0.525) for the direct one; 100% of 5,000 replicates are positive. The collapse is not a seed artefact."""),
    171: ("A stratified re-analysis is less", """A stratified re-analysis is less comforting (Appendix C, Table C2). The cliff persists in the IC50-only (Δ = 0.451, n = 70) and single-measurement (Δ = 0.523, n = 93) strata, but reverses in the small mixed-assay stratum (Δ = −0.158, n = 30, 13 evaluable folds); the replicated stratum (n = 7) cannot be fitted. Assay-type confounding cannot be excluded there, so the cliff is robust for the IC50-dominated bulk and unresolved for the heterogeneous remainder."""),
    175: ("Error rises monotonically as the", """Error rises monotonically with distance from the training set: RMSE is 0.70 above similarity 0.7 (n = 263), 0.94 at 0.6–0.7, 0.96 at 0.5–0.6 and 0.95–1.12 below 0.5 (Figure 5a). The absolute effect is modest — about 0.3 log units — but monotone and actionable: above 0.6 a prediction can be quoted with error comparable to assay noise; below 0.6 it cannot."""),
    178: ("The random-split control behaves", """The random-split control behaves as theory predicts: 0.79 and 0.89 empirical coverage at nominal 0.80 and 0.90. Under scaffold splitting coverage drops to 0.73 and 0.87, a mild exchangeability violation; outside the applicability domain the 80% interval covers only 54%. Intervals are honest for compounds resembling the training set and optimistic otherwise — the opposite of the intuition that error bars are conservative."""),
    181: ("The practical rule that follows", """The practical rule is a two-gate decision: compute the similarity, and quote a conformal interval only when it exceeds 0.6. Below that threshold the model may still rank a series internally, but its numeric interval should not enter a synthesis decision."""),
    182: ("A natural repair is class-conditional", """A natural repair is class-conditional (Mondrian) conformal calibration: bin out-of-fold residuals by how well each training compound is represented, and let each bin use its own quantile. Appendix C, Table C3 and Figure 6 show out-of-domain coverage at 80% rising from 0.54 to 0.74 and at 90% from 0.77 to 1.00, with in-domain coverage essentially unchanged (0.77 to 0.74; 0.88 to 0.88)."""),
    183: ("This is a diagnostic, not a deployable", """This is a diagnostic, not a deployable fix. Only 27 of roughly 280 training compounds fall in the out-of-domain calibration bin, so its quantile is noisy: at 90% it over-covers completely (1.00) and more than doubles interval width from 2.40 to 5.01 log units. Conditional calibration buys nominal coverage by spending precision — the trade-off this report argues should be reported rather than hidden behind one reassuring number."""),
    189: ("The raw time-split model has RMSE", """The raw time-split model has RMSE 2.07 and R² −0.82, with mean signed error +1.18 log units (single fit at seed 0; the five-seed mean in §5.2 is −0.79 with RMSE 2.05): it under-predicts post-2018 potency by over an order of magnitude (Figure 7b). Recalibrating one intercept on a quarter of the test set reduces RMSE to 1.64 and R² to −0.15, but rank correlation is unchanged at −0.12. The error is almost entirely a level shift: correcting the level restores absolute accuracy without restoring ranking ability. A chemist recalibrating on a few new compounds would get correct potency estimates while still ranking the series at random."""),
    195: ("All three campaigns quickly reach", """All three campaigns quickly reach the affinity model's ceiling (best predicted pFAP 9.6–9.7, Figure 8b) — itself a warning, since an ensemble trained on 384 compounds cannot separate a genuine 10 nM inhibitor from an implausible one. The campaigns differ in what they sacrifice. Without selection pressure the population drifts from pFAP 7.35 to 6.44 in 24 generations (Figure 8a); under affinity-only selection potency climbs while chemistry degenerates, with median QED collapsing to 0.25 and molecular weight inflating to 548 Da as polar atoms are traded for hydrophobic bulk."""),
    196: ("The multi-objective campaign pays", """The multi-objective campaign pays a small affinity price (median novel pFAP 8.24 in Table 9 versus 9.36) and returns the most usable chemistry: 115 Murcko scaffolds against 80–83 for the controls, median QED 0.61, synthetic accessibility 2.93 and molecular weight 413 Da, within the clinical FAPI range. Its 61% pass rate matches the random control (60%) — the objectives show up in diversity and developability rather than pass rate, exactly what a Pareto front should do."""),
    197: ("Two features of Table 9 need to be", """Two features of Table 9 need care. First, the affinity-only 100% pass rate follows from the metric, not better chemistry: selection drives all 369 novel molecules above pFAP 8, and since margin is pFAP minus the largest paralogue prediction, a high pFAP almost automatically meets margin ≥ 1. The same circularity explains why adding selectivity and developability pressure lowers the pass rate to 61%: the optimizer can no longer spend its whole budget on pFAP. What that budget bought matters more than the pass rate — median QED 0.61, MW 413 Da, 115 scaffolds at pFAP 8.24, versus QED 0.25, MW 548 Da, 83 scaffolds for affinity-only. The top-ranked designs are quinoline-carboxamides with modified amide linkers (Table 10); Section 5.11 and Appendix B explain why their warhead is not the clinical FAPI warhead."""),
    203: ("A potency model returns a molecule", """A potency model returns a molecule that binds; a radiopharmaceutical must also be made and labelled. Both questions were put to the six top-ranked designs using only rule-based calculations (Appendix B). BRICS disconnection gives 7–10 fragments (median 9), an estimated 7–10-step route and synthetic-accessibility scores of 2.97–4.40 (median 3.24), the range of ordinary medicinal chemistry. All six combine a quinoline carboxylic acid with a proline-derived amine, the same logic as published FAPIs. These are disconnection heuristics, not validated routes; no reaction in Table B2 has been run."""),
    204: ("The radiochemical audit is where", """The radiochemical audit is where the designs fail. None carries a chelator or a free conjugation handle: every heteroatom is already in an amide or the warhead. 68Ga and 177Lu need a chelator, 18F a displaceable leaving group; neither is present. Grafting DOTA onto the nearest carboxylic acid moves median molecular weight from 498.6 to 885.0 Da (878.9–949.9) and median TPSA from 126.0 to 247.7 Å². The training set has median MW 401.4 Da, 95th percentile 557.3 Da, and only 2.3% above 800 Da. The conjugated ligand therefore leaves the applicability domain of the model that scored it: the optimised quantity and the deliverable molecule are not the same object."""),
    205: ("The warhead mismatch compounds this", """The warhead mismatch compounds this. Five of six designs carry an α-diketone or ketoamide electrophile, whereas clinical FAPI tracers use a prolyl nitrile forming a reversible covalent adduct with the catalytic serine; only T10-6 recovers that chemotype. A model trained on a few hundred compounds dominated by one warhead class will reward a different, more reactive one, because both occupy the same descriptor neighbourhood. The fix is not a better regressor: the objective must include linker, chelator and labelling chemistry, and candidates must be re-scored after conjugation — a change of design space, and the most useful output of this audit."""),
    207: ("Three findings deserve emphasis.", """Three findings deserve emphasis. First, the evaluation protocol changes the conclusion more than the algorithm: the same XGBoost is excellent (R² 0.75) or useless (−0.79) depending only on how data are divided, so any FAP prediction claim without a stated split is unfalsifiable. This is not new in machine learning, but is rarely applied to radiopharmaceutical targets, where datasets are small and pressure to report a good number is high."""),
    208: ("Second, the error floor set by", """Second, the assay-noise floor (0.67 log units) is close to the best achievable RMSE (0.70 under random splitting). The model is limited by enzyme-assay reproducibility, not architecture. Further modelling on this dataset has a low ceiling; the highest-value intervention is a consistent centralised assay campaign for a few hundred FAP inhibitors, not a better neural network."""),
    209: ("Third, selectivity is the quantity", """Third, selectivity decides whether a FAP ligand advances, and it generalises worst. Predicting ΔpActivity within a series is easy; across series it is not currently possible with 100 dual-measured compounds. Design campaigns should treat selectivity as a filter to verify experimentally, not an objective to optimise in silico."""),
    211: ("•\u2003The dataset is one snapshot", """•\u2003The dataset is one ChEMBL snapshot, inheriting publication bias towards potent compounds, inconsistent assays and an over-representation of the quinoline warhead after 2018."""),
    212: ("•\u2003Only ligand-based descriptors", """•\u2003Only ligand-based descriptors were used; no docking, pharmacophore or structural information entered the models, which cannot distinguish a sub-pocket match from a physicochemical coincidence."""),
    213: ("•\u2003The generative campaign scores", """•\u2003The generative campaign is scored by the same models whose limits are documented above. Predicted affinities above pActivity 9 mean 'inside the model's confident region', not '10 nM'."""),
    214: ("•\u2003The applicability-domain filter", """•\u2003The 0.40 applicability-domain filter constrains novelty by construction; a more adventurous campaign would generate more diverse but less assessable chemistry."""),
    215: ("•\u2003Temporal analysis uses publication", """•\u2003Temporal analysis uses publication year as a proxy for availability; a true prospective test needs an external, later dataset."""),
    216: ("•\u2003Five of the six highest-ranked", """•\u2003Five of six top-ranked designs (Table 10) carry an alpha-diketone / ketoamide electrophile instead of the clinical prolyl nitrile warhead. pFAP above 9.5 therefore marks the model's confident region, not a validated SAR series; neither potency nor synthetic tractability follows (Appendix B)."""),
    217: ("•\u2003No protein-structure, docking", """•\u2003No docking, pharmacophore, free-energy or structure-based calculation was performed; the brief lists these as optional, not required. Adding a docking score would introduce a second, uncalibrated confidence scale and contradict this report's audit framing. Any future structure-based work should be used qualitatively and reported with the same applicability-domain and failure checks."""),
    218: ("•\u2003The selectivity cliff is robust", """•\u2003The selectivity cliff is robust in the IC50-dominated bulk of the paired data but reverses in a 30-compound mixed-assay stratum (Appendix C, Table C2); assay-type confounding there cannot be excluded."""),

    # ---- 6. Conclusions -------------------------------------------------
    220: ("This project asked what machine", """This project asked what machine learning can deliver for FAP radiopharmaceutical design from the public bioactivity record, including where the answer is negative. A curated set of 384 inhibitors supports R² 0.68 and screening AUROC 0.94 under scaffold splitting, meeting the pre-registered triage criterion. The same model fails under a time split (R² −0.79), a failure traceable to a documented change of chemical series rather than algorithmic weakness. Paralogue transfer is neutral; selectivity over PREP is learnable within a series (AUROC 0.84) but anti-correlated across scaffolds (0.38). Conformal intervals are valid under interpolation and optimistic outside the applicability domain, so any quoted uncertainty requires a similarity gate."""),
    221: ("The four research questions are", """The four research questions are answered, though not as a typical QSAR paper would report. RQ1: accuracy is protocol-dependent enough to change the sign of R². RQ2: transfer is neutral and selectivity does not generalise across scaffolds. RQ3: intervals are trustworthy only inside the applicability domain. RQ4: multi-objective selection improves diversity and drug-likeness (115 scaffolds, median QED 0.61) while single-objective optimisation degenerates chemically (QED 0.25); neither can manufacture information the scoring model lacks."""),
    222: ("The broader contribution is methodological", """The broader contribution is methodological. For targets with a few hundred compounds, a drifting literature and heterogeneous assays, the deliverable should be a calibrated decision rule — use the model here, not there — rather than a single predicted potency. Building it requires three cheap analyses: protocol comparison, applicability domain and conformal calibration."""),

    # ---- Appendices B / C ----------------------------------------------
    259: ("This appendix closes the gap", """This appendix closes the gap left in Section 1.2 with a rule-based structural analysis, not a verified synthesis: BRICS disconnections, reaction assignments and step counts are graph heuristics from src/12_design_feasibility.py, with no reaction, yield or supplier check. Descriptors are RDKit calculations; chelator figures are stoichiometric estimates for a DOTA mono-amide (404.4 Da, C16H28N4O8; conjugation adds 386.4 Da)."""),
    262: ("The labelling handles lead to the", """The labelling handles lead to the same conclusion. 18F needs a displaceable leaving group on an aryl or alkyl carbon; the fluorines here are stable aryl-F or gem-difluoro substituents. 68Ga and 177Lu need a macrocyclic chelator, yet every nitrogen and oxygen in the six structures is already in an amide or the electrophilic warhead, with no free amine or carboxylate for coupling. All three routes require structural modification first, producing a molecule unlike the 384 compounds defining the applicability domain. Tumour retention and normal-tissue background are therefore out of scope: they are pharmacokinetic properties of a conjugate this model has never seen. The next step is narrow — synthesise three to five chelator-bearing analogues, measure affinity and logD, and check the 0.6 similarity gate."""),
    265: ("The medicinal-chemistry answer", """An alpha-diketone is accessible (oxidation of the corresponding alpha-hydroxy ketone, then amide couplings onto a quinoline-4-carboxamide and a proline-derived amine), but it is not the clinical FAPI warhead. Clinical tracers use a prolyl nitrile forming a reversible covalent adduct with the catalytic serine; only the sixth design in Table 10 retains it. A high predicted pFAP for the first five means the optimizer found the model's confident region for a different chemotype, not that it rediscovered FAPI-04. This appendix offers triage and checkable questions, not a claimed synthesis."""),
    267: ("All three analyses below reuse", """All three analyses below reuse the ChEMBL release-37 tables from src/13_robustness.py; no external dataset or new model class was introduced. They test, rather than decorate, the two strongest claims — the selectivity cliff and the out-of-domain conformal failure."""),
    274: ("Read together, these tables strengthen", """Read together, the tables strengthen the first claim and qualify the second. The paired interval excludes zero in every replicate, so the cliff is not seed noise; but it reverses in a 30-compound mixed-assay stratum, leaving assay-type confounding a live limitation. Mondrian calibration lifts out-of-domain coverage without changing in-domain coverage, yet uses a 27-member bin whose 90% quantile over-covers (1.00) at more than twice the interval width. It is a diagnostic of where exchangeability fails, not an interval to quote in a synthesis meeting — the conclusion Section 5.6 reaches by another route."""),
}


# Second-round typographic touch-ups applied after the bulk pass.
# Keyed by index; prefix is the text the paragraph carries on entry.
TOUCHUPS: dict[int, tuple[str, str]] = {
    # U+2308 / U+2309 have no glyph in the body font and print as braces
    128: ("Uncertainty uses split conformal prediction.",
          """Uncertainty uses split conformal prediction. Out-of-fold residuals come from five-fold cross-validation inside the training set; the conformal half-width is the smallest absolute residual at rank ceil((n+1)(1−α)). Intervals are evaluated on held-out compounds at nominal 80% and 90%, with coverage stratified by applicability-domain class. A random-split experiment is the positive control: with exchangeable data, coverage should match the nominal level."""),
    # the long package name sat at a line end and stretched the justified first line
    254: ("The complete project is submitted as a single package named",
          """The complete project is submitted as a single package; 03_Code/ inside it holds the source code and process evidence, and 04_Results/ holds every table and figure cited here. The package root is Submission_HuangYong_P2600222/."""),
}


def apply_touchups(doc) -> int:
    changed = 0
    for idx, (prefix, new) in TOUCHUPS.items():
        cur = doc.paragraphs[idx].text.strip()
        if cur == new.strip():
            continue
        assert cur.startswith(prefix), (
            f"touch-up {idx} mismatch: {cur[:70]!r}")
        set_text(doc.paragraphs[idx], new)
        changed += 1
    return changed


def references(doc) -> Counter:
    """Multiset of every numbered cross-reference in the prose."""
    text = "\n".join(p.text for p in doc.paragraphs)
    return Counter(re.findall(r"(?:Table|Figure|Appendix)\s+[A-C]?\d+|\[\d+(?:[–-]\d+)?\]", text))


def word_count(doc) -> int:
    return sum(len(p.text.split()) for p in doc.paragraphs)


def body_words(doc) -> int:
    """Prose only: excludes captions, table titles and the reference list."""
    total = 0
    for p in doc.paragraphs:
        t = p.text.strip()
        if not t:
            continue
        if re.match(r"^(Table|Figure)\s+[A-C]?\d+\.", t) or re.match(r"^\[\d+\]", t):
            continue
        total += len(t.split())
    return total


def set_text(par, text: str) -> None:
    """Preserve the first run's character formatting (see 25_fix_appendix_paths.py)."""
    if not par.runs:
        par.add_run(text)
        return
    par.runs[0].text = text
    for run in par.runs[1:]:
        run.text = ""


def verify(doc) -> None:
    assert len(doc.paragraphs) == 275, "paragraph count changed"
    for idx, (_prefix, new) in {**PATCHES, **TOUCHUPS}.items():
        got = doc.paragraphs[idx].text.strip()
        assert got == new.strip(), f"paragraph {idx} not in condensed form"


def apply(do_pdf: bool) -> None:
    from docx import Document

    doc = Document(str(DOCX))
    before_refs = references(doc)
    before_words = word_count(doc)
    before_body = body_words(doc)
    assert len(doc.paragraphs) == 275, "unexpected paragraph count"

    pending: list[int] = []
    for idx, (prefix, new) in PATCHES.items():
        cur = doc.paragraphs[idx].text.strip()
        if cur == new.strip():
            continue
        # already carrying its touch-up version (e.g. paragraph 128)
        if idx in TOUCHUPS and cur == TOUCHUPS[idx][1].strip():
            continue
        assert cur.startswith(prefix), (
            f"paragraph {idx} does not match the expected original: {cur[:70]!r}")
        assert doc.paragraphs[idx].style.name == "Normal", f"paragraph {idx} is not body prose"
        pending.append(idx)

    if not pending:
        n = apply_touchups(doc)
        if n:
            doc.save(str(DOCX))
            print(f"  touch-ups       : {n}")
        print("  already condensed; bulk pass skipped")
        verify(doc)
        if do_pdf:
            export_pdf()
        return

    backup = Path(tempfile.gettempdir()) / "aidd_report_before_condense.docx"
    shutil.copy2(DOCX, backup)

    for idx in pending:
        set_text(doc.paragraphs[idx], PATCHES[idx][1])
    apply_touchups(doc)

    after_refs = references(doc)
    assert after_refs == before_refs, (
        f"cross-references changed: lost {before_refs - after_refs}, "
        f"added {after_refs - before_refs}")
    doc.save(str(DOCX))

    reread = Document(str(DOCX))
    verify(reread)
    after_words = word_count(reread)
    after_body = body_words(reread)
    print(f"  paragraphs      : 275 (unchanged)")
    print(f"  rewritten       : {len(pending)}/{len(PATCHES)}")
    print(f"  all words       : {before_words:,} -> {after_words:,} "
          f"(-{before_words - after_words:,}, -{(before_words - after_words) / before_words:.1%})")
    print(f"  body prose only : {before_body:,} -> {after_body:,} "
          f"(-{before_body - after_body:,}, -{(before_body - after_body) / before_body:.1%})")
    print(f"  cross-refs      : {sum(before_refs.values())} tokens, identical")
    print(f"  backup          : {backup}")

    if do_pdf:
        export_pdf()


def export_pdf() -> None:
    import win32com.client as win32

    tmp = Path(tempfile.gettempdir()) / f"aidd_report_{int(time.time())}.pdf"
    word = win32.DispatchEx("Word.Application")
    word.Visible = False
    try:
        word.DisplayAlerts = 0
        doc = word.Documents.Open(str(DOCX), False, True, False)
        doc.Fields.Update()
        for sec in doc.Sections:
            for i in (1, 2, 3):
                for coll in (sec.Footers, sec.Headers):
                    try:
                        coll(i).Range.Fields.Update()
                    except Exception:
                        pass
        doc.Repaginate()
        print(f"  Word pages      : {doc.ComputeStatistics(2)}")
        doc.ExportAsFixedFormat(str(tmp), 17, False, 0, 0, 0, 0, 0,
                                True, True, 0, True, True, False)
        doc.Close(False)
    finally:
        word.Quit()
    shutil.copy2(tmp, PDF)
    tmp.unlink(missing_ok=True)
    print(f"  PDF exported    : {PDF.name} ({PDF.stat().st_size:,} bytes)")


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--check", action="store_true", help="verify only, write nothing")
    ap.add_argument("--no-pdf", action="store_true", help="rewrite the docx but skip PDF export")
    args = ap.parse_args()

    from docx import Document
    doc = Document(str(DOCX))
    if args.check:
        verify(doc)
        print(f"  condensed state OK | words {word_count(doc):,} | "
              f"cross-refs {sum(references(doc).values())}")
    else:
        apply(not args.no_pdf)


if __name__ == "__main__":
    main()
