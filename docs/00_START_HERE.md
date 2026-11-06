# AIDD8122 Final Project — Submission Package

**Title:** Trust Before Throughput: Benchmarking, Uncertainty-Aware Screening and
Multi-Objective Generative Design of FAP-Selective Radioligand Warheads

| | |
|---|---|
| Student | Huang Yong (黄勇) |
| Student ID | P2600222 |
| Group | Group 2 — presentation 24 November 2026 |
| Programme | PhD in Artificial Intelligence Driven Drug Discovery, Macao Polytechnic University |
| Course | AIDD8122 Advanced Topics in Artificial Intelligence & Drug Discovery |

---

## What is in this folder

This index describes the **submission package**. In the public GitHub repository
the identical material is laid out as `src/`, `data/`, `results/`, `report/` and
`AIDD8122_Final_Presentation_HuangYong/`; `README.md` there documents that view.

```
00_START_HERE.md                 this file

01_Report/
    AIDD8122_Final_Report_HuangYong_P2600222_v2.pdf     report, 30 pages, live TOC (submission copy)
    AIDD8122_Final_Report_HuangYong_P2600222_v2.docx    editable source, real Word TOC field

02_Presentation/
    AIDD8122_Final_Presentation_HuangYong_10min.pdf     14 slides, timed to 10:00 (submission copy)
    AIDD8122_Final_Presentation_HuangYong_10min.pptx    editable source, timed speaker notes
    Speaker_Notes_10min.md                              printable script, 10 min 00 s

03_Code/
    README.md                 one-command reproduction + table-to-script map
    requirements.txt          pinned versions, exactly those used for the results
    src/01 … src/17           numbered pipeline, in execution order
    src/aidd_utils.py         featurisation, splits, metrics, plotting style
    logs/                     environment fingerprint, file manifest, reproduction log
    data/raw/                 the three ChEMBL release-37 JSON exports (query strings in 01)
    data/processed/           the 972 → 438 → 384 curation chain, per target

04_Results/
    FIGURE_INDEX.md           report figure/table number → file → producing script
    figures/                  fig0–fig8, as they appear in the report
    tables/                   every number quoted in Tables 1–10 and Appendices B–C

05_Data/
    fap_curated.csv · prep_curated.csv · dpp4_curated.csv    curated, model-ready tables
```

Figure file numbers are offset by one from the report numbers (`fig0_workflow.png`
is Figure 1); `04_Results/FIGURE_INDEX.md` gives the full mapping, including which
script produces which table.

## How to reproduce

```bash
pip install -r requirements.txt
python src/01_fetch_chembl.py                            # ~2 min, needs www.ebi.ac.uk
python src/02_curate.py
python src/03_benchmark.py --target FAP --seeds 5
python src/04_transfer.py --seeds 20
python src/05_analysis.py
python src/06_generative.py --generations 24 --pop 96    # ~30 min on CPU
python src/07_figures.py && python src/08_workflow.py
python src/09_build_report.py                            # rebuilds the .docx
```

The report pipeline continues with `10_add_notes.py` (speaker notes), the
verification and robustness scripts `11`–`13`, the reconciliation script `14`,
the content patch `15`, the evidence bundle `16` and the figure-restoration step
`17`. The last two rebuild artefacts and are not needed to reproduce any number.

`data/raw/` is included, so steps 01 and 02 can be skipped; the JSON exports are
the exact files used, retrieved 3 November 2026.

## Headline numbers (all traceable to `04_Results/tables/`)

| Finding | Number | Source file |
|---|---|---|
| Best model, random split (XGBoost) | R² 0.75, RMSE 0.70 | `benchmark_FAP_combined_summary.csv` |
| Same model, scaffold split | R² 0.68, screening AUROC 0.94 | same |
| Same model, time split (train ≤ 2018) | R² −0.79, EF10 0.96 | same |
| Median replicate SD — the error floor | 0.67 log units | `curation_log.csv` |
| Selectivity AUROC, within / across scaffolds | 0.84 / 0.38 (20 seeds) | `selectivity_cv.csv` |
| Paired bootstrap of that gap | Δ 0.463 [0.392, 0.532] | `selectivity_bootstrap.csv` |
| Stratified re-analysis | cliff holds in IC50-only, reverses in mixed-assay (n = 30) | `selectivity_stratified.csv` |
| Conformal coverage at nominal 80%, in / out of domain | 0.77 / 0.54 | `conformal.csv` |
| Mondrian (conditional) calibration, out of domain | 0.54 → 0.74 at 80%, width 2.40 → 5.01 at 90% | `mondrian_conformal.csv` |
| Systematic offset of post-2018 literature | +1.18 log units | `temporal.csv` |
| Multi-objective campaign | 361 novel, 115 scaffolds, median QED 0.61 | `generative_summary.json` |
| Design feasibility triage | 0/6 carry a chelator; DOTA conjugation 498.6 → 885.0 Da | `radiolabel_feasibility.csv`, `retrosynthesis.csv` |

## Data provenance

- ChEMBL REST API, **release 37 (May 2026)**; records retrieved 3 November 2026.
- Targets: FAP `CHEMBL4683`, PREP `CHEMBL3202`, DPP4 `CHEMBL284`.
- Curation chain for FAP: 972 raw → 510 exact-relation → 438 in range → 384 unique
  structures / **161** Murcko scaffolds (acyclic molecules are not counted as a
  scaffold); every filter step is logged in `curation_log.csv`.
- All 29 references carry a DOI verified against Crossref, except the RDKit
  software citation, which uses the Zenodo DOI of the exact release (2026.09.1).

## Code and data availability

Appendix A of the report carries a clickable link to the public repository
**https://github.com/hy123huangyong/aidd8122-fap-radioligand**.  The URL is stored in `repo_url.txt`; to change it, run
`python src/20_set_repo_url.py <new-url>` (do not rebuild with
`09_build_report.py`, which would discard post-build corrections).

Local git history exists in the project root (`AIDD8122_Final_Project/.git`).
`.gitignore` excludes caches, slide-building scratch state and this package
(it is derived).

## Checklist before submission

- [ ] Report PDF and presentation PDF are the two files the guidelines ask for
- [x] Appendix A links to the live public code repository: https://github.com/hy123huangyong/aidd8122-fap-radioligand
- [ ] `03_Code/` contains `src/`, `requirements.txt`, `README.md`; `04_Results/tables/` holds every quoted number
- [ ] `git status` is clean and the last commit matches the submitted PDF
- [ ] Presentation rehearsed once at 10 minutes using the notes in `02_Presentation/`
- [ ] Group 2 slot: 24 November 2026
