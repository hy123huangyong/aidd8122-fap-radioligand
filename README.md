# Trust Before Throughput — FAP-Selective Radioligand Warheads (AIDD8122 Final Project)

**Author:** Huang Yong (P2600222) · AIDD8122, Macao Polytechnic University · November 2026

This repository reproduces every number, figure and table in the report
*"Trust Before Throughput: Benchmarking, Uncertainty-Aware Screening and
Multi-Objective Generative Design of FAP-Selective Radioligand Warheads"*.

The project asks what machine learning can actually deliver for FAP-targeted
radiopharmaceutical design given the public ChEMBL record — and, more
importantly, where the answer is negative.

## Repository layout

```
src/
  01_fetch_chembl.py     retrieve FAP / PREP / DPP4 bioactivities from the ChEMBL API
  02_curate.py           filtering pipeline (972 -> 384 FAP records) + curation log
  03_benchmark.py        5 models x 3 split protocols (random / scaffold / time)
  04_transfer.py         paralogue transfer features + selectivity protocols (direct vs indirect)
  05_analysis.py         applicability domain, conformal coverage, temporal drift, figures
  06_generative.py       graph GA + NSGA-II multi-objective design (3 campaigns)
  07_figures.py          benchmark and selectivity figures
  08_workflow.py         workflow schematic
  09_build_report.py     assembles the .docx report from the template
  10_add_notes.py        writes the timed speaker notes into the presentation deck
  11_reference_dois.py   verifies all 29 references against Crossref and attaches DOIs
  12_design_feasibility.py  BRICS retrosynthetic triage + radiolabelling audit (Appendix B)
  13_robustness.py       paired bootstrap, stratified re-analysis, Mondrian conformal (Appendix C)
  14_sync_scaffold_summary.py  unifies the scaffold-count convention across tables and figures
  15_patch_report.py     applies the content corrections to 09_build_report.py (idempotent)
  16_evidence_bundle.py  writes logs/ (environment fingerprint, manifest, reconciliation log)
  17_restore_images.py   restores full-resolution figures after Word saves the .docx
  pack_submission.py     assembles Submission_HuangYong_P2600222/ and its .zip
  aidd_utils.py          featurisation, splits, metrics, plotting style
data/raw/                ChEMBL records as retrieved (JSON)
data/processed/          curated per-target tables (CSV)
results/tables/          every number reported in the document (CSV/JSON)
results/figures/         every figure in the document (PNG)
report/                  final report (.docx / .pdf)
```

## Reproduction

```bash
pip install -r requirements.txt
python src/01_fetch_chembl.py     # needs internet access to www.ebi.ac.uk
python src/02_curate.py
python src/03_benchmark.py --target FAP --seeds 5
python src/04_transfer.py --seeds 20
python src/05_analysis.py
python src/06_generative.py --generations 24 --pop 96   # ~30 min on CPU
python src/07_figures.py && python src/08_workflow.py
python src/09_build_report.py
python src/11_reference_dois.py && python src/12_design_feasibility.py && python src/13_robustness.py
```

The PDF is then produced by `src/build_report_pdf.ps1`, which inserts the Word
table-of-contents field, exports the PDF and restores the full-resolution figures
that Word downsamples on save. `src/14`–`src/17` and `pack_submission.py` are
build/report utilities rather than analysis steps; they are documented in
`docs/00_START_HERE.md`.

## Headline results

| Finding | Number |
| --- | --- |
| Best model under random split (XGBoost) | R² 0.75, RMSE 0.70 |
| Same model under scaffold split | R² 0.68, AUROC 0.94 |
| Same model under time split (train ≤ 2018) | R² −0.79, EF10 0.96 |
| Median replicate SD of FAP assays (error floor) | 0.67 log units |
| Selectivity AUROC, within / across scaffolds | 0.84 / 0.38 |
| Conformal coverage at nominal 80%, in / out of domain | 0.77 / 0.54 |
| Temporal offset of post-2018 literature | +1.18 log units |
| Multi-objective design campaign | 361 novel molecules, 115 scaffolds, median QED 0.61 |
| Paired bootstrap of the cliff | Δ AUROC 0.463 [0.392, 0.532] over 20 seeds |
| Mondrian calibration, out of domain | coverage 0.54 → 0.74 at 80%, width 2.40 → 5.01 at 90% |
| Radiochemical feasibility of the six designs | 0/6 carry a chelator; 498.6 → 856.9 Da after DOTA |

## Data sources

- ChEMBL REST API (ChEMBL_37, released May 2026): FAP (CHEMBL4683), PREP (CHEMBL3202), DPP4 (CHEMBL284).
- All records were retrieved on 2026-11-03; curation criteria are documented in the report (§4.2).
