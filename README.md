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
```

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

## Data sources

- ChEMBL REST API (ChEMBL_37, released May 2026): FAP (CHEMBL4683), PREP (CHEMBL3202), DPP4 (CHEMBL284).
- All records were retrieved on 2026-11-03; curation criteria are documented in the report (§4.2).
