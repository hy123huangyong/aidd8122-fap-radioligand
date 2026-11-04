# Figure and table index

File names follow the order in which the analysis scripts write them, which is
**not** the same as the figure numbers in the report. Use this table to match a
report figure to the code that produced it.

| Report | File in `04_Results/figures/` | Produced by | Content |
|---|---|---|---|
| Figure 1 | `fig0_workflow.png` | `src/08_workflow.py` | Project workflow schematic |
| Figure 2 | `fig1_dataset.png` | `src/07_figures.py` | Curated datasets: potency coverage, size, era |
| Figure 3 | `fig2_benchmark.png` | `src/03_benchmark.py` → `07_figures.py` | RMSE / R² / Spearman by split protocol and model |
| Figure 4 | `fig3_transfer_selectivity.png` | `src/04_transfer.py` → `07_figures.py` | Paralogue transfer and selectivity |
| Figure 5 | `fig4_applicability_domain.png` | `src/05_analysis.py` | Error vs maximum Tanimoto similarity |
| Figure 6 | `fig5_temporal_drift.png` | `src/05_analysis.py` | Random-split interpolation vs 2019–2024 extrapolation |
| Figure 7 | `fig7_generative.png` | `src/06_generative.py` | Generative campaigns: selection pressure, archive, QED |

There is no `fig6_*` file: the generative figure was written after the numbering
of the analysis figures had been fixed, so it carries the number `7` to match its
report label.

## Tables

The report's Tables 1–10 are all built from `04_Results/tables/`:

| Report | Source file(s) |
|---|---|
| Table 1 — Data sources and endpoints | `dataset_summary.csv`, `curation_log.csv` |
| Table 2 — Benchmark across protocols | `benchmark_FAP_combined_summary.csv` |
| Table 3 — Feature ablation | `ablation_features_FAP.csv` |
| Table 4 — Paralogue transfer | `transfer_runs.csv` |
| Table 5 — Selectivity protocols | `selectivity_cv.csv`, `selectivity_ci.csv` |
| Table 6 — Selectivity overview | `selectivity_overview.csv` |
| Table 7 — Applicability-domain bins | `ad_bins.csv` |
| Table 8 — Conformal coverage | `conformal.csv` |
| Table 9 — Temporal drift | `temporal.csv` |
| Table 10 — Generative campaigns | `generative_summary.json`, `generative_archive.csv` |

`benchmark_FAP_combined_preds.parquet` and `transfer_preds.parquet` hold the
per-compound predictions behind every metric, so any number in the report can be
recomputed without refitting.
