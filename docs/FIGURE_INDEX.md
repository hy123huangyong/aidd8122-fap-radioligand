# Figure and table index

File names follow the order in which the analysis scripts write them, which is
**not** the same as the figure numbers in the report. Use this table to match a
report figure to the code that produced it. All eight figures are shipped at
their native resolution (2,414–3,075 px wide).

## Figures

| Report | File in `04_Results/figures/` | Produced by | Content |
|---|---|---|---|
| Figure 1 | `fig0_workflow.png` | `src/08_workflow.py` | Project workflow schematic |
| Figure 2 | `fig1_dataset.png` | `src/07_figures.py` | Curated datasets: potency coverage, era shift, scaffold diversity |
| Figure 3 | `fig2_benchmark.png` | `src/03_benchmark.py` → `07_figures.py` | RMSE / predicted-vs-measured / early enrichment |
| Figure 4 | `fig3_transfer_selectivity.png` | `src/04_transfer.py` → `07_figures.py` | Paralogue transfer and selectivity AUROC / ΔpActivity |
| Figure 5 | `fig4_applicability_domain.png` | `src/05_analysis.py` | Error vs maximum Tanimoto similarity |
| Figure 6 | `fig5_temporal_drift.png` | `src/05_analysis.py` | Interpolation vs 2019–2024 extrapolation and recalibration |
| Figure 7 | `fig7_generative.png` | `src/06_generative.py` | Generative campaigns: selection pressure, archive, QED |
| Figure 8 | `fig8_mondrian.png` | `src/13_robustness.py` | Global vs class-conditional conformal coverage |

There is no `fig6_*` file: the generative figure was written after the numbering
of the analysis figures had been fixed, so it carries the number `7` to match its
report label. Figure 8 was added in the optimisation pass.

## Tables

The report's Tables 1–10 are built directly from `04_Results/tables/` by
`src/09_build_report.py`. Note that **Table 6 is the applicability-domain bins**,
not the selectivity overview (that descriptive table is `selectivity_overview.csv`
and is quoted in §5.5 text only).

| Report | Source file(s) | Produced by |
|---|---|---|
| Table 1 — Composition of the three curated datasets | `dataset_summary.csv`, `curation_log.csv` | `02_curate.py`, `14_sync_scaffold_summary.py` |
| Table 2 — Five models × three split protocols | `benchmark_FAP_combined_runs.csv` | `03_benchmark.py` |
| Table 3 — Feature ablation | `ablation_features_FAP.csv` | `03_benchmark.py` |
| Table 4 — Paralogue transfer | `transfer_runs.csv` | `04_transfer.py` |
| Table 5 — Selectivity protocols (100 dual-measured FAP/PREP) | `selectivity_cv.csv`, `selectivity_ci.csv` | `04_transfer.py` |
| Table 6 — Error by similarity bin | `ad_bins.csv` | `05_analysis.py` |
| Table 7 — Empirical conformal coverage | `conformal.csv` | `05_analysis.py` |
| Table 8 — Temporal drift and intercept recalibration | `temporal.csv` | `05_analysis.py` |
| Table 9 — Three generative campaigns | `generative_archive.csv` | `06_generative.py` |
| Table 10 — Six highest-ranked novel SMILES | `generative_top15_ga.csv` | `06_generative.py` |
| §3.3 success criteria (unnumbered) | derived from Tables 2 and 7 | `09_build_report.py` |
| Table B1 — Radiolabelling feasibility | `radiolabel_feasibility.csv` | `12_design_feasibility.py` |
| Table B2 — Retrosynthetic triage | `retrosynthesis.csv` | `12_design_feasibility.py` |
| Table C1 — Paired bootstrap | `selectivity_bootstrap.csv` | `13_robustness.py` |
| Table C2 — Stratified selectivity | `selectivity_stratified.csv` | `13_robustness.py` |
| Table C3 — Mondrian conformal | `mondrian_conformal.csv` | `13_robustness.py` |
| §7 references — 29 DOIs | `reference_dois.csv` | `11_reference_dois.py` (Crossref / DataCite) |

`benchmark_FAP_combined_preds.parquet` and `transfer_preds.parquet` hold the
per-compound predictions behind every metric, so any number in the report can be
recomputed without refitting. `03_Code/logs/reproduction_log.md` contains a numeric
snapshot read back from these tables and the full table → script → command map.
