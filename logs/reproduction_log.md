# Reproduction log — AIDD8122 Final Project (Huang Yong, P2600222)

Generated 2026-11-05T02:22:48+00:00.
This file snapshots the **committed** result tables; it is not a record of a
past console session. Run the commands below to regenerate those files.

## Environment

See `environment.txt`. Python 3.13.14, ChEMBL release 37,
RDKit 2026.09.1, XGBoost 3.4.1, scikit-learn 1.9.1, NumPy 2.5.4.

## Fixed random seeds

- Benchmark / transfer / analysis: seeds 0-4 (5 seeds).
- Selectivity protocols: seeds 0-19 (20 seeds).
- Generative campaign: seeds fixed in `src/06_generative.py`; all three arms
  share the identical 24-generation / population-96 budget.
- Paired bootstrap: 5,000 replicates, numpy `default_rng(20261105)`.

## Execution order

```bash
pip install -r requirements.txt
python src/01_fetch_chembl.py                      # data/raw/*.json
python src/02_curate.py                            # data/processed/*.csv, curation_log.csv, dataset_summary.csv
python src/14_sync_scaffold_summary.py             # reconcile scaffold counts (empty-string rule)
python src/03_benchmark.py --target FAP --seeds 5 # Table 2, 3; benchmark_*.csv
python src/04_transfer.py --seeds 20              # Table 4, 5, 6; selectivity_*.csv
python src/05_analysis.py                         # Table 7, 8, 9; conformal.csv, temporal.csv, ad_bins.csv
python src/06_generative.py --generations 24 --pop 96
python src/07_figures.py && python src/08_workflow.py
python src/11_reference_dois.py                   # reference_dois.csv (Crossref)
python src/12_design_feasibility.py               # retrosynthesis.csv, radiolabel_feasibility.csv
python src/13_robustness.py                       # Table C1-C3, Figure 8
python src/15_patch_report.py                     # idempotent; verifies all text patches applied
python src/16_evidence_bundle.py                  # this log
python src/09_build_report.py                     # final .docx (Word then updates TOC and exports PDF)
python src/pack_submission.py                     # submission package
```

## Table -> script map (the acceptance check in the optimisation plan)

| Report table | Producing script | Primary file |
|---|---|---|
| Table 1 | 02_curate.py (+14) | dataset_summary.csv, curation_log.csv |
| Table 2 | 03_benchmark.py | benchmark_FAP_combined_runs.csv |
| Table 3 | 03_benchmark.py | ablation_features_FAP.csv |
| Table 4 | 04_transfer.py | transfer_runs.csv |
| Table 5 | 04_transfer.py | selectivity_cv.csv, selectivity_ci.csv |
| Table 6 | 05_analysis.py | ad_bins.csv |
| Table 7 | 05_analysis.py | conformal.csv |
| Table 8 | 05_analysis.py | temporal.csv |
| Table 9 | 06_generative.py | generative_archive.csv |
| Table 10 | 06_generative.py | generative_top15_ga.csv |
| Table B1, B2 | 12_design_feasibility.py | radiolabel_feasibility.csv, retrosynthesis.csv |
| Table C1, C2, C3 | 13_robustness.py | selectivity_bootstrap.csv, selectivity_stratified.csv, mondrian_conformal.csv |

Per-compound predictions are in `benchmark_FAP_combined_preds.parquet` and
`transfer_preds.parquet`, so every metric can be recomputed without refitting.

## Numeric snapshot read back from the committed tables

- Table 1: FAP 384 compounds / 161 Murcko scaffolds; DPP4 4338 / 1596.
- Table 2: XGBoost random: R2 0.75, RMSE 0.71 (5 seeds).
- Table 2: XGBoost scaffold: R2 0.68, RMSE 0.83 (5 seeds).
- Table 2: XGBoost time: R2 -0.79, RMSE 2.05 (5 seeds).
- Table 8: raw time split RMSE 2.07, R2 -0.82; mean offset +1.18. (single fit at seed 0; the five-seed mean in Table 2 is -0.79 / 2.05)
- Table 9: ga: 361 novel, median novel pFAP 8.24, pass 61%, median QED 0.61.
- Table 9: affinity_only: 369 novel, median novel pFAP 9.36, pass 100%, median QED 0.25.
- Table 9: random: 300 novel, median novel pFAP 8.58, pass 60%, median QED 0.52.
- Table 10: 6 SMILES, pFAP 9.50-9.62, MW 493-564 Da.
- Table 7: nominal 80% coverage, in / out of domain: 0.765 / 0.543 (5 seeds).
- Table 7: nominal 90% coverage, in / out of domain: 0.884 / 0.771 (5 seeds).
- Table C1: indirect paired delta 0.463 [0.392, 0.532].
- Table C1: direct paired delta 0.454 [0.382, 0.525].

## Known caveats (recorded, not hidden)

- Mondrian calibration uses only ~27 out-of-domain training compounds; its
  90% quantile over-covers (1.00) and doubles the interval width. It is a
  sensitivity analysis, not a deployable interval.
- The 30-compound mixed-assay stratum reverses the selectivity cliff;
  assay-type confounding there is unresolved.
- Retrosynthetic step counts are BRICS-based heuristics, not experimental routes.
- No docking, MD, FEP or protein-structure prediction was performed (route A).
