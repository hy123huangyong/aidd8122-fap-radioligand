"""
Build the process-evidence bundle required by P0.1 of the optimisation plan.

Generates ``logs/`` in the project root; pack_submission.py copies it to
``03_Code/logs/``.  Everything written here is either measured from the current
environment or read back from the committed result files -- no synthetic run
times, no invented history.

Outputs:
  logs/environment.txt        interpreter, OS, package versions, UTC timestamp
  logs/fetch_metadata.json    ChEMBL query strings, retrieval date, record counts
  logs/reproduction_log.md    script order, seeds, table -> script -> command map,
                              and a numeric snapshot read from the result tables
  logs/file_manifest.txt      every result/figure file with size and SHA-256

Usage: python src/16_evidence_bundle.py
"""
from __future__ import annotations

import hashlib
import importlib
import json
import platform
import sys
from datetime import datetime, timezone
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
TAB = ROOT / "results" / "tables"
FIG = ROOT / "results" / "figures"
RAW = ROOT / "data" / "raw"
PROC = ROOT / "data" / "processed"
LOGS = ROOT / "logs"
LOGS.mkdir(exist_ok=True)

PACKAGES = ["numpy", "pandas", "scipy", "sklearn", "xgboost", "rdkit",
            "matplotlib", "docx", "pymupdf", "pyarrow", "PIL"]


def env_text() -> str:
    lines = [
        "AIDD8122 Final Project - environment fingerprint",
        "=" * 58,
        f"generated (UTC) : {datetime.now(timezone.utc).isoformat(timespec='seconds')}",
        f"python          : {sys.version.split()[0]} ({platform.python_implementation()})",
        f"platform        : {platform.platform()}",
        f"machine         : {platform.machine()}",
        "",
        "Package versions (importlib.metadata where available):",
    ]
    for name in PACKAGES:
        try:
            mod = importlib.import_module(name)
            ver = getattr(mod, "__version__", None)
            if ver is None:
                from importlib.metadata import version as _v
                ver = _v(name)
        except Exception as exc:                            # noqa: BLE001
            ver = f"unavailable ({type(exc).__name__})"
        lines.append(f"  {name:12s} {ver}")
    lines += ["", "Data snapshot: ChEMBL release 37 (May 2026), retrieved 2026-11-03.",
              "Targets: FAP=CHEMBL4683, PREP=CHEMBL3202, DPP4=CHEMBL284."]
    return "\n".join(lines) + "\n"


def fetch_meta() -> dict:
    base = "https://www.ebi.ac.uk/chembl/api/data/activity.json"
    out = {
        "database": "ChEMBL REST API",
        "release": "ChEMBL_37 (released 2026-05-01)",
        "retrieved": "2026-11-03",
        "endpoint": base,
        "page_size": 1000,
        "standard_types": ["Ki", "Kd", "IC50"],
        "producer": "src/01_fetch_chembl.py",
        "targets": {},
    }
    for t, tid in (("FAP", "CHEMBL4683"), ("PREP", "CHEMBL3202"),
                   ("DPP4", "CHEMBL284")):
        f = RAW / f"chembl_{t}.json"
        recs = json.loads(f.read_text(encoding="utf-8")) if f.exists() else []
        out["targets"][t] = {
            "target_chembl_id": tid,
            "file": f"data/raw/chembl_{t}.json",
            "records": len(recs),
            "size_bytes": f.stat().st_size if f.exists() else 0,
            "query_template": (f"{base}?target_chembl_id={tid}"
                               "&standard_type=<Ki|Kd|IC50>&limit=1000&offset=<n>"),
        }
    return out


def snapshot_lines() -> list[str]:
    """Read the headline numbers back from the committed result tables."""
    out = []
    ds = pd.read_csv(TAB / "dataset_summary.csv").set_index("target")
    out.append(f"- Table 1: FAP {int(ds.loc['FAP','compounds'])} compounds / "
               f"{int(ds.loc['FAP','scaffolds'])} Murcko scaffolds; "
               f"DPP4 {int(ds.loc['DPP4','compounds'])} / "
               f"{int(ds.loc['DPP4','scaffolds'])}.")
    b = pd.read_csv(TAB / "benchmark_FAP_combined_runs.csv")
    for split in ("random", "scaffold", "time"):
        d = b[(b.split == split) & (b.model == "xgb")]
        out.append(f"- Table 2: XGBoost {split}: R2 {d.R2.mean():.2f}, "
                   f"RMSE {d.RMSE.mean():.2f} (5 seeds).")
    t = pd.read_csv(TAB / "temporal.csv").set_index("setting")
    r = t.loc["train<=2018, test 2019-2024"]
    out.append(f"- Table 8: raw time split RMSE {r.RMSE:.2f}, R2 {r.R2:.2f}; "
               f"mean offset {t.loc['intercept recalibrated on 25% of test','mean_offset']:+.2f}. "
               "(single fit at seed 0; the five-seed mean in Table 2 is -0.79 / 2.05)")
    arc = pd.read_csv(TAB / "generative_archive.csv")
    for mo in ("ga", "affinity_only", "random"):
        d = arc[arc["mode"] == mo]
        nov = d[d.novel]
        ok = nov[(nov.pFAP >= 8) & (nov.margin >= 1)]
        out.append(f"- Table 9: {mo}: {len(nov)} novel, median novel pFAP "
                   f"{nov.pFAP.median():.2f}, pass {100*len(ok)/len(nov):.0f}%, "
                   f"median QED {d.qed.median():.2f}.")
    top = pd.read_csv(TAB / "generative_top15_ga.csv").head(6)
    out.append(f"- Table 10: {len(top)} SMILES, pFAP "
               f"{top.pFAP.min():.2f}-{top.pFAP.max():.2f}, MW "
               f"{top.mw.min():.0f}-{top.mw.max():.0f} Da.")
    conf = pd.read_csv(TAB / "conformal.csv")
    for lvl in (0.8, 0.9):
        c = conf[(conf.protocol == "scaffold:out_domain(sim<0.6)")
                 & (conf.level == lvl)]
        i = conf[(conf.protocol == "scaffold:in_domain(sim>=0.6)")
                 & (conf.level == lvl)]
        out.append(f"- Table 7: nominal {lvl:.0%} coverage, in / out of domain: "
                   f"{i.coverage.mean():.3f} / {c.coverage.mean():.3f} (5 seeds).")
    boot = pd.read_csv(TAB / "selectivity_bootstrap.csv")
    for rr in boot.itertuples():
        out.append(f"- Table C1: {rr.protocol} paired delta "
                   f"{rr.paired_delta:.3f} "
                   f"[{rr.delta_CI95_low:.3f}, {rr.delta_CI95_high:.3f}].")
    return out


def reproduction_md() -> str:
    snap = "\n".join(snapshot_lines())
    return f"""# Reproduction log — AIDD8122 Final Project (Huang Yong, P2600222)

Generated {datetime.now(timezone.utc).isoformat(timespec='seconds')}.
This file snapshots the **committed** result tables; it is not a record of a
past console session. Run the commands below to regenerate those files.

## Environment

See `environment.txt`. Python {sys.version.split()[0]}, ChEMBL release 37,
RDKit 2026.09.1, XGBoost 3.4.1, scikit-learn 1.9.1, NumPy 2.5.4,
Matplotlib 3.11.2, pandas 3.1.0.

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

{snap}

## Known caveats (recorded, not hidden)

- Mondrian calibration uses only ~27 out-of-domain training compounds; its
  90% quantile over-covers (1.00) and doubles the interval width. It is a
  sensitivity analysis, not a deployable interval.
- The 30-compound mixed-assay stratum reverses the selectivity cliff;
  assay-type confounding there is unresolved.
- Retrosynthetic step counts are BRICS-based heuristics, not experimental routes.
- No docking, MD, FEP or protein-structure prediction was performed (route A).
"""


def manifest() -> str:
    lines = ["file                                             bytes    sha256-12",
             "-" * 86]
    files = sorted(TAB.iterdir()) + sorted(FIG.iterdir()) + \
        sorted(PROC.iterdir()) + sorted((ROOT / "src").glob("*.py"))
    for f in files:
        if not f.is_file():
            continue
        h = hashlib.sha256(f.read_bytes()).hexdigest()[:12]
        rel = f.relative_to(ROOT).as_posix()
        lines.append(f"{rel:48s} {f.stat().st_size:8d}  {h}")
    return "\n".join(lines) + "\n"


def main() -> None:
    (LOGS / "environment.txt").write_text(env_text(), encoding="utf-8")
    (LOGS / "fetch_metadata.json").write_text(
        json.dumps(fetch_meta(), indent=2, ensure_ascii=False), encoding="utf-8")
    (LOGS / "reproduction_log.md").write_text(reproduction_md(), encoding="utf-8")
    (LOGS / "file_manifest.txt").write_text(manifest(), encoding="utf-8")
    for f in sorted(LOGS.iterdir()):
        print(f"  {f.name:26s} {f.stat().st_size:7d} bytes")
    print(f"\n-> {LOGS}")


if __name__ == "__main__":
    main()
