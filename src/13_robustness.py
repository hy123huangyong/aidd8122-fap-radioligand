"""
Robustness analyses added in the optimisation pass (route A).

Part A  Paired bootstrap confidence intervals for the selectivity "cliff".
        Pairs the random- and scaffold-split AUROC obtained on the SAME seed,
        so the CI refers to the within-seed difference rather than to two
        independent intervals.

Part B  Stratified re-analysis of the selectivity experiment.  Re-runs the
        identical 04_transfer.py protocol inside strata defined by assay type,
        publication era and replicate status, to check whether the cliff is
        driven by one subgroup.  No new data source is used.

Part C  Mondrian (class-conditional) conformal prediction as a sensitivity
        analysis for the out-of-domain coverage failure in section 5.6.
        Calibration residuals are binned by each training compound's maximum
        Tanimoto similarity to the OTHER training compounds; test compounds
        use the bin matching their own max similarity to the training set.
        The binning variable is therefore defined identically on both sides
        and is available at prediction time (no labels needed).

Outputs:
  results/tables/selectivity_bootstrap.csv
  results/tables/selectivity_stratified.csv
  results/tables/mondrian_conformal.csv
  results/figures/fig8_mondrian.png

Usage: python src/13_robustness.py
"""
from __future__ import annotations

import sys
from pathlib import Path

import numpy as np
import pandas as pd
from rdkit import DataStructs
from sklearn.metrics import roc_auc_score
from sklearn.model_selection import KFold
from xgboost import XGBRegressor

sys.path.insert(0, str(Path(__file__).resolve().parent))
import aidd_utils as U                                    # noqa: E402

ROOT = Path(__file__).resolve().parents[1]
TAB = ROOT / "results" / "tables"
FIG = ROOT / "results" / "figures"
AD_THR = 0.60                                            # single AD threshold used report-wide
B = 5000
SEEDS = list(range(20))


def xgb(seed: int, n: int = 700, depth: int = 6) -> XGBRegressor:
    return XGBRegressor(
        n_estimators=n, learning_rate=0.03, max_depth=depth, subsample=0.8,
        colsample_bytree=0.4, reg_lambda=1.0, min_child_weight=3,
        n_jobs=4, random_state=seed, tree_method="hist",
    )


# --------------------------------------------------------------------- Part A
def paired_bootstrap() -> pd.DataFrame:
    cv = pd.read_csv(TAB / "selectivity_cv.csv")
    rng = np.random.default_rng(20261105)
    rows = []
    for proto in ("indirect", "direct"):
        piv = (cv[cv.protocol == proto]
               .pivot(index="seed", columns="split", values="selectivity_AUROC")
               .dropna())
        r, s = piv["random"].to_numpy(), piv["scaffold"].to_numpy()
        delta = r - s
        idx = rng.integers(0, len(delta), size=(B, len(delta)))
        bd = delta[idx].mean(axis=1)
        br = r[idx].mean(axis=1)
        bs = s[idx].mean(axis=1)
        rows.append({
            "protocol": proto,
            "n_paired_seeds": len(delta),
            "AUROC_random": round(float(r.mean()), 3),
            "random_CI95": f"[{np.percentile(br, 2.5):.3f}, {np.percentile(br, 97.5):.3f}]",
            "AUROC_scaffold": round(float(s.mean()), 3),
            "scaffold_CI95": f"[{np.percentile(bs, 2.5):.3f}, {np.percentile(bs, 97.5):.3f}]",
            "paired_delta": round(float(delta.mean()), 3),
            "delta_CI95_low": round(float(np.percentile(bd, 2.5)), 3),
            "delta_CI95_high": round(float(np.percentile(bd, 97.5)), 3),
            "pct_delta_positive": round(100 * float((bd > 0).mean()), 1),
        })
    out = pd.DataFrame(rows)
    out.to_csv(TAB / "selectivity_bootstrap.csv", index=False)
    return out


# --------------------------------------------------------------------- Part B
def _auroc(delta_te, p) -> float:
    lab = (delta_te >= 1).astype(int)
    return float(roc_auc_score(lab, p)) if 0 < lab.sum() < len(lab) else np.nan


def stratified() -> pd.DataFrame:
    fap = U.load_target("FAP").reset_index(drop=True)
    prep = U.load_target("PREP").reset_index(drop=True)
    dual = fap.merge(prep[["inchikey", "pchembl"]], on="inchikey",
                     suffixes=("", "_prep")).reset_index(drop=True)
    delta_all = (dual.pchembl.to_numpy() - dual.pchembl_prep.to_numpy())

    strata = {
        "all (reference)": np.ones(len(dual), dtype=bool),
        "assay type: IC50 only": (dual.types == "IC50").to_numpy(),
        "assay type: mixed / other": (dual.types != "IC50").to_numpy(),
        "era: 2004-2018": (dual.year_max < 2019).to_numpy(),
        "era: 2019-2024": (dual.year_max >= 2019).to_numpy(),
        "replicated (>1 assay)": (dual.n_measure > 1).to_numpy(),
        "single measurement": (dual.n_measure <= 1).to_numpy(),
    }
    rows = []
    for name, mask in strata.items():
        idx = np.where(mask)[0]
        if len(idx) < 20:
            rows.append({"stratum": name, "n": len(idx), "note": "too few compounds to fit"})
            continue
        X = U.featurize(dual.smiles.iloc[idx].tolist(), "combined")
        d = delta_all[idx]
        scaf = dual.scaffold.iloc[idx].tolist()
        by_split = {"random": [], "scaffold": []}
        for seed in SEEDS:
            for split in ("random", "scaffold"):
                if split == "random":
                    tr, te = U.random_split(d, 0.25, seed)
                else:
                    tr, te = U.scaffold_split(scaf, 0.25, seed)
                if len(te) < 8 or len(tr) < 12:
                    continue
                lab_tr = (d[tr] >= 1).astype(int)
                if lab_tr.sum() == 0 or lab_tr.sum() == len(lab_tr):
                    continue
                mf, mp = xgb(seed, 400, 5), xgb(seed, 400, 5)
                mf.fit(X[tr], dual.pchembl.to_numpy()[idx][tr])
                mp.fit(X[tr], dual.pchembl_prep.to_numpy()[idx][tr])
                a = _auroc(d[te], mf.predict(X[te]) - mp.predict(X[te]))
                if not np.isnan(a):
                    by_split[split].append(a)
        rr, ss = by_split["random"], by_split["scaffold"]
        rows.append({
            "stratum": name,
            "n": len(idx),
            "n_folds_random": len(rr),
            "AUROC_random": round(float(np.mean(rr)), 3) if rr else np.nan,
            "n_folds_scaffold": len(ss),
            "AUROC_scaffold": round(float(np.mean(ss)), 3) if ss else np.nan,
            "cliff_delta": round(float(np.mean(rr) - np.mean(ss)), 3) if rr and ss else np.nan,
            "note": "AUROC unstable when stratum n < 30; interpret as a consistency check",
        })
    out = pd.DataFrame(rows)
    out.to_csv(TAB / "selectivity_stratified.csv", index=False)
    return out


# --------------------------------------------------------------------- Part C
def mondrian() -> tuple[pd.DataFrame, pd.DataFrame]:
    fap = U.load_target("FAP").reset_index(drop=True)
    X = U.featurize(fap.smiles.tolist(), "combined")
    y = fap.pchembl.to_numpy()
    smi = fap.smiles.to_numpy()
    rows, sanity = [], []

    for seed in range(5):
        tr, te = U.scaffold_split(fap.scaffold.tolist(), 0.2, seed)
        m = xgb(seed)
        m.fit(X[tr], y[tr])
        pred = m.predict(X[te])

        oof = np.zeros(len(tr))
        kf = KFold(n_splits=5, shuffle=True, random_state=seed)
        for a, b in kf.split(X[tr]):
            mm = xgb(seed)
            mm.fit(X[tr][a], y[tr][a])
            oof[b] = mm.predict(X[tr][b])
        resid = y[tr] - oof

        # internal (leave-one-out) similarity among the training compounds
        fps = [U.morgan_fp(s) for s in smi[tr]]
        mat = np.array([DataStructs.BulkTanimotoSimilarity(f, fps) for f in fps])
        np.fill_diagonal(mat, -1.0)
        sim_in = mat.max(axis=1)
        sim_te = U.max_tanimoto(smi[te].tolist(), smi[tr].tolist())

        for alpha in (0.20, 0.10):
            level = 1 - alpha
            q_glob = U.conformal_quantile(resid, alpha)
            q_in = U.conformal_quantile(resid[sim_in >= AD_THR], alpha) \
                if (sim_in >= AD_THR).sum() >= 5 else q_glob
            q_out = U.conformal_quantile(resid[sim_in < AD_THR], alpha) \
                if (sim_in < AD_THR).sum() >= 5 else q_glob

            for tag, mask in (("overall", np.ones(len(te), dtype=bool)),
                              ("in_domain", sim_te >= AD_THR),
                              ("out_domain", sim_te < AD_THR)):
                if mask.sum() < 3:
                    continue
                cov_g = float(np.mean(np.abs(y[te][mask] - pred[mask]) <= q_glob))
                q_m = np.where(sim_te[mask] >= AD_THR, q_in, q_out)
                cov_m = float(np.mean(np.abs(y[te][mask] - pred[mask]) <= q_m))
                rows.append({
                    "seed": seed, "level": level, "domain": tag, "n": int(mask.sum()),
                    "coverage_global": cov_g, "width_global": 2 * q_glob,
                    "coverage_mondrian": cov_m, "width_mondrian": float(np.mean(2 * q_m)),
                    "q_global": q_glob, "q_in": q_in, "q_out": q_out,
                })
            sanity.append({"seed": seed, "level": level,
                           "coverage_global_all": float(np.mean(np.abs(y[te] - pred) <= q_glob)),
                           "n_train_in": int((sim_in >= AD_THR).sum()),
                           "n_train_out": int((sim_in < AD_THR).sum())})

    det = pd.DataFrame(rows)
    det.to_csv(TAB / "mondrian_conformal.csv", index=False)
    return det, pd.DataFrame(sanity)


def figure(det: pd.DataFrame) -> None:
    U.set_style()
    import matplotlib.pyplot as plt
    g = det.groupby(["level", "domain"])[["coverage_global", "coverage_mondrian",
                                          "width_global", "width_mondrian"]].mean().reset_index()
    fig, axes = plt.subplots(1, 2, figsize=(8.2, 3.1))
    for ax, level in zip(axes, [0.8, 0.9]):
        sub = g[g.level == level]
        x = np.arange(len(sub))
        ax.bar(x - 0.18, sub.coverage_global, 0.36, label="global split conformal",
               color=U.PALETTE["PREP"])
        ax.bar(x + 0.18, sub.coverage_mondrian, 0.36, label="Mondrian (AD-conditional)",
               color=U.PALETTE["FAP"])
        ax.axhline(level, ls="--", c="k", lw=1, label=f"nominal {level:.0%}")
        ax.set_xticks(x)
        ax.set_xticklabels([d.replace("_", " ") for d in sub.domain], fontsize=8)
        ax.set_ylim(0, 1.05)
        ax.set_ylabel("empirical coverage")
        ax.set_title(f"a  nominal {level:.0%}" if level == 0.8 else f"b  nominal {level:.0%}",
                     loc="left")
        ax.legend(fontsize=7, loc="lower right")
    fig.suptitle("Figure 8. Class-conditional conformal calibration (5 seeds, scaffold split)",
                 fontsize=9, y=1.02)
    fig.tight_layout()
    FIG.mkdir(parents=True, exist_ok=True)
    fig.savefig(FIG / "fig8_mondrian.png", dpi=330)
    plt.close(fig)


def main() -> None:
    a = paired_bootstrap()
    print("=== A. paired bootstrap (20 seeds) ===")
    print(a.to_string(index=False))

    b = stratified()
    print("\n=== B. stratified selectivity re-analysis ===")
    print(b.to_string(index=False))

    det, sanity = mondrian()
    print("\n=== C. Mondrian conformal: global vs conditional (mean over 5 seeds) ===")
    print(det.groupby(["level", "domain"])[["coverage_global", "coverage_mondrian",
                                            "width_global", "width_mondrian"]]
          .mean().round(3).to_string())
    print("\ncalibration-bin sizes / global-coverage sanity check")
    print(sanity.round(3).to_string(index=False))
    figure(det)
    print("\nfigure -> results/figures/fig8_mondrian.png")


if __name__ == "__main__":
    main()
