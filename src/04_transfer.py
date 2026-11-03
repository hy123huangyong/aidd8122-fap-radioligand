"""
Step 4 -- Transfer from the paralogue family, and prediction of selectivity.

(a) Transfer.  PREP and DPP4 have 1.3x and 11x more compounds than FAP. A shared
    neural encoder is not an option at n=384, so transfer enters as *features*:
    paralogue models are trained on their own datasets, and their predictions for
    each FAP compound (plus the chemical similarity to each paralogue training set)
    are appended to the FAP feature block.

(b) Selectivity.  Two protocols are compared for the therapeutically decisive
    quantity pFAP - pPREP:
        indirect : difference of the two single-target models
        direct   : a dedicated model trained on the 100 dual-measured compounds
    Both are evaluated with random and scaffold cross-validation, because a
    selectivity model that only works inside known scaffolds is not usable for
    scaffold hopping.

Usage: python src/04_transfer.py --seeds 5
"""
from __future__ import annotations

import argparse

import numpy as np
import pandas as pd
from sklearn.metrics import roc_auc_score
from xgboost import XGBRegressor

import aidd_utils as U


def xgb(seed: int, n_estimators: int = 700, depth: int = 6) -> XGBRegressor:
    return XGBRegressor(
        n_estimators=n_estimators, learning_rate=0.03, max_depth=depth, subsample=0.8,
        colsample_bytree=0.4, reg_lambda=1.0, min_child_weight=3,
        n_jobs=4, random_state=seed, tree_method="hist",
    )


def build_transfer_features(fap: pd.DataFrame, seed: int) -> pd.DataFrame:
    """Paralogue-prediction + similarity features for every FAP compound."""
    feats = pd.DataFrame(index=fap.index)
    Xq = U.featurize(fap.smiles.tolist(), "combined")
    for name in ("PREP", "DPP4"):
        par = U.load_target(name)
        model = xgb(seed, n_estimators=400)
        model.fit(U.featurize(par.smiles.tolist(), "combined"), par.pchembl.to_numpy())
        feats[f"pred_{name}"] = model.predict(Xq)
        feats[f"sim_{name}"] = U.max_tanimoto(fap.smiles.tolist(), par.smiles.tolist())
    return feats


def transfer_experiment(seeds: int) -> None:
    fap = U.load_target("FAP")
    y = fap.pchembl.to_numpy()
    X_base = U.featurize(fap.smiles.tolist(), "combined")

    rows, preds = [], []
    for seed in U.SEEDS[:seeds]:
        extra = build_transfer_features(fap, seed)
        X_tr = np.hstack([X_base, extra.to_numpy(dtype=np.float32)])

        for split in ("random", "scaffold"):
            if split == "random":
                tr, te = U.random_split(y, 0.2, seed)
            else:
                tr, te = U.scaffold_split(fap.scaffold.tolist(), 0.2, seed)

            for variant, X in (("baseline", X_base),
                               ("stacked_transfer", X_tr),
                               ("transfer_nosim", X_tr[:, :-2])):
                model = xgb(seed)
                model.fit(X[tr], y[tr])
                p = model.predict(X[te])
                row = {"variant": variant, "split": split, "seed": seed,
                       "n_train": len(tr), "n_test": len(te)}
                row.update(U.regression_metrics(y[te], p))
                row.update({f"screen_{k}": v for k, v in U.screening_metrics(y[te], p).items()})
                rows.append(row)
                preds.append(pd.DataFrame({
                    "inchikey": fap.inchikey.to_numpy()[te], "y_true": y[te],
                    "y_pred": p, "split": split, "model": variant, "seed": seed,
                }))

    res = pd.DataFrame(rows)
    res.to_csv(U.TAB / "transfer_runs.csv", index=False)
    pd.concat(preds).to_parquet(U.TAB / "transfer_preds.parquet", index=False)
    cols = ["RMSE", "MAE", "R2", "Spearman", "screen_AUROC", "screen_EF10"]
    print(res.groupby(["split", "variant"])[cols].agg(["mean", "std"]).round(3).to_string())


def selectivity_experiment(seeds: int = 5) -> None:
    fap, prep, dpp4 = U.load_target("FAP"), U.load_target("PREP"), U.load_target("DPP4")
    dual = fap.merge(prep[["inchikey", "pchembl"]], on="inchikey", suffixes=("", "_prep"))
    print(f"dual FAP/PREP compounds: {len(dual)} ({dual.scaffold.nunique()} scaffolds)")

    X = U.featurize(dual.smiles.tolist(), "combined")
    y_f, y_p = dual.pchembl.to_numpy(), dual.pchembl_prep.to_numpy()
    delta = y_f - y_p
    rows = []

    for seed in range(seeds):
        for split in ("random", "scaffold"):
            if split == "random":
                tr, te = U.random_split(delta, 0.25, seed)
            else:
                tr, te = U.scaffold_split(dual.scaffold.tolist(), 0.25, seed)
            if len(te) < 12:
                continue

            # indirect: difference of two independently trained single-target models
            mf, mp = xgb(seed, 400, 5), xgb(seed, 400, 5)
            mf.fit(X[tr], y_f[tr])
            mp.fit(X[tr], y_p[tr])
            p_ind = mf.predict(X[te]) - mp.predict(X[te])

            # direct: one model on the difference
            md = xgb(seed, 400, 5)
            md.fit(X[tr], delta[tr])
            p_dir = md.predict(X[te])

            for protocol, p in (("indirect", p_ind), ("direct", p_dir)):
                row = {"protocol": protocol, "split": split, "seed": seed,
                       "n_test": len(te), "n_selective": int((delta[te] >= 1).sum())}
                m = U.regression_metrics(delta[te], p)
                row.update({f"delta_{k}": v for k, v in m.items()})
                lab = (delta[te] >= 1).astype(int)
                row["selectivity_AUROC"] = (
                    float(roc_auc_score(lab, p)) if 0 < lab.sum() < len(lab) else np.nan
                )
                rows.append(row)

    out = pd.DataFrame(rows)
    out.to_csv(U.TAB / "selectivity_cv.csv", index=False)
    cols = ["delta_R2", "delta_Spearman", "delta_MAE", "selectivity_AUROC"]
    print(out.groupby(["split", "protocol"])[cols].agg(["mean", "std"]).round(3).to_string())

    # 95% CI of the mean across seeds (t-interval via normal approximation)
    ci = out.groupby(["split", "protocol"])[cols].agg(["mean", "std", "count"])
    ci.columns = [f"{a}_{b}" for a, b in ci.columns]
    for c in cols:
        ci[f"{c}_ci95"] = 1.96 * ci[f"{c}_std"] / np.sqrt(ci[f"{c}_count"])
    ci.round(3).to_csv(U.TAB / "selectivity_ci.csv")

    # descriptive statistics of the selectivity problem itself
    desc = []
    for name, other in (("PREP", prep), ("DPP4", dpp4)):
        d2 = fap.merge(other[["inchikey", "pchembl"]], on="inchikey", suffixes=("", "_o"))
        dd = d2.pchembl - d2.pchembl_o
        desc.append({
            "pair": f"FAP/{name}", "n": len(d2), "scaffolds": d2.scaffold.nunique(),
            "delta_mean": round(float(dd.mean()), 2),
            "frac_selective_ge1": round(float((dd >= 1).mean()), 3),
            "frac_nonselective_le0": round(float((dd <= 0).mean()), 3),
        })
    pd.DataFrame(desc).to_csv(U.TAB / "selectivity_overview.csv", index=False)
    print(pd.DataFrame(desc).to_string(index=False))


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--seeds", type=int, default=5)
    ap.add_argument("--skip-transfer", action="store_true")
    args = ap.parse_args()
    if not args.skip_transfer:
        transfer_experiment(args.seeds)
    selectivity_experiment(args.seeds)
