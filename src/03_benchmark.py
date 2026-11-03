"""
Step 3 -- Benchmark of descriptor/fingerprint models under three split strategies.

Why three splits: random splits leak analogues into the test set and are known to
inflate QSAR performance; scaffold splits emulate "new chemotype" forecasting;
a temporal split emulates prospective use (train on the literature up to a year,
predict everything published afterwards).

Models
    mean      : training-set mean (reference floor)
    ridge     : L2 linear model on Morgan fingerprints
    rf        : random forest on Morgan + MACCS + 15 physicochemical descriptors
    xgb       : gradient-boosted trees on the same feature block
    svr       : RBF support-vector regression on scaled descriptors

Usage
    python src/03_benchmark.py --target FAP --seeds 5
    python src/03_benchmark.py --target DPP4 --seeds 2
"""
from __future__ import annotations

import argparse
from pathlib import Path

import numpy as np
import pandas as pd
from joblib import Parallel, delayed
from sklearn.ensemble import RandomForestRegressor
from sklearn.pipeline import make_pipeline
from sklearn.preprocessing import StandardScaler
from sklearn.linear_model import Ridge
from sklearn.svm import SVR
from xgboost import XGBRegressor

import aidd_utils as U

MODELS = ("mean", "ridge", "rf", "xgb", "svr")


def build_model(name: str, seed: int, n_features: int):
    if name == "mean":
        return None
    if name == "ridge":
        return make_pipeline(StandardScaler(with_mean=False), Ridge(alpha=10.0))
    if name == "rf":
        return RandomForestRegressor(
            n_estimators=500, max_features=0.33, min_samples_leaf=2,
            n_jobs=-1, random_state=seed,
        )
    if name == "xgb":
        return XGBRegressor(
            n_estimators=700, learning_rate=0.03, max_depth=6, subsample=0.8,
            colsample_bytree=0.4, reg_lambda=1.0, min_child_weight=3,
            n_jobs=4, random_state=seed, tree_method="hist",
        )
    if name == "svr":
        return make_pipeline(StandardScaler(), SVR(C=10.0, epsilon=0.3, gamma="scale"))
    raise ValueError(name)


def run_one(df: pd.DataFrame, X: np.ndarray, split: str, seed: int, model: str) -> dict:
    y = df.pchembl.to_numpy()
    if split == "random":
        tr, te = U.random_split(y, 0.2, seed)
    elif split == "scaffold":
        tr, te = U.scaffold_split(df.scaffold.tolist(), 0.2, seed)
    elif split == "time":
        tr, te = U.time_split(df.year_max.to_numpy(), 2018)
    else:
        raise ValueError(split)

    if len(te) < 10 or len(tr) < 50:
        return {}

    est = build_model(model, seed, X.shape[1])
    if est is None:
        pred = np.full(len(te), y[tr].mean())
    else:
        est.fit(X[tr], y[tr])
        pred = est.predict(X[te])

    row = {
        "target": df.target.iloc[0], "split": split, "model": model, "seed": seed,
        "n_train": len(tr), "n_test": len(te),
    }
    row.update(U.regression_metrics(y[te], pred))
    row.update({f"screen_{k}": v for k, v in U.screening_metrics(y[te], pred).items()})
    return {"row": row, "pred": pd.DataFrame({
        "inchikey": df.inchikey.to_numpy()[te], "smiles": df.smiles.to_numpy()[te],
        "y_true": y[te], "y_pred": pred, "split": split, "model": model, "seed": seed,
    })}


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--target", default="FAP")
    ap.add_argument("--seeds", type=int, default=5)
    ap.add_argument("--feature", default="combined")
    args = ap.parse_args()

    df = U.load_target(args.target)
    X = U.featurize(df.smiles.tolist(), args.feature)
    print(f"{args.target}: {len(df)} compounds, X={X.shape}")

    seeds = U.SEEDS[: args.seeds]
    jobs = [(split, seed, m) for split in ("random", "scaffold", "time")
            for seed in seeds for m in MODELS]
    if args.target != "FAP":  # keep big targets cheap
        jobs = [(s, sd, m) for s, sd, m in jobs if s != "time"]

    out = Parallel(n_jobs=6, verbose=5)(delayed(run_one)(df, X, s, sd, m) for s, sd, m in jobs)
    rows = [o["row"] for o in out if o]
    preds = pd.concat([o["pred"] for o in out if o], ignore_index=True)

    res = pd.DataFrame(rows)
    tag = f"{args.target}_{args.feature}"
    res.to_csv(U.TAB / f"benchmark_{tag}_runs.csv", index=False)
    preds.to_parquet(U.TAB / f"benchmark_{tag}_preds.parquet", index=False)

    cols = ["RMSE", "MAE", "R2", "Spearman", "screen_AUROC", "screen_EF10"]
    agg = res.groupby(["split", "model"])[cols].agg(["mean", "std"]).round(3)
    agg.to_csv(U.TAB / f"benchmark_{tag}_summary.csv")
    print(agg.to_string())

    # feature ablation with the strongest tree model
    abl_rows = []
    for feat in ("morgan", "maccs", "desc", "combined"):
        Xf = U.featurize(df.smiles.tolist(), feat)
        for seed in seeds:
            r = run_one(df, Xf, "scaffold", seed, "xgb")
            if r:
                abl_rows.append({**r["row"], "features": feat})
    abl = pd.DataFrame(abl_rows)
    abl.to_csv(U.TAB / f"ablation_features_{args.target}.csv", index=False)
    print(abl.groupby("features")[["RMSE", "R2", "Spearman"]].mean().round(3).to_string())


if __name__ == "__main__":
    main()
