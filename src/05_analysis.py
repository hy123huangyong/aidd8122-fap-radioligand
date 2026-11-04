"""
Step 5 -- Trust analysis: applicability domain, conformal uncertainty, temporal drift.

Three questions a medicinal chemist will ask before using a model:

  1. Can I trust a prediction for *this* molecule?      -> applicability domain
  2. How wide is the error bar on that prediction?      -> conformal prediction
  3. What happens when the literature moves on?         -> temporal drift + recalibration

Outputs: results/tables/ad_bins.csv, conformal.csv, temporal.csv
         results/figures/fig1..fig6
"""
from __future__ import annotations

import numpy as np
import pandas as pd
from sklearn.model_selection import KFold
from xgboost import XGBRegressor

import aidd_utils as U

U.set_style()
import matplotlib.pyplot as plt


def xgb(seed: int, n: int = 700) -> XGBRegressor:
    return XGBRegressor(
        n_estimators=n, learning_rate=0.03, max_depth=6, subsample=0.8,
        colsample_bytree=0.4, reg_lambda=1.0, min_child_weight=3,
        n_jobs=4, random_state=seed, tree_method="hist",
    )


# ------------------------------------------------------------------ dataset图
def fig_dataset(fap: pd.DataFrame, prep: pd.DataFrame, dpp4: pd.DataFrame) -> None:
    fig, axes = plt.subplots(1, 3, figsize=(9.2, 2.7))

    ax = axes[0]
    bins = np.arange(3.8, 10.4, 0.4)
    for name, df, col in (("FAP", fap, U.PALETTE["FAP"]),
                          ("PREP", prep, U.PALETTE["PREP"]),
                          ("DPP4", dpp4, U.PALETTE["DPP4"])):
        ax.hist(df.pchembl, bins=bins, histtype="step", lw=1.4, density=True,
                color=col, label=f"{name} (n={len(df)})")
    ax.axvline(7.0, ls=":", c="k", lw=0.8)
    ax.text(7.08, ax.get_ylim()[1] * 0.03, "100 nM", fontsize=7, rotation=90)
    ax.set_xlabel("pActivity (-log$_{10}$ M)")
    ax.set_ylabel("density")
    ax.legend(loc="upper right", fontsize=6.8, handlelength=1.4)
    ax.set_title("a  Potency coverage", loc="left")

    ax = axes[1]
    fap = fap.copy()
    fap["era"] = np.where(fap.year_max >= 2019, "2019-2024", "2004-2018")
    for era, c, m in (("2004-2018", U.PALETTE["grey"], "o"), ("2019-2024", U.PALETTE["accent"], "^")):
        d = fap[fap.era == era]
        ax.scatter(d.MW, d.pchembl, s=14, alpha=0.75, c=c, marker=m, label=f"{era} (n={len(d)})",
                   linewidths=0)
    ax.set_xlabel("molecular weight (Da)")
    ax.set_ylabel("pActivity")
    ax.legend(loc="lower right", fontsize=7)
    ax.set_title("b  Two publication eras", loc="left")

    ax = axes[2]
    counts = [fap.scaffold.nunique(), prep.scaffold.nunique(), dpp4.scaffold.nunique()]
    ax.bar(["FAP", "PREP", "DPP4"], counts,
           color=[U.PALETTE["FAP"], U.PALETTE["PREP"], U.PALETTE["DPP4"]], width=0.6)
    for i, (c, n) in enumerate(zip(counts, [len(fap), len(prep), len(dpp4)])):
        ax.text(i, c + 25, f"{c}\nscaffolds\n({n} cpds)", ha="center", fontsize=6.5)
    ax.set_ylabel("unique Murcko scaffolds")
    ax.set_ylim(0, max(counts) * 1.42)
    ax.set_title("c  Scaffold diversity", loc="left")
    fig.tight_layout()
    fig.savefig(U.FIG / "fig1_dataset.png")
    plt.close(fig)


# --------------------------------------------------------------- AD + conformal
def ad_and_conformal(fap: pd.DataFrame, seeds: int = 5) -> None:
    y = fap.pchembl.to_numpy()
    X = U.featurize(fap.smiles.tolist(), "combined")

    ad_rows, conf_rows = [], []
    for seed in U.SEEDS[:seeds]:
        for protocol in ("scaffold", "random"):
            if protocol == "scaffold":
                tr, te = U.scaffold_split(fap.scaffold.tolist(), 0.2, seed)
            else:
                tr, te = U.random_split(y, 0.2, seed)

            model = xgb(seed)
            model.fit(X[tr], y[tr])
            pred = model.predict(X[te])

            sim = U.max_tanimoto(fap.smiles.to_numpy()[te].tolist(),
                                 fap.smiles.to_numpy()[tr].tolist())
            if protocol == "scaffold":
                ad_rows.append(pd.DataFrame({
                    "seed": seed, "sim": sim, "abs_err": np.abs(pred - y[te]),
                    "y_true": y[te], "y_pred": pred,
                }))

            # conformal calibration on out-of-fold residuals inside the training set
            oof = np.zeros(len(tr))
            kf = KFold(n_splits=5, shuffle=True, random_state=seed)
            for a, b in kf.split(X[tr]):
                m = xgb(seed)
                m.fit(X[tr][a], y[tr][a])
                oof[b] = m.predict(X[tr][b])
            resid = y[tr] - oof
            for alpha in (0.20, 0.10):
                q = U.conformal_quantile(resid, alpha)
                lo, hi = pred - q, pred + q
                covered = (y[te] >= lo) & (y[te] <= hi)
                conf_rows.append({
                    "protocol": protocol, "seed": seed, "level": 1 - alpha, "q": q,
                    "coverage": float(np.mean(covered)),
                    "width": 2 * q,
                })
                # conditional coverage inside / outside the applicability domain
                for tag, mask in (("in_domain(sim>=0.6)", sim >= 0.6),
                                  ("out_domain(sim<0.6)", sim < 0.6)):
                    if mask.sum() >= 5:
                        conf_rows.append({
                            "protocol": f"{protocol}:{tag}", "seed": seed,
                            "level": 1 - alpha, "q": q,
                            "coverage": float(np.mean(covered[mask])),
                            "width": 2 * q,
                        })

    ad = pd.concat(ad_rows, ignore_index=True)
    ad["bin"] = pd.cut(ad.sim, bins=[0, 0.3, 0.4, 0.5, 0.6, 0.7, 1.0],
                       labels=["<0.3", "0.3-0.4", "0.4-0.5", "0.5-0.6", "0.6-0.7", ">0.7"])
    bins = ad.groupby("bin", observed=True).agg(
        n=("abs_err", "size"), rmse=("abs_err", lambda s: float(np.sqrt((s ** 2).mean()))),
        mae=("abs_err", "mean"), sim_median=("sim", "median"),
    ).reset_index()
    bins.to_csv(U.TAB / "ad_bins.csv", index=False)
    print(bins.to_string(index=False))

    conf = pd.DataFrame(conf_rows)
    conf.to_csv(U.TAB / "conformal.csv", index=False)
    print(conf.groupby(["protocol", "level"])[["coverage", "width"]].mean().round(2).to_string())

    # figure: error vs similarity + coverage/error trade-off
    fig, axes = plt.subplots(1, 2, figsize=(7.4, 3.0))
    ax = axes[0]
    ax.scatter(ad.sim, ad.abs_err, s=13, alpha=0.45, c=U.PALETTE["FAP"], linewidths=0)
    med = ad.groupby("bin", observed=True).abs_err.median()
    ax.plot(bins.sim_median, med.values, "o-", c=U.PALETTE["accent"], lw=1.4, ms=4,
            label="median |error| per similarity bin")
    ax.axhline(0.5, ls="--", c="k", lw=0.8)
    ax.text(0.34, 0.56, "0.5 log unit", fontsize=7)
    ax.set_xlabel("max Tanimoto similarity to the training set")
    ax.set_ylabel("|predicted - measured| (log units)")
    ax.legend(loc="upper left", fontsize=7)
    ax.set_title("a  Applicability domain", loc="left")

    ax = axes[1]
    labels = ["<0.55", "0.55-0.70", ">0.70"]
    masks = [ad.sim < 0.55, (ad.sim >= 0.55) & (ad.sim < 0.70), ad.sim >= 0.70]
    rmse_b = [float(np.sqrt((ad.abs_err[m] ** 2).mean())) for m in masks]
    n_b = [int(m.sum()) for m in masks]
    bars = ax.bar(labels, rmse_b, color=[U.PALETTE["PREP"], U.PALETTE["accent"], U.PALETTE["FAP"]],
                  width=0.55)
    overall = float(np.sqrt((ad.abs_err ** 2).mean()))
    ax.axhline(overall, ls="--", c="k", lw=0.8)
    ax.text(2.42, overall + 0.02, "overall", fontsize=7, ha="right")
    for b_, n_ in zip(bars, n_b):
        ax.text(b_.get_x() + b_.get_width() / 2, b_.get_height() + 0.02,
                f"n={n_}", ha="center", fontsize=7)
    ax.set_ylabel("RMSE (log units)")
    ax.set_ylim(0, max(rmse_b) * 1.25)
    ax.set_title("b  Error by applicability class", loc="left")
    fig.tight_layout()
    fig.savefig(U.FIG / "fig4_applicability_domain.png")
    plt.close(fig)


# ------------------------------------------------------------------ time drift
def temporal(fap: pd.DataFrame) -> None:
    y = fap.pchembl.to_numpy()
    X = U.featurize(fap.smiles.tolist(), "combined")
    tr, te = U.time_split(fap.year_max.to_numpy(), 2018)

    model = xgb(0)
    model.fit(X[tr], y[tr])
    pred = model.predict(X[te])

    raw = U.regression_metrics(y[te], pred)
    offset = float(np.mean(y[te][: len(te) // 4] - pred[: len(te) // 4]))
    recal = U.regression_metrics(y[te], pred + offset)
    print(f"time split: raw R2={raw['R2']:.3f} RMSE={raw['RMSE']:.3f} | "
          f"offset={offset:+.2f} -> R2={recal['R2']:.3f} RMSE={recal['RMSE']:.3f}")

    pd.DataFrame([
        {"setting": "train<=2018, test 2019-2024", **raw, "mean_offset": 0.0},
        {"setting": "intercept recalibrated on 25% of test", **recal, "mean_offset": offset},
    ]).to_csv(U.TAB / "temporal.csv", index=False)

    # figure
    fig, axes = plt.subplots(1, 2, figsize=(7.4, 3.0))
    ax = axes[0]
    rnd_tr, rnd_te = U.random_split(y, 0.2, 0)
    m2 = xgb(0)
    m2.fit(X[rnd_tr], y[rnd_tr])
    p2 = m2.predict(X[rnd_te])
    ax.scatter(y[rnd_te], p2, s=16, c=U.PALETTE["grey"], alpha=0.8, linewidths=0,
               label=f"random split (R$^2$={U.regression_metrics(y[rnd_te], p2)['R2']:.2f})")
    ax.plot([4, 10], [4, 10], "k--", lw=0.8)
    ax.set_xlabel("measured pActivity")
    ax.set_ylabel("predicted pActivity")
    ax.legend(loc="upper left", fontsize=7)
    ax.set_title("a  Interpolation", loc="left")

    ax = axes[1]
    ax.scatter(y[te], pred, s=18, c=U.PALETTE["PREP"], alpha=0.8, linewidths=0,
               label=f"time split (R$^2$={raw['R2']:.2f})")
    ax.scatter(y[te], pred + offset, s=12, c=U.PALETTE["accent"], alpha=0.8, linewidths=0,
               marker="^", label=f"recalibrated (R$^2$={recal['R2']:.2f})")
    ax.plot([4, 10], [4, 10], "k--", lw=0.8)
    ax.set_xlabel("measured pActivity")
    ax.set_ylabel("predicted pActivity")
    ax.legend(loc="upper left", fontsize=7)
    ax.set_title("b  Extrapolation to a new era", loc="left")
    fig.tight_layout()
    fig.savefig(U.FIG / "fig5_temporal_drift.png")
    plt.close(fig)


# ------------------------------------------------------------- generative fig
def generative() -> None:
    traj = pd.read_csv(U.TAB / "generative_trajectory.csv")
    arc = pd.read_csv(U.TAB / "generative_archive.csv")

    fig, axes = plt.subplots(1, 3, figsize=(9.4, 2.8))

    # a  population fitness under the three regimes
    ax = axes[0]
    for mode, lab, col in (("ga", "multi-objective", U.PALETTE["FAP"]),
                           ("affinity_only", "affinity only", U.PALETTE["accent"]),
                           ("random", "random", U.PALETTE["grey"])):
        d = traj[traj["mode"] == mode]
        ax.plot(d.generation, d.median_pFAP, lw=1.4, color=col, label=lab)
    ax.set_xlabel("generation")
    ax.set_ylabel("median pFAP of the breeding population")
    ax.set_ylim(6.0, 9.3)
    ax.legend(loc="lower left", fontsize=7)
    ax.set_title("a  Selection pressure", loc="left")

    # b  predicted affinity vs selectivity margin
    ax = axes[1]
    for mode, col, m in (("ga", U.PALETTE["FAP"], "o"),
                         ("affinity_only", U.PALETTE["accent"], "^"),
                         ("random", U.PALETTE["grey"], "s")):
        d = arc[arc["mode"] == mode]
        ax.scatter(d.pFAP, d.margin, s=9, alpha=0.5, c=col, marker=m, linewidths=0, label=mode)
    ax.axvline(8.0, ls=":", c="k", lw=0.8)
    ax.axhline(1.0, ls=":", c="k", lw=0.8)
    ax.set_xlim(6.4, 9.9)
    ax.set_ylim(-2.4, 4.4)
    ax.set_xlabel("predicted pFAP")
    ax.set_ylabel("predicted margin (pFAP $-$ max paralogue)")
    ax.legend(loc="lower right", fontsize=6.8, framealpha=0.9, handletextpad=0.4)
    ax.set_title("b  Archived molecules", loc="left")

    # c  drug-likeness pays for single-objective greed
    ax = axes[2]
    order = ["ga", "affinity_only", "random"]
    labs = ["multi-\nobjective", "affinity\nonly", "random"]
    data = [arc[arc["mode"] == mo].qed for mo in order]
    bp = ax.boxplot(data, tick_labels=labs, widths=0.5, patch_artist=True, showfliers=False,
                    medianprops={"color": "k", "lw": 1.0})
    for patch, mo in zip(bp["boxes"], order):
        patch.set_facecolor(U.PALETTE.get({"ga": "FAP", "affinity_only": "accent",
                                           "random": "grey"}[mo]))
        patch.set_alpha(0.75)
        patch.set_edgecolor("#555555")
    med = [float(np.median(d)) for d in data]
    for i, v in enumerate(med, start=1):
        ax.text(i + 0.28, v, f"{v:.2f}", fontsize=7, va="center")
    ax.set_ylabel("QED (drug-likeness)")
    ax.set_title("c  Developability", loc="left")

    fig.tight_layout()
    fig.savefig(U.FIG / "fig7_generative.png")
    plt.close(fig)


def main() -> None:
    fap, prep, dpp4 = U.load_target("FAP"), U.load_target("PREP"), U.load_target("DPP4")
    fig_dataset(fap, prep, dpp4)
    ad_and_conformal(fap)
    temporal(fap)
    generative()
    print("figures written to", U.FIG)


if __name__ == "__main__":
    main()
