"""
Step 7 -- Publication figures for the benchmark and the selectivity study.

fig2  model benchmark: accuracy under three split protocols + virtual-screening
      enrichment + calibration of the best model
fig3  transfer from the paralogue family and the selectivity problem

All numbers are read from results/tables, so the figures can be regenerated
without refitting anything.
"""
from __future__ import annotations

import matplotlib
import matplotlib.transforms

import numpy as np
import pandas as pd
from scipy import stats

import aidd_utils as U

U.set_style()
import matplotlib.pyplot as plt


MODEL_ORDER = ["mean", "ridge", "svr", "rf", "xgb"]
MODEL_LABEL = {"mean": "mean", "ridge": "ridge", "svr": "SVR", "rf": "RF", "xgb": "XGBoost"}
MODEL_COLOR = {
    "mean": "#b9b9b9", "ridge": "#8fb8c9", "svr": "#5b8fa8", "rf": "#2f6f8b", "xgb": "#123f52",
}
SPLIT_LABEL = {"random": "Random", "scaffold": "Scaffold", "time": "Time"}


def fig_benchmark() -> None:
    runs = pd.read_csv(U.TAB / "benchmark_FAP_combined_runs.csv")
    preds = pd.read_parquet(U.TAB / "benchmark_FAP_combined_preds.parquet")

    fig, axes = plt.subplots(1, 3, figsize=(9.4, 2.75))

    # ---- a  RMSE across split protocols
    ax = axes[0]
    splits = ["random", "scaffold", "time"]
    width = 0.16
    for k, model in enumerate(MODEL_ORDER):
        d = runs[runs.model == model]
        means = [d[d.split == s].RMSE.mean() for s in splits]
        sds = [d[d.split == s].RMSE.std() for s in splits]
        x = np.arange(len(splits)) + (k - 2) * width
        ax.bar(x, means, width, yerr=sds, capsize=1.6, color=MODEL_COLOR[model],
               label=MODEL_LABEL[model], error_kw={"lw": 0.7, "elinewidth": 0.7})
    ax.set_xticks(np.arange(len(splits)))
    ax.set_xticklabels([SPLIT_LABEL[s] for s in splits])
    ax.set_ylabel("RMSE (log units)")
    ax.set_ylim(0, 2.75)
    ax.legend(ncol=5, loc="upper center", bbox_to_anchor=(0.5, -0.20), fontsize=6.2,
              columnspacing=0.8, handlelength=0.9, handletextpad=0.3, borderaxespad=0)
    ax.set_title("a  Accuracy by split protocol", loc="left")

    # ---- b  calibration of the best model
    ax = axes[1]
    for split, c, m in (("random", U.PALETTE["grey"], "o"), ("time", U.PALETTE["accent"], "^")):
        d = preds[(preds.model == "xgb") & (preds.split == split)]
        ax.scatter(d.y_true, d.y_pred, s=13, alpha=0.65, c=c, marker=m, linewidths=0,
                   label=f"{SPLIT_LABEL[split]} (R$^2$="
                         f"{U.regression_metrics(d.y_true, d.y_pred)['R2']:.2f})")
    ax.plot([4, 10], [4, 10], "k--", lw=0.8)
    ax.text(9.6, 9.75, "y = x", fontsize=7, ha="right")
    ax.set_xlabel("measured pActivity")
    ax.set_ylabel("predicted pActivity")
    ax.set_xlim(3.8, 10.2)
    ax.set_ylim(3.4, 10.4)
    ax.legend(loc="upper left", fontsize=6.0, handlelength=1.0, handletextpad=0.25,
              borderpad=0.2, frameon=True, facecolor="white", edgecolor="none",
              framealpha=1.0)
    ax.set_title("b  XGBoost calibration", loc="left")

    # ---- c  early enrichment
    ax = axes[2]
    ef = runs.groupby(["split", "model"]).screen_EF10.agg(["mean", "std"])
    for k, model in enumerate(MODEL_ORDER[1:], start=1):
        xs, ys, es = [], [], []
        for j, s in enumerate(splits):
            if (s, model) in ef.index:
                xs.append(j + (k - 1.5) * 0.18)
                ys.append(ef.loc[(s, model), "mean"])
                es.append(ef.loc[(s, model), "std"])
        ax.bar(xs, ys, 0.17, yerr=es, capsize=1.6, color=MODEL_COLOR[model],
               label=MODEL_LABEL[model], error_kw={"lw": 0.7, "elinewidth": 0.7})
    ax.axhline(1.0, ls="--", c="k", lw=0.8)
    ax.text(0.015, 1.10, "no enrichment", fontsize=6.2, ha="left", va="bottom",
            transform=matplotlib.transforms.blended_transform_factory(ax.transAxes, ax.transData),
            bbox={"facecolor": "white", "edgecolor": "none", "pad": 0.8})
    ax.set_xticks(np.arange(len(splits)))
    ax.set_xticklabels([SPLIT_LABEL[s] for s in splits])
    ax.set_ylabel("EF$_{10}$")
    ax.set_ylim(0, 2.65)
    ax.legend(ncol=4, loc="upper center", bbox_to_anchor=(0.5, -0.20), fontsize=6.2,
              handlelength=0.9, handletextpad=0.3, columnspacing=0.7, borderaxespad=0)
    ax.set_title("c  Early enrichment", loc="left")

    fig.tight_layout(rect=(0, 0.06, 1, 1))
    fig.savefig(U.FIG / "fig2_benchmark.png")
    plt.close(fig)


def fig_transfer_selectivity() -> None:
    runs = pd.read_csv(U.TAB / "transfer_runs.csv")
    sel = pd.read_csv(U.TAB / "selectivity_cv.csv")
    dual = pd.read_csv(U.TAB / "selectivity_overview.csv")

    fig, axes = plt.subplots(1, 3, figsize=(9.4, 2.75))

    # ---- a  paired effect of paralogue transfer
    ax = axes[0]
    base = runs[runs.variant == "baseline"].set_index(["split", "seed"]).RMSE
    tr = runs[runs.variant == "stacked_transfer"].set_index(["split", "seed"]).RMSE
    delta = (base - tr).dropna()          # positive = transfer is better
    positions = {"random": 0, "scaffold": 1}
    for i, (split, pos) in enumerate(positions.items()):
        vals = delta.loc[split].to_numpy()
        ax.scatter(np.full(len(vals), pos) + np.linspace(-0.07, 0.07, len(vals)),
                   vals, s=16, color=U.PALETTE["FAP"], zorder=3, linewidths=0)
        mu, se = vals.mean(), vals.std(ddof=1) / np.sqrt(len(vals))
        ax.errorbar(pos + 0.22, mu, yerr=1.96 * se, fmt="D", ms=4.5,
                    color=U.PALETTE["accent"], capsize=2.5, lw=1.0, zorder=4)
        p = stats.wilcoxon(vals).pvalue if len(vals) >= 5 else np.nan
        ax.text(pos, max(vals.max(), 0) + 0.012, f"p={p:.2f}", ha="center", fontsize=6.8)
    ax.axhline(0, ls="--", c="k", lw=0.8)
    ax.set_xticks(list(positions.values()))
    ax.set_xticklabels([SPLIT_LABEL[s] for s in positions])
    ax.set_xlim(-0.5, 1.6)
    ax.set_ylabel("RMSE$_{baseline}$ - RMSE$_{transfer}$")
    ax.set_title("a  Paralogue transfer", loc="left")

    # ---- b  selectivity is learnable only inside known scaffolds
    ax = axes[1]
    order = [("random", "indirect"), ("random", "direct"),
             ("scaffold", "indirect"), ("scaffold", "direct")]
    colors = [U.PALETTE["PREP"], "#e8b4a8", U.PALETTE["FAP"], "#9dc0cf"]
    for i, ((split, proto), c) in enumerate(zip(order, colors)):
        v = sel[(sel.split == split) & (sel.protocol == proto)].selectivity_AUROC.dropna()
        ax.bar(i, v.mean(), 0.62, yerr=1.96 * v.std(ddof=1) / np.sqrt(len(v)),
               color=c, capsize=2.2, error_kw={"lw": 0.8, "elinewidth": 0.8})
        ax.text(i, 0.06, f"n={len(v)}", ha="center", fontsize=6.2, color="white")
    ax.axhline(0.5, ls="--", c="k", lw=0.9)
    ax.text(3.45, 0.525, "chance", fontsize=6.8, ha="right")
    ax.set_xticks(range(4))
    ax.set_xticklabels(["indirect", "direct", "indirect", "direct"], fontsize=7.4)
    ax.text(0.5, -0.28, "Random", ha="center", fontsize=8, transform=ax.get_xaxis_transform())
    ax.text(2.5, -0.28, "Scaffold", ha="center", fontsize=8, transform=ax.get_xaxis_transform())
    ax.set_ylim(0, 1.0)
    ax.set_ylabel("AUROC (selective vs not)")
    ax.set_title("b  Selectivity prediction", loc="left")

    # ---- c  how selective are the dual-measured compounds?
    ax = axes[2]
    fap = U.load_target("FAP")
    prep = U.load_target("PREP")
    dpp4 = U.load_target("DPP4")
    d1 = fap.merge(prep[["inchikey", "pchembl"]], on="inchikey", suffixes=("", "_o"))
    d2 = fap.merge(dpp4[["inchikey", "pchembl"]], on="inchikey", suffixes=("", "_o"))
    bins = np.arange(-4, 4.4, 0.5)
    ax.hist(d1.pchembl - d1.pchembl_o, bins=bins, histtype="step", lw=1.3,
            color=U.PALETTE["PREP"], density=True, label=f"vs PREP (n={len(d1)})")
    ax.hist(d2.pchembl - d2.pchembl_o, bins=bins, histtype="step", lw=1.3,
            color=U.PALETTE["DPP4"], density=True, label=f"vs DPP4 (n={len(d2)})")
    ax.axvline(1.0, ls=":", c="k", lw=0.9)
    ax.set_ylim(0, 0.58)
    ax.text(1.06, 0.99, "10-fold selective", fontsize=6.2, ha="left", va="top",
            transform=ax.get_xaxis_transform(),
            bbox={"facecolor": "white", "edgecolor": "none", "pad": 0.8})
    ax.set_xlabel("$\\Delta$pActivity (FAP - paralogue)")
    ax.set_ylabel("density")
    ax.legend(loc="upper left", fontsize=6.0, handlelength=1.0, handletextpad=0.25,
              borderpad=0.2, frameon=True, facecolor="white", edgecolor="none",
              framealpha=1.0)
    ax.set_title("c  Selectivity margin", loc="left")

    fig.tight_layout(rect=(0, 0.07, 1, 1))
    fig.savefig(U.FIG / "fig3_transfer_selectivity.png")
    plt.close(fig)


if __name__ == "__main__":
    fig_benchmark()
    fig_transfer_selectivity()
    print("fig2, fig3 written")
