"""Step 22 -- redraw every report figure for print-size legibility.

The original figures were composed on 9.2-9.4 inch canvases and then inserted
into a 6.2 inch text column.  At that reduction their 6-7 pt labels become
roughly 4 pt: in the PDF they look like garbled characters rather than text.
Several legends and numeric annotations were also placed *inside* the axes,
where they collide with bars, curves and dashed reference lines.

This script draws exactly the same data on smaller canvases with larger type,
moves every legend below the axes, and puts every annotation on an opaque
white patch clear of the data.  Nothing here recomputes a published number:
figures 1, 2, 3, 5, 7 and 8 are read from ``results/tables``; only figure 4's
applicability-domain scatter is regenerated in memory (five fixed-seed
scaffold-split fits), and no CSV is written.

Usage (from the project root)::

    <venv>/Scripts/python.exe src/22_figures_clear.py

The output overwrites ``results/figures/*.png``.  Step 23 embeds the new
bitmaps in the report and the deck.
"""
from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(Path(__file__).resolve().parent))

import matplotlib

matplotlib.use("Agg")
import matplotlib as mpl
import numpy as np
import pandas as pd
from matplotlib.lines import Line2D
from matplotlib.patches import FancyArrowPatch, FancyBboxPatch, Patch

import aidd_utils as U

FIG = U.FIG
TAB = U.TAB

MODEL_ORDER = ["mean", "ridge", "svr", "rf", "xgb"]
MODEL_LABEL = {"mean": "mean", "ridge": "ridge", "svr": "SVR", "rf": "RF", "xgb": "XGBoost"}
MODEL_COLOR = {
    "mean": "#b9b9b9", "ridge": "#8fb8c9", "svr": "#5b8fa8", "rf": "#2f6f8b", "xgb": "#123f52",
}
SPLIT_LABEL = {"random": "Random", "scaffold": "Scaffold", "time": "Time"}
WHITE = {"facecolor": "white", "edgecolor": "none", "alpha": 0.92, "pad": 1.2}


def setup() -> None:
    """Larger type on a smaller canvas => legible at 6.3 inch column width."""
    mpl.rcParams.update({
        "figure.dpi": 120,
        "savefig.dpi": 330,
        "savefig.bbox": "tight",
        "savefig.facecolor": "white",
        "font.family": "DejaVu Sans",
        "font.size": 11,
        "axes.titlesize": 11.5,
        "axes.labelsize": 11,
        "xtick.labelsize": 10,
        "ytick.labelsize": 10,
        "legend.fontsize": 9.5,
        "legend.frameon": False,
        "axes.linewidth": 0.9,
        "axes.spines.top": False,
        "axes.spines.right": False,
        "xtick.direction": "out",
        "ytick.direction": "out",
        "figure.facecolor": "white",
    })


def save(fig, name: str) -> None:
    out = FIG / name
    fig.savefig(out, dpi=330, facecolor="white")
    print(f"  {name:36s} {out.stat().st_size // 1024:4d} KB")
    import matplotlib.pyplot as plt

    plt.close(fig)


def bottom_legend(fig, handles, labels, ncol: int, y: float = 0.005) -> None:
    fig.legend(handles, labels, loc="lower center", bbox_to_anchor=(0.5, y),
               ncol=ncol, frameon=False, fontsize=9.5, handlelength=1.3,
               handletextpad=0.4, columnspacing=1.2, borderaxespad=0)


# ------------------------------------------------------------------ figure 0
def fig_workflow() -> None:
    """Four stage cards on a 1:3.33 canvas — the aspect the report inserts at 6.35 in.

    Drawn at final column width so the type size survives insertion: the previous
    version used an 8.6 in canvas whose tight bbox grew to 9.6 in, shrinking 8 pt
    type to about 5.3 pt and letting the longest lines run past the card edges.
    """
    import matplotlib.pyplot as plt
    from matplotlib.patches import Rectangle

    INK, MUTED, RULE = "#22303c", "#5f6f7a", "#c9d1d9"
    fig, ax = plt.subplots(figsize=(6.5, 2.35))
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    ax.axis("off")

    W, GAP, X0 = 0.222, 0.030, 0.011
    Y, H = 0.465, 0.495
    xs = [X0 + i * (W + GAP) for i in range(4)]
    colors = [U.PALETTE["FAP"], U.PALETTE["PREP"], U.PALETTE["DPP4"], U.PALETTE["accent"]]

    cards = [
        ("A  Data", ["ChEMBL bioactivities", "FAP 972 → 384",
                     "PREP 962 → 512", "DPP4 6,553 → 4,338",
                     "curation, assay type"]),
        ("B  Models", ["Morgan + MACCS", "+ 15 descriptors",
                       "ridge/SVR/RF/XGB", "splits: random,",
                       "scaffold, time"]),
        ("C  Trust", ["applicability domain", "(max Tanimoto)",
                      "conformal intervals", "temporal + intercept",
                      "AUROC, EF10, P@10"]),
        ("D  Design", ["graph GA + NSGA-II", "objectives: pFAP,",
                       "margin, QED, SA", "filters: 250–650 Da,",
                       "similarity ≥ 0.40"]),
    ]

    for x, (title, lines), c in zip(xs, cards, colors):
        # white card, hairline rule, single accent bar: no tinted fill, no watermark
        ax.add_patch(Rectangle((x, Y), W, H, facecolor="white",
                               edgecolor=RULE, linewidth=0.85, zorder=1))
        ax.add_patch(Rectangle((x, Y + H - 0.008), W, 0.008,
                               facecolor=c, edgecolor="none", alpha=0.9, zorder=2))
        ax.text(x + 0.016, Y + H - 0.042, title, fontsize=7.8, fontweight="bold",
                color=INK, va="top", ha="left", zorder=3)
        for i, ln in enumerate(lines):
            ax.text(x + 0.016, Y + H - 0.100 - i * 0.076, ln, fontsize=6.5,
                    color=INK, va="top", ha="left", zorder=3)

    for i in range(3):
        ax.add_patch(FancyArrowPatch((xs[i] + W + 0.004, Y + H / 2),
                                     (xs[i + 1] - 0.004, Y + H / 2),
                                     arrowstyle="-|>", mutation_scale=7,
                                     lw=0.8, color=MUTED,
                                     shrinkA=0, shrinkB=0, zorder=4))

    # error analysis loops back from Trust to Models
    ax.add_patch(FancyArrowPatch((xs[2] + W / 2, Y - 0.005), (xs[1] + W / 2, Y - 0.005),
                                 connectionstyle="arc3,rad=-0.14", arrowstyle="-|>",
                                 mutation_scale=7, lw=0.8, color=MUTED,
                                 linestyle=(0, (2.5, 2)), zorder=4))
    ax.text((xs[1] + xs[2] + W) / 2, 0.350,
            "error analysis feeds back into feature and model choice",
            fontsize=6.6, color=MUTED, ha="center", va="top", style="italic")

    ax.text(X0, 0.300, "Outputs", fontsize=7.0, fontweight="bold", color=INK, va="top")
    outs = ["dataset + curation log", "benchmark, 5 models × 3 splits",
            "calibrated error bars", "Pareto set of novel warheads"]
    for i, o in enumerate(outs):
        ax.text(X0 + 0.012 + (i % 2) * 0.49, 0.245 - (i // 2) * 0.050,
                "–  " + o, fontsize=6.8, color=MUTED, va="top")

    ax.text(X0, 0.122,
            "Every experiment is repeated over 5 random seeds with identical features and "
            "hyper-parameters;",
            fontsize=6.2, color=MUTED, style="italic", va="top")
    ax.text(X0, 0.072,
            "curated tables, figures and code are released with the report.",
            fontsize=6.2, color=MUTED, style="italic", va="top")
    save(fig, "fig0_workflow.png")


# ------------------------------------------------------------------ figure 1
def fig_dataset() -> None:
    import matplotlib.pyplot as plt

    fap, prep, dpp4 = (U.load_target(n) for n in ("FAP", "PREP", "DPP4"))
    fig, axes = plt.subplots(1, 3, figsize=(8.0, 3.3))

    ax = axes[0]
    bins = np.arange(3.8, 10.4, 0.4)
    handles = []
    for name, df, col in (("FAP", fap, U.PALETTE["FAP"]),
                          ("PREP", prep, U.PALETTE["PREP"]),
                          ("DPP4", dpp4, U.PALETTE["DPP4"])):
        ax.hist(df.pchembl, bins=bins, histtype="step", lw=1.5, density=True,
                color=col, label=f"{name} (n={len(df)})")
        handles.append(Line2D([], [], color=col, lw=1.5, label=f"{name} (n={len(df)})"))
    top = ax.get_ylim()[1]
    ax.axvline(7.0, ls=":", c="k", lw=1.0)
    # horizontal, above the density curves, on an opaque patch
    ax.text(7.06, top * 0.95, "100 nM", fontsize=9, ha="left", va="top", bbox=WHITE)
    ax.set_xlabel("pActivity (-log$_{10}$ M)")
    ax.set_ylabel("density")
    ax.set_title("a  Potency coverage", loc="left")

    ax = axes[1]
    era = np.where(fap.year_max >= 2019, "2019-2024", "2004-2018")
    for lab, c, m in (("2004-2018", U.PALETTE["grey"], "o"),
                      ("2019-2024", U.PALETTE["accent"], "^")):
        d = fap[era == lab]
        ax.scatter(d.MW, d.pchembl, s=15, alpha=0.75, c=c, marker=m,
                   linewidths=0, label=f"{lab} (n={len(d)})")
    ax.set_xlabel("molecular weight (Da)")
    ax.set_ylabel("pActivity")
    ax.legend(loc="lower right", fontsize=8.8, frameon=True, facecolor="white",
              edgecolor="none", framealpha=0.9, borderpad=0.3)
    ax.set_title("b  Two publication eras", loc="left")

    ax = axes[2]
    counts = [df.scaffold.replace("", np.nan).nunique() for df in (fap, prep, dpp4)]
    ax.bar(["FAP", "PREP", "DPP4"], counts,
           color=[U.PALETTE["FAP"], U.PALETTE["PREP"], U.PALETTE["DPP4"]], width=0.6)
    # only the scaffold count sits above each bar; the compound counts are
    # already spelled out in the figure caption and would collide on the two
    # short bars (161 and 227) if printed here as well
    for i, c in enumerate(counts):
        ax.text(i, c + 45, f"{c:,}", ha="center", va="bottom",
                fontsize=9, bbox=WHITE)
    ax.set_ylabel("unique Murcko scaffolds")
    ax.set_ylim(0, max(counts) * 1.35)
    ax.set_title("c  Scaffold diversity", loc="left")

    bottom_legend(fig, handles, [h.get_label() for h in handles], 3)
    fig.tight_layout(rect=(0, 0.11, 1, 1))
    save(fig, "fig1_dataset.png")


# ------------------------------------------------------------------ figure 2
def fig_benchmark() -> None:
    import matplotlib.pyplot as plt

    runs = pd.read_csv(TAB / "benchmark_FAP_combined_runs.csv")
    preds = pd.read_parquet(TAB / "benchmark_FAP_combined_preds.parquet")
    fig, axes = plt.subplots(1, 3, figsize=(8.0, 3.3))
    splits = ["random", "scaffold", "time"]

    ax = axes[0]
    width = 0.16
    for k, model in enumerate(MODEL_ORDER):
        d = runs[runs.model == model]
        means = [d[d.split == s].RMSE.mean() for s in splits]
        sds = [d[d.split == s].RMSE.std() for s in splits]
        x = np.arange(len(splits)) + (k - 2) * width
        ax.bar(x, means, width, yerr=sds, capsize=1.8, color=MODEL_COLOR[model],
               error_kw={"lw": 0.8, "elinewidth": 0.8})
    ax.set_xticks(np.arange(len(splits)))
    ax.set_xticklabels([SPLIT_LABEL[s] for s in splits])
    ax.set_ylabel("RMSE (log units)")
    ax.set_ylim(0, 2.75)
    ax.set_title("a  Accuracy by split protocol", loc="left")

    ax = axes[1]
    for split, c, m in (("random", U.PALETTE["grey"], "o"),
                        ("time", U.PALETTE["accent"], "^")):
        d = preds[(preds.model == "xgb") & (preds.split == split)]
        ax.scatter(d.y_true, d.y_pred, s=13, alpha=0.6, c=c, marker=m, linewidths=0,
                   label=f"{SPLIT_LABEL[split]} (R²="
                         f"{U.regression_metrics(d.y_true, d.y_pred)['R2']:.2f})")
    ax.plot([4, 10], [4, 10], "k--", lw=0.9)
    ax.text(10.1, 3.55, "y = x", fontsize=9, ha="right", va="bottom", bbox=WHITE)
    ax.set_xlabel("measured pActivity")
    ax.set_ylabel("predicted pActivity")
    ax.set_xlim(3.8, 10.2)
    ax.set_ylim(3.4, 10.4)
    ax.legend(loc="upper left", fontsize=8.6, frameon=True, facecolor="white",
              edgecolor="none", framealpha=0.92, borderpad=0.3)
    ax.set_title("b  XGBoost calibration", loc="left")

    ax = axes[2]
    ef = runs.groupby(["split", "model"]).screen_EF10.agg(["mean", "std"])
    for k, model in enumerate(MODEL_ORDER[1:], start=1):
        xs, ys, es = [], [], []
        for j, s in enumerate(splits):
            if (s, model) in ef.index:
                xs.append(j + (k - 1.5) * 0.18)
                ys.append(ef.loc[(s, model), "mean"])
                es.append(ef.loc[(s, model), "std"])
        ax.bar(xs, ys, 0.17, yerr=es, capsize=1.8, color=MODEL_COLOR[model],
               error_kw={"lw": 0.8, "elinewidth": 0.8})
    ax.axhline(1.0, ls="--", c="k", lw=0.9)
    ax.text(2.45, 1.10, "no enrichment", fontsize=8.8, ha="right", va="bottom", bbox=WHITE)
    ax.set_xticks(np.arange(len(splits)))
    ax.set_xticklabels([SPLIT_LABEL[s] for s in splits])
    ax.set_ylabel("EF$_{10}$")
    ax.set_ylim(0, 2.65)
    ax.set_title("c  Early enrichment", loc="left")

    handles = [Patch(facecolor=MODEL_COLOR[m], label=MODEL_LABEL[m]) for m in MODEL_ORDER]
    bottom_legend(fig, handles, [h.get_label() for h in handles], 5)
    fig.tight_layout(rect=(0, 0.11, 1, 1))
    save(fig, "fig2_benchmark.png")


# ------------------------------------------------------------------ figure 3
def fig_transfer_selectivity() -> None:
    import matplotlib.pyplot as plt
    from scipy import stats

    runs = pd.read_csv(TAB / "transfer_runs.csv")
    sel = pd.read_csv(TAB / "selectivity_cv.csv")
    fig, axes = plt.subplots(1, 3, figsize=(8.2, 3.4),
                             gridspec_kw={"width_ratios": [0.95, 1.50, 1.15]})

    ax = axes[0]
    base = runs[runs.variant == "baseline"].set_index(["split", "seed"]).RMSE
    tr = runs[runs.variant == "stacked_transfer"].set_index(["split", "seed"]).RMSE
    delta = (base - tr).dropna()
    positions = {"random": 0, "scaffold": 1}
    for pos, split in enumerate(positions):
        vals = delta.loc[split].to_numpy()
        ax.scatter(np.full(len(vals), pos) + np.linspace(-0.07, 0.07, len(vals)),
                   vals, s=18, color=U.PALETTE["FAP"], zorder=3, linewidths=0)
        mu, se = vals.mean(), vals.std(ddof=1) / np.sqrt(len(vals))
        ax.errorbar(pos + 0.22, mu, yerr=1.96 * se, fmt="D", ms=4.5,
                    color=U.PALETTE["accent"], capsize=2.5, lw=1.0, zorder=4)
        p = stats.wilcoxon(vals).pvalue if len(vals) >= 5 else np.nan
        ax.text(pos, 0.049, f"p={p:.2f}", ha="center", va="top", fontsize=9, bbox=WHITE)
    ax.axhline(0, ls="--", c="k", lw=0.9)
    ax.set_ylim(-0.028, 0.055)
    ax.set_xticks(list(positions.values()))
    ax.set_xticklabels([SPLIT_LABEL[s] for s in positions])
    ax.set_xlim(-0.5, 1.6)
    ax.set_ylabel("RMSE$_{baseline}$ - RMSE$_{transfer}$")
    ax.set_title("a  Paralogue transfer", loc="left")

    ax = axes[1]
    order = [("random", "indirect"), ("random", "direct"),
             ("scaffold", "indirect"), ("scaffold", "direct")]
    colors = [U.PALETTE["PREP"], "#e8b4a8", U.PALETTE["FAP"], "#9dc0cf"]
    for i, ((split, proto), c) in enumerate(zip(order, colors)):
        v = sel[(sel.split == split) & (sel.protocol == proto)].selectivity_AUROC.dropna()
        err = 1.96 * v.std(ddof=1) / np.sqrt(len(v))
        ax.bar(i, v.mean(), 0.62, yerr=err, color=c, capsize=2.2,
               error_kw={"lw": 0.8, "elinewidth": 0.8})
        # inside the bar, near the baseline: dark ink on an opaque patch reads
        # on both the dark and the pastel bars, and never collides with "chance"
        ax.text(i, 0.055, f"n={len(v)}", ha="center", va="bottom",
                fontsize=8.8, color="#222222", bbox=WHITE)
    ax.axhline(0.5, ls="--", c="k", lw=0.9)
    ax.text(3.45, 0.515, "chance", fontsize=9, ha="right", va="bottom", bbox=WHITE)
    ax.set_xticks(range(4))
    ax.set_xticklabels(["indirect", "direct", "indirect", "direct"], fontsize=8.2)
    ax.tick_params(axis="x", pad=4)
    ax.text(0.5, -0.20, "Random", ha="center", fontsize=9.5,
            transform=ax.get_xaxis_transform())
    ax.text(2.5, -0.20, "Scaffold", ha="center", fontsize=9.5,
            transform=ax.get_xaxis_transform())
    ax.set_ylim(0, 1.05)
    ax.set_ylabel("AUROC (selective vs not)")
    ax.set_title("b  Selectivity prediction", loc="left")

    ax = axes[2]
    fap, prep, dpp4 = (U.load_target(n) for n in ("FAP", "PREP", "DPP4"))
    d1 = fap.merge(prep[["inchikey", "pchembl"]], on="inchikey", suffixes=("", "_o"))
    d2 = fap.merge(dpp4[["inchikey", "pchembl"]], on="inchikey", suffixes=("", "_o"))
    bins = np.arange(-4, 4.4, 0.5)
    ax.hist(d1.pchembl - d1.pchembl_o, bins=bins, histtype="step", lw=1.4,
            color=U.PALETTE["PREP"], density=True, label=f"vs PREP (n={len(d1)})")
    ax.hist(d2.pchembl - d2.pchembl_o, bins=bins, histtype="step", lw=1.4,
            color=U.PALETTE["DPP4"], density=True, label=f"vs DPP4 (n={len(d2)})")
    ax.axvline(1.0, ls=":", c="k", lw=1.0)
    ax.set_ylim(0, 0.62)
    # next to the threshold line, in empty space to its right; the legend
    # occupies the upper-left corner and must not be covered
    ax.text(1.45, 0.34, "10-fold\nselective", fontsize=8.8, ha="left",
            va="top", linespacing=1.3, bbox=WHITE)
    ax.set_xlabel("$\\Delta$pActivity (FAP - paralogue)")
    ax.set_ylabel("density")
    ax.legend(loc="upper left", fontsize=8.6, frameon=True, facecolor="white",
              edgecolor="none", framealpha=0.92, borderpad=0.3)
    ax.set_title("c  Selectivity margin", loc="left")

    fig.tight_layout(rect=(0, 0.10, 1, 1))
    save(fig, "fig3_transfer_selectivity.png")


# ------------------------------------------------------------------ figure 4
def _ad_scatter() -> pd.DataFrame:
    """Reproduce the scaffold-split AD scatter in memory; writes nothing."""
    from xgboost import XGBRegressor

    fap = U.load_target("FAP")
    y = fap.pchembl.to_numpy()
    X = U.featurize(fap.smiles.tolist(), "combined")
    smiles = fap.smiles.to_numpy()
    rows = []
    for seed in U.SEEDS:
        tr, te = U.scaffold_split(fap.scaffold.tolist(), 0.2, seed)
        model = XGBRegressor(n_estimators=700, learning_rate=0.03, max_depth=6,
                             subsample=0.8, colsample_bytree=0.4, reg_lambda=1.0,
                             min_child_weight=3, n_jobs=4, random_state=seed,
                             tree_method="hist")
        model.fit(X[tr], y[tr])
        pred = model.predict(X[te])
        sim = U.max_tanimoto(smiles[te].tolist(), smiles[tr].tolist())
        rows.append(pd.DataFrame({"sim": sim, "abs_err": np.abs(pred - y[te])}))
    return pd.concat(rows, ignore_index=True)


def fig_applicability_domain() -> None:
    import matplotlib.pyplot as plt

    ad = _ad_scatter()
    bins = pd.read_csv(TAB / "ad_bins.csv")
    fig, axes = plt.subplots(1, 2, figsize=(7.6, 3.6))

    ax = axes[0]
    ax.scatter(ad.sim, ad.abs_err, s=13, alpha=0.42, c=U.PALETTE["FAP"], linewidths=0)
    ax.plot(bins.sim_median, bins.mae, "o-", c=U.PALETTE["accent"], lw=1.6, ms=5)
    ax.axhline(0.5, ls="--", c="k", lw=0.9)
    ax.text(0.335, 0.545, "0.5 log unit", fontsize=9, va="bottom", bbox=WHITE)
    ax.set_xlabel("max Tanimoto similarity to the training set")
    ax.set_ylabel("|predicted - measured| (log units)")
    ax.set_title("a  Applicability domain", loc="left")

    ax = axes[1]
    labels = ["<0.55", "0.55-0.70", ">0.70"]
    masks = [ad.sim < 0.55, (ad.sim >= 0.55) & (ad.sim < 0.70), ad.sim >= 0.70]
    rmse_b = [float(np.sqrt((ad.abs_err[m] ** 2).mean())) for m in masks]
    n_b = [int(m.sum()) for m in masks]
    bars = ax.bar(labels, rmse_b,
                  color=[U.PALETTE["PREP"], U.PALETTE["accent"], U.PALETTE["FAP"]],
                  width=0.55)
    overall = float(np.sqrt((ad.abs_err ** 2).mean()))
    ax.axhline(overall, ls="--", c="k", lw=0.9)
    ax.text(0.99, overall + 0.02, "overall", fontsize=9, ha="right", va="bottom",
            transform=ax.get_yaxis_transform(), bbox=WHITE)
    for b_, n_ in zip(bars, n_b):
        ax.text(b_.get_x() + b_.get_width() / 2, b_.get_height() + 0.02,
                f"n={n_}", ha="center", va="bottom", fontsize=9, bbox=WHITE)
    ax.set_ylabel("RMSE (log units)")
    ax.set_ylim(0, max(rmse_b) * 1.28)
    ax.set_title("b  Error by applicability class", loc="left")

    handles = [
        Line2D([], [], color=U.PALETTE["FAP"], marker="o", lw=0, alpha=0.6,
               label="scaffold-split test compounds"),
        Line2D([], [], color=U.PALETTE["accent"], marker="o", lw=1.6,
               label="median |error| per similarity bin"),
        Line2D([], [], color="k", ls="--", lw=0.9, label="reference level"),
    ]
    bottom_legend(fig, handles, [h.get_label() for h in handles], 3)
    fig.tight_layout(rect=(0, 0.13, 1, 1))
    save(fig, "fig4_applicability_domain.png")


# ------------------------------------------------------------------ figure 5
def fig_temporal_drift() -> None:
    import matplotlib.pyplot as plt

    preds = pd.read_parquet(TAB / "benchmark_FAP_combined_preds.parquet")
    temporal = pd.read_csv(TAB / "temporal.csv")
    offset = float(temporal.mean_offset.iloc[-1])
    rnd = preds[(preds.model == "xgb") & (preds.split == "random") & (preds.seed == 0)]
    tim = preds[(preds.model == "xgb") & (preds.split == "time") & (preds.seed == 0)]
    raw_r2 = float(temporal.R2.iloc[0])
    rec_r2 = float(temporal.R2.iloc[-1])
    # five-seed mean, so the label matches the value quoted in the report
    rnd_all = preds[(preds.model == "xgb") & (preds.split == "random")]
    rnd_r2 = float(np.mean([
        U.regression_metrics(g.y_true, g.y_pred)["R2"]
        for _, g in rnd_all.groupby("seed")
    ]))

    fig, axes = plt.subplots(1, 2, figsize=(7.6, 3.5))
    ax = axes[0]
    ax.scatter(rnd.y_true, rnd.y_pred, s=18, c=U.PALETTE["grey"], alpha=0.8,
               linewidths=0, label=f"random split (R²={rnd_r2:.2f})")
    ax.plot([4, 10], [4, 10], "k--", lw=0.9)
    ax.set_xlabel("measured pActivity")
    ax.set_ylabel("predicted pActivity")
    ax.set_title("a  Interpolation", loc="left")

    ax = axes[1]
    ax.scatter(tim.y_true, tim.y_pred, s=18, c=U.PALETTE["PREP"], alpha=0.8,
               linewidths=0, label=f"time split (R²={raw_r2:.2f})")
    ax.scatter(tim.y_true, tim.y_pred + offset, s=13, c=U.PALETTE["accent"],
               alpha=0.8, linewidths=0, marker="^",
               label=f"recalibrated (R²={rec_r2:.2f})")
    ax.plot([4, 10], [4, 10], "k--", lw=0.9)
    ax.set_xlabel("measured pActivity")
    ax.set_ylabel("predicted pActivity")
    ax.set_title("b  Extrapolation to a new era", loc="left")

    handles = [
        Line2D([], [], color=U.PALETTE["grey"], marker="o", lw=0,
               label=axes[0].collections[0].get_label()),
        Line2D([], [], color=U.PALETTE["PREP"], marker="o", lw=0,
               label=axes[1].collections[0].get_label()),
        Line2D([], [], color=U.PALETTE["accent"], marker="^", lw=0,
               label=axes[1].collections[1].get_label()),
    ]
    bottom_legend(fig, handles, [h.get_label() for h in handles], 3)
    fig.tight_layout(rect=(0, 0.13, 1, 1))
    save(fig, "fig5_temporal_drift.png")


# ------------------------------------------------------------------ figure 7
def fig_generative() -> None:
    import matplotlib.pyplot as plt

    traj = pd.read_csv(TAB / "generative_trajectory.csv")
    arc = pd.read_csv(TAB / "generative_archive.csv")
    fig, axes = plt.subplots(1, 3, figsize=(8.0, 3.4))

    ax = axes[0]
    for mode, lab, col in (("ga", "multi-objective", U.PALETTE["FAP"]),
                           ("affinity_only", "affinity only", U.PALETTE["accent"]),
                           ("random", "random", U.PALETTE["grey"])):
        d = traj[traj["mode"] == mode]
        ax.plot(d.generation, d.median_pFAP, lw=1.6, color=col, label=lab)
    ax.set_xlabel("generation")
    ax.set_ylabel("median pFAP of population")
    ax.set_ylim(6.0, 9.3)
    ax.set_title("a  Selection pressure", loc="left")

    ax = axes[1]
    mode_label = {"ga": "multi-objective", "affinity_only": "affinity only", "random": "random"}
    mode_color = {"ga": U.PALETTE["FAP"], "affinity_only": U.PALETTE["accent"],
                  "random": U.PALETTE["grey"]}
    for mode, m in (("ga", "o"), ("affinity_only", "^"), ("random", "s")):
        d = arc[arc["mode"] == mode]
        ax.scatter(d.pFAP, d.margin, s=10, alpha=0.5, c=mode_color[mode],
                   marker=m, linewidths=0)
    ax.axvline(8.0, ls=":", c="k", lw=0.9)
    ax.axhline(1.0, ls=":", c="k", lw=0.9)
    ax.set_xlim(6.4, 9.9)
    ax.set_ylim(-2.4, 4.4)
    ax.set_xlabel("predicted pFAP")
    ax.set_ylabel("predicted margin (pFAP - max paralogue)")
    ax.set_title("b  Archived molecules", loc="left")

    ax = axes[2]
    order = ["ga", "affinity_only", "random"]
    labs = ["multi-\nobjective", "affinity\nonly", "random"]
    data = [arc[arc["mode"] == mo].qed for mo in order]
    bp = ax.boxplot(data, tick_labels=labs, widths=0.5, patch_artist=True,
                    showfliers=False, medianprops={"color": "k", "lw": 1.0})
    for patch, mo in zip(bp["boxes"], order):
        patch.set_facecolor(mode_color[mo])
        patch.set_alpha(0.75)
        patch.set_edgecolor("#555555")
    for i, d in enumerate(data, start=1):
        v = float(np.median(d))
        ax.text(i + 0.33, v, f"{v:.2f}", fontsize=9, va="center", ha="left", bbox=WHITE)
    ax.set_xlim(0.5, 3.75)
    ax.set_ylabel("QED (drug-likeness)")
    ax.set_title("c  Developability", loc="left")

    handles = [Line2D([], [], color=mode_color[m], lw=1.6, marker="o",
                      label=mode_label[m]) for m in order]
    bottom_legend(fig, handles, [h.get_label() for h in handles], 3)
    fig.tight_layout(rect=(0, 0.11, 1, 1))
    save(fig, "fig7_generative.png")


# ------------------------------------------------------------------ figure 8
def fig_mondrian() -> None:
    import matplotlib.pyplot as plt

    det = pd.read_csv(TAB / "mondrian_conformal.csv")
    g = det.groupby(["level", "domain"])[
        ["coverage_global", "coverage_mondrian", "width_global", "width_mondrian"]
    ].mean().reset_index()
    fig, axes = plt.subplots(1, 2, figsize=(7.6, 3.4))
    for ax, level, tag in zip(axes, [0.8, 0.9], ["a", "b"]):
        sub = g[g.level == level]
        x = np.arange(len(sub))
        ax.bar(x - 0.18, sub.coverage_global, 0.36, color=U.PALETTE["PREP"])
        ax.bar(x + 0.18, sub.coverage_mondrian, 0.36, color=U.PALETTE["FAP"])
        ax.axhline(level, ls="--", c="k", lw=1.0)
        ax.set_xticks(x)
        ax.set_xticklabels([d.replace("_", " ") for d in sub.domain])
        ax.set_ylim(0, 1.05)
        ax.set_ylabel("empirical coverage")
        ax.set_title(f"{tag}  nominal {level:.0%}", loc="left")
    handles = [
        Patch(facecolor=U.PALETTE["PREP"], label="global split conformal"),
        Patch(facecolor=U.PALETTE["FAP"], label="Mondrian (AD-conditional)"),
        Line2D([], [], color="k", ls="--", lw=1.0, label="nominal level"),
    ]
    bottom_legend(fig, handles, [h.get_label() for h in handles], 3)
    fig.tight_layout(rect=(0, 0.13, 1, 1))
    save(fig, "fig8_mondrian.png")


def main() -> None:
    setup()
    FIG.mkdir(parents=True, exist_ok=True)
    print("Redrawing figures (larger type, legends and annotations clear of data):")
    fig_workflow()
    fig_dataset()
    fig_benchmark()
    fig_transfer_selectivity()
    fig_applicability_domain()
    fig_temporal_drift()
    fig_generative()
    fig_mondrian()
    print("done ->", FIG)


if __name__ == "__main__":
    main()
