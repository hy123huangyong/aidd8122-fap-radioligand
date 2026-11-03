"""Shared helpers: data loading, featurisation, split strategies, metrics, plotting."""
from __future__ import annotations

import json
from pathlib import Path

import numpy as np
import pandas as pd
from rdkit import Chem, DataStructs, RDLogger
from rdkit.Chem import Descriptors, MACCSkeys, rdMolDescriptors
from rdkit.Chem import rdFingerprintGenerator
from sklearn.metrics import mean_absolute_error, mean_squared_error, r2_score, roc_auc_score
from scipy.stats import pearsonr, spearmanr

RDLogger.DisableLog("rdApp.*")

ROOT = Path(__file__).resolve().parents[1]
PROC = ROOT / "data" / "processed"
TAB = ROOT / "results" / "tables"
FIG = ROOT / "results" / "figures"
MODELS = ROOT / "results" / "models"
for p in (TAB, FIG, MODELS):
    p.mkdir(parents=True, exist_ok=True)

SEEDS = [0, 1, 2, 3, 4]
TARGETS = ["FAP", "PREP", "DPP4"]

# ---------------------------------------------------------------- featurisation
_MORGAN = rdFingerprintGenerator.GetMorganGenerator(radius=2, fpSize=2048)
_DESC_NAMES = [
    "MW", "cLogP", "TPSA", "HBD", "HBA", "RotB", "Rings", "AromRings",
    "FracCsp3", "HeavyAtoms", "Nitrile", "Amide", "CarboxylicAcid", "N+O", "Halogen",
]


def _descriptor_vector(mol: Chem.Mol) -> list[float]:
    from rdkit.Chem import Lipinski

    return [
        Descriptors.MolWt(mol),
        Descriptors.MolLogP(mol),
        rdMolDescriptors.CalcTPSA(mol),
        Lipinski.NumHDonors(mol),
        Lipinski.NumHAcceptors(mol),
        rdMolDescriptors.CalcNumRotatableBonds(mol),
        rdMolDescriptors.CalcNumRings(mol),
        rdMolDescriptors.CalcNumAromaticRings(mol),
        rdMolDescriptors.CalcFractionCSP3(mol),
        mol.GetNumHeavyAtoms(),
        len(mol.GetSubstructMatches(Chem.MolFromSmarts("[CX2]#N"))),
        len(mol.GetSubstructMatches(Chem.MolFromSmarts("[NX3][CX3](=[OX1])"))),
        len(mol.GetSubstructMatches(Chem.MolFromSmarts("[CX3](=[OX1])[OX2H1]"))),
        sum(1 for a in mol.GetAtoms() if a.GetSymbol() in ("N", "O")),
        sum(1 for a in mol.GetAtoms() if a.GetSymbol() in ("F", "Cl", "Br", "I")),
    ]


def featurize(smiles: list[str], kind: str = "combined") -> np.ndarray:
    """kind: 'morgan' | 'maccs' | 'desc' | 'wl' | 'combined'."""
    mols = [Chem.MolFromSmiles(s) for s in smiles]
    parts = []
    if kind in ("morgan", "combined", "combined_wl"):
        parts.append(np.array(_MORGAN.GetFingerprints(mols), dtype=np.float32))
    if kind in ("maccs", "combined", "combined_wl"):
        parts.append(np.array([np.array(MACCSkeys.GenMACCSKeys(m), dtype=np.float32) for m in mols]))
    if kind in ("desc", "combined", "combined_wl"):
        parts.append(np.array([_descriptor_vector(m) for m in mols], dtype=np.float32))
    if kind in ("wl", "combined_wl"):
        parts.append(wl_features(mols))
    return np.hstack(parts) if len(parts) > 1 else parts[0]


# ------------------------------------------------- Weisfeiler-Lehman features
_WL_DIM = 1024


def _hash(label) -> int:
    import hashlib

    return int(hashlib.md5(str(label).encode()).hexdigest()[:8], 16)


def wl_features(mols: list[Chem.Mol], depth: int = 2, dim: int = _WL_DIM) -> np.ndarray:
    """Hashed Weisfeiler-Lehman subtree counts (graph-topology-aware representation)."""
    out = np.zeros((len(mols), dim), dtype=np.float32)
    for r, mol in enumerate(mols):
        labels = [
            (a.GetSymbol(), a.GetDegree(), a.GetFormalCharge(), a.GetIsAromatic())
            for a in mol.GetAtoms()
        ]
        for _ in range(depth + 1):
            for lab in labels:
                out[r, _hash(lab) % dim] += 1.0
            new = []
            for atom in mol.GetAtoms():
                nb = sorted(labels[n.GetIdx()] for n in atom.GetNeighbors())
                new.append((labels[atom.GetIdx()], tuple(nb)))
            labels = [tuple(map(str, lab)) for lab in new]
        out[r] /= max(1.0, mol.GetNumAtoms())
    return out


# ------------------------------------------------------- conformal prediction
def conformal_quantile(residuals: np.ndarray, alpha: float) -> float:
    """Split-conformal half-width for a target miscoverage alpha."""
    n = len(residuals)
    k = int(np.ceil((n + 1) * (1 - alpha)))
    return float(np.sort(np.abs(residuals))[min(k, n) - 1])


def descriptor_frame(smiles: list[str]) -> pd.DataFrame:
    mols = [Chem.MolFromSmiles(s) for s in smiles]
    return pd.DataFrame([_descriptor_vector(m) for m in mols], columns=_DESC_NAMES)


def morgan_fp(smiles: str):
    return _MORGAN.GetFingerprint(Chem.MolFromSmiles(smiles))


def max_tanimoto(query_smiles: list[str], reference_smiles: list[str]) -> np.ndarray:
    ref = [morgan_fp(s) for s in reference_smiles]
    out = []
    for s in query_smiles:
        sims = DataStructs.BulkTanimotoSimilarity(morgan_fp(s), ref)
        out.append(max(sims) if sims else np.nan)
    return np.array(out)


# ---------------------------------------------------------------------- splits
def random_split(y: np.ndarray, frac: float, seed: int) -> tuple[np.ndarray, np.ndarray]:
    """Random train/test split, stratified over pActivity bins."""
    rng = np.random.default_rng(seed)
    bins = pd.qcut(pd.Series(y).rank(method="first"), 5, labels=False)
    test_idx = []
    for b in np.unique(bins):
        idx = np.where(bins == b)[0]
        rng.shuffle(idx)
        test_idx.extend(idx[: max(1, int(round(frac * len(idx))))])
    test_idx = np.array(sorted(test_idx))
    train_idx = np.setdiff1d(np.arange(len(y)), test_idx)
    return train_idx, test_idx


def scaffold_split(
    scaffolds: list[str], frac: float, seed: int, balance: bool = True
) -> tuple[np.ndarray, np.ndarray]:
    """Bemis-Murcko scaffold split; whole scaffolds go to one side of the split."""
    rng = np.random.default_rng(seed)
    groups: dict[str, list[int]] = {}
    for i, sc in enumerate(scaffolds):
        groups.setdefault(sc if sc else "", []).append(i)
    keys = list(groups)
    if balance:
        # large scaffolds first so the test fraction is reached with few groups,
        # but shuffle inside each size bucket so that the seed changes the split
        buckets: dict[int, list[str]] = {}
        for k in keys:
            buckets.setdefault(int(np.floor(np.log2(len(groups[k])))), []).append(k)
        keys = []
        for b in sorted(buckets, reverse=True):
            ks = buckets[b]
            rng.shuffle(ks)
            keys.extend(ks)
    else:
        rng.shuffle(keys)

    n = len(scaffolds)
    test_idx: list[int] = []
    for k in keys:
        if len(test_idx) / n >= frac:
            break
        test_idx.extend(groups[k])
    test_idx = np.array(sorted(test_idx))
    train_idx = np.setdiff1d(np.arange(n), test_idx)
    return train_idx, test_idx


def time_split(years: np.ndarray, cutoff: int) -> tuple[np.ndarray, np.ndarray]:
    train_idx = np.where(years <= cutoff)[0]
    test_idx = np.where(years > cutoff)[0]
    return train_idx, test_idx


# --------------------------------------------------------------------- metrics
def regression_metrics(y_true: np.ndarray, y_pred: np.ndarray) -> dict:
    return {
        "RMSE": float(np.sqrt(mean_squared_error(y_true, y_pred))),
        "MAE": float(mean_absolute_error(y_true, y_pred)),
        "R2": float(r2_score(y_true, y_pred)),
        "Pearson": float(pearsonr(y_true, y_pred)[0]) if len(y_true) > 2 else np.nan,
        "Spearman": float(spearmanr(y_true, y_pred)[0]) if len(y_true) > 2 else np.nan,
    }


def screening_metrics(y_true: np.ndarray, y_pred: np.ndarray, thr: float = 7.0) -> dict:
    """Classification view of a virtual screen: actives are pAct >= thr."""
    mask = (y_true >= thr) | (y_true <= 5.5)
    yt = (y_true[mask] >= thr).astype(int)
    yp = y_pred[mask]
    out = {"n_eval": int(mask.sum()), "n_active": int(yt.sum())}
    if yt.sum() == 0 or yt.sum() == len(yt):
        out.update({"AUROC": np.nan, "EF10": np.nan, "Prec10": np.nan})
        return out
    out["AUROC"] = float(roc_auc_score(yt, yp))
    k = max(1, int(round(0.10 * len(yt))))
    order = np.argsort(-yp)
    hits = yt[order[:k]].sum()
    out["EF10"] = float(hits / (k * yt.mean()))
    out["Prec10"] = float(hits / k)
    return out


def save_json(obj, path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(obj, indent=2, default=float), encoding="utf-8")


# --------------------------------------------------------------------- plotting
def set_style() -> None:
    import matplotlib as mpl

    mpl.rcParams.update(
        {
            "figure.dpi": 120,
            "savefig.dpi": 330,
            "font.family": "DejaVu Sans",
            "font.size": 9,
            "axes.titlesize": 10,
            "axes.labelsize": 9,
            "axes.linewidth": 0.9,
            "axes.spines.top": False,
            "axes.spines.right": False,
            "xtick.direction": "out",
            "ytick.direction": "out",
            "legend.frameon": False,
            "legend.fontsize": 8,
            "figure.facecolor": "white",
            "savefig.facecolor": "white",
            "savefig.bbox": "tight",
        }
    )


PALETTE = {
    "FAP": "#1f6f8b",
    "PREP": "#c0504d",
    "DPP4": "#7f9a3c",
    "grey": "#8a8a8a",
    "accent": "#e0a526",
}


def load_target(name: str) -> pd.DataFrame:
    return pd.read_csv(PROC / f"{name.lower()}_curated.csv")
