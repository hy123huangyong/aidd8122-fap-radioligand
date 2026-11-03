"""
Step 6 -- Multi-objective generative design of FAP-selective warheads.

A graph-based genetic algorithm (GraphGA-style: fragment crossover + graph
mutation, Jensen 2019) is used to explore the FAPI warhead space under four
objectives that a radiochemist actually cares about:

    f1  predicted FAP affinity            (ensemble of 5 XGBoost models)
    f2  predicted selectivity margin      pFAP - max(pPREP, pDPP4)
    f3  drug-likeness                     QED
    f4  synthetic accessibility           -SA/10   (cheap chemistry is good)

Hard filters: valid structure, 250-650 Da, applicability-domain similarity
>= 0.40 to the FAP training set. Selection is NSGA-II (non-dominated sort +
crowding distance), so the output is a Pareto front rather than a single
"best" molecule.

Three campaigns are compared, all with the same evaluation budget:
    ga            multi-objective (f1-f4)
    affinity_only single-objective (f1), NSGA-II degenerates to a hill climber
    random        random mutation of the same seeds, no selection

Usage: python src/06_generative.py --generations 25 --pop 96
"""
from __future__ import annotations

import argparse
import json
import os
import sys
from dataclasses import dataclass

import numpy as np
import pandas as pd
from rdkit import Chem, RDConfig, RDLogger
from rdkit.Chem import Descriptors, QED, rdMolDescriptors
from rdkit.Chem.Scaffolds import MurckoScaffold
from xgboost import XGBRegressor

sys.path.append(os.path.join(RDConfig.RDContribDir, "SA_Score"))
import sascorer  # noqa: E402

import aidd_utils as U  # noqa: E402

RDLogger.DisableLog("rdApp.*")

FAP_CUT = 7.0
KNOWN_KEYS: set[str] = set()
MW_MIN, MW_MAX = 250.0, 650.0
AD_MIN = 0.40


# --------------------------------------------------------------------- scoring
@dataclass
class Molecule:
    smiles: str
    pFAP: float
    sdFAP: float
    margin: float
    qed: float
    sa: float
    ad: float
    mw: float
    objectives: tuple = ()

    @property
    def key(self) -> str:
        return Chem.MolToInchiKey(Chem.MolFromSmiles(self.smiles))


class Scorer:
    """Ensemble affinity + selectivity + developability scoring."""

    def __init__(self, seeds: list[int] = U.SEEDS):
        self.models: dict[str, list[XGBRegressor]] = {}
        self.train = {}
        for name in U.TARGETS:
            df = U.load_target(name)
            X = U.featurize(df.smiles.tolist(), "combined")
            self.train[name] = (df, X)
            self.models[name] = []
            for s in seeds:
                m = XGBRegressor(
                    n_estimators=700, learning_rate=0.03, max_depth=6, subsample=0.8,
                    colsample_bytree=0.4, reg_lambda=1.0, min_child_weight=3,
                    n_jobs=4, random_state=s, tree_method="hist",
                )
                m.fit(X, df.pchembl.to_numpy())
                self.models[name].append(m)
        self.fap_fps_smiles = self.train["FAP"][0].smiles.tolist()

    def predict(self, smiles: list[str]) -> dict[str, np.ndarray]:
        X = U.featurize(smiles, "combined")
        out = {}
        for name, models in self.models.items():
            preds = np.vstack([m.predict(X) for m in models])
            out[name] = preds.mean(0)
            out[f"{name}_sd"] = preds.std(0)
        return out

    def score(self, smiles: str) -> Molecule | None:
        mol = Chem.MolFromSmiles(smiles)
        if mol is None:
            return None
        mw = Descriptors.MolWt(mol)
        if not (MW_MIN <= mw <= MW_MAX):
            return None
        smi = Chem.MolToSmiles(mol)
        pred = self.predict([smi])
        pFAP = float(pred["FAP"][0])
        margin = float(pred["FAP"][0] - max(pred["PREP"][0], pred["DPP4"][0]))
        ad = float(U.max_tanimoto([smi], self.fap_fps_smiles)[0])
        if ad < AD_MIN:
            return None
        return Molecule(
            smiles=smi, pFAP=pFAP, sdFAP=float(pred["FAP_sd"][0]), margin=margin,
            qed=float(QED.qed(mol)), sa=float(sascorer.calculateScore(mol)), ad=ad, mw=mw,
        )


# ------------------------------------------------------------------- operators
def _join_fragments(fa: Chem.Mol, fb: Chem.Mol) -> Chem.Mol | None:
    """Join two fragments that each carry exactly one dummy atom."""
    da = [a for a in fa.GetAtoms() if a.GetAtomicNum() == 0]
    db = [a for a in fb.GetAtoms() if a.GetAtomicNum() == 0]
    if len(da) != 1 or len(db) != 1:
        return None
    na = da[0].GetNeighbors()[0].GetIdx()
    nb = db[0].GetNeighbors()[0].GetIdx() + fa.GetNumAtoms()
    combo = Chem.RWMol(Chem.CombineMols(fa, fb))
    try:
        combo.AddBond(na, nb, Chem.BondType.SINGLE)
        for idx in sorted([da[0].GetIdx(), db[0].GetIdx() + fa.GetNumAtoms()], reverse=True):
            combo.RemoveAtom(idx)
        mol = combo.GetMol()
        Chem.SanitizeMol(mol)
        return mol
    except Exception:
        return None


def crossover(a: Chem.Mol, b: Chem.Mol, rng: np.random.Generator) -> Chem.Mol | None:
    ba = [x.GetIdx() for x in a.GetBonds() if not x.IsInRing()]
    bb = [x.GetIdx() for x in b.GetBonds() if not x.IsInRing()]
    if not ba or not bb:
        return None
    fa = Chem.GetMolFrags(
        Chem.FragmentOnBonds(a, [int(rng.choice(ba))], addDummies=True), asMols=True
    )
    fb = Chem.GetMolFrags(
        Chem.FragmentOnBonds(b, [int(rng.choice(bb))], addDummies=True), asMols=True
    )
    if len(fa) < 2 or len(fb) < 2:
        return None
    # keep the larger fragment of each parent: keeps drug-like size
    fa = sorted(fa, key=lambda m: m.GetNumHeavyAtoms(), reverse=True)
    fb = sorted(fb, key=lambda m: m.GetNumHeavyAtoms(), reverse=True)
    return _join_fragments(fa[0], fb[0])


def _mutate_atom(mol: Chem.RWMol, rng: np.random.Generator) -> bool:
    atoms = [a for a in mol.GetAtoms() if a.GetAtomicNum() in (6, 7, 8, 16) and not a.IsInRing()]
    if not atoms:
        return False
    atom = atoms[int(rng.integers(len(atoms)))]
    new_z = int(rng.choice([6, 7, 8]))
    if new_z == atom.GetAtomicNum():
        return False
    atom.SetAtomicNum(new_z)
    atom.SetFormalCharge(0)
    atom.SetNoImplicit(False)
    atom.SetNumExplicitHs(0)
    return True


def _mutate_bond(mol: Chem.RWMol, rng: np.random.Generator) -> bool:
    bonds = [b for b in mol.GetBonds()
             if not b.IsInRing() and b.GetBondType() in (Chem.BondType.SINGLE, Chem.BondType.DOUBLE)
             and b.GetBeginAtom().GetAtomicNum() in (6, 7, 8) and b.GetEndAtom().GetAtomicNum() in (6, 7, 8)]
    if not bonds:
        return False
    b = bonds[int(rng.integers(len(bonds)))]
    b.SetBondType(Chem.BondType.DOUBLE if b.GetBondType() == Chem.BondType.SINGLE else Chem.BondType.SINGLE)
    return True


def _mutate_add(mol: Chem.RWMol, rng: np.random.Generator) -> bool:
    cands = [a for a in mol.GetAtoms()
             if a.GetAtomicNum() == 6 and a.GetTotalNumHs() > 0 and a.GetDegree() < 4]
    if not cands:
        return False
    atom = cands[int(rng.integers(len(cands)))]
    z = int(rng.choice([6, 7, 8, 9]))
    new = mol.AddAtom(Chem.Atom(z))
    mol.AddBond(atom.GetIdx(), new, Chem.BondType.SINGLE)
    return True


def _mutate_delete(mol: Chem.RWMol, rng: np.random.Generator) -> bool:
    cands = [a for a in mol.GetAtoms()
             if a.GetDegree() == 1 and not a.IsInRing() and a.GetAtomicNum() != 6]
    if not cands:
        cands = [a for a in mol.GetAtoms()
                 if a.GetDegree() == 1 and not a.IsInRing() and mol.GetNumAtoms() > 12]
    if not cands:
        return False
    mol.RemoveAtom(cands[int(rng.integers(len(cands)))].GetIdx())
    return True


def mutate(mol: Chem.Mol, rng: np.random.Generator) -> Chem.Mol | None:
    rw = Chem.RWMol(mol)
    ops = [_mutate_atom, _mutate_bond, _mutate_add, _mutate_delete]
    rng.shuffle(ops)
    for op in ops:
        try:
            if op(rw, rng):
                cand = rw.GetMol()
                Chem.SanitizeMol(cand)
                if cand.GetNumHeavyAtoms() >= 8:
                    return cand
        except Exception:
            continue
        rw = Chem.RWMol(mol)
    return None


# -------------------------------------------------------------------- NSGA-II
def dominates(a: tuple, b: tuple) -> bool:
    return all(x >= y for x, y in zip(a, b)) and any(x > y for x, y in zip(a, b))


def fast_non_dominated_sort(objs: list[tuple]) -> list[list[int]]:
    fronts: list[list[int]] = [[]]
    S = [[] for _ in objs]
    n = [0] * len(objs)
    rank = [0] * len(objs)
    for p in range(len(objs)):
        for q in range(len(objs)):
            if p == q:
                continue
            if dominates(objs[p], objs[q]):
                S[p].append(q)
            elif dominates(objs[q], objs[p]):
                n[p] += 1
        if n[p] == 0:
            rank[p] = 0
            fronts[0].append(p)
    i = 0
    while fronts[i]:
        nxt = []
        for p in fronts[i]:
            for q in S[p]:
                n[q] -= 1
                if n[q] == 0:
                    rank[q] = i + 1
                    nxt.append(q)
        i += 1
        fronts.append(nxt)
    return fronts[:-1]


def crowding(objs: list[tuple], idx: list[int]) -> dict[int, float]:
    if len(idx) <= 2:
        return {i: np.inf for i in idx}
    dist = {i: 0.0 for i in idx}
    m = len(objs[0])
    for k in range(m):
        order = sorted(idx, key=lambda i: objs[i][k])
        lo, hi = objs[order[0]][k], objs[order[-1]][k]
        dist[order[0]] = dist[order[-1]] = np.inf
        if hi - lo < 1e-12:
            continue
        for j in range(1, len(order) - 1):
            dist[order[j]] += (objs[order[j + 1]][k] - objs[order[j - 1]][k]) / (hi - lo)
    return dist


def nsga2_select(pop: list[Molecule], n: int) -> list[Molecule]:
    objs = [m.objectives for m in pop]
    fronts = fast_non_dominated_sort(objs)
    out: list[Molecule] = []
    for front in fronts:
        if len(out) + len(front) <= n:
            out.extend(pop[i] for i in front)
        else:
            dist = crowding(objs, front)
            take = sorted(front, key=lambda i: dist[i], reverse=True)[: n - len(out)]
            out.extend(pop[i] for i in take)
            break
    return out


# ---------------------------------------------------------------------- driver
def objectives(m: Molecule, mode: str) -> tuple:
    if mode == "affinity_only":
        return (m.pFAP,)
    return (m.pFAP, m.margin, m.qed, -m.sa / 10.0)


def run_campaign(mode: str, generations: int, pop_size: int, scorer: Scorer,
                 seed_smiles: list[str], rng: np.random.Generator) -> tuple[pd.DataFrame, list[Molecule]]:
    pop: list[Molecule] = []
    for s in rng.choice(seed_smiles, size=min(len(seed_smiles), pop_size), replace=False):
        mol = scorer.score(str(s))
        if mol:
            mol.objectives = objectives(mol, mode)
            pop.append(mol)
    archive: list[Molecule] = list(pop)
    traj = []

    for gen in range(generations):
        children: list[Molecule] = []
        n_try = 0
        while len(children) < pop_size and n_try < pop_size * 30:
            n_try += 1
            p1 = pop[int(rng.integers(len(pop)))]
            p2 = pop[int(rng.integers(len(pop)))]
            m1, m2 = Chem.MolFromSmiles(p1.smiles), Chem.MolFromSmiles(p2.smiles)
            if mode != "random" and rng.random() < 0.5:
                cand = crossover(m1, m2, rng)
                if cand is None:
                    continue
                if rng.random() < 0.5:
                    cand = mutate(cand, rng) or cand
            else:
                cand = mutate(m1, rng)
                if cand is None:
                    continue
            sc = scorer.score(Chem.MolToSmiles(cand))
            if sc:
                sc.objectives = objectives(sc, mode)
                children.append(sc)

        if mode == "random":  # no selection pressure: parents are replaced blindly
            pop = children[:pop_size] if len(children) >= pop_size else children + pop[: pop_size - len(children)]
        else:
            pop = nsga2_select(pop + children, pop_size)
        archive = nsga2_select(archive + children, 4 * pop_size)

        best_pfap = max(m.pFAP for m in archive)
        best_margin = max(m.margin for m in archive)
        novel = [m for m in archive if m.key not in KNOWN_KEYS]
        traj.append({"mode": mode, "generation": gen + 1,
                     "best_pFAP": best_pfap, "best_margin": best_margin,
                     "best_novel_pFAP": max([m.pFAP for m in novel], default=np.nan),
                     "best_novel_margin": max([m.margin for m in novel], default=np.nan),
                     "n_novel": len(novel),
                     "n_children": len(children),
                     "median_pFAP": float(np.median([m.pFAP for m in pop]))})
        print(f"  [{mode}] gen {gen + 1:2d}: best pFAP {best_pfap:.2f}, "
              f"best margin {best_margin:.2f}, children {len(children)}")

    return pd.DataFrame(traj), archive


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--generations", type=int, default=25)
    ap.add_argument("--pop", type=int, default=96)
    args = ap.parse_args()

    print("fitting scoring ensembles ...")
    for name in U.TARGETS:
        KNOWN_KEYS.update(U.load_target(name).inchikey.tolist())
    scorer = Scorer()

    fap = U.load_target("FAP")
    seeds = fap.loc[fap.pchembl >= FAP_CUT, "smiles"].tolist()
    print(f"seed set: {len(seeds)} FAP actives (pActivity >= {FAP_CUT})")

    all_traj, all_pareto, summary = [], [], {}
    for i, mode in enumerate(("ga", "affinity_only", "random")):
        rng = np.random.default_rng(100 + i)
        traj, archive = run_campaign(mode, args.generations, args.pop, scorer, seeds, rng)
        all_traj.append(traj)

        df = pd.DataFrame([{
            "smiles": m.smiles, "pFAP": m.pFAP, "sdFAP": m.sdFAP, "margin": m.margin,
            "qed": m.qed, "sa": m.sa, "ad": m.ad, "mw": m.mw, "mode": mode,
        } for m in archive])
        df["scaffold"] = [
            MurckoScaffold.MurckoScaffoldSmiles(mol=Chem.MolFromSmiles(s))
            for s in df.smiles
        ]
        df["novel"] = ~df.smiles.map(lambda x: Chem.MolToInchiKey(Chem.MolFromSmiles(x))).isin(KNOWN_KEYS)
        all_pareto.append(df)
        nov = df[df.novel]
        nov.sort_values(["pFAP", "margin"], ascending=False).head(15).to_csv(
            U.TAB / f"generative_top15_{mode}.csv", index=False)

        summary[mode] = {
            "n_archive": len(df),
            "n_novel": int(df.novel.sum()),
            "best_novel_pFAP": float(nov.pFAP.max()) if len(nov) else None,
            "best_novel_margin": float(nov.margin.max()) if len(nov) else None,
            "median_novel_pFAP": float(nov.pFAP.median()) if len(nov) else None,
            "median_novel_margin": float(nov.margin.median()) if len(nov) else None,
            "frac_novel_pFAP8_margin1": float(((nov.pFAP >= 8) & (nov.margin >= 1)).mean())
            if len(nov) else None,
            "n_unique_scaffolds": int(df.scaffold.nunique()),
            "best_pFAP": float(df.pFAP.max()),
            "best_margin": float(df.margin.max()),
            "best_margin_at_pFAP8": float(df.loc[df.pFAP >= 8, "margin"].max())
            if (df.pFAP >= 8).any() else None,
            "median_qed": float(df.qed.median()),
            "median_sa": float(df.sa.median()),
        }

    traj = pd.concat(all_traj, ignore_index=True)
    traj.to_csv(U.TAB / "generative_trajectory.csv", index=False)
    pareto = pd.concat(all_pareto, ignore_index=True)
    pareto.to_csv(U.TAB / "generative_archive.csv", index=False)
    U.save_json(summary, U.TAB / "generative_summary.json")
    print(json.dumps(summary, indent=2))


if __name__ == "__main__":
    main()
