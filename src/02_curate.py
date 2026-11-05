"""
Step 2 -- Curation of the ChEMBL activity records into modelling-ready tables.

Rules (each rule's effect is written to results/tables/curation_log.csv):

  1. keep only records with an exact standard relation ("=")
  2. keep only standard units of nM and a positive standard value
  3. drop records flagged by ChEMBL curation (data_validity_comment)
  4. drop records flagged as potential duplicates
  5. keep biochemical assays only (assay_type B)
  6. convert to pActivity = 9 - log10(value in nM)
  7. standardise the structure with RDKit (cleanup, largest fragment, uncharge)
  8. de-duplicate on InChIKey, keeping the median pActivity of replicates
  9. restrict to 150-1500 Da (drops fragments and undecoded peptide material)

Every compound also gets a three-level activity class used later for the
screening metrics: active (pAct >= 7), inactive (pAct <= 5.5), else intermediate.

Outputs
    data/processed/{fap,prep,dpp4}_curated.csv
    results/tables/curation_log.csv
    results/tables/dataset_summary.csv
"""
from __future__ import annotations

import json
from pathlib import Path

import numpy as np
import pandas as pd
from rdkit import Chem, RDLogger
from rdkit.Chem import Descriptors, rdMolDescriptors
from rdkit.Chem.MolStandardize import rdMolStandardize
from rdkit.Chem.Scaffolds import MurckoScaffold

RDLogger.DisableLog("rdApp.*")

ROOT = Path(__file__).resolve().parents[1]
RAW = ROOT / "data" / "raw"
PROC = ROOT / "data" / "processed"
TAB = ROOT / "results" / "tables"
PROC.mkdir(parents=True, exist_ok=True)
TAB.mkdir(parents=True, exist_ok=True)

TARGETS = {"FAP": "CHEMBL4683", "PREP": "CHEMBL3202", "DPP4": "CHEMBL284"}
LFC = rdMolStandardize.LargestFragmentChooser()
UNCHARGER = rdMolStandardize.Uncharger()


def standardise(smiles: str) -> str | None:
    mol = Chem.MolFromSmiles(smiles)
    if mol is None:
        return None
    try:
        mol = rdMolStandardize.Cleanup(mol)
        mol = LFC.choose(mol)
        mol = UNCHARGER.uncharge(mol)
        Chem.SanitizeMol(mol)
    except Exception:
        return None
    if mol.GetNumHeavyAtoms() < 5:
        return None
    return Chem.MolToSmiles(mol, isomericSmiles=True)


def curate_one(name: str) -> tuple[pd.DataFrame, list[dict]]:
    recs = json.loads((RAW / f"chembl_{name}.json").read_text(encoding="utf-8"))
    log = [{"target": name, "step": "raw records", "n": len(recs)}]
    df = pd.DataFrame(recs)

    df = df[df["standard_relation"] == "="]
    log.append({"target": name, "step": "exact relation only", "n": len(df)})

    df = df[(df["standard_units"] == "nM") & (df["standard_value"].astype(float) > 0)]
    log.append({"target": name, "step": "nM, positive value", "n": len(df)})

    df = df[df["data_validity_comment"].isna()]
    log.append({"target": name, "step": "no ChEMBL validity flag", "n": len(df)})

    df = df[~df["potential_duplicate"].fillna(False).astype(bool)]
    log.append({"target": name, "step": "no potential duplicate", "n": len(df)})

    df = df[df["assay_type"] == "B"]
    log.append({"target": name, "step": "biochemical assay", "n": len(df)})

    df["pchembl"] = 9.0 - np.log10(df["standard_value"].astype(float))
    df = df[(df["pchembl"] >= 3.0) & (df["pchembl"] <= 12.0)]
    log.append({"target": name, "step": "pActivity in [3,12]", "n": len(df)})

    df = df[df["canonical_smiles"].notna() & (df["canonical_smiles"] != "")]
    df["smiles_std"] = [standardise(s) for s in df["canonical_smiles"]]
    df = df[df["smiles_std"].notna()].copy()
    log.append({"target": name, "step": "RDKit compatible", "n": len(df)})

    df["MW"] = [Descriptors.MolWt(Chem.MolFromSmiles(s)) for s in df["smiles_std"]]
    df = df[(df["MW"] >= 150) & (df["MW"] <= 1500)]
    log.append({"target": name, "step": "MW 150-1500 Da", "n": len(df)})

    df["inchikey"] = [Chem.MolToInchiKey(Chem.MolFromSmiles(s)) for s in df["smiles_std"]]

    agg = (
        df.groupby("inchikey")
        .agg(
            smiles=("smiles_std", "first"),
            pchembl=("pchembl", "median"),
            pchembl_sd=("pchembl", "std"),
            n_measure=("pchembl", "size"),
            n_assays=("assay_chembl_id", "nunique"),
            types=("standard_type", lambda s: "+".join(sorted(set(s)))),
            year_max=("document_year", "max"),
            chembl_ids=("molecule_chembl_id", lambda s: ";".join(sorted(set(s)))),
        )
        .reset_index()
    )
    agg["pchembl_sd"] = agg["pchembl_sd"].fillna(0.0)

    mols = [Chem.MolFromSmiles(s) for s in agg["smiles"]]
    agg["MW"] = [Descriptors.MolWt(m) for m in mols]
    agg["clogp"] = [Descriptors.MolLogP(m) for m in mols]
    agg["tpsa"] = [rdMolDescriptors.CalcTPSA(m) for m in mols]
    agg["hbd"] = [rdMolDescriptors.CalcNumHBD(m) for m in mols]
    agg["rotb"] = [rdMolDescriptors.CalcNumRotatableBonds(m) for m in mols]
    agg["rings"] = [rdMolDescriptors.CalcNumRings(m) for m in mols]
    agg["scaffold"] = [
        MurckoScaffold.MurckoScaffoldSmiles(mol=m, includeChirality=False) for m in mols
    ]
    agg["target"] = name
    agg["activity_class"] = np.where(
        agg["pchembl"] >= 7.0, "active", np.where(agg["pchembl"] <= 5.5, "inactive", "intermediate")
    )
    log.append({"target": name, "step": "unique structures", "n": len(agg)})

    sd = agg.loc[agg["n_measure"] > 1, "pchembl_sd"]
    log.append(
        {
            "target": name,
            "step": f"median SD of replicates (n={len(sd)})",
            "n": round(float(sd.median()), 3) if len(sd) else float("nan"),
        }
    )
    return agg, log


def main() -> None:
    all_logs = []
    frames = {}
    for name in TARGETS:
        df, log = curate_one(name)
        df.to_csv(PROC / f"{name.lower()}_curated.csv", index=False)
        frames[name] = df
        all_logs.extend(log)
        print(
            f"{name}: {len(df)} compounds, "
            f"{df.scaffold.replace('', np.nan).nunique()} scaffolds, "
            f"pAct {df.pchembl.min():.2f}-{df.pchembl.max():.2f} (median {df.pchembl.median():.2f})"
        )

    pd.DataFrame(all_logs).to_csv(TAB / "curation_log.csv", index=False)

    rows = []
    for name, df in frames.items():
        rows.append(
            {
                "target": name,
                "records_used": int(df.n_measure.sum()),
                "compounds": len(df),
                # MurckoScaffoldSmiles returns "" for acyclic compounds; an
                # empty string is NOT a scaffold.  Exclude it so this count
                # matches the CSV round-trip (where "" is read back as NaN)
                # and Figure 2c, which reads data/processed/*.csv.
                "scaffolds": int(df.scaffold.replace("", np.nan).nunique()),
                "pAct_mean": round(df.pchembl.mean(), 2),
                "pAct_sd": round(df.pchembl.std(), 2),
                "pAct_min": round(df.pchembl.min(), 2),
                "pAct_max": round(df.pchembl.max(), 2),
                "pct_active": round(100 * (df.activity_class == "active").mean(), 1),
                "median_MW": round(df.MW.median(), 1),
                "median_clogp": round(df.clogp.median(), 2),
                "replicates_pct": round(100 * (df.n_measure > 1).mean(), 1),
            }
        )
    summary = pd.DataFrame(rows)
    summary.to_csv(TAB / "dataset_summary.csv", index=False)
    print(summary.to_string(index=False))

    fap, prep, dpp4 = frames["FAP"], frames["PREP"], frames["DPP4"]
    both = (
        set(fap.inchikey) & set(prep.inchikey),
        set(fap.inchikey) & set(dpp4.inchikey),
        set(prep.inchikey) & set(dpp4.inchikey),
    )
    print(f"overlap FAP/PREP={len(both[0])} FAP/DPP4={len(both[1])} PREP/DPP4={len(both[2])}")


if __name__ == "__main__":
    main()
