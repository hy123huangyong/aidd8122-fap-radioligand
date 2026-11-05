"""
Retrosynthetic disconnection analysis and radiolabelling feasibility for the
six top-ranked generative designs in Table 10 of the AIDD8122 report.

IMPORTANT — scope statement (mirrored in the report):
  This is a *rule-based structural analysis*, not a literature-verified
  synthetic route.  BRICS disconnections and the step-count estimate are
  heuristics computed from the molecular graph; no reaction has been run, no
  yield or availability has been checked experimentally.  The radiolabelling
  columns report calculated descriptors and stoichiometric estimates for a
  chelator conjugate, not measured values.

Usage:  python src/12_design_feasibility.py
Output: results/tables/retrosynthesis.csv
        results/tables/radiolabel_feasibility.csv
"""
from __future__ import annotations

from pathlib import Path

import pandas as pd
from rdkit import Chem
from rdkit.Chem import BRICS, Descriptors, rdMolDescriptors

ROOT = Path(__file__).resolve().parents[1]
TAB = ROOT / "results" / "tables"
TOP = TAB / "generative_top15_ga.csv"

# Structural alerts used for the warhead / handle analysis.
PATTERNS = {
    "alpha_dicarbonyl": Chem.MolFromSmarts("[CX3](=[OX1])[CX3](=[OX1])"),
    "nitrile": Chem.MolFromSmarts("[CX2]#[NX1]"),
    "amide": Chem.MolFromSmarts("[NX3][CX3](=[OX1])"),
    "quinoline": Chem.MolFromSmarts("c1ccc2ncccc2c1"),
    "quinoline_alt": Chem.MolFromSmarts("c1ccnc2ccccc12"),
    "aryl_halide": Chem.MolFromSmarts("[c][F,Cl,Br,I]"),
    "alkyl_F": Chem.MolFromSmarts("[CX4][F]"),
    "carboxylic_acid": Chem.MolFromSmarts("[CX3](=[OX1])[OX2H1]"),
    "free_amine": Chem.MolFromSmarts("[NX3;H1,H2;!$(N[C,S,P]=[O,S])]"),
    "gem_diol": Chem.MolFromSmarts("[CX4]([OX2H1])([OX2H1])"),
    "macrocyclic_chelator": Chem.MolFromSmarts("N1CCNCCNCCNCC1"),
}

# DOTA free acid, used only for a stoichiometric "what if" estimate.
DOTA = Chem.MolFromSmiles("OC(=O)CN1CCN(CC(=O)O)CCN(CC(=O)O)CCN1CC(=O)O")
DOTA_MW = Descriptors.MolWt(DOTA)
DOTA_TPSA = rdMolDescriptors.CalcTPSA(DOTA)
# mono-amide formation loses one H2O from the acid + one H from the amine
CONJ_MW_ADD = DOTA_MW - 18.015
WATER = 18.015


def has(mol, key) -> bool:
    return mol.HasSubstructMatch(PATTERNS[key])


def warhead(mol) -> str:
    if has(mol, "alpha_dicarbonyl"):
        return "alpha-diketone / ketoamide (reversible covalent electrophile)"
    if has(mol, "nitrile"):
        return "prolyl nitrile (clinical FAPI-type warhead)"
    return "unclassified"


def disconnections(mol) -> tuple[int, str]:
    try:
        frags = sorted(BRICS.BRICSDecompose(mol))
    except Exception:                                   # noqa: BLE001
        return 0, "BRICS failed"
    # A BRICS fragment carries [*] attachment points; report the count and the
    # two largest recognisable building blocks by heavy-atom count.
    cleaned = []
    for f in frags:
        try:
            m = Chem.MolFromSmiles(f.replace("[*]", "[H]"))
            cleaned.append((m.GetNumHeavyAtoms() if m else 0, f))
        except Exception:                               # noqa: BLE001
            cleaned.append((0, f))
    cleaned.sort(reverse=True)
    return len(frags), "; ".join(f for _, f in cleaned[:3])


def key_reactions(mol) -> str:
    rx = []
    if has(mol, "amide"):
        rx.append("amide coupling (carboxylic acid + amine)")
    if has(mol, "alpha_dicarbonyl"):
        rx.append("oxidation of alpha-hydroxy ketone to alpha-diketone")
    if has(mol, "nitrile"):
        rx.append("dehydration of primary amide to nitrile")
    if has(mol, "aryl_halide"):
        rx.append("aryl halide installed by electrophilic halogenation")
    if has(mol, "alkyl_F"):
        rx.append("aliphatic F introduced by deoxyfluorination")
    return "; ".join(rx) if rx else "no standard disconnection identified"


def label_handles(mol) -> tuple[str, str]:
    """Return (available handles, chelator conjugation site)."""
    handles = []
    if has(mol, "macrocyclic_chelator"):
        handles.append("macrocyclic chelator already present")
    if has(mol, "free_amine"):
        handles.append("free amine")
    if has(mol, "carboxylic_acid"):
        handles.append("free carboxylic acid")
    if has(mol, "aryl_halide"):
        handles.append("aryl halide (cold F/Cl; needs activated leaving group for 18F SNAr)")
    if has(mol, "alkyl_F"):
        handles.append("aliphatic C-F (stable; needs a displaced leaving group for 18F)")
    if not handles:
        handles.append("none — all heteroatoms are engaged in amides/warhead")
    site = ("free amine / carboxylate available for chelator coupling"
            if has(mol, "free_amine") or has(mol, "carboxylic_acid")
            else "no free conjugation handle; derivatisation required")
    return "; ".join(handles), site


def main() -> None:
    top = pd.read_csv(TOP).head(6).reset_index(drop=True)
    retro_rows, label_rows = [], []

    for i, r in top.iterrows():
        mol = Chem.MolFromSmiles(r.smiles)
        if mol is None:
            raise ValueError(f"invalid SMILES in Table 10 row {i + 1}")
        n_frag, frag_str = disconnections(mol)
        wh = warhead(mol)
        handles, site = label_handles(mol)

        # Heuristic step estimate: one coupling per BRICS junction plus one
        # functional-group interconversion for each specialised warhead.
        steps = max(2, n_frag - 1) + (1 if has(mol, "alpha_dicarbonyl") else 0) \
                + (1 if has(mol, "nitrile") else 0)
        commercial = ("quinoline-4-carboxylic acid, proline-derived amine and a "
                      "short linker are commercial building blocks"
                      if has(mol, "quinoline") or has(mol, "quinoline_alt")
                      else "building-block availability not established")

        retro_rows.append({
            "id": f"T10-{i + 1}",
            "warhead": wh,
            "brics_fragments": n_frag,
            "largest_fragments": frag_str,
            "key_reactions": key_reactions(mol),
            "estimated_steps": steps,
            "starting_materials": commercial,
            "SA_score": round(float(r.sa), 2),
            "route_status": "rule-based disconnection only; no experimental verification",
        })

        mw = Descriptors.MolWt(mol)
        tpsa = rdMolDescriptors.CalcTPSA(mol)
        label_rows.append({
            "id": f"T10-{i + 1}",
            "MW_Da": round(mw, 1),
            "cLogP": round(Descriptors.MolLogP(mol), 2),
            "TPSA": round(tpsa, 1),
            "HBD": rdMolDescriptors.CalcNumHBD(mol),
            "rotatable_bonds": rdMolDescriptors.CalcNumRotatableBonds(mol),
            "warhead": "alpha-diketone" if has(mol, "alpha_dicarbonyl") else
                       ("prolyl nitrile" if has(mol, "nitrile") else "other"),
            "chelator_present": "yes" if has(mol, "macrocyclic_chelator") else "no",
            "available_handles": handles,
            "conjugation_site": site,
            "MW_after_DOTA_est": round(mw + CONJ_MW_ADD, 1),
            "TPSA_after_DOTA_est": round(tpsa + DOTA_TPSA - 2 * 20.2, 1),
            "note": ("68Ga/177Lu require a chelator; 18F requires a displaceable "
                     "leaving group. Neither is present in the generated structure."),
        })

    retro = pd.DataFrame(retro_rows)
    lab = pd.DataFrame(label_rows)
    retro.to_csv(TAB / "retrosynthesis.csv", index=False)
    lab.to_csv(TAB / "radiolabel_feasibility.csv", index=False)

    print(f"DOTA free acid: MW {DOTA_MW:.1f} Da, TPSA {DOTA_TPSA:.1f} A^2; "
          f"mono-amide conjugate adds ~{CONJ_MW_ADD:.1f} Da\n")
    print("=== retrosynthesis ===")
    print(retro[["id", "warhead", "brics_fragments", "estimated_steps", "SA_score"]].to_string(index=False))
    print("\n=== radiolabel feasibility ===")
    print(lab[["id", "MW_Da", "cLogP", "TPSA", "warhead", "chelator_present",
               "MW_after_DOTA_est"]].to_string(index=False))
    print(f"\nchelator present: {(lab.chelator_present == 'yes').sum()}/6")
    print(f"median MW: {lab.MW_Da.median():.1f} -> after DOTA {lab.MW_after_DOTA_est.median():.1f} Da")
    print(f"median TPSA: {lab.TPSA.median():.1f} -> after DOTA {lab.TPSA_after_DOTA_est.median():.1f} A^2")


if __name__ == "__main__":
    main()
