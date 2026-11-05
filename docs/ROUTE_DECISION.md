# Route decision record — AIDD8122 Final Project optimisation

**Date:** 5 November 2026 · **Student:** Huang Yong (P2600222)
**Input:** `AIDD8122_R报告优化方案_20261105.md`, decision tree in section 3
**Source checked:** `AIDD8122_Final_Project_Guidelines.pdf` (3 pages, full text read)

## Decision: ROUTE A (P0 + P1, no P2 structural work)

The optimisation plan makes the process brief-driven: route B (docking / structure
prediction / MD) is triggered only when the brief explicitly names structure-based
work. It does not.

### What the brief actually says

| Brief clause | Wording | Implication |
|---|---|---|
| Possible directions | "…such as molecular property prediction, **molecular docking, protein structure prediction**, or other related AIDD tasks." | Listed as *examples of possible directions*, not required deliverables. |
| Project details 2 | "Develop, adapt, fine-tune, or **systematically evaluate** suitable machine learning … methods" | Our three-protocol audit satisfies this directly. |
| Project details 3 | "comparative experiments, or **ablation studies**" | Feature ablation, three split protocols, direct/indirect selectivity, three generative arms. |
| Marking: Methods (10%) | "data, methods, and computational workflow … workflow diagram" | Figure 1 + §4. |
| Marking: Results (20%) | "tables and figures … baselines … strengths and limitations" | Tables 1–10, limitations §5.11, Appendices B/C. |
| Marking: References (5%) | "cited consistently and appropriately … MPU project report format" | 29 references, all with verified DOIs. |
| Presentation | "10-minute presentation, followed by a 5-minute Q&A" | 15-slide deck with 14:50 timed notes. |
| Submission | "final report (PDF) and presentation slides (PDF) … link to an accessible GitHub repository" | **One open item — see below.** |

No clause requires inverse synthesis, retrosynthesis, pose prediction, MD or free
energy. The plan's P1.1 (retrosynthesis) is therefore treated as a low-cost
**coverage** enhancement, not a brief requirement.

### Consequence

- **Executed:** P0.1 evidence bundle; P0.2 argument repairs; P0.3 layout/metadata;
  P0.4 DOIs; P1.1 retrosynthetic triage (Appendix B); P1.2 radioligand feasibility
  (§1.2 + Appendix B); P1.3 paired bootstrap + stratified check (Appendix C);
  P1.4 Mondrian conformal sensitivity (§5.6 + Appendix C).
- **Deliberately not executed (P2):** receptor preparation, redocking, cross-docking,
  short MD. Rationale recorded in §5.11 and the optimisation plan §2.3: a new
  docking score would create a second uncalibrated confidence scale and contradict
  the report's "trust before throughput" argument.
- **Explicitly excluded (P3):** AlphaFold3 full-length modelling, long MD / FEP / TI,
  docking score as a QSAR feature, MM-PBSA validation of Table 10.
- **Preserved unchanged:** ChEMBL release 37 and the three target IDs, the two 2025
  clinical references, Table 10 MW/QED values, Figure 2b sample sizes, and the
  §5.7 single-seed versus five-seed note.

### Honest negative results kept in the report

1. Affinity-only 100% pass rate is a **circular metric**, not superior chemistry.
2. Success criterion C3 is **partially achieved**, not "met".
3. The mixed-assay stratum (n = 30) **reverses** the selectivity cliff; assay-type
   confounding remains unresolved.
4. Mondrian calibration **mitigates but does not repair** out-of-domain coverage;
   its 90% quantile over-covers (1.00) at more than double the interval width.
5. All six top designs lack a chelator and five use a non-clinical warhead.

### One open item before submission

The brief requires a link to an accessible GitHub repository. Appendix A currently
points to the submitted `03_Code/` evidence bundle (no dangling placeholder URL).
Once the repository exists, write its URL to `repo_url.txt` in the project root
(or set `AIDD_REPO_URL`), then rebuild with `python src/09_build_report.py` and
re-export the PDF. No other text needs editing.
