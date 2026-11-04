"""
Build the AIDD8122 final report (.docx) from the project template.

The template supplies the cover page, page geometry and heading styles; this
script fills the cover, inserts a real TOC field, writes the body, embeds the
figures and tables, and leaves the document ready for Word to update fields and
export to PDF.

Usage: python src/09_build_report.py
"""
from __future__ import annotations

import json
from pathlib import Path

import pandas as pd
from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Inches, Pt, RGBColor

ROOT = Path(__file__).resolve().parents[1]
TAB = ROOT / "results" / "tables"
FIG = ROOT / "results" / "figures"
OUT = ROOT / "report" / "AIDD8122_Final_Report_HuangYong.docx"
TEMPLATE = Path("E:/AIDD/8122最终作业/AIDD8122_Final_Project_Report_Template.docx")

TITLE = ("Trust Before Throughput: Benchmarking, Uncertainty-Aware Screening and "
         "Multi-Objective Generative Design of FAP-Selective Radioligand Warheads")
STUDENT = "Huang Yong"
SID = "P2600222"
DATE = "16/11/2026"

BODY_FONT = "Times New Roman"


# --------------------------------------------------------------------- helpers
def set_run(run, size=12, bold=False, italic=False, color=None):
    run.font.name = BODY_FONT
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic
    if color:
        run.font.color.rgb = RGBColor(*color)
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rfonts.set(qn(attr), BODY_FONT)


def para(doc, text="", size=12, bold=False, italic=False, align="justify",
         space_after=6, space_before=0, line=1.5, indent=None, keep=False):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.space_after = Pt(space_after)
    pf.space_before = Pt(space_before)
    pf.line_spacing = line
    if indent is not None:
        pf.first_line_indent = Cm(indent)
    pf.alignment = {"justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
                    "left": WD_ALIGN_PARAGRAPH.LEFT,
                    "center": WD_ALIGN_PARAGRAPH.CENTER,
                    "right": WD_ALIGN_PARAGRAPH.RIGHT}[align]
    pf.keep_with_next = keep
    if text:
        # allow **bold** fragments
        parts = text.split("**")
        for i, chunk in enumerate(parts):
            if not chunk:
                continue
            r = p.add_run(chunk)
            set_run(r, size=size, bold=bold or (i % 2 == 1), italic=italic)
    return p


def heading(doc, text, level=1, page_break=False):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.keep_with_next = True
    pf.page_break_before = page_break
    pf.space_before = Pt(14 if level == 1 else 10)
    pf.space_after = Pt(6 if level == 1 else 4)
    pf.line_spacing = 1.2
    r = p.add_run(text)
    set_run(r, size=14 if level == 1 else 12, bold=True)
    # outline level so the TOC field picks it up
    ppr = p._p.get_or_add_pPr()
    ol = OxmlElement("w:outlineLvl")
    ol.set(qn("w:val"), str(level - 1))
    ppr.append(ol)
    if level == 1:
        p.paragraph_format.page_break_before = False
    return p


def figure(doc, path, caption, width=6.2):
    p = doc.add_paragraph()
    p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.keep_with_next = True
    p.add_run().add_picture(str(path), width=Inches(width))
    c = doc.add_paragraph()
    c.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    c.paragraph_format.space_after = Pt(10)
    c.paragraph_format.line_spacing = 1.0
    parts = caption.split("**")
    for i, chunk in enumerate(parts):
        if not chunk:
            continue
        r = c.add_run(chunk)
        set_run(r, size=9.5, bold=(i % 2 == 1))


def _shade(cell, hex_fill):
    tcpr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_fill)
    tcpr.append(shd)


def bullet(doc, text, size=11.5):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.space_after = Pt(3)
    pf.line_spacing = 1.3
    pf.left_indent = Cm(0.9)
    pf.first_line_indent = Cm(-0.45)
    pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    r = p.add_run("\u2022\u2003")
    set_run(r, size=size)
    parts = text.split("**")
    for i, chunk in enumerate(parts):
        if not chunk:
            continue
        rr = p.add_run(chunk)
        set_run(rr, size=size, bold=(i % 2 == 1))
    return p


def table(doc, header, rows, caption=None, widths=None, size=9, align=None, fixed=False):
    if caption:
        c = doc.add_paragraph()
        c.paragraph_format.space_before = Pt(10)
        c.paragraph_format.space_after = Pt(3)
        c.paragraph_format.line_spacing = 1.0
        c.paragraph_format.keep_with_next = True
        parts = caption.split("**")
        for i, chunk in enumerate(parts):
            if not chunk:
                continue
            r = c.add_run(chunk)
            set_run(r, size=9.5, bold=(i % 2 == 1))

    t = doc.add_table(rows=1, cols=len(header))
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr = t.rows[0].cells
    for i, h in enumerate(header):
        hdr[i].text = ""
        p = hdr[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(1)
        p.paragraph_format.space_before = Pt(1)
        r = p.add_run(str(h))
        set_run(r, size=size, bold=True)
        _shade(hdr[i], "E9EFF3")
    # repeat header row across pages
    trpr = t.rows[0]._tr.get_or_add_trPr()
    th = OxmlElement("w:tblHeader")
    th.set(qn("w:val"), "true")
    trpr.append(th)

    for row in rows:
        cells = t.add_row().cells
        for i, v in enumerate(row):
            cells[i].text = ""
            p = cells[i].paragraphs[0]
            p.paragraph_format.space_after = Pt(1)
            p.paragraph_format.space_before = Pt(1)
            p.alignment = (WD_ALIGN_PARAGRAPH.LEFT if i == 0
                           else (align[i] if align else WD_ALIGN_PARAGRAPH.CENTER))
            r = p.add_run(str(v))
            set_run(r, size=size)

    if widths:
        if fixed:
            t.autofit = False
            layout = OxmlElement("w:tblLayout")
            layout.set(qn("w:type"), "fixed")
            t._tbl.tblPr.append(layout)
            for i, w in enumerate(widths):
                t.columns[i].width = Cm(w)
        for row in t.rows:
            for i, w in enumerate(widths):
                row.cells[i].width = Cm(w)
    # spacing paragraph after the table
    sp = doc.add_paragraph()
    sp.paragraph_format.space_after = Pt(4)
    sp.paragraph_format.space_before = Pt(0)
    return t


def toc_field(doc):
    p = doc.add_paragraph()
    r = p.add_run()
    fld_begin = OxmlElement("w:fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = r' TOC \o "1-2" \h \z \u '
    fld_sep = OxmlElement("w:fldChar")
    fld_sep.set(qn("w:fldCharType"), "separate")
    placeholder = OxmlElement("w:t")
    placeholder.text = "Right-click and choose Update Field to build the table of contents."
    fld_end = OxmlElement("w:fldChar")
    fld_end.set(qn("w:fldCharType"), "end")
    r._r.append(fld_begin)
    r._r.append(instr)
    r._r.append(fld_sep)
    r._r.append(placeholder)
    r._r.append(fld_end)
    return p


# ------------------------------------------------------------------ data loads
def load_numbers() -> dict:
    d = {}
    ds = pd.read_csv(TAB / "dataset_summary.csv")
    d["dataset"] = ds
    d["bench"] = pd.read_csv(TAB / "benchmark_FAP_combined_runs.csv")
    d["ablation"] = pd.read_csv(TAB / "ablation_features_FAP.csv")
    d["transfer"] = pd.read_csv(TAB / "transfer_runs.csv")
    d["sel"] = pd.read_csv(TAB / "selectivity_cv.csv")
    d["sel_ci"] = pd.read_csv(TAB / "selectivity_ci.csv")
    d["sel_over"] = pd.read_csv(TAB / "selectivity_overview.csv")
    d["ad"] = pd.read_csv(TAB / "ad_bins.csv")
    d["conf"] = pd.read_csv(TAB / "conformal.csv")
    d["temporal"] = pd.read_csv(TAB / "temporal.csv")
    d["curation"] = pd.read_csv(TAB / "curation_log.csv")
    gen = TAB / "generative_summary.json"
    d["gen"] = json.loads(gen.read_text(encoding="utf-8")) if gen.exists() else {}
    top = TAB / "generative_top15_ga.csv"
    d["gen_top"] = pd.read_csv(top) if top.exists() else pd.DataFrame()
    arch = TAB / "generative_archive.csv"
    d["gen_arc"] = pd.read_csv(arch) if arch.exists() else pd.DataFrame()
    return d


def fmt(x, n=2, sign=False):
    if pd.isna(x):
        return "--"
    s = f"{x:+.{n}f}" if sign else f"{x:.{n}f}"
    return s


# ------------------------------------------------------------------------ body
def build_body(doc, D):
    ds = D["dataset"].set_index("target")

    # ---------------------------------------------------------------- abstract
    heading(doc, "Abstract", 1, page_break=True)
    para(doc,
         "Fibroblast activation protein (FAP) is one of the most actively pursued targets in "
         "nuclear medicine, yet the public structure–activity record available to guide the next "
         "generation of FAP inhibitors is small, assay-heterogeneous and chemically non-stationary. "
         "This project asks a deliberately narrow question: given that record, what can machine "
         "learning actually deliver for radiopharmaceutical design, and where should a medicinal "
         "chemist refuse to trust it? We curated 972 ChEMBL bioactivities into 384 unique FAP "
         "inhibitors (162 Murcko scaffolds) and benchmarked five models under random, scaffold and "
         "time splits. Random splitting suggests excellent accuracy (XGBoost R² = 0.75), but a time "
         "split that mimics prospective use collapses to R² = −0.79, because the post-2019 "
         "literature describes a different chemical series (median molecular weight +80 Da, median "
         "potency +2.0 log units). Selectivity against the paralogues PREP and DPP4 is predictable "
         "inside known scaffolds (AUROC 0.84) but anti-correlated on novel scaffolds (AUROC 0.38). "
         "Split-conformal intervals are valid under random splitting (79% and 89% empirical "
         "coverage at nominal 80% and 90%) but cover only 54% of out-of-domain compounds at the 80% "
         "level. Finally, a graph genetic algorithm with NSGA-II selection was used to design novel "
         "warheads under four objectives; gating candidates by applicability domain and by predicted "
         "selectivity changed the output more than the choice of optimizer. We argue that for "
         "targets of this size the useful deliverable of AI-driven drug discovery is not a single "
         "predicted number but a calibrated decision rule that says when a prediction may be acted on.")
    para(doc, "**Keywords:** fibroblast activation protein; radiopharmaceuticals; bioactivity "
              "prediction; scaffold split; conformal prediction; multi-objective optimization",
         space_after=4)

    # ------------------------------------------------------------ introduction
    heading(doc, "§1. Introduction", 1)
    heading(doc, "1.1 A stromal target that reached the clinic in five years", 2)
    para(doc,
         "In desmoplastic tumours such as pancreatic ductal adenocarcinoma, breast carcinoma and "
         "several sarcomas, cancer-associated fibroblasts (CAFs) can constitute the majority of the "
         "tumour mass and deposit the collagen-rich stroma that limits drug penetration and "
         "correlates with poor prognosis. Fibroblast activation protein (FAP) is a type II "
         "transmembrane serine protease of the S9B prolyl-oligopeptidase family that is displayed on "
         "the surface of activated fibroblasts and is largely absent from healthy tissues [9], which "
         "makes it an unusually clean stromal address. Because its catalytic site accepts small, "
         "drug-like inhibitors, FAP became the first stromal target for which a purely small-molecule "
         "theranostic pair was realised: a quinoline-based inhibitor bearing a prolyl warhead can be "
         "labelled with 68Ga or 18F for imaging and with 90Y or 177Lu for therapy, with only the "
         "chelator changed [1,6].")
    para(doc,
         "The clinical trajectory has been steep. FAPI-04, optimised from a series of quinoline "
         "inhibitors in 2018 [1], produced high-contrast PET images in a broad panel of tumour entities [3,4,10], "
         "and its successors FAPI-46 and FAPI-74 extended the usable activity and improved tumour "
         "retention [5,7]. Several hundred patients have since been imaged, and therapeutic analogues "
         "such as 90Y-FAPI-46, 177Lu-FAP-2286 and the albumin-binding 177Lu-LNC1004 have entered "
         "early-phase trials [6–8]. For a nuclear medicine group, this is now a ligand-design problem rather "
         "than a target-validation problem.")
    heading(doc, "1.2 What limits the current generation of ligands", 2)
    para(doc,
         "Imaging performance has not translated proportionally into therapeutic benefit. In the "
         "published FAPI radioligand-therapy series, disease control is reported in roughly 20–80% of "
         "patients while objective responses remain uncommon (about 21% in the largest phase II cohort "
         "to date), with haematological toxicity as the practical constraint [7,8,29]. Retention in tumour is shorter than for prostate-specific membrane antigen "
         "(PSMA) ligands, and uptake in normal fibroblasts of the salivary glands, pancreas and "
         "marrow sets a background that cannot be removed by dosimetry alone. The medicinal-chemistry "
         "levers are well defined — the warhead (nitrile versus boronic acid), linker length and "
         "rigidity, the chelator, and for albumin-binding variants an Evans blue motif — but each "
         "combination must be synthesised, radiolabelled, tested against the enzyme, and imaged in "
         "mice. A single design–test cycle costs weeks of chemistry and a full animal experiment.")
    para(doc,
         "Two selectivity problems dominate this optimisation. The first is selectivity against the "
         "closest paralogues, prolyl endopeptidase (PREP) and dipeptidyl peptidase IV (DPP4/CD26), "
         "which share the catalytic fold and are widely expressed; a ligand that inhibits DPP4 will "
         "carry a systemic off-target liability that no dosimetry can compensate. The second is "
         "selectivity against FAP expressed in healthy tissue, which is a pharmacokinetic rather than "
         "an affinity problem and is not addressed here. This project concerns the first.")
    heading(doc, "1.3 Why machine learning, and why the usual evaluation is not enough", 2)
    para(doc,
         "Public bioactivity data for FAP are an order of magnitude smaller than for kinase targets: "
         "ChEMBL contains a few hundred unique FAP inhibitors, contributed by sixty-odd publications, "
         "assayed across 78 distinct assays [11,28]. Published QSAR models of FAP "
         "inhibitors typically report random-split coefficients of determination of 0.7–0.8 [12,14], which "
         "invites the conclusion that affinity prediction is solved. For a chemist the relevant "
         "question is different: will the model rank a new chemotype correctly, and will its error "
         "bars hold when the literature moves on? Random splitting answers neither question, because "
         "near-duplicate analogues of the test compounds remain in the training set.")
    para(doc,
         "This project therefore treats the evaluation protocol as the object of study rather than "
         "as a formality. We benchmark the same models under three protocols that correspond to three "
         "real decisions — interpolation inside a known series, scaffold hopping to a new series, and "
         "prospective use on compounds published after the model was trained — and we ask under which "
         "of these the output is actionable.")
    heading(doc, "1.4 Contributions", 2)
    for b in [
        "A reproducible curation of the public FAP, PREP and DPP4 bioactivity record, with an "
        "explicit audit trail from 972 raw FAP records to 384 unique compounds, and a quantification "
        "of assay noise that sets the error floor of any model trained on it.",
        "A three-protocol benchmark (random, scaffold, time) of five model families with identical "
        "features and hyperparameters, reported together with virtual-screening metrics rather than "
        "regression metrics alone.",
        "A quantitative test of whether paralogue bioactivity transfers to FAP, and of whether FAP "
        "selectivity over PREP and DPP4 can be predicted at all outside the training scaffolds.",
        "A trust analysis that combines applicability domain, split-conformal intervals and temporal "
        "recalibration, and localises the conditions under which a predicted interval is honest.",
        "A multi-objective generative design campaign (graph genetic algorithm with NSGA-II "
        "selection) whose output is compared against single-objective and random controls, with "
        "novelty and applicability explicitly tracked.",
    ]:
        bullet(doc, b)

    # -------------------------------------------------------- literature review
    heading(doc, "§2. Literature Review", 1)
    heading(doc, "2.1 The FAP ligand series", 2)
    para(doc,
         "The modern FAP inhibitor series originates from quinoline-2-carboxamides decorated with a "
         "prolyl nitrile warhead, a linker and a chelating group. Lindner and colleagues optimised "
         "fifteen candidates and identified FAPI-04 as the first clinical tracer [1], with improved "
         "serum stability and a clear preference for FAP over CD26 relative to its predecessor "
         "FAPI-02. Loktev and colleagues independently established that such inhibitors can image "
         "CAF-rich tumours with high contrast [2], and Kratochwil and colleagues documented uptake across "
         "28 tumour entities [3], which established FAP as a pan-tumour stromal target. FAPI-46 improved "
         "tumour retention and reduced washout [7], and FAPI-74 allowed 18F-AlF labelling through a "
         "one-step aluminium fluoride method [5], which decouples the tracer from a 68Ga generator.")
    para(doc,
         "The therapeutic branch followed quickly. 90Y-FAPI-46 was tested in advanced sarcoma [7], "
         "177Lu-FAP-2286 became the first DOTA-bearing FAP ligand taken into a first-in-human "
         "dosimetry study [6], and 177Lu-LNC1004 introduced an Evans blue albumin-binding motif to extend "
         "circulation time [8]. Across these series the pattern is consistent: imaging contrast is "
         "excellent, objective response rates are low, and the field attributes the gap to short "
         "intratumoural retention and to the stromal, non-malignant nature of the target. Whatever "
         "the biological explanation, the practical consequence for design is that affinity alone is "
         "an insufficient objective, and selectivity against the paralogues must be designed in "
         "rather than measured afterwards.")
    heading(doc, "2.2 Machine learning for target-specific potency", 2)
    para(doc,
         "On ChEMBL-scale benchmarks, gradient-boosted trees over circular fingerprints remain "
         "competitive with, and frequently superior to, graph neural networks, particularly when the "
         "number of measurements per target is in the hundreds [26]. Mayr and colleagues showed on 1,300 "
         "ChEMBL targets that deep architectures only overtake classical methods when data are "
         "abundant [12], and van Tilborg and colleagues demonstrated that much of the apparent advantage "
         "of learned representations disappears once activity cliffs and scaffold novelty are "
         "accounted for [13]. For a target like FAP, with fewer than 500 unique compounds, a well-tuned "
         "tree ensemble over combined fingerprints and physicochemical descriptors is therefore the "
         "rational baseline, and any deep model must justify itself against it.")
    para(doc,
         "A second line of work uses transfer learning across related targets. Because PREP and DPP4 "
         "belong to the same protease family and contribute 1.3-fold and 11-fold more compounds than "
         "FAP, their data are an obvious auxiliary source. Reported gains from such transfer are "
         "typically small and highly dependent on whether the auxiliary task shares chemotypes with "
         "the target task; when the auxiliary data are dominated by a different chemical series, "
         "transfer can be neutral or harmful.")
    heading(doc, "2.3 Evaluation protocols: why random splits flatter", 2)
    para(doc,
         "The gap between random and scaffold splits is one of the best-documented reproducibility "
         "problems in molecular machine learning. MoleculeNet popularised scaffold splitting precisely "
         "because random splits allow near-analogues of test compounds into training [14]. Wallach and "
         "Heifets showed that many ligand-based benchmarks reward memorisation rather than "
         "generalisation [15], and Sheridan argued for time-split cross-validation as the closest "
         "available proxy for prospective prediction [16]. For a target whose literature is dominated by "
         "one chemotype introduced at a known date, as is the case for FAP after 2018, the time split "
         "is not a technicality: it is the only protocol that reproduces the situation in which a "
         "medicinal chemist would actually use the model.")
    heading(doc, "2.4 Uncertainty and the applicability domain", 2)
    para(doc,
         "Point predictions without an error estimate are of limited use in a synthesis meeting. Two "
         "complementary tools exist. Applicability-domain analysis asks whether a query compound is "
         "represented by the training set, typically through a similarity threshold or a bounding "
         "box in descriptor space, and has been reviewed extensively for QSAR models [17]. Conformal "
         "prediction instead constructs a distribution-free interval around each prediction from "
         "held-out residuals, with a finite-sample coverage guarantee [18] that holds under exchangeability "
         "of calibration and test data. The exchangeability assumption is exactly what scaffold and "
         "temporal splits violate, which makes conformal prediction an informative diagnostic rather "
         "than a universal remedy: if coverage collapses, the data are not exchangeable and the "
         "interval should not be quoted.")
    heading(doc, "2.5 Generative molecular design", 2)
    para(doc,
         "Generative design for small molecules has converged on two families: graph-based genetic "
         "algorithms operating directly on molecular graphs through fragment crossover and graph "
         "mutation, and deep generative models trained on large corpora [19]. Jensen's graph GA remains a "
         "strong baseline because it requires no training data and optimises arbitrary scoring "
         "functions. When several properties must be balanced — affinity, selectivity, "
         "drug-likeness, synthetic accessibility — multi-objective evolutionary algorithms such as "
         "NSGA-II [20] are the natural choice, because they return a Pareto front instead of a single "
         "scalarised optimum and thereby expose the trade-offs to the chemist. Their weakness in a "
         "drug-discovery setting is that the scoring function is itself a model, so the optimizer "
         "will exploit its errors unless novelty and applicability are enforced explicitly.")
    heading(doc, "2.6 Positioning of this work", 2)
    para(doc,
         "The literature supplies mature tools but few honest audits of those tools on small "
         "radiopharmaceutical targets. This project does not propose a new architecture. It asks "
         "whether the standard toolkit, applied carefully to FAP, produces anything a radiochemist "
         "should act on, and it reports the cases where the answer is no.")

    # ------------------------------------------------------------- research qs
    heading(doc, "§3. Research Questions and Objectives", 1)
    heading(doc, "3.1 Research questions", 2)
    for q, body in [
        ("RQ1.", "How accurate is a FAP affinity model under protocols that reflect real use — "
                 "interpolation within a known series, scaffold hopping to a new series, and "
                 "prospective prediction on later literature?"),
        ("RQ2.", "Does bioactivity from the paralogue family (PREP, DPP4) transfer to FAP, and can "
                 "FAP selectivity over those paralogues be predicted outside the training scaffolds?"),
        ("RQ3.", "Under which conditions are the model's uncertainty estimates trustworthy, and can "
                 "an applicability-domain gate restore calibration?"),
        ("RQ4.", "What does a multi-objective generative algorithm add beyond enumerating known FAP "
                 "inhibitors, and does explicit selectivity pressure change the chemical output?"),
    ]:
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(5)
        p.paragraph_format.line_spacing = 1.4
        p.paragraph_format.left_indent = Cm(0.75)
        p.paragraph_format.first_line_indent = Cm(-0.75)
        r = p.add_run(q + "  ")
        set_run(r, size=12, bold=True)
        r2 = p.add_run(body, )
        set_run(r2, size=12)

    heading(doc, "3.2 Objectives", 2)
    para(doc,
         "Four objectives follow directly. O1: build and document a curated, deduplicated dataset for "
         "FAP, PREP and DPP4 with an auditable filtering log and an estimate of assay noise. O2: "
         "benchmark representative model families under the three protocols and report both "
         "regression and virtual-screening metrics. O3: quantify the value of paralogue transfer and "
         "of direct versus indirect selectivity modelling. O4: determine the operating region in "
         "which predictions and intervals can be trusted, and use that region as a hard constraint in "
         "a generative design campaign.")
    heading(doc, "3.3 Scope and success criteria", 2)
    para(doc,
         "The project is deliberately restricted to ligand-based modelling of the FAP warhead region; "
         "it does not model the chelator, the radiometal, or in vivo pharmacokinetics, and it does not "
         "perform docking or free-energy calculations. Success was defined in advance as: (i) a "
         "scaffold-split R² above 0.5 with a screening AUROC above 0.85 for the best model, which "
         "would justify use as a triage filter; (ii) a measurable and statistically defensible "
         "difference between modelling protocols; and (iii) conformal intervals whose empirical "
         "coverage can be predicted from the applicability domain. All three were met, although the "
         "third only conditionally.")

    # ------------------------------------------------------------------ methods
    heading(doc, "§4. Methods and Materials", 1)
    para(doc,
         "The workflow has four stages — data, models, trust, design — and each stage produces a table "
         "that the next stage consumes (Figure 1). Every experiment was repeated over five random "
         "seeds with identical feature blocks and hyperparameters, and all figures and tables in this "
         "report are regenerated by the accompanying code.")
    figure(doc, FIG / "fig0_workflow.png",
           "**Figure 1.** Project workflow. Bioactivity records are curated into per-target tables "
           "(A), represented by combined fingerprints and descriptors and modelled under three split "
           "protocols (B), audited with applicability-domain, conformal and temporal analyses (C), and "
           "finally used as scoring functions inside a multi-objective graph genetic algorithm (D). "
           "The dashed arrow indicates that error analysis feeds back into feature and model choice.")

    heading(doc, "4.1 Data acquisition", 2)
    para(doc,
         "Bioactivity records were retrieved programmatically from the ChEMBL REST API (release 37, May 2026) "
         "for three targets: FAP (CHEMBL4683), PREP (CHEMBL3202) and DPP4 (CHEMBL284). Only records "
         "with an exact activity relation and a standard type of Ki, IC50 or Kd were requested, with "
         "standard units in nanomolar, retaining ChEMBL's own validity annotations [28]. For FAP this "
         "returned 972 records, for PREP 962 and for DPP4 "
         "6,553. Records were retrieved together with assay type, assay description, document year, "
         "validity flags and the canonical SMILES of the tested molecule.")

    heading(doc, "4.2 Curation", 2)
    para(doc,
         "Curation followed a fixed sequence, and the number of records surviving each step is logged "
         "(Table 1). Records were dropped if they carried a ChEMBL data-validity comment, were flagged "
         "as potential duplicates, came from a non-biochemical assay (cell-based or ADME assays), fell "
         "outside pActivity 3–12, failed RDKit parsing, or lay outside 150–1,500 Da. Multiple "
         "measurements of the same InChIKey were collapsed to their median pActivity, and the number "
         "of contributing measurements and the standard deviation were retained. Murcko scaffolds [25] were "
         "computed for scaffold grouping and splitting. The FAP table is the modelling target; PREP "
         "and DPP4 serve as auxiliary and off-target references.")
    para(doc,
         "Assay noise was estimated directly from the replicate structure of the data: for every "
         "compound measured more than once, the standard deviation of its replicate pActivity values "
         "was computed, and the median across compounds is reported. This quantity sets a floor on "
         "the root-mean-square error that any model trained on these labels can achieve, and it is "
         "used throughout the discussion as the reference against which model error is judged.")

    heading(doc, "4.3 Chemical representation", 2)
    para(doc,
         "Each molecule is represented by a concatenation of three blocks: 2,048-bit Morgan [26] "
         "fingerprints with radius 2, 167-bit MACCS keys, and 15 physicochemical descriptors "
         "(molecular weight, cLogP, TPSA, hydrogen-bond donors and acceptors, rotatable bonds, ring "
         "and aromatic-ring counts, fraction of sp3 carbons, heavy-atom count, and counts of nitrile, "
         "amide, carboxylic acid, N+O and halogen groups). The descriptor block was included because "
         "the FAP warhead is defined by functional groups rather than by scaffold shape, and because "
         "descriptors extrapolate more gracefully than fingerprints when the scaffold changes.")

    heading(doc, "4.4 Models and hyperparameters", 2)
    para(doc,
         "Five estimators were compared: a mean predictor (the trivial baseline), ridge regression, "
         "support-vector regression with an RBF kernel, a random forest [24] with 500 trees, and gradient "
         "boosting (XGBoost [23]) with 700 trees, learning rate 0.03, maximum depth 6, subsample 0.8, "
         "column subsample 0.4, L2 penalty 1.0 and minimum child weight 3. Hyperparameters were fixed "
         "a priori rather than tuned per split, because tuning on the test protocol would erase the "
         "very difference the benchmark is designed to measure. All models were fitted on the same "
         "feature matrix and the same training indices.")

    heading(doc, "4.5 Evaluation protocols", 2)
    para(doc,
         "Three protocols were used. Random splitting holds out 20% of compounds uniformly, which "
         "measures interpolation and is the protocol reported by most published QSAR studies. "
         "Scaffold splitting groups compounds by Murcko scaffold and holds out 20% of scaffolds, "
         "which measures the ability to rank a new chemical series and is the relevant protocol for "
         "scaffold hopping. Time splitting trains on compounds whose most recent source document is "
         "from 2018 or earlier and tests on 2019–2024, which measures prospective performance and "
         "reproduces the situation of a model trained before the FAPI series matured. Each protocol "
         "was repeated over five seeds; scaffold and random splits differ between seeds, whereas the "
         "time split is deterministic.")

    heading(doc, "4.6 Regression and screening metrics", 2)
    para(doc,
         "Regression quality is reported as RMSE, MAE, R², Pearson and Spearman correlation. Because "
         "the practical use of such a model is triage rather than quantitation, screening metrics were "
         "computed on a focused evaluation set containing only clear actives (pActivity ≥ 7, i.e. "
         "IC50 ≤ 100 nM) and clear inactives (pActivity ≤ 5.5), excluding the ambiguous middle. On "
         "this set we report the area under the ROC curve (AUROC), the enrichment factor in the top "
         "10% of ranked compounds (EF10) and the precision in the top 10%.")

    heading(doc, "4.7 Applicability domain and conformal prediction", 2)
    para(doc,
         "The applicability domain is defined by chemical similarity: for each test compound, the "
         "maximum Tanimoto similarity to any training compound is computed over Morgan fingerprints. "
         "Test compounds are then binned by this similarity and the RMSE of each bin is reported, "
         "which yields an operational rule (a similarity threshold) rather than a binary claim.")
    para(doc,
         "Uncertainty is quantified by split conformal prediction. Out-of-fold residuals are obtained "
         "by five-fold cross-validation inside the training set, and the conformal half-width is the "
         "⌈(n+1)(1−α)⌉-th smallest absolute residual. Intervals are then evaluated on the held-out "
         "compounds at nominal 80% and 90% levels, and coverage is additionally stratified by "
         "applicability-domain class. A random-split experiment serves as a positive control: if the "
         "procedure is implemented correctly and the data are exchangeable, coverage should match the "
         "nominal level there.")

    heading(doc, "4.8 Temporal drift and recalibration", 2)
    para(doc,
         "For the time split, the raw model is compared with a recalibrated version in which a single "
         "intercept correction is estimated on the first quarter of the test set — simulating a small "
         "prospective calibration campaign — and applied to the remainder. This separates two failure "
         "modes: a shift in the overall potency level, which an intercept can absorb, from a loss of "
         "ranking information, which it cannot.")

    heading(doc, "4.9 Multi-objective generative design", 2)
    para(doc,
         "Design used a graph genetic algorithm in the style of Jensen [19]: parents are recombined by "
         "breaking one acyclic bond in each parent and joining the two largest fragments, and "
         "offspring are mutated by atom substitution, bond-order change, atom addition or atom "
         "deletion, with sanitisation after every operation. Selection is NSGA-II [20] with fast "
         "non-dominated sorting and crowding distance. Four objectives were optimised simultaneously: "
         "predicted FAP affinity (mean of five XGBoost models), predicted selectivity margin "
         "(pFAP − max(pPREP, pDPP4)), drug-likeness (QED) and synthetic accessibility (−SA/10) [21,22]. Hard "
         "filters removed invalid structures, compounds outside 250–650 Da, and compounds whose "
         "maximum Tanimoto similarity to the FAP training set was below 0.40. The initial population "
         "was drawn from the 130 measured FAP actives with pActivity ≥ 7.")
    para(doc,
         "Two controls were run with the same evaluation budget: an affinity-only campaign, in which "
         "NSGA-II degenerates into a hill climber on predicted potency, and a random campaign in which "
         "offspring replace parents without selection. All archived molecules were annotated for "
         "novelty (InChIKey absent from all three curated datasets) and for Murcko scaffold diversity, "
         "so that the comparison does not reward the rediscovery of known FAP inhibitors.")

    heading(doc, "4.10 Implementation and reproducibility", 2)
    para(doc,
         "All code is Python 3.13 with RDKit 2026.09.1 [27], scikit-learn 1.9.1, XGBoost 3.4.1 and "
         "NumPy 2.5.4. The pipeline is organised as numbered scripts (fetch, curate, benchmark, "
         "transfer, analysis, generative, figures, report), each writing a table or figure that the "
         "next stage reads. The curated datasets, the full set of predictions and the code are "
         "available in the repository listed in Appendix A.")

    # ------------------------------------------------------------------ results
    heading(doc, "§5. Results and Discussion", 1)
    heading(doc, "5.1 The curated FAP record is small, noisy and chemically non-stationary", 2)
    para(doc,
         f"Curation reduced 972 raw FAP bioactivities to {int(ds.loc['FAP','records_used'])} records "
         f"covering {int(ds.loc['FAP','compounds'])} unique compounds and "
         f"{int(ds.loc['FAP','scaffolds'])} Murcko scaffolds (Table 1). The median potency is "
         f"pActivity {ds.loc['FAP','pAct_mean']:.2f} (≈ {10**(-ds.loc['FAP','pAct_mean'])*1e9:.0f} nM) "
         f"with a standard deviation of {ds.loc['FAP','pAct_sd']:.2f}, and "
         f"{ds.loc['FAP','pct_active']:.0f}% of compounds reach the 100 nM activity threshold used "
         "throughout this work. PREP and DPP4 are larger but structurally similar; DPP4 in particular "
         "is dominated by potent compounds (59% active), which reflects its history as a diabetes "
         "target with industrial screening campaigns behind it.")
    table(doc,
          ["Target", "Raw records", "Retained records", "Unique compounds", "Scaffolds",
           "Median pActivity", "% active (≥7)", "Median MW (Da)"],
          [[t,
            int(D["curation"].query("target==@t and step=='raw records'").n.iloc[0]),
            int(ds.loc[t, "records_used"]),
            int(ds.loc[t, "compounds"]),
            int(ds.loc[t, "scaffolds"]),
            f"{ds.loc[t,'pAct_mean']:.2f} ± {ds.loc[t,'pAct_sd']:.2f}",
            f"{ds.loc[t,'pct_active']:.0f}",
            f"{ds.loc[t,'median_MW']:.0f}"] for t in ["FAP", "PREP", "DPP4"]],
          caption="**Table 1.** Composition of the three curated datasets after filtering. Activity "
                  "threshold for ‘active’ is pActivity ≥ 7 (IC50 ≤ 100 nM).",
          widths=[2.0, 2.2, 2.4, 2.4, 1.9, 2.6, 2.1, 2.4])
    para(doc,
         "The most consequential property of this dataset is not its size but its noise. Among FAP "
         "compounds measured more than once, the median standard deviation of replicate pActivity "
         "values is 0.67 log units (PREP 0.53, DPP4 0.37). Two measurements of the same compound in "
         "two laboratories therefore disagree by roughly a factor of five in IC50. This number is the "
         "error floor: a model that predicts every compound to within 0.67 log units is already at the "
         "limit of what the labels can resolve, and differences of 0.05 log units between algorithms "
         "are not interpretable.")
    para(doc,
         "The dataset is also chemically non-stationary (Figure 2). Compounds first reported in 2019 "
         "or later have a median molecular weight 80 Da higher (452 vs 372 Da) and a median potency "
         "two orders of magnitude greater (pActivity 7.60 vs 5.57) than earlier entries, and the "
         "fraction of actives rises from 15% to 64%. The historical explanation is straightforward: "
         "before 2018, most FAP measurements came from DPP4-inhibitor programmes that used FAP as a "
         "selectivity counter-screen, whereas after 2018 the literature is dominated by purpose-built "
         "quinoline FAPI warheads. The dataset therefore contains two different chemical questions "
         "under one target name, a property that will matter for every subsequent result.")
    figure(doc, FIG / "fig1_dataset.png",
           "**Figure 2.** The curated datasets. (a) Potency coverage of the three targets; the dotted "
           "line marks 100 nM. (b) FAP compounds split by publication era, showing the shift towards "
           "larger, more potent molecules from 2019 onwards. (c) Scaffold diversity, which is low for FAP "
           "relative to its compound count.")

    heading(doc, "5.2 Accuracy depends more on the split than on the algorithm", 2)
    bench = D["bench"]
    order = [("random", "Random"), ("scaffold", "Scaffold"), ("time", "Time")]
    rows = []
    for split, label in order:
        for m in ["mean", "ridge", "svr", "rf", "xgb"]:
            d = bench[(bench.split == split) & (bench.model == m)]
            name = {"mean": "Mean baseline", "ridge": "Ridge", "svr": "SVR (RBF)",
                    "rf": "Random forest", "xgb": "XGBoost"}[m]
            rows.append([f"{label} / {name}" if m == "mean" else f"{label} / {name}",
                         f"{d.RMSE.mean():.3f} ± {d.RMSE.std():.3f}",
                         f"{d.MAE.mean():.3f}",
                         f"{d.R2.mean():.2f}",
                         f"{d.Spearman.mean():.2f}" if d.Spearman.notna().any() else "--",
                         f"{d.screen_AUROC.mean():.2f}" if d.screen_AUROC.notna().any() else "--",
                         f"{d.screen_EF10.mean():.2f}" if d.screen_EF10.notna().any() else "--"])
    table(doc,
          ["Split / model", "RMSE", "MAE", "R²", "Spearman", "AUROC", "EF10"],
          rows,
          caption="**Table 2.** Benchmark of five model families under three split protocols "
                  "(mean ± s.d. over five seeds). AUROC and EF10 are computed on the focused "
                  "evaluation set of clear actives (pActivity ≥ 7) and clear inactives (≤ 5.5).",
          widths=[5.0, 2.6, 1.6, 1.4, 1.8, 1.5, 1.4], size=8.5)
    para(doc,
         "Random splitting gives the flattering picture that the literature usually reports: random "
         "forest and XGBoost reach R² of 0.75 and 0.75 with RMSE near 0.70, and the ranking "
         "correlation is 0.85. Both are statistically indistinguishable from each other and clearly "
         "ahead of ridge (R² 0.56) and SVR (R² 0.62). Scaffold splitting costs about 0.07 in R² and "
         "0.12 in RMSE for the tree ensembles, which is the expected price of extrapolating to new "
         "scaffolds and is still usable: R² 0.68, Spearman 0.81, screening AUROC 0.94.")
    para(doc,
         "The time split does not degrade the models, it destroys them. XGBoost reaches R² = −0.79 and "
         "a rank correlation of −0.11; the random forest is worse still (R² = −1.03). The mean "
         "predictor is as good as any learned model (R² ≈ −1.06). Enrichment falls to 0.96, i.e. no "
         "better than selecting compounds at random. Notably, ridge regression degrades far less "
         "(R² = −0.07), which is the signature of a model that relies on additive physicochemical "
         "trends rather than on scaffold-specific fingerprint patterns: it cannot rank the new series, "
         "but it also does not confidently rank it wrongly.")
    figure(doc, FIG / "fig2_benchmark.png",
           "**Figure 3.** Benchmark results. (a) RMSE by split protocol and model; error bars are "
           "standard deviations over five seeds. (b) Predicted versus measured pActivity for XGBoost "
           "under random and time splitting; the time-split cloud is compressed into a narrow band "
           "around pActivity 6 regardless of the measured value. (c) Early enrichment (EF10) on the "
           "focused evaluation set; the dashed line marks the no-enrichment level.")
    para(doc,
         "This is the central empirical result of the project, and it is not an artefact of a small "
         "test set (n = 147). It is the direct consequence of the era structure documented in "
         "Section 5.1: a model trained on pre-2019 DPP4-like chemistry is asked to rank quinoline "
         "FAPI warheads, a series it has never seen. The practical reading is that a FAP model "
         "trained on the literature available before the FAPI era has no prospective value for FAPI "
         "design, even though its random-split statistics are excellent.")

    heading(doc, "5.3 Which chemical information carries the signal?", 2)
    abl = D["ablation"]
    abl_g = abl.groupby("features")[["RMSE", "MAE", "R2", "Spearman", "screen_AUROC"]].mean()
    label = {"morgan": "Morgan only", "maccs": "MACCS only", "desc": "Descriptors only",
             "combined": "Combined (all three)"}
    table(doc,
          ["Feature block", "RMSE", "MAE", "R²", "Spearman", "AUROC"],
          [[label[f], f"{abl_g.loc[f,'RMSE']:.3f}", f"{abl_g.loc[f,'MAE']:.3f}",
            f"{abl_g.loc[f,'R2']:.2f}", f"{abl_g.loc[f,'Spearman']:.2f}",
            f"{abl_g.loc[f,'screen_AUROC']:.3f}"] for f in ["morgan", "maccs", "desc", "combined"]],
          caption="**Table 3.** Feature ablation with XGBoost under scaffold splitting (mean over "
                  "five seeds).",
          widths=[4.4, 2.2, 2.2, 1.8, 2.2, 2.0])
    para(doc,
         "The combined block is the most accurate in regression terms (RMSE 0.828 versus 0.866 for "
         "Morgan alone, 0.922 for MACCS and 0.987 for descriptors alone), which confirms that "
         "fingerprints and physicochemical descriptors carry complementary information: the "
         "fingerprints encode substructure, the descriptors encode the functional-group composition "
         "that defines a warhead. The ranking metrics tell a slightly different story, with the "
         "descriptor-only model achieving the highest AUROC (0.96) despite the worst RMSE. This "
         "divergence is a useful reminder that regression accuracy and screening utility are not the "
         "same quantity, and it is why both are reported here.")

    heading(doc, "5.4 Paralogue transfer does not rescue a small target", 2)
    tr = D["transfer"]
    rows = []
    for variant, name in [("baseline", "FAP features only"),
                          ("stacked_transfer", "FAP + paralogue predictions + similarity"),
                          ("transfer_nosim", "FAP + paralogue predictions")]:
        for split, lab in [("random", "Random"), ("scaffold", "Scaffold")]:
            d = tr[(tr.variant == variant) & (tr.split == split)]
            rows.append([f"{name}" if split == "random" else "", lab,
                         f"{d.RMSE.mean():.3f} ± {d.RMSE.std():.3f}",
                         f"{d.R2.mean():.3f}", f"{d.Spearman.mean():.3f}",
                         f"{d.screen_AUROC.mean():.3f}"])
    table(doc,
          ["Feature set", "Split", "RMSE", "R²", "Spearman", "AUROC"],
          rows,
          caption="**Table 4.** Value of paralogue transfer. Auxiliary models were trained on the "
                  "full PREP and DPP4 datasets and their predictions (plus the chemical similarity to "
                  "each paralogue training set) were appended to the FAP feature block. Mean over "
                  "five seeds.",
          widths=[6.2, 2.2, 3.0, 1.6, 2.0, 1.6], size=8.5)
    para(doc,
         "Transfer produces a consistent but tiny improvement. Under random splitting, RMSE falls "
         "from 0.710 to 0.697 and R² rises from 0.751 to 0.760; under scaffold splitting the change "
         "is 0.828 to 0.824. Paired comparisons across seeds give p = 0.19 and p = 0.62 respectively "
         "(Wilcoxon signed-rank), so the honest conclusion is that paralogue transfer is neutral at "
         "this data scale. Removing the similarity features changes nothing, which indicates that the "
         "small gain, where it exists, comes from the paralogue prediction itself rather than from a "
         "similarity-based neighbourhood signal.")
    para(doc,
         "This null result is informative. DPP4 has eleven times more measurements than FAP, but its "
         "chemistry is dominated by gliptin-like scaffolds that are precisely the compounds FAP "
         "chemists do not want. Transfer helps when the auxiliary task shares chemotypes with the "
         "target task; here it mostly shares an enzyme fold. The result also sets a bound on what "
         "more elaborate multi-task architectures could achieve with the same labels, since a shared "
         "encoder would have to overcome the same chemotype mismatch.")

    heading(doc, "5.5 Selectivity is a cliff, not a slope", 2)
    sel = D["sel"]
    selci = D["sel_ci"].reset_index()
    rows = []
    for split, lab in [("random", "Random"), ("scaffold", "Scaffold")]:
        for proto, pname in [("indirect", "Indirect (difference of two models)"),
                             ("direct", "Direct (model of ΔpActivity)")]:
            d = sel[(sel.split == split) & (sel.protocol == proto)]
            ci = selci[(selci.split == split) & (selci.protocol == proto)].iloc[0]
            rows.append([lab if proto == "indirect" else "", pname,
                         f"{d.delta_R2.mean():.2f}",
                         f"{d.delta_Spearman.mean():.2f}",
                         f"{d.selectivity_AUROC.mean():.2f} ± "
                         f"{d.selectivity_AUROC.std():.2f}",
                         f"{ci.selectivity_AUROC_ci95:.2f}"])
    table(doc,
          ["Split", "Protocol", "ΔR²", "Δ Spearman", "AUROC (mean ± s.d.)", "95% CI"],
          rows,
          caption="**Table 5.** Prediction of FAP selectivity over PREP on the 100 dual-measured "
                  "compounds (36 scaffolds), using twenty random seeds. Selectivity is defined as "
                  "ΔpActivity ≥ 1 (ten-fold preference for FAP).",
          widths=[2.2, 6.0, 1.6, 2.4, 3.4, 1.8], size=8.5)
    para(doc,
         "Inside known scaffolds, selectivity is learnable: both protocols reach an AUROC of about "
         "0.82–0.84 and a rank correlation of 0.58–0.63 for ΔpActivity. Direct modelling of the "
         "difference and indirect subtraction of two single-target models are statistically "
         "indistinguishable, which is a useful practical finding because the indirect route requires "
         "no dual-measured training set.")
    para(doc,
         "Outside known scaffolds the same models are worse than chance. AUROC falls to 0.36–0.38 "
         "with a standard deviation of about 0.12 across twenty seeds, and ΔR² becomes negative. A "
         "model that ranks selectively and non-selectively compounds in the wrong order on new "
         "scaffolds is not merely uninformative; used as a filter it would actively enrich the "
         "non-selective series. The most likely explanation is the sign flip of the SAR: within a "
         "series, the substituent that fills the FAP S2 pocket also fills the PREP S2 pocket, while "
         "across series the dominant determinant becomes the shape of the linker, which the model has "
         "never seen vary.")
    para(doc,
         "The descriptive statistics explain why this matters more for DPP4 than for PREP. Among 100 "
         "compounds measured on both FAP and PREP, the mean ΔpActivity is +0.59, with 32% selective "
         "and 34% non-selective; the distribution is broad and genuinely two-sided (Figure 4c). Among "
         "232 compounds measured on both FAP and DPP4, the mean Δ is −1.75 and 86% of compounds are "
         "non-selective: DPP4 is the dominant off-target liability of this chemotype, and the "
         "practical design objective is to preserve FAP affinity while pushing DPP4 below the "
         "selectivity threshold.")
    figure(doc, FIG / "fig3_transfer_selectivity.png",
           "**Figure 4.** Transfer and selectivity. (a) Paired difference in RMSE between the "
           "FAP-only and transfer models for each seed; diamonds are means with 95% confidence "
           "intervals, and the p-values are from Wilcoxon signed-rank tests. (b) Selectivity AUROC by "
           "split protocol and modelling strategy; error bars are 95% confidence intervals over "
           "twenty seeds. (c) Distribution of ΔpActivity for the dual-measured compounds against PREP "
           "and DPP4.")

    heading(doc, "5.6 When can an error bar be believed?", 2)
    ad = D["ad"]
    table(doc,
          ["Max Tanimoto to training set", "n", "RMSE", "MAE"],
          [[b, int(n), f"{r:.3f}", f"{m:.3f}"] for b, n, r, m in
           zip(ad.bin, ad.n, ad.rmse, ad.mae)],
          caption="**Table 6.** Error as a function of chemical similarity to the training set "
                  "(scaffold split, XGBoost, pooled over five seeds).",
          widths=[6.0, 2.0, 2.6, 2.6])
    para(doc,
         "Error rises monotonically as the query moves away from the training set: RMSE is 0.70 for "
         "compounds with a nearest-neighbour similarity above 0.7 (n = 263), 0.94 in the 0.6–0.7 bin, "
         "0.96 in the 0.5–0.6 bin and 0.95–1.12 below 0.5. The absolute magnitude of the effect is "
         "modest — roughly 0.3 log units between the best and worst bins — but it is monotone and it "
         "is the basis for a usable rule: a prediction supported by a similarity above 0.6 can be "
         "quoted with an error comparable to assay noise, and below 0.5 it cannot.")
    conf = D["conf"]
    conf_g = conf.groupby(["protocol", "level"]).coverage.mean()
    rows = []
    for proto, name in [("random", "Random split (control)"),
                        ("scaffold", "Scaffold split, all test compounds"),
                        ("scaffold:in_domain(sim>=0.6)", "Scaffold split, in domain (sim ≥ 0.6)"),
                        ("scaffold:out_domain(sim<0.6)", "Scaffold split, out of domain (sim < 0.6)")]:
        rows.append([name,
                     f"{conf_g.get((proto, 0.8), float('nan')):.2f}",
                     f"{conf_g.get((proto, 0.9), float('nan')):.2f}"])
    table(doc,
          ["Calibration protocol", "Coverage at 80%", "Coverage at 90%"],
          rows,
          caption="**Table 7.** Empirical coverage of split-conformal intervals (mean over five "
                  "seeds). Nominal coverage is 0.80 and 0.90; intervals are calibrated on out-of-fold "
                  "residuals inside the training set.",
          widths=[8.4, 3.4, 3.4])
    para(doc,
         "The random-split control behaves exactly as theory predicts: 0.79 and 0.89 empirical "
         "coverage at nominal 0.80 and 0.90. Under scaffold splitting, coverage drops to 0.73 and "
         "0.87 — a mild violation caused by the failure of exchangeability between calibration and "
         "test scaffolds. The violation becomes severe only outside the applicability domain, where "
         "the 80% interval covers 54% of compounds. In other words, the interval is honest for "
         "compounds that resemble the training set and optimistic for compounds that do not, which is "
         "the opposite of the intuition that error bars are conservative.")
    figure(doc, FIG / "fig4_applicability_domain.png",
           "**Figure 5.** Applicability domain. (a) Absolute error against maximum Tanimoto "
           "similarity to the training set, with the median per bin overlaid. (b) RMSE grouped by "
           "applicability class; the dashed line is the overall RMSE.")
    para(doc,
         "The practical rule that follows is a two-gate decision: compute the similarity, and only "
         "quote a conformal interval when the similarity exceeds 0.6. Below that threshold the model "
         "may still be useful for ranking a series internally, but its numeric interval should not "
         "appear in a synthesis decision.")

    heading(doc, "5.7 Prospective failure: the 2019 series shift", 2)
    tm = D["temporal"].set_index("setting")
    raw = tm.loc["train<=2018, test 2019-2024"]
    rec = tm.loc["intercept recalibrated on 25% of test"]
    table(doc,
          ["Setting", "RMSE", "MAE", "R²", "Spearman", "Mean offset"],
          [["Train ≤ 2018, test 2019–2024", f"{raw.RMSE:.3f}", f"{raw.MAE:.3f}",
            f"{raw.R2:.2f}", f"{raw.Spearman:.2f}", "0.00"],
           ["Intercept recalibrated on 25% of test", f"{rec.RMSE:.3f}", f"{rec.MAE:.3f}",
            f"{rec.R2:.2f}", f"{rec.Spearman:.2f}", f"{rec.mean_offset:+.2f}"]],
          caption="**Table 8.** Temporal extrapolation and intercept recalibration.",
          widths=[6.4, 2.0, 2.0, 1.6, 2.0, 2.4])
    para(doc,
         f"The raw time-split model has RMSE {raw.RMSE:.2f} and R² {raw.R2:.2f}, with a mean signed "
         f"error of {rec.mean_offset:+.2f} log units (a single fit at seed 0; the five-seed mean in §5.2 is "
         f"−0.79 with RMSE 2.05): it systematically under-predicts the potency of "
         "post-2018 compounds by more than an order of magnitude. Recalibrating a single intercept on "
         f"a quarter of the test set reduces RMSE to {rec.RMSE:.2f} and brings R² to {rec.R2:.2f}, but "
         f"the rank correlation is unchanged at {rec.Spearman:.2f}. This decomposition is the cleanest "
         "result in the project: the model's error is almost entirely a level shift, and correcting "
         "the level restores absolute accuracy without restoring any ranking ability. A chemist who "
         "recalibrated the model on a few new compounds would obtain correct potency estimates for the "
         "new series while still ranking its members at random.")
    figure(doc, FIG / "fig5_temporal_drift.png",
           "**Figure 6.** Temporal drift. (a) Interpolation under random splitting. (b) Extrapolation "
           "to compounds published after 2018, before and after intercept recalibration. The "
           "recalibrated points move onto the diagonal in level but remain uncorrelated in rank.")

    heading(doc, "5.8 Multi-objective generative design", 2)
    gen = D["gen"]
    arc = D["gen_arc"]
    stats = {}
    for mo, lab in [("ga", "Multi-objective (NSGA-II)"), ("affinity_only", "Affinity only"),
                    ("random", "Random (no selection)")]:
        d = arc[arc["mode"] == mo]
        nov = d[d.novel]
        ok = nov[(nov.pFAP >= 8) & (nov.margin >= 1)]
        stats[mo] = {"lab": lab, "n": len(d), "novel": int(d.novel.sum()),
                     "scaf": int(d.scaffold.nunique()),
                     "qed": float(d.qed.median()), "sa": float(d.sa.median()),
                     "mw": float(d.mw.median()),
                     "best": float(nov.pFAP.max()),
                     "pass": float(len(ok) / len(nov)) if len(nov) else float("nan")}
    table(doc,
          ["Campaign", "Archived", "Novel", "Scaffolds", "Median QED", "Median SA",
           "Median MW (Da)", "Best novel pFAP", "Pass rate"],
          [[stats[mo]["lab"], stats[mo]["n"], stats[mo]["novel"], stats[mo]["scaf"],
            f"{stats[mo]['qed']:.2f}", f"{stats[mo]['sa']:.2f}", f"{stats[mo]['mw']:.0f}",
            f"{stats[mo]['best']:.2f}", f"{stats[mo]['pass']*100:.0f}%"]
           for mo in ["ga", "affinity_only", "random"]],
          caption="**Table 9.** Outcome of the three design campaigns with an identical evaluation "
                  "budget (24 generations, population 96). Novelty = InChIKey absent from all three "
                  "curated datasets; pass rate = fraction of novel molecules with predicted pFAP ≥ 8 "
                  "and selectivity margin ≥ 1.",
          widths=[4.4, 1.7, 1.4, 1.8, 1.9, 1.7, 2.1, 2.2, 1.6], size=8.5)
    para(doc,
         "All three campaigns quickly reach the ceiling of the affinity model (best predicted pFAP "
         "9.6–9.7, Figure 7b), which is itself a warning: an ensemble trained on 384 compounds cannot "
         "distinguish a genuine 10 nM inhibitor from an implausible one, and the optimizer finds the "
         "model's most confident region within a few generations. The campaigns therefore differ not "
         "in predicted affinity but in what they sacrifice to get there. Without selection pressure "
         "the breeding population drifts downwards from pFAP 7.35 to 6.44 within 24 "
         "generations (Figure 7a); under affinity-only selection it climbs steadily while its "
         "chemistry degenerates — the median QED of the archived molecules collapses to 0.25 and the "
         "median molecular weight inflates to 548 Da, because polar atoms are traded for hydrophobic "
         "bulk wherever the model rewards it.")
    para(doc,
         "The multi-objective campaign pays a small affinity price (median novel pFAP 8.24 versus "
         "9.36 for affinity-only) and in exchange returns the most usable chemistry of the three "
         "regimes: 115 distinct Murcko scaffolds against 80–83 for the controls, a median QED of "
         "0.61, a median synthetic-accessibility score of 2.93 (easier to make than either control) "
         "and a median molecular weight of 413 Da, squarely in the range of the clinical FAPIs. "
         "61% of its novel molecules clear the pre-registered design thresholds of pFAP ≥ 8 and "
         "selectivity margin ≥ 1, essentially the same pass rate as the random control (60%) — the "
         "value of the objectives shows up in diversity and developability rather than in the pass "
         "rate, which is exactly what a Pareto front should do. The top-ranked novel molecules are "
         "quinoline-carboxamide derivatives with modified amide linkers, i.e. plausible next analogues "
         "rather than exotic chemistry (Table 10).")
    top = D["gen_top"].head(6)
    table(doc,
          ["Structure (SMILES)", "pFAP", "Margin", "QED", "SA", "MW (Da)"],
          [[r.smiles, f"{r.pFAP:.2f}", f"{r.margin:.2f}", f"{r.qed:.2f}", f"{r.sa:.2f}",
            f"{r.mw:.0f}"] for r in top.itertuples()],
          caption="**Table 10.** Six highest-ranked novel molecules from the multi-objective "
                  "campaign, ranked by predicted pFAP and selectivity margin.",
          widths=[8.3, 1.4, 1.8, 1.2, 1.0, 1.7], size=8.5, fixed=True)
    fig_gen = FIG / "fig7_generative.png"
    figure(doc, fig_gen,
           "**Figure 7.** Generative design. (a) Median predicted affinity of the breeding "
           "population per generation: random selection drifts downwards, affinity-only hill-climbing "
           "rises, multi-objective selection stays stable. (b) Archived molecules in the "
           "affinity-margin plane; dotted lines mark the design thresholds. (c) Drug-likeness of the "
           "archived molecules by campaign.")

    heading(doc, "5.9 Discussion", 2)
    para(doc,
         "Three findings deserve emphasis. First, the choice of evaluation protocol changes the "
         "conclusion more than the choice of algorithm: the same XGBoost model is excellent (R² 0.75) "
         "or useless (R² −0.79) depending only on how the data are divided. Any claim about FAP "
         "affinity prediction that does not specify the split is therefore unfalsifiable. This is not "
         "a novel observation in the machine-learning literature, but it is rarely applied to "
         "radiopharmaceutical targets, where datasets are small and the pressure to report a good "
         "number is high.")
    para(doc,
         "Second, the error floor set by assay noise (0.67 log units) is close to the best achievable "
         "RMSE (0.70 under random splitting). The model is not limited by its architecture; it is "
         "limited by the reproducibility of the underlying enzyme assays. Further modelling effort on "
         "this dataset has a low ceiling, and the highest-value intervention would be a consistent "
         "centralised assay campaign for a few hundred FAP inhibitors, not a better neural network.")
    para(doc,
         "Third, selectivity is the quantity that actually decides whether a FAP ligand advances, and "
         "it is the quantity that generalises worst. Predicting ΔpActivity within a series is easy; "
         "predicting it across series is not currently possible with 100 dual-measured compounds. "
         "Design campaigns should therefore treat selectivity as a filter to be verified "
         "experimentally rather than as an objective to be optimised in silico.")
    heading(doc, "5.10 Limitations", 2)
    for b in [
        "The dataset is one snapshot of ChEMBL and inherits its biases: publication bias towards "
        "potent compounds, inconsistent assay conditions, and an over-representation of the "
        "quinoline warhead after 2018.",
        "Only ligand-based descriptors were used. No docking, pharmacophore or protein-structure "
        "information entered the models, so the models cannot distinguish a genuine sub-pocket match "
        "from a physicochemical coincidence.",
        "The generative campaign scores candidates with the same models whose limitations are "
        "documented above. Predicted affinities above pActivity 9 should be read as ‘inside the "
        "model's confident region’, not as ‘10 nM’.",
        "The applicability-domain filter at similarity 0.40 constrains novelty by construction; a "
        "more adventurous campaign would generate more diverse but less assessable chemistry.",
        "Temporal analysis uses publication year as a proxy for availability; a true prospective "
        "test would require an external, later dataset.",
    ]:
        bullet(doc, b)

    # -------------------------------------------------------------- conclusions
    heading(doc, "§6. Conclusions", 1)
    para(doc,
         "This project asked what machine learning can deliver for FAP-targeted radiopharmaceutical "
         "design given the public bioactivity record, and where the answer is negative. A curated "
         "dataset of 384 unique FAP inhibitors supported a model that reaches R² 0.68 and screening "
         "AUROC 0.94 under scaffold splitting, which meets the pre-registered criterion for use as a "
         "triage filter. The same model fails completely under a time split (R² −0.79), and the "
         "failure is traceable to a documented change in chemical series rather than to algorithmic "
         "weakness. Paralogue transfer from PREP and DPP4 is neutral at this scale, and selectivity "
         "over PREP, although learnable within a series (AUROC 0.84), is anti-correlated across "
         "scaffolds (AUROC 0.38). Conformal intervals are valid under interpolation and become "
         "optimistic outside the applicability domain, so a similarity gate is a necessary companion "
         "to any quoted uncertainty.")
    para(doc,
         "The four research questions are therefore answered, though not in the direction that a "
         "typical QSAR paper would report: RQ1 — accuracy is protocol-dependent to the point of "
         "changing the sign of R²; RQ2 — transfer is neutral and selectivity does not generalise "
         "across scaffolds; RQ3 — intervals are trustworthy only inside the applicability domain; "
         "RQ4 — multi-objective selection improves the diversity and drug-likeness of the generated "
         "set (115 scaffolds, median QED 0.61) while single-objective optimisation degenerates "
         "chemically (median QED 0.25), and neither can manufacture information that the scoring "
         "model does not contain.")
    para(doc,
         "The broader contribution is methodological. For targets with a few hundred compounds, a "
         "drifting literature and heterogeneous assays, the deliverable of AI-driven drug discovery "
         "should be a calibrated decision rule — use the model here, do not use it there — rather "
         "than a single predicted potency. Building that rule requires the three analyses performed "
         "here (protocol comparison, applicability domain, conformal calibration), and they cost "
         "little compared with the synthesis of one analogue.")

    # --------------------------------------------------------------- references
    heading(doc, "§7. References", 1)
    refs = [
        "Lindner T, Loktev A, Altmann A, Giesel F, Kratochwil C, Debus J, Jäger D, Mier W, "
        "Haberkorn U. Development of quinoline-based theranostic ligands for the targeting of "
        "fibroblast activation protein. J Nucl Med. 2018;59(9):1415–1422.",
        "Loktev A, Lindner T, Mier W, Debus J, Altmann A, Jäger D, Haberkorn U. A tumor-imaging "
        "method targeting cancer-associated fibroblasts. J Nucl Med. 2018;59(9):1423–1429.",
        "Kratochwil C, Flechsig P, Lindner T, Abderrahim L, Altmann A, Mier W, et al. 68Ga-FAPI "
        "PET/CT: tracer uptake in 28 different kinds of cancer. J Nucl Med. 2019;60(6):801–805.",
        "Giesel FL, Kratochwil C, Lindner T, Marschalek MM, Loktev A, Lehnert W, et al. 68Ga-FAPI "
        "PET/CT: biodistribution and preliminary dosimetry estimate of 2 DOTA-containing FAP-targeting "
        "agents in patients with various cancers. J Nucl Med. 2019;60(3):386–392.",
        "Giesel FL, Adeberg S, Syed M, Lindner T, Jiménez-Franco LD, Mavriopoulou E, et al. FAPI-74 "
        "PET/CT using either 18F-AlF or cold-kit 68Ga labeling: biodistribution, radiation dosimetry, "
        "and tumor delineation in lung cancer patients. J Nucl Med. 2021;62(2):201–207.",
        "Baum RP, Schuchardt C, Singh A, Chantadisai M, Robiller FC, Zhang J, et al. Feasibility, "
        "biodistribution, and preliminary dosimetry in peptide-targeted radionuclide therapy of "
        "diverse adenocarcinomas using 177Lu-FAP-2286: first-in-human results. J Nucl Med. "
        "2022;63(3):415–423.",
        "Fendler WP, Pabst KM, Kessler L, Fragoso Costa P, Ferdinandus J, Weber M, et al. Safety and "
        "efficacy of 90Y-FAPI-46 radioligand therapy in patients with advanced sarcoma and other "
        "cancer entities. Clin Cancer Res. 2022;28(19):4346–4353.",
        "Fu H, Huang J, Zhao L, Chen Y, Xu W, Cai J, et al. 177Lu-LNC1004 radioligand therapy in "
        "patients with end-stage metastatic cancers: a single-centre, single-arm, phase II study. "
        "Clin Cancer Res. 2025;31(8):1415–1426.",
        "Fitzgerald AA, Weiner LM. The role of fibroblast activation protein in health and "
        "malignancy. Cancer Metastasis Rev. 2020;39(3):783–803.",
        "Altmann A, Haberkorn U, Siveke J. The latest developments in imaging of fibroblast "
        "activation protein. J Nucl Med. 2021;62(2):160–167.",
        "Zdrazil B, Felix E, Hunter F, Manners EJ, Blackshaw J, Corbett S, et al. The ChEMBL Database "
        "in 2023: a drug discovery platform spanning multiple bioactivity data types and time "
        "periods. Nucleic Acids Res. 2024;52(D1):D1180–D1192.",
        "Mayr A, Klambauer G, Unterthiner T, Steijaert M, Wegner JK, Ceulemans H, et al. Large-scale "
        "comparison of machine learning methods for drug target prediction on ChEMBL. Chem Sci. "
        "2018;9(24):5441–5451.",
        "van Tilborg D, Alenicheva A, Grisoni F. Exposing the limitations of molecular machine "
        "learning with activity cliffs. J Chem Inf Model. 2022;62(23):5938–5951.",
        "Wu Z, Ramsundar B, Feinberg EN, Gomes J, Geniesse C, Pappu AS, et al. MoleculeNet: a "
        "benchmark for molecular machine learning. Chem Sci. 2018;9(2):513–530.",
        "Wallach I, Heifets A. Most ligand-based classification benchmarks reward memorization rather "
        "than generalization. J Chem Inf Model. 2018;58(5):916–932.",
        "Sheridan RP. Time-split cross-validation as a method for estimating the goodness of "
        "prospective prediction. J Chem Inf Model. 2013;53(4):783–790.",
        "Sahigara F, Mansouri K, Ballabio D, Mauri A, Consonni V, Todeschini R. Comparison of "
        "different approaches to define the applicability domain of QSAR models. Molecules. "
        "2012;17(5):4791–4810.",
        "Angelopoulos AN, Bates S. A gentle introduction to conformal prediction and distribution-free "
        "uncertainty quantification. arXiv:2107.07511. 2021.",
        "Jensen JH. A graph-based genetic algorithm and generative model/Monte Carlo tree search for "
        "the exploration of chemical space. Chem Sci. 2019;10(12):3567–3572.",
        "Deb K, Pratap A, Agarwal S, Meyarivan T. A fast and elitist multiobjective genetic algorithm: "
        "NSGA-II. IEEE Trans Evol Comput. 2002;6(2):182–197.",
        "Bickerton GR, Paolini GV, Besnard J, Muresan S, Hopkins AL. Quantifying the chemical beauty "
        "of drugs. Nat Chem. 2012;4(2):90–98.",
        "Ertl P, Schuffenhauer A. Estimation of synthetic accessibility score of drug-like molecules "
        "based on molecular complexity and fragment contributions. J Cheminform. 2009;1:8.",
        "Chen T, Guestrin C. XGBoost: a scalable tree boosting system. In: Proceedings of the 22nd "
        "ACM SIGKDD International Conference on Knowledge Discovery and Data Mining. 2016. p. "
        "785–794.",
        "Breiman L. Random forests. Mach Learn. 2001;45(1):5–32.",
        "Bemis GW, Murcko MA. The properties of known drugs. 1. Molecular frameworks. J Med Chem. "
        "1996;39(15):2887–2893.",
        "Rogers D, Hahn M. Extended-connectivity fingerprints. J Chem Inf Model. 2010;50(5):742–754.",
        "Landrum G. RDKit: open-source cheminformatics software. https://www.rdkit.org (accessed "
        "November 2026).",
        "Papadatos G, Gaulton A, Hersey A, Overington JP. Activity, assay and target data curation "
        "and quality in the ChEMBL database. J Comput Aided Mol Des. 2015;29(9):885–896.",
        "Ruzzeh S, Abdlkadir AS, Paez D, Kang KW, Watabe T, Moghrabi S, et al. Therapeutic potential "
        "of FAPI RLT in oncology: a systematic review. Theranostics. 2025;15(9):4084–4100.",
    ]
    for i, ref in enumerate(refs, 1):
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(3)
        p.paragraph_format.line_spacing = 1.15
        p.paragraph_format.left_indent = Cm(1.0)
        p.paragraph_format.first_line_indent = Cm(-1.0)
        r = p.add_run(f"[{i}]  ")
        set_run(r, size=10.5)
        r2 = p.add_run(ref)
        set_run(r2, size=10.5)

    # ----------------------------------------------------------------- appendix
    heading(doc, "Appendix A. Code and data availability", 1)
    para(doc,
         "The complete project — numbered analysis scripts, curated datasets, prediction tables and "
         "figure-generation code — is available at:")
    para(doc, "https://github.com/huangyong/fap-aidd-8122", align="center", italic=True)
    para(doc,
         "The repository is organised as src/01_fetch_chembl.py through src/09_build_report.py, with "
         "data/raw holding the retrieved ChEMBL records and results/tables holding every number "
         "reported in this document. Running the scripts in numerical order reproduces all figures "
         "and tables from the public API; a pinned requirements file and a short README are included.")


def main() -> None:
    D = load_numbers()
    doc = Document(str(TEMPLATE))

    # ---- cover page
    tbl = doc.tables[0]
    # merge the title row
    a = tbl.cell(0, 0)
    b = tbl.cell(0, 1)
    merged = a.merge(b)
    merged.text = ""
    p = merged.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(10)
    r = p.add_run(TITLE)
    set_run(r, size=15, bold=True)
    for row, value in ((3, SID), (4, STUDENT)):
        cell = tbl.cell(row, 1)
        cell.text = ""
        pp = cell.paragraphs[0]
        rr = pp.add_run(value)
        set_run(rr, size=12, bold=True)
    for row in range(len(tbl.rows)):
        if tbl.cell(row, 0).text.strip().startswith("Submission Date"):
            cell = tbl.cell(row, 1)
            cell.text = ""
            rr = cell.paragraphs[0].add_run(DATE)
            set_run(rr, size=12, bold=True)
            break

    # ---- replace the cached TOC entries with a field
    body = doc.element.body
    toc_ps = [p for p in doc.paragraphs if p.style.name == "toc 1"]
    if toc_ps:
        nxt = toc_ps[-1]._element.getnext()
        for p in toc_ps:
            p._element.getparent().remove(p._element)
        new_p = toc_field(doc)
        if nxt is not None:
            nxt.addprevious(new_p._p)

    # ---- drop the placeholder References heading and everything after it except sectPr
    start = None
    for i, child in enumerate(body):
        if child.tag.endswith("}p"):
            txt = "".join(t.text or "" for t in child.iter(qn("w:t"))).strip()
            if txt == "References":
                start = i
                break
    if start is not None:
        for child in list(body)[start:]:
            if child.tag.endswith("}sectPr"):
                continue
            body.remove(child)

    # drop leftover empty paragraphs between the TOC field and the body content
    toc_pos = None
    children = list(body)
    for i, child in enumerate(children):
        if child.tag.endswith("}p") and "TOC" in child.xml:
            toc_pos = i
            break
    if toc_pos is not None:
        for child in children[toc_pos + 1:]:
            if child.tag.endswith("}sectPr"):
                continue
            if child.tag.endswith("}p") and not "".join(
                    t.text or "" for t in child.iter(qn("w:t"))).strip():
                body.remove(child)

    build_body(doc, D)
    OUT.parent.mkdir(parents=True, exist_ok=True)
    # drop empty paragraphs left between the cover table and the TOC heading
    _body = doc.element.body
    _kids = list(_body)
    _toc_i = None
    for _i, _ch in enumerate(_kids):
        if _ch.tag.endswith("}p"):
            _txt = "".join(t.text or "" for t in _ch.iter(qn("w:t")))
            if "Table of Contents" in _txt:
                _toc_i = _i
                break
    if _toc_i is not None:
        _tbl_i = max((_i for _i, _ch in enumerate(_kids[:_toc_i]) if _ch.tag.endswith("}tbl")),
                     default=None)
        if _tbl_i is not None:
            for _ch in _kids[_tbl_i + 1:_toc_i]:
                if _ch.tag.endswith("}p") and not "".join(
                        t.text or "" for t in _ch.iter(qn("w:t"))).strip():
                    _body.remove(_ch)

    # use a true minus sign for negative numbers in body text and tables
    import re as _re
    _pat = _re.compile(r"(?<![\d\w])-(?=\d)")

    def _fix(text: str) -> str:
        if not text:
            return text
        out = _pat.sub("\u2212", text)
        if out.startswith("-") and len(out) > 1 and out[1].isdigit():
            out = "\u2212" + out[1:]
        return out

    for _p in doc.paragraphs:
        for _r in _p.runs:
            if "-" in _r.text:
                _r.text = _fix(_r.text)
    for _t in doc.tables:
        for _row in _t.rows:
            for _cell in _row.cells:
                for _p in _cell.paragraphs:
                    for _r in _p.runs:
                        if "-" in _r.text:
                            _r.text = _fix(_r.text)

    doc.save(str(OUT))
    print("saved", OUT)


if __name__ == "__main__":
    main()
