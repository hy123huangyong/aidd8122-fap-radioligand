# -*- coding: utf-8 -*-
"""Assemble the submission package from the working tree.

Builds ``Submission_HuangYong_P2600222/`` with the v2 report, the 10-minute
presentation and its notes, the numbered source code, the raw ChEMBL exports,
the intermediate curated tables, the process evidence (logs, environment
fingerprint, reproduction log, file manifest), every results table/figure
quoted in the report, and the route-decision record.

Usage
-----
    python src/pack_submission.py            # build / refresh the package
    python src/pack_submission.py --clean    # remove it first, then rebuild
    python src/pack_submission.py --no-zip   # skip the .zip archive
"""
from __future__ import annotations

import argparse
import shutil
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
PKG = ROOT / "Submission_HuangYong_P2600222"

REPORT = ROOT / "report"
DECK = ROOT / "AIDD8122_Final_Presentation_HuangYong"
DECK_STEM = "AIDD8122_Final_Presentation_HuangYong_10min"
REPORT_STEM = "AIDD8122_Final_Report_HuangYong_P2600222_v2"

LAYOUT = {
    "01_Report": [
        (REPORT / f"{REPORT_STEM}.pdf", f"{REPORT_STEM}.pdf"),
        (REPORT / f"{REPORT_STEM}.docx", f"{REPORT_STEM}.docx"),
    ],
    "02_Presentation": [
        (DECK / f"{DECK_STEM}.pdf", f"{DECK_STEM}.pdf"),
        (DECK / f"{DECK_STEM}.pptx", f"{DECK_STEM}.pptx"),
    ],
    "03_Code": [
        (ROOT / "README.md", "README.md"),
        (ROOT / "requirements.txt", "requirements.txt"),
    ],
    "05_Data": [
        (ROOT / "data" / "processed" / name, name)
        for name in ("fap_curated.csv", "prep_curated.csv", "dpp4_curated.csv")
    ],
}


def copy_files() -> list[str]:
    copied = []

    # package-level indices (final deliverables only)
    for name, dest in (("00_START_HERE.md", "00_START_HERE.md"),
                       ("FIGURE_INDEX.md", "04_Results/FIGURE_INDEX.md")):
        index = ROOT / "docs" / name
        if index.exists():
            target = PKG / dest
            target.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(index, target)
            copied.append(dest)

    for sub, items in LAYOUT.items():
        dest_dir = PKG / sub
        dest_dir.mkdir(parents=True, exist_ok=True)
        for src, name in items:
            if not src.exists():
                print(f"  ! missing {src.relative_to(ROOT)}", file=sys.stderr)
                continue
            shutil.copy2(src, dest_dir / name)
            copied.append(f"{sub}/{name}")

    # numbered source code
    src_dir = PKG / "03_Code" / "src"
    src_dir.mkdir(parents=True, exist_ok=True)
    publishing_helpers = {"authorize_device.py", "publish_repo.py"}
    for py in sorted((ROOT / "src").glob("*.py")):
        if py.name in publishing_helpers:
            continue
        shutil.copy2(py, src_dir / py.name)
        copied.append(f"03_Code/src/{py.name}")

    # style assets needed by 24_restyle_mpu_deck.py (crest, wordmark, schematics)
    asset_src = ROOT / "assets"
    if asset_src.exists():
        asset_dir = PKG / "assets"
        asset_dir.mkdir(parents=True, exist_ok=True)
        for f in sorted(asset_src.glob("*.png")):
            shutil.copy2(f, asset_dir / f.name)
            copied.append(f"assets/{f.name}")

    # P0.1 evidence bundle: logs + raw API exports + intermediate tables
    for sub, source, pattern in (
        ("03_Code/logs", ROOT / "logs", "*"),
        ("03_Code/data/raw", ROOT / "data" / "raw", "*.json"),
        ("03_Code/data/processed", ROOT / "data" / "processed", "*.csv"),
    ):
        if not source.exists():
            continue
        dest_dir = PKG / sub
        dest_dir.mkdir(parents=True, exist_ok=True)
        for f in sorted(source.glob(pattern)):
            if f.is_file():
                shutil.copy2(f, dest_dir / f.name)
                copied.append(f"{sub}/{f.name}")

    # figures and tables
    for sub, pattern in (("04_Results/figures", "*.png"),
                         ("04_Results/tables", "*.csv"),
                         ("04_Results/tables", "*.json"),
                         ("04_Results/tables", "*.parquet")):
        dest_dir = PKG / sub
        dest_dir.mkdir(parents=True, exist_ok=True)
        for f in sorted((ROOT / "results" / sub.split("/")[-1]).glob(pattern)):
            shutil.copy2(f, dest_dir / f.name)
            copied.append(f"{sub}/{f.name}")
    return copied


def export_notes() -> None:
    """Dump the speaker notes to a printable markdown file."""
    try:
        from pptx import Presentation
    except ImportError:
        print("  ! python-pptx not installed, skipping speaker notes", file=sys.stderr)
        return
    deck = DECK / f"{DECK_STEM}.pptx"
    if not deck.exists():
        return
    prs = Presentation(deck)
    pages = [(s.notes_slide.notes_text_frame.text.strip() if s.has_notes_slide else "")
             for s in prs.slides]
    total = sum(len(p) for p in pages)          # count follows the deck, not a constant
    lines = ["# AIDD8122 汇报讲稿（大白话版）\n",
             "Huang Yong · P2600222 · Group 2 · 12 November 2026\n",
             "每页都是能直接照着念的白话，已经同步写进 .pptx 的演讲者备注栏；这份是打印用的。\n",
             f"全篇约 {round(total, -2):.0f} 字，正常语速十分钟。时间紧就每页只念开头一两句和最后结论，\n",
             "中间的数字解释留给自己看懂就行。开讲前先翻最后的「提问环节可能问到」。\n"]
    for i, text in enumerate(pages, 1):
        lines.append(f"\n---\n\n## {i:02d}\n\n```\n{text}\n```\n")

    # 白话 Q&A 只附在打印版，不塞进 8 秒的末页备注栏
    sys.path.insert(0, str(Path(__file__).resolve().parent))
    from deck_notes import QA
    lines.append("\n---\n\n" + QA)

    out = PKG / "02_Presentation" / "Speaker_Notes_10min.md"
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text("".join(lines), encoding="utf-8")
    print(f"  + 02_Presentation/{out.name}")


def make_zip() -> Path:
    archive = ROOT / PKG.name
    out = shutil.make_archive(str(archive), "zip", str(PKG))
    return Path(out)


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--clean", action="store_true", help="delete the package first")
    ap.add_argument("--no-zip", action="store_true", help="skip the .zip archive")
    args = ap.parse_args()

    if args.clean and PKG.exists():
        shutil.rmtree(PKG)
        print(f"removed {PKG.name}")

    copied = copy_files()
    export_notes()

    total = sum(f.stat().st_size for f in PKG.rglob("*") if f.is_file())
    n_files = sum(1 for f in PKG.rglob("*") if f.is_file())
    print(f"\n{len(copied)} files copied -> {PKG}")
    print(f"package: {n_files} files, {total / 1e6:.1f} MB")
    if not args.no_zip:
        z = make_zip()
        print(f"archive: {z.name} ({z.stat().st_size / 1e6:.1f} MB)")


if __name__ == "__main__":
    main()
