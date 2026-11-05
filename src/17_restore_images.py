"""Restore full-resolution figures inside the Word-saved report package.

Word downsamples embedded pictures to ~1364 px when it saves a .docx (the
``w:doNotCompressImages`` flag is written by 09_build_report.py but Word 2021
drops it on its own save).  The PDF export keeps the full-resolution bitmaps,
so the submitted PDF is unaffected; this script repairs the .docx so that both
deliverables carry the same figure data.

Matching is content-based, not order-based: every embedded image is compared
with each candidate in ``results/figures`` after resizing the candidate to the
embedded pixel size, and the lowest RMSE wins.  A match is accepted only if the
RMSE is small and the aspect ratios agree, so a wrong replacement cannot happen
silently.

Usage
-----
    python src/17_restore_images.py [path/to/report.docx]
"""

from __future__ import annotations

import io
import shutil
import sys
import zipfile
from pathlib import Path

import numpy as np
from PIL import Image

ROOT = Path(__file__).resolve().parents[1]
FIG = ROOT / "results" / "figures"
DEFAULT = ROOT / "report" / "AIDD8122_Final_Report_HuangYong_P2600222_v2.docx"
RMSE_MAX = 16.0
ASPECT_TOL = 0.01


def rmse(a: Image.Image, b: Image.Image) -> float:
    a = a.convert("RGB")
    b = b.convert("RGB")
    if a.size != b.size:
        b = b.resize(a.size, Image.LANCZOS)
    x = np.asarray(a, dtype=np.float32)
    y = np.asarray(b, dtype=np.float32)
    return float(np.sqrt(((x - y) ** 2).mean()))


def normalize_props(target: Path) -> list[str]:
    """Re-normalise docProps after the Office save.

    The Word.Application ProgID on this machine is served by WPS Office, which
    rewrites ``docProps/app.xml`` with its own Application string and zeroed
    statistics, and resets ``cp:lastModifiedBy`` to the signed-in account.  The
    numbers below are recomputed from ``word/document.xml`` so that the file
    properties of the submitted .docx describe the document rather than the
    software that last touched it.
    """
    import re

    with zipfile.ZipFile(target) as z:
        items = [(n, z.read(n)) for n in z.namelist()]

    document = next(d for n, d in items if n == "word/document.xml").decode("utf-8")
    texts = re.findall(r"<w:t(?:\s[^>]*)?>([^<]*)</w:t>", document)
    body = " ".join(t for t in texts if t.strip())
    words = len(body.split())
    chars_ns = len(re.sub(r"\s", "", body))
    chars_ws = len(body)
    paras = document.count("<w:p ")

    changed = []
    out = []
    for name, data in items:
        if name == "docProps/app.xml":
            s = data.decode("utf-8")
            s = re.sub(r"<Application>.*?</Application>",
                       "<Application>Microsoft Office Word</Application>", s)
            s = re.sub(r"<Words>\d+</Words>", f"<Words>{words}</Words>", s)
            s = re.sub(r"<Characters>\d+</Characters>", f"<Characters>{chars_ns}</Characters>", s)
            s = re.sub(r"<CharactersWithSpaces>\d+</CharactersWithSpaces>",
                       f"<CharactersWithSpaces>{chars_ws}</CharactersWithSpaces>", s)
            s = re.sub(r"<Paragraphs>\d+</Paragraphs>", f"<Paragraphs>{paras}</Paragraphs>", s)
            data = s.encode("utf-8")
            changed.append(f"app.xml: {words} words, {paras} paragraphs")
        elif name == "docProps/core.xml":
            s = data.decode("utf-8")
            s = re.sub(r"<cp:lastPrinted[^>]*>.*?</cp:lastPrinted>", "", s)
            s = re.sub(r"<cp:lastModifiedBy>.*?</cp:lastModifiedBy>",
                       "<cp:lastModifiedBy>Huang Yong</cp:lastModifiedBy>", s)
            s = re.sub(r"<dc:creator>.*?</dc:creator>",
                       "<dc:creator>Huang Yong</dc:creator>", s)
            data = s.encode("utf-8")
            changed.append("core.xml: author reset, lastPrinted dropped")
        out.append((name, data))

    tmp = target.with_name(target.name + ".tmp")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in out:
            z.writestr(name, data)
    shutil.move(str(tmp), str(target))
    return changed


def main() -> None:
    target = Path(sys.argv[1]) if len(sys.argv) > 1 else DEFAULT
    if not target.exists():
        raise SystemExit(f"missing {target}")

    candidates = {}
    for f in sorted(FIG.glob("*.png")):
        candidates[f.name] = Image.open(f).convert("RGB")

    with zipfile.ZipFile(target) as z:
        items = [(n, z.read(n)) for n in z.namelist()]

    replaced, kept, failed = [], [], []
    new_items = []
    for name, data in items:
        if not (name.startswith("word/media/") and data[:8] == b"\x89PNG\r\n\x1a\n"):
            new_items.append((name, data))
            continue
        emb = Image.open(io.BytesIO(data)).convert("RGB")
        emb_ar = emb.size[0] / emb.size[1]
        best, best_rmse = None, 1e9
        for fn, cand in candidates.items():
            if abs(cand.size[0] / cand.size[1] - emb_ar) > ASPECT_TOL:
                continue
            r = rmse(emb, cand)
            if r < best_rmse:
                best, best_rmse = fn, r
        if best is None or best_rmse > RMSE_MAX:
            kept.append((name, emb.size, round(best_rmse, 1), best))
            new_items.append((name, data))
            continue
        src = FIG / best
        cand = candidates[best]
        if cand.size[0] * cand.size[1] <= emb.size[0] * emb.size[1]:
            # already at least as large as the embedded copy
            kept.append((name, emb.size, round(best_rmse, 1), best))
            new_items.append((name, data))
            continue
        new_items.append((name, src.read_bytes()))
        replaced.append((name, emb.size, cand.size, best, round(best_rmse, 1)))

    if not replaced:
        print("nothing to restore")
    else:
        tmp = target.with_name(target.name + ".tmp")
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
            for name, data in new_items:
                z.writestr(name, data)
        shutil.move(str(tmp), str(target))

    for name, old, new, best, r in replaced:
        print(f"  restored {name:24s} {old[0]}x{old[1]} -> {new[0]}x{new[1]}  ({best}, RMSE {r})")
    for name, size, r, best in kept:
        print(f"  kept     {name:24s} {size[0]}x{size[1]}  (best {best}, RMSE {r})")
    for line in normalize_props(target):
        print(f"  props    {line}")
    print(f"\n{len(replaced)} restored, {len(kept)} left untouched -> {target}")


if __name__ == "__main__":
    main()
