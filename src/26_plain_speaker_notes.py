# -*- coding: utf-8 -*-
"""把白话中文口播备注写进现有 10 分钟 PPT。

只动演讲者备注栏：不重建幻灯片、不重导 PDF、不碰版式与正文文字。
备注内容以 src/deck_notes.py 为唯一来源，重跑多少次结果都一样。

用法：
    python src/26_plain_speaker_notes.py            # 写入并自检
    python src/26_plain_speaker_notes.py --check    # 只检查，不写文件
"""
from __future__ import annotations

import argparse
import hashlib
import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(Path(__file__).resolve().parent))

from deck_notes import NOTES  # noqa: E402
from pptx import Presentation  # noqa: E402

PPTX = (ROOT / "AIDD8122_Final_Presentation_HuangYong"
        / "AIDD8122_Final_Presentation_HuangYong_10min.pptx")

# 书面腔 / AI 腔高频套路词，写进备注就算不合格
BANNED = (
    "综上", "值得注意的是", "综上所述", "不难看出", "众所周知", "本节将",
    "赋能", "闭环", "系统性", "换言之", "不可证伪", "可迁移的是协议",
    "目标函数就是化学", "能辩护的误差棒", "决定性意义", "奠定", "旨在",
)


def slide_texts(prs) -> list[list[str]]:
    """幻灯片可见文字，用于证明正文一个字没动。"""
    return [[sh.text_frame.text for sh in s.shapes if sh.has_text_frame]
            for s in prs.slides]


def digest(obj) -> str:
    return hashlib.sha256(json.dumps(obj, ensure_ascii=False).encode("utf-8")).hexdigest()


def lint(prs) -> None:
    assert len(prs.slides) == len(NOTES) == 13, "页数或备注条数不是 13"
    total = 0
    for i, slide in enumerate(prs.slides, 1):
        got = slide.notes_slide.notes_text_frame.text.strip() if slide.has_notes_slide else ""
        want = NOTES[i].strip()
        assert got == want, f"第 {i} 页备注与 deck_notes.py 不一致"
        # 备注里不允许出现时间戳或时长：节奏由讲者自己把握
        for bad in ("［", "］", "分钟", "小时"):
            assert bad not in got, f"第 {i} 页仍含时间表述：{bad}"
        assert any("\u4e00" <= c <= "\u9fff" for c in got), f"第 {i} 页不是中文"
        total += len(got)
    all_text = "\n".join(NOTES.values())
    hits = [w for w in BANNED if w in all_text]
    assert not hits, f"发现书面/AI 套话: {hits}"
    # 口播标记：白话版应当有这些日常说法，而不是通篇书面语
    markers = ("说白了", "请一定看一眼", "我也没装", "挺土的规矩",
               "好家伙", "省得老师问", "一点忙都帮不上", "摇头")
    hit = [m for m in markers if m in all_text]
    assert len(hit) >= 4, f"白话口播标记太少: {hit}"
    print(f"  slides        : {len(prs.slides)}")
    print(f"  notes pages   : {sum(1 for s in prs.slides if s.has_notes_slide)}/13")
    print(f"  notes chars   : {total:,}")
    print(f"  banned phrases: 0")
    print(f"  plain markers : {len(hit)}/8  {hit}")
    print(f"  time refs     : 0 (no timestamps, no durations)")


def apply() -> None:
    prs = Presentation(str(PPTX))
    before = digest(slide_texts(prs))
    for i, slide in enumerate(prs.slides, 1):
        slide.notes_slide.notes_text_frame.text = NOTES[i].strip()
    prs.save(str(PPTX))

    after = Presentation(str(PPTX))
    assert digest(slide_texts(after)) == before, "幻灯片正文被改动了，立即停止"
    print(f"  PPTX updated  : {PPTX.name}")
    print(f"  body text sha : {before[:16]} (unchanged)")
    lint(after)


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--check", action="store_true", help="只校验，不写文件")
    args = ap.parse_args()
    if args.check:
        lint(Presentation(str(PPTX)))
    else:
        apply()


if __name__ == "__main__":
    main()
