# v2 提交前修正记录（2026-11-05）

依据 `AIDD8122_v2提交前评估_20261105.md` 的 6 项 P0 与相关 P1，已全部处理。

| 编号 | 问题 | 处理 | 证据 |
|---|---|---|---|
| P0-1 | 21% 被当作客观缓解率（实为 3/4 级血液学毒性） | §1.2 改为 ORR 14%（4/28）、DCR 46%，21% 移至毒性句；PPT 第 3 页大数字改 14%，标签改 `objective response (4/28); DCR 46%`，来源行标 `21% grade 3/4 haematological tox` | Fu 2025（177Lu-LNC1004，n=28）ITT：4 PR、13 SD；6/28 例 3–4 级血液学毒性 |
| P0-2 | DOTA 分子量偏低 28 Da | `src/12_design_feasibility.py` 的 DOTA SMILES 少一个乙撑桥（10 元环 376.37 → 正确 cyclen 404.42）；重算残基 386.41；Table B1 六行各 +28 Da；§1.2/§5.9/Appendix B 全文同步 | RDKit：C16H28N4O8，MW 404.42；Table B1 = 881.0/919.0/889.0/878.9/878.9/949.9；中位 498.6 → 885.0 Da，TPSA 247.7 不变 |
| P0-3 | §5.6 整段（含 Figure 8 图注）重复两次 | 删除引用 `Table C2` 的第一遍，保留 `Table C3` 版本 | 全文 `A natural repair is class-conditional` 与 `Figure 8.` 各 1 次 |
| P0-4 | 页脚 `Page 1 of 1` | `build_report_pdf.ps1` 增加页眉/页脚 story 域刷新；PDF 逐页核对 | PDF 30 页，页脚依次 `Page 1 of 30` … `Page 30 of 30` |
| P0-5 | §3.3 C3 证据 0.76/0.50 与 Table 7 不符 | 改为 0.77 / 0.54 | Table 7、Table C3 均为 0.765/0.543（四舍五入 0.77/0.54） |
| P0-6 | Table 8 `Mean offset` 列与正文冲突 | 列名改 `Offset applied (log)`，表注写明原始模型 mean signed error +1.18、校正施加 +1.18、残余 0.00 | Table 8 两行：0.00 / +1.18 |

## P1

- **§4.10 脚本命名**：删除 “content corrections of this optimisation pass”“figure-restoration step” 等过程性表述，改为中性描述（report assets / evidence bundle / asset-integrity checks）。
- **PPT 口径**：第 11 页补 `warheads only, not radioligands`；第 8 页 86%/Δ−1.75 标注 `vs DPP4` 并交代 mixed-assay 分层（n=30）反转；第 3 页文字压缩到约 60 秒口播量；第 8/11 页溢出的长句已重排。
- **演讲备注**：15 页全部写入备注，并改为标准 notes body 占位符（PowerPoint 备注窗格/Presenter view 可见），总计时 14:50。
- **附录 A 与代码包**：`Submission_HuangYong_P2600222/03_Code/` 实际存在，含 01–17 主线脚本、`fix_*` 工具、`logs/`、`data/raw`、`data/processed`、`results/`。

## 仍需本人确认

1. 封面 16/11/2026（提交截止）与 PPT 24/11/2026（Group 2 汇报日）是否为两个不同日期；若 brief 只要求一个日期，需统一。
2. brief 的字数/页数上限：v2 正文约 7,772 词、全文 30 页；若限 8,000 词或 20 页，优先压缩 §2 与 §5.10 Discussion。
3. 报告 Appendix A 目前写的是提交包内 `03_Code/`，若课程要求可访问的 GitHub 链接，需要补仓库 URL 并重出 PDF。
