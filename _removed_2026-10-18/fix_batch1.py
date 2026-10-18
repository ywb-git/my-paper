import io
import sys

sys.stdout.reconfigure(encoding="utf-8")

FABER = "2×10⁶ 化学空间（ML 预测，MAE≈0.1 eV/atom）；约 10⁴ DFT 真值"

edits = [
    # 1. Elpasolite 表述（6 处）
    ("开题报告02_材料问答中的选择性拒答_2026-10-18.md",
     "| Elpasolite | 约 200 万成分 | 闭集问答，真值完整 |",
     f"| Elpasolite | {FABER} | 闭集问答只用 DFT 子集 |"),
    ("开题报告04_材料逆向设计的运行间稳定性_2026-10-18.md",
     "| Elpasolite | 约 200 万成分，完整真值 | 闭集，可精确计算覆盖率 |",
     f"| Elpasolite | {FABER} | 闭集评测只用 DFT 子集 |"),
    ("开题报告05_材料设计的交互式多轮修正_2026-10-18.md",
     "| Elpasolite | 约 200 万成分 | 闭集多轮任务 |",
     f"| Elpasolite | {FABER} | 多轮任务只用 DFT 子集 |"),
    ("开题报告07_可合成性引导的材料生成_2026-10-18.md",
     "| Elpasolite | 约 200 万 | 闭集验证 |",
     f"| Elpasolite | {FABER} | 闭集验证只用 DFT 子集 |"),
    ("开题报告08_材料逆向设计的反事实敏感性_2026-10-18.md",
     "| Elpasolite | 约 200 万成分，完整真值 | 闭集，可精确计算 ρ |",
     f"| Elpasolite | {FABER} | 闭集评测只用 DFT 子集 |"),
    ("开题报告08_材料逆向设计的反事实敏感性_2026-10-18.md",
     "**数据可行性。** Elpasolite 提供完整真值，使 ρ 可精确计算，这是本课题相对于开放集的最大优势。",
     "**数据可行性。** Elpasolite 的约 10⁴ DFT 子集提供真值，使 ρ 可在该子集上精确计算；2×10⁶ 化学空间的其余部分是 ML 预测（MAE≈0.1 eV/atom），只用于生成质量与开放集评测。"),

    # 8. 重复章节号
    ("开题报告04_材料逆向设计的运行间稳定性_2026-10-18.md",
     "### 4.4 技术路线",
     "### 4.5 技术路线"),
    ("开题报告08_材料逆向设计的反事实敏感性_2026-10-18.md",
     "### 4.5 技术路线",
     "### 4.6 技术路线"),
]

for path, old, new in edits:
    with io.open(path, encoding="utf-8") as f:
        text = f.read()
    if old not in text:
        print(f"[MISS] {path}: {old[:40]}")
        continue
    text = text.replace(old, new, 1)
    with io.open(path, "w", encoding="utf-8") as f:
        f.write(text)
    print(f"[OK] {path}: {old[:36]} -> ...")
