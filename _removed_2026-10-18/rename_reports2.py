import glob
import os
import sys

sys.stdout.reconfigure(encoding="utf-8")

# 旧编号 -> 新编号（实用性优先排序）
# 01 冲突 -> 02
# 02 拒答 -> 01
# 03 稳定性 -> 03
# 04 失败样本 -> 04
# 05 溯源 -> 07
# 06 可合成 -> 05
# 07 多轮 -> 06
# 08 反事实 -> 08
mapping = {
    "开题报告01_材料知识冲突的识别与消解_2026-10-18": "开题报告02_材料知识冲突的识别与消解_2026-10-18",
    "开题报告02_材料问答中的选择性拒答_2026-10-18": "开题报告01_材料问答中的选择性拒答_2026-10-18",
    "开题报告03_材料逆向设计的运行间稳定性_2026-10-18": "开题报告03_材料逆向设计的运行间稳定性_2026-10-18",
    "开题报告04_闭环材料发现中的失败样本利用_2026-10-17": "开题报告04_闭环材料发现中的失败样本利用_2026-10-17",
    "开题报告05_材料推荐的知识溯源_2026-10-18": "开题报告07_材料推荐的知识溯源_2026-10-18",
    "开题报告06_可合成性引导的材料生成_2026-10-18": "开题报告05_可合成性引导的材料生成_2026-10-18",
    "开题报告07_材料设计的交互式多轮修正_2026-10-18": "开题报告06_材料设计的交互式多轮修正_2026-10-18",
    "开题报告08_材料逆向设计的反事实敏感性_2026-10-18": "开题报告08_材料逆向设计的反事实敏感性_2026-10-18",
}

tmp = {}
for old, new in mapping.items():
    if old == new:
        continue
    for ext in (".md", ".docx"):
        if os.path.exists(old + ext):
            tmp_path = "_tmp2_" + old + ext
            os.rename(old + ext, tmp_path)
            tmp[tmp_path] = new + ext
            print(f"stage1 {old+ext} -> {tmp_path}")

for src, dst in tmp.items():
    if os.path.exists(dst):
        print(f"[WARN] exists, skip: {dst}")
        continue
    os.rename(src, dst)
    print(f"stage2 {src} -> {dst}")

print("---- 结果 ----")
for f in sorted(glob.glob("开题报告*.md")):
    print(" ", f)
