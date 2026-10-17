import re
import glob
import sys

sys.stdout.reconfigure(encoding="utf-8")

REQUIRED = [
    "摘要", "1 研究背景", "2 国内外研究现状", "3 研究内容", "4 研究方法",
    "5 实验设计", "6 特色与创新", "7 预期成果", "8 进度安排",
    "9 研究基础", "10 可行性", "11 风险", "参考文献",
]

for f in sorted(glob.glob("开题报告0*.md")):
    t = open(f, encoding="utf-8").read()
    issues = []
    for s in REQUIRED:
        if s not in t:
            issues.append(f"缺章节:{s}")
    if "**算法创新" not in t:
        issues.append("无算法创新段")
    refs = set(int(x) for x in re.findall(r"^\[(\d+)\]", t, re.M))
    body = t.split("## 参考文献")[0]
    cited = set()
    for g in re.findall(r"\[(\d+(?:,\s*\d+)*)\]", body):
        cited.update(int(x) for x in g.split(","))
    if refs - cited:
        issues.append(f"未引用:{sorted(refs - cited)}")
    if cited - refs:
        issues.append(f"无条目:{sorted(cited - refs)}")
    for p in ["TBD", "TODO", "待定"]:
        if p in t:
            issues.append(f"占位符:{p}")
    if "查新结论" not in t:
        issues.append("缺查新结论")
    print(f)
    print("   问题:", "; ".join(issues) if issues else "无")
    print("   字符:", len(re.findall(r"[\u4e00-\u9fff]", t)))
