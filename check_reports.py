import re
import glob
import sys
from collections import Counter

sys.stdout.reconfigure(encoding="utf-8")

REQUIRED = [
    "摘要", "1 研究背景", "2 国内外研究现状", "3 研究内容", "4 研究方法",
    "5 实验设计", "6 特色与创新", "7 预期成果", "8 进度安排",
    "9 研究基础", "10 可行性", "11 风险", "参考文献",
]

BLOG_DOMAINS = ["substack.com", "medium.com", "zhihu.com", "csdn.net"]
TRANSITIVE_MARKERS = ["转引自", "转引"]


def check_line(line):
    """对单条参考文献做静态质量检查。"""
    problems = []
    if any(k in line for k in TRANSITIVE_MARKERS):
        problems.append("转引")
    for d in BLOG_DOMAINS:
        if d in line:
            problems.append(f"博客来源:{d}")
    for m in re.finditer(r"arXiv:\s*([^\s（），。；、]+)", line):
        num = m.group(1).rstrip(".,;)]")
        if not re.fullmatch(r"\d{4}\.\d{4,5}(v\d+)?", num):
            problems.append(f"arXiv编号可疑:{num}")
    return problems


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

    heads = re.findall(r"^### (\d+\.\d+)", t, re.M)
    dup = [h for h, c in Counter(heads).items() if c > 1]
    if dup:
        issues.append(f"重复章节号:{dup}")

    refs_block = t.split("## 参考文献")[-1]
    for line in refs_block.splitlines():
        problems = check_line(line)
        if problems:
            issues.append(f"{','.join(problems)}: {line.strip()[:46]}")

    print(f)
    print("   问题:", "; ".join(issues) if issues else "无")
    print("   字符:", len(re.findall(r"[\u4e00-\u9fff]", t)))
