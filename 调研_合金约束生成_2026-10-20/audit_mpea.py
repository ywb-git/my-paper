#!/usr/bin/env python3
"""审计公开 MPEA CSV；仅统计，不训练预测器或验证新合金。"""
import argparse
import csv
import hashlib
import json
import math
import re
from collections import Counter, defaultdict
from pathlib import Path

SOURCE_COMMIT = "d31e2510cccc3fb9e79d43ad62652f64ef8c2c5e"
SOURCE_URL = f"https://raw.githubusercontent.com/CitrineInformatics/MPEA_dataset/{SOURCE_COMMIT}/MPEA_dataset.csv"
EXPECTED_SHA256 = "c5504d93fd324d1be26cf814d0694ee9ee95578d68a0b38aba6813b939fa2c5d"
KEYS = {
    "formula": "FORMULA", "doi": "REFERENCE: doi",
    "process": "PROPERTY: Processing method", "test": "PROPERTY: Type of test",
    "temperature": "PROPERTY: Test temperature ($^\\circ$C)",
    "ys": "PROPERTY: YS (MPa)", "uts": "PROPERTY: UTS (MPa)", "hv": "PROPERTY: HV",
    "elongation": "PROPERTY: Elongation (%)", "grain": "PROPERTY: grain size ($\\mu$m)",
    "density_exp": "PROPERTY: Exp. Density (g/cm$^3$)",
}


def number(value):
    try:
        x = float(value)
        return x if math.isfinite(x) else None
    except (ValueError, TypeError):
        return None


def composition(formula):
    compact = re.sub(r"\s+", "", formula)
    tokens = list(re.finditer(r"([A-Z][a-z]?)(\d+(?:\.\d*)?|\.\d+)?", compact))
    if not tokens or "".join(t.group(0) for t in tokens) != compact:
        raise ValueError(formula)
    values = defaultdict(float)
    for token in tokens:
        values[token.group(1)] += float(token.group(2) or 1)
    total = sum(values.values())
    if total <= 0 or any(v <= 0 for v in values.values()):
        raise ValueError(formula)
    return {e: v / total for e, v in values.items()}


def signature(comp):
    return tuple(sorted((e, round(v, 8)) for e, v in comp.items()))


def summary(rows):
    return {
        "records": len(rows),
        "unique_compositions": len({signature(r["composition"]) for r in rows}),
        "unique_doi": len({r["doi"] for r in rows if r["doi"]}),
        "process_counts": dict(Counter(r["process"] for r in rows)),
        "test_type_counts": dict(Counter(r["test"] for r in rows)),
        "with_ys": sum(r["ys"] is not None for r in rows),
        "with_ys_and_elongation": sum(r["ys"] is not None and r["elongation"] is not None for r in rows),
        "with_grain_size": sum(r["grain"] is not None for r in rows),
    }


def audit(path):
    data = path.read_bytes()
    digest = hashlib.sha256(data).hexdigest()
    assert digest == EXPECTED_SHA256, "源版本校验失败，不能套用已有统计"
    with path.open(encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        fields = reader.fieldnames
        assert set(KEYS.values()).issubset(fields)
        raw = list(reader)
    rows, failures = [], []
    for index, record in enumerate(raw, 2):
        try:
            comp = composition(record[KEYS["formula"]])
        except ValueError:
            failures.append({"csv_line": index, "formula": record[KEYS["formula"]]})
            continue
        row = {key: record[column].strip() for key, column in KEYS.items()}
        row["composition"] = comp
        for key in ("temperature", "ys", "uts", "elongation", "grain", "density_exp", "hv"):
            row[key] = number(row[key])
        rows.append(row)
    room = [r for r in rows if r["temperature"] is not None and 20 <= r["temperature"] <= 30]
    tensile = [r for r in room if r["test"] == "T" and r["ys"] is not None]
    narrow = [r for r in tensile if set(r["composition"]) <= {"Al", "Co", "Cr", "Fe", "Mn", "Ni"}]
    pairs = []
    for i, a in enumerate(tensile):
        for b in tensile[i+1:]:
            if not a["doi"] or a["doi"] != b["doi"] or a["process"] != b["process"] or a["temperature"] != b["temperature"]:
                continue
            distance = sum(abs(a["composition"].get(e, 0) - b["composition"].get(e, 0)) for e in a["composition"].keys() | b["composition"].keys())
            if 1e-8 < distance <= .20:
                pairs.append((a["doi"], a["formula"], b["formula"]))
    condition_groups = defaultdict(list)
    for r in tensile:
        condition_groups[r["process"]].append(r)
    return {
        "source": {"url": SOURCE_URL, "commit": SOURCE_COMMIT, "bytes": len(data), "sha256": digest},
        "definitions": {
            "room_temperature": "20–30摄氏度，含端点；以数据原值为准",
            "tensile_ys": "test=T且YS是有限数字；不使用压缩YS或硬度换算",
            "composition_signature": "原子比例归一化后按元素排序、8位小数；不将工艺差异去重",
            "local_pair": "同DOI、同粗工艺字符串、同测试温度且不同成分，原子分数L1距离≤0.20；仅可用性统计，不证明工艺或晶粒相同",
        },
        "raw_records": len(raw), "columns": fields, "formula_parse_failures": failures,
        "all_parsed": summary(rows), "room_temperature": summary(room),
        "room_tensile_ys": summary(tensile),
        "room_cast_hardness": summary([r for r in room if r["process"] == "CAST" and r["hv"] is not None]),
        "room_tensile_ys_by_process": {key: summary(value) for key, value in sorted(condition_groups.items())},
        "room_tensile_ys_AlCoCrFeMnNi_subset": summary(narrow),
        "local_pair_count": len(pairs), "local_pair_unique_doi": len({p[0] for p in pairs}),
        "local_pair_examples": pairs[:8],
        "boundaries": ["粗工艺标签不等于固定热处理、冷加工量或晶粒状态", "记录数和派生配对数不等于独立实验数", "未验证预测器或任何新候选的真实力学性能"],
    }


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("input", type=Path)
    parser.add_argument("output", type=Path)
    args = parser.parse_args()
    assert composition("Al0.3 Co1 Cr1 Fe1 Ni1")["Al"] == .3 / 4.3
    assert signature(composition("Fe1 Ni1")) == signature(composition("Ni50 Fe50"))
    assert number("nan") is None
    result = audit(args.input)
    args.output.write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({key: result[key] for key in ("raw_records", "formula_parse_failures", "room_tensile_ys", "room_tensile_ys_by_process", "room_tensile_ys_AlCoCrFeMnNi_subset", "local_pair_count", "local_pair_unique_doi")}, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
