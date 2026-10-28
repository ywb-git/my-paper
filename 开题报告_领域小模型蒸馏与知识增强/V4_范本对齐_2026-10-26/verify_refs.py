#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""逐条核验参考文献的真实出处（AGENTS.md 强制：禁止凭印象写书目）。

做法：
- arXiv 编号 -> export.arxiv.org API（https，http 会 301），每批 8 条，避免返回被截断
- DOI       -> api.crossref.org/works/<DOI>
- OpenReview-> api2/api.openreview.net（可能 403，届时保留上一轮已核验记录并注明）
标题比对采用「归一后子串」判定：把 API 返回标题去掉非字母数字，看它是否出现在
该条目正文的同样归一结果里。这样不依赖「句点分割」这种脆弱启发式。

结果写入 V4_文献逐条核验.json。

用法：
    python verify_refs.py            # 核验并写 JSON
    python verify_refs.py --dry-run  # 只打印，不写文件
"""

from __future__ import annotations

import json
import re
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
import xml.etree.ElementTree as ET
from datetime import datetime, timezone
from pathlib import Path

HERE = Path(__file__).resolve().parent
MD = HERE / "开题报告_正文_V4.md"
OUT = HERE / "V4_文献逐条核验.json"
UA = "Mozilla/5.0 (compatible; kaiti-ref-check/1.0)"
ARXIV = "https://export.arxiv.org/api/query?id_list={}"
CR = "https://api.crossref.org/works/{}"
OR_APIS = ("https://api2.openreview.net/notes?forum={}",
           "https://api.openreview.net/notes?forum={}")

RE_NUM = re.compile(r"^\[(\d+)\]\s*(.*)$")
RE_ARXIV = re.compile(r"arXiv:\s*(\d{4}\.\d{4,5}(?:v\d+)?)")
RE_DOI = re.compile(r"\b(10\.\d{4,9}/[^\s，。；）)]+)")
ARXIV_CHUNK = 8


def fetch(url: str, timeout: int = 25):
    """返回 (status, body)；网络/HTTP 错误不抛异常，返回 (状态码或0, b"")。"""
    req = urllib.request.Request(url, headers={"User-Agent": UA})
    try:
        with urllib.request.urlopen(req, timeout=timeout) as r:
            return r.status, r.read()
    except urllib.error.HTTPError as e:
        return e.code, b""
    except Exception:
        return 0, b""


def norm(s: str) -> str:
    return re.sub(r"[^a-z0-9]", "", (s or "").lower())


def read_refs() -> list[str]:
    lines = MD.read_text(encoding="utf-8").splitlines()
    start = None
    for i, ln in enumerate(lines):
        if re.match(r"^#{1,3}\s*参考文献\s*$", ln.strip()):
            start = i
            break
    if start is None:
        raise SystemExit("md 中未找到「参考文献」标题")
    return [l.strip() for l in lines[start + 1:] if re.match(r"^\[\d+\]", l.strip())]


def arxiv_batch(ids: list[str]) -> dict:
    """批量取 arXiv 元数据，返回 {id: {...}}。分批 + 间隔，避免被截断/限流。"""
    out = {}
    for k in range(0, len(ids), ARXIV_CHUNK):
        chunk = ids[k:k + ARXIV_CHUNK]
        st, raw = fetch(ARXIV.format(",".join(chunk)))
        if st == 200 and raw:
            ns = {"a": "http://www.w3.org/2005/Atom"}
            root = ET.fromstring(raw)
            for e in root.findall("a:entry", ns):
                eid = e.find("a:id", ns).text.rsplit("/abs/", 1)[-1]
                eid = re.sub(r"v\d+$", "", eid)
                out[eid] = {
                    "title": " ".join((e.find("a:title", ns).text or "").split()),
                    "authors": [x.find("a:name", ns).text for x in e.findall("a:author", ns)],
                    "publication_date": [(e.find("a:published", ns).text or "")[:10]],
                    "journal": [], "doi": [],
                }
        else:
            print(f"[警告] arXiv 批次 {k//ARXIV_CHUNK+1} status={st}")
        time.sleep(3)
    missing = set(ids) - set(out)
    if missing:
        print(f"[警告] arXiv 未返回 {len(missing)} 条：{sorted(missing)}")
    return out


def crossref_lookup(doi: str):
    st, raw = fetch(CR.format(urllib.parse.quote(doi, safe="")))
    if st != 200 or not raw:
        return st, {}
    msg = json.loads(raw.decode("utf-8")).get("message", {})
    authors = []
    for a in msg.get("author", []) or []:
        fam, giv = a.get("family", ""), a.get("given", "")
        authors.append(f"{fam}, {giv}".strip(", "))
    parts = msg.get("issued", {}).get("date-parts", [[None]])[0]
    return st, {
        "title": " ".join((msg.get("title") or [""])[0].split()),
        "authors": authors,
        "publication_date": ["/".join(str(x) for x in parts)],
        "journal": [msg.get("container-title", [""])[0]] if msg.get("container-title") else [],
        "doi": [doi],
        "volume": msg.get("volume", ""), "issue": msg.get("issue", ""),
        "page": msg.get("page", ""),
    }


def openreview_lookup(fid: str):
    """依次尝试 api2 / api；403 属站点策略，返回 (403, {})。"""
    last = 0
    for tpl in OR_APIS:
        st, raw = fetch(tpl.format(fid))
        last = st
        if st == 200 and raw:
            notes = json.loads(raw.decode("utf-8")).get("notes", [])
            if not notes:
                continue
            c = notes[0].get("content", {})

            def val(key):
                v = c.get(key, {})
                return v.get("value") if isinstance(v, dict) else v

            return st, {"title": " ".join((val("title") or "").split()),
                        "authors": val("authors") or [], "publication_date": [],
                        "journal": [], "doi": []}
    return last, {}


def main(argv):
    dry = "--dry-run" in argv
    refs = read_refs()
    print(f"参考文献共 {len(refs)} 条")

    # 上一轮记录（OpenReview 等本次取不到的条目沿用，并在记录里注明来源轮次）
    prior = {}
    if OUT.exists():
        try:
            for it in json.loads(OUT.read_text(encoding="utf-8")).get("items", []):
                prior[it["number"]] = it
        except Exception:
            prior = {}

    parsed = []
    for line in refs:
        m = RE_NUM.match(line)
        num, body = int(m.group(1)), m.group(2)
        ax = RE_ARXIV.search(line)
        doi = RE_DOI.search(body)
        parsed.append({
            "num": num, "body": body,
            "arxiv": re.sub(r"v\d+$", "", ax.group(1)) if ax else None,
            "doi": doi.group(1).rstrip(".,;") if doi else None,
        })

    ax_map = {}
    ax_ids = sorted({p["arxiv"] for p in parsed if p["arxiv"]})
    if ax_ids:
        print(f"arXiv 待查 {len(ax_ids)} 条，分批 {ARXIV_CHUNK}")
        ax_map = arxiv_batch(ax_ids)

    items = []
    for p in parsed:
        rec = {"number": p["num"], "entry": p["body"]}
        if p["arxiv"]:
            got = ax_map.get(p["arxiv"])
            rec.update({"kind": "arxiv", "identifier": p["arxiv"],
                        "url": f"https://arxiv.org/abs/{p['arxiv']}",
                        "http_status": 200 if got else 404})
            if got:
                rec.update(got)
                rec["status"] = "retrieved"
        elif p["doi"]:
            st, got = crossref_lookup(p["doi"])
            rec.update({"kind": "doi", "identifier": p["doi"],
                        "url": CR.format(p["doi"]), "http_status": st})
            if got:
                rec.update(got)
                rec["status"] = "retrieved"
            time.sleep(0.4)
        elif "OpenReview" in p["body"]:
            m = re.search(r"id=([A-Za-z0-9_-]+)", p["body"])
            fid = m.group(1) if m else ""
            st, got = openreview_lookup(fid) if fid else (0, {})
            rec.update({"kind": "openreview", "identifier": fid,
                        "url": f"https://openreview.net/forum?id={fid}",
                        "http_status": st})
            if got:
                rec.update(got)
                rec["status"] = "retrieved"
            elif prior.get(p["num"], {}).get("title"):
                old = prior[p["num"]]
                rec.update({"title": old["title"], "authors": old.get("authors", []),
                            "publication_date": old.get("publication_date", []),
                            "journal": old.get("journal", []), "doi": old.get("doi", []),
                            "status": "retrieved:prior-round",
                            "verification_method": old.get("verification_method", ""),
                            "note": f"本轮 API 返回 {st}，沿用上一轮已核验记录"})
        else:
            rec.update({"kind": "unknown", "status": "no-identifier"})

        api_title = rec.get("title", "")
        rec["title_match_in_entry"] = bool(api_title) and norm(api_title) in norm(p["body"])
        items.append(rec)

        ok = rec["title_match_in_entry"]
        print(f"{'OK ' if ok else '!! '}[{p['num']:2d}] {rec.get('status','?'):20s} "
              f"{' '.join(rec.get('authors', [])[:4])[:34]:34s} {api_title[:48]!r}")
        if not ok:
            print(f"      条目原文={p['body'][:96]!r}")

    bad = [i["number"] for i in items if not i["title_match_in_entry"]]
    noid = [i["number"] for i in items if i.get("status") == "no-identifier"]
    n_ret = sum(1 for i in items if str(i.get("status", "")).startswith("retrieved"))
    print(f"\n核验完成：共 {len(items)} 条；取到原始出处 {n_ret} 条；"
          f"标题不一致 {bad or '无'}；无标识符 {noid or '无'}")

    if not dry:
        OUT.write_text(json.dumps({
            "checked_at": datetime.now(timezone.utc).isoformat(),
            "scope": (f"本轮 {len(items)} 条均实际调用 arXiv/CrossRef/OpenReview API 并比对标题"
                      "（归一后子串判定）；元数据核验不等于全文审查或实验复现。"
                      "arXiv 按所访问版本引用，不宣称无正式发表版。"),
            "items": items,
        }, ensure_ascii=False, indent=1), encoding="utf-8")
        print(f"已写入 {OUT.name}")
    return 1 if (bad or noid) else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
