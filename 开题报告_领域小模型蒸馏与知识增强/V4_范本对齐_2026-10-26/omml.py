#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""极简 LaTeX 子集 -> OMML（Word 原生公式）转换。

只实现本报告实际用到的语法，遇到不认识的命令**直接抛错**，不做静默降级：
静默把一个公式渲染成看似合理的错误形式，比报错更糟。

支持：字母/数字/运算符、``_`` ``^`` 上下标（含 {..} 分组）、``\\sum`` 带下限、
``\\frac{}{}``、``\\left( .. \\right)`` 定界符、``\\max \\log \\exp`` 等函数名、
若干希腊字母与关系符、``\\mathrm{..}`` 直立文字、``'`` 撇号。

对外接口：
    latex_to_omath(latex: str) -> lxml.etree._Element   # <m:oMath>
"""

from __future__ import annotations

import re

from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# ---------------------------------------------------------------- 词法

GREEK = {
    "alpha": "α", "beta": "β", "gamma": "γ", "delta": "δ", "epsilon": "ε",
    "varepsilon": "ε", "zeta": "ζ", "eta": "η", "theta": "θ", "lambda": "λ",
    "mu": "μ", "nu": "ν", "pi": "π", "rho": "ρ", "sigma": "σ", "tau": "τ",
    "phi": "φ", "varphi": "φ", "chi": "χ", "psi": "ψ", "omega": "ω",
    "Gamma": "Γ", "Delta": "Δ", "Theta": "Θ", "Lambda": "Λ", "Pi": "Π",
    "Sigma": "Σ", "Phi": "Φ", "Omega": "Ω",
    # 数学字母（本身直立语义但按数学斜体渲染，与 Word 默认一致）
    "ell": "ℓ", "hbar": "ℏ",
}

SYMBOL = {
    "le": "≤", "leq": "≤", "ge": "≥", "geq": "≥", "ne": "≠", "neq": "≠",
    "in": "∈", "notin": "∉", "cup": "∪", "cap": "∩", "cdot": "·",
    "times": "×", "pm": "±", "forall": "∀", "exists": "∃", "mid": "|",
    "to": "→", "rightarrow": "→", "leftarrow": "←", "infty": "∞",
    "partial": "∂", "approx": "≈", "sim": "∼", "propto": "∝",
    "subset": "⊂", "subseteq": "⊆", "ldots": "…", "cdots": "⋯",
}

SPACE = {",": "\u2009", ";": "\u2005", "!": "", " ": " ",
         "quad": "\u2003", "qquad": "\u2003\u2003"}

FUNCS = {"max", "min", "log", "ln", "exp", "arg", "sin", "cos", "lim",
         "sup", "inf", "det", "argmax", "argmin"}

NARY = {"sum": "∑", "prod": "∏", "int": "∫"}

DELIMS = {"{": "{", "}": "}", "(": "(", ")": ")", "[": "[", "]": "]",
          "|": "|", "\\{": "{", "\\}": "}", "\\|": "‖", ".": ""}

RE_CMD = re.compile(r"\\[a-zA-Z]+")
RE_NUM = re.compile(r"\d+(?:\.\d+)?")


class LatexError(ValueError):
    pass


def tokenize(s: str):
    """切成 (kind, value) 序列。kind ∈ cmd/num/chr。"""
    toks = []
    i = 0
    while i < len(s):
        c = s[i]
        if c == "\\":
            m = RE_CMD.match(s, i)
            if m:
                toks.append(("cmd", m.group(0)[1:]))   # 去掉反斜杠
                i = m.end()
            else:
                toks.append(("cmd", s[i + 1]))          # \, \; \{ \} 等单字符命令
                i += 2
            continue
        m = RE_NUM.match(s, i)
        if m:
            toks.append(("num", m.group(0)))
            i = m.end()
            continue
        if c in "{}_^":
            toks.append(("chr", c))
            i += 1
            continue
        if c == " ":
            i += 1
            continue
        toks.append(("chr", c))
        i += 1
    return toks


# ---------------------------------------------------------------- 语法树


def parse(tokens):
    """返回节点列表。节点是 dict：{'t': 类型, ...}。"""
    pos = 0

    def peek(off=0):
        return tokens[pos + off] if pos + off < len(tokens) else (None, None)

    def take():
        nonlocal pos
        t = tokens[pos]
        pos += 1
        return t

    def parse_group():
        """解析一个 {..} 或单个原子，作为上下标/分母的内容。"""
        nonlocal pos
        k, v = peek()
        if (k, v) == ("chr", "{"):
            take()
            body = parse_seq(stop={"}"})
            if peek() != ("chr", "}"):
                raise LatexError("花括号未闭合")
            take()
            return body
        atom = parse_atom()
        return [atom] if atom else []

    def parse_scripts(base):
        sub = sup = None
        while peek()[0] == "chr" and peek()[1] in ("_", "^"):
            k, v = take()
            g = parse_group()
            if v == "_":
                sub = g
            else:
                sup = g
        if sub is None and sup is None:
            return base
        if base.get("t") == "nary":
            if sub is not None:
                base["sub"] = sub
            if sup is not None:
                base["sup"] = sup
            return base
        kind = "subsup" if (sub is not None and sup is not None) else ("sub" if sub else "sup")
        return {"t": kind, "base": [base], "sub": sub, "sup": sup}

    def parse_atom():
        k, v = peek()
        if k is None:
            return None
        if (k, v) == ("chr", "{"):
            return {"t": "grp", "body": parse_group()}
        if k == "cmd":
            take()
            if v in NARY:
                return {"t": "nary", "chr": NARY[v], "sub": None, "sup": None}
            if v == "frac":
                num = parse_group()
                den = parse_group()
                return {"t": "frac", "num": num, "den": den}
            if v == "left":
                beg = take()
                begch = DELIMS.get("\\" + beg[1] if beg[0] == "cmd" else beg[1])
                if begch is None:
                    raise LatexError(f"不支持的左定界符 {beg!r}")
                body = parse_seq(stop={"right"})
                if peek()[0] != "cmd" or peek()[1] != "right":
                    raise LatexError("\\left 缺少配对的 \\right")
                take()
                end = take()
                endch = DELIMS.get("\\" + end[1] if end[0] == "cmd" else end[1])
                if endch is None:
                    raise LatexError(f"不支持的右定界符 {end!r}")
                return {"t": "d", "beg": begch, "end": endch, "body": body}
            if v == "mathrm":
                body = parse_group()
                return {"t": "up", "body": body}
            if v == "text":
                body = parse_group()
                return {"t": "up", "body": body}
            if v in FUNCS:
                return {"t": "fn", "name": v}
            if v in GREEK:
                return {"t": "run", "text": GREEK[v]}
            if v in SYMBOL:
                return {"t": "run", "text": SYMBOL[v]}
            if v in SPACE:
                return {"t": "run", "text": SPACE[v]} if SPACE[v] else None
            # 单字符命令：\, \; \{ \} \| \（空格）等
            if ("\\" + v) in SPACE:
                sp = SPACE["\\" + v]
                return {"t": "run", "text": sp} if sp else None
            if ("\\" + v) in DELIMS:
                return {"t": "run", "text": DELIMS["\\" + v]}
            raise LatexError(f"不支持的 LaTeX 命令 \\{v}")
        take()
        if k == "num":
            return {"t": "run", "text": v}
        if v == "-":
            return {"t": "run", "text": "−"}
        return {"t": "run", "text": v}

    def parse_seq(stop=frozenset()):
        out = []
        while True:
            k, v = peek()
            if k is None:
                break
            if k == "chr" and v in stop:
                break
            if k == "cmd" and v in stop:
                break
            atom = parse_atom()
            if atom is None:
                continue
            # 撇号并入前一个 run，避免上下标挂到撇号上
            while peek() == ("chr", "'") and atom.get("t") == "run":
                take()
                atom["text"] += "′"
            atom = parse_scripts(atom)
            # 求和号的被加式取“其后整个序列”：OMML 的 m:nary 必须有 m:e，
            # 留空会被渲染成一个占位方框（实测 LibreOffice 就是如此）。
            if atom.get("t") == "nary" and atom.get("body") is None:
                atom["body"] = parse_seq(stop=stop)
            out.append(atom)
        return out

    nodes = parse_seq()
    if pos != len(tokens):
        raise LatexError(f"解析在 {tokens[pos:]!r} 处中断（可能有未配对的花括号）")
    return nodes


# ---------------------------------------------------------------- OMML 生成


def _el(tag, **attrs):
    e = OxmlElement(tag)
    for k, v in attrs.items():
        e.set(qn(k), v)
    return e


def _run(text: str, upright=False):
    r = _el("m:r")
    if upright:
        rpr = _el("m:rPr")
        sty = _el("m:sty")
        sty.set(qn("m:val"), "p")
        rpr.append(sty)
        r.append(rpr)
    t = _el("m:t")
    t.text = text
    r.append(t)
    return r


def _wrap(tag, nodes, upright_ctx=False):
    """生成 <m:tag><m:e>nodes</m:e></m:tag> 结构。"""
    e = _el(tag)
    inner = _el("m:e")
    _emit_into(inner, nodes, upright_ctx)
    e.append(inner)
    return e


def _emit_into(parent, nodes, upright_ctx=False):
    for n in nodes or []:
        _emit(parent, n, upright_ctx)


def _slot(parent, tag, nodes, upright_ctx=False):
    s = _el(tag)
    _emit_into(s, nodes, upright_ctx)
    parent.append(s)
    return s


def _emit(parent, n, upright_ctx=False):
    t = n["t"]
    if t == "run":
        parent.append(_run(n["text"], upright=upright_ctx))
    elif t == "grp":
        _emit_into(parent, n["body"], upright_ctx)
    elif t == "sub":
        el = _el("m:sSub")
        _slot(el, "m:e", n["base"], upright_ctx)
        _slot(el, "m:sub", n["sub"], upright_ctx)
        parent.append(el)
    elif t == "sup":
        el = _el("m:sSup")
        _slot(el, "m:e", n["base"], upright_ctx)
        _slot(el, "m:sup", n["sup"], upright_ctx)
        parent.append(el)
    elif t == "subsup":
        el = _el("m:sSubSup")
        _slot(el, "m:e", n["base"], upright_ctx)
        _slot(el, "m:sub", n["sub"], upright_ctx)
        _slot(el, "m:sup", n["sup"], upright_ctx)
        parent.append(el)
    elif t == "nary":
        el = _el("m:nary")
        pr = _el("m:naryPr")
        chr_ = _el("m:chr"); chr_.set(qn("m:val"), n["chr"]); pr.append(chr_)
        loc = _el("m:limLoc"); loc.set(qn("m:val"), "undOvr"); pr.append(loc)
        if n.get("sup") is None:
            hide = _el("m:supHide"); hide.set(qn("m:val"), "1"); pr.append(hide)
        el.append(pr)
        _slot(el, "m:sub", n.get("sub") or [], upright_ctx)
        _slot(el, "m:sup", n.get("sup") or [], upright_ctx)
        _slot(el, "m:e", n.get("body") or [], upright_ctx)
        parent.append(el)
    elif t == "frac":
        el = _el("m:f")
        _slot(el, "m:num", n["num"], upright_ctx)
        _slot(el, "m:den", n["den"], upright_ctx)
        parent.append(el)
    elif t == "d":
        el = _el("m:d")
        pr = _el("m:dPr")
        if n["beg"]:
            b = _el("m:begChr"); b.set(qn("m:val"), n["beg"]); pr.append(b)
        if n["end"]:
            e = _el("m:endChr"); e.set(qn("m:val"), n["end"]); pr.append(e)
        el.append(pr)
        _slot(el, "m:e", n["body"], upright_ctx)
        parent.append(el)
    elif t == "fn":
        parent.append(_run(n["name"], upright=True))
    elif t == "up":
        _emit_into(parent, n["body"], upright_ctx=True)
    else:
        raise LatexError(f"未知节点类型 {t}")


def latex_to_omath(latex: str):
    """LaTeX 子集 -> <m:oMath> 元素（不抛 MX 之外的异常类型）。"""
    nodes = parse(tokenize(latex))
    om = _el("m:oMath")
    _emit_into(om, nodes)
    return om


# ---------------------------------------------------------------- 自检

if __name__ == "__main__":
    from lxml import etree

    CASES = {
        "1": r"u_{i\ell j}=h_{i\ell j}\cdot s_{\varphi}(q,j,i\ell)",
        "2": r"\max F(x)=\sum_{j}w_{j}\max\left(\{u_{i\ell j}:x_{i\ell}=1\}\cup\{0\}\right)",
        "3": r"\sum_{\ell}x_{i\ell}\le 1;\quad x_{i\ell}\le z_{a},\ \forall a\in D_{i\ell}",
        "4": r"\sum_{i\ell}c_{i\ell}x_{i\ell}+\sum_{a}d_{a}z_{a}\le B",
        "5": r"\max G(A)=\sum_{i}v_{i}p_{i}\max\left(\{r_{ai}:a\in A\}\cup\{0\}\right)",
        "6": r"\sum_{a\in A}t_{a}\le H",
        "7": r"L=L_{\mathrm{SFT}}+\lambda_{r}L_{\mathrm{rank}}+\lambda_{c}L_{\mathrm{cond}}",
        "8": r"L_{\mathrm{SFT}}(\theta)=-\sum_{(q,E,y)}\sum_{t}\log\pi_{\theta}(y_{t}\mid q,E,y_{<t})",
        "9": (r"L_{\mathrm{rank}}=-\log\frac{\exp\left(s_{\theta}(q,d^{+})/\tau\right)}"
              r"{\sum_{d\in D(q)}\exp\left(s_{\theta}(q,d)/\tau\right)}"),
        "10": (r"L_{\mathrm{cond}}=\max\left(0,m-s_{\theta}(q,d^{+})+s_{\theta}(q,d'^{+})\right)"
               r"+\max\left(0,m-s_{\theta}(q',d'^{+})+s_{\theta}(q',d^{+})\right)"),
    }
    ok = 0
    for k, v in CASES.items():
        el = latex_to_omath(v)
        xml = etree.tostring(el, pretty_print=False).decode()
        # 基本自检：必须含 oMath，且无未转义的 & 或 "<" 破坏 XML
        assert xml.startswith("<m:oMath"), xml[:80]
        assert "None" not in xml
        print(f"式({k}) OK  len={len(xml):5d}  {xml[:88]}...")
        ok += 1
    print(f"\n自检通过 {ok}/{len(CASES)}")
