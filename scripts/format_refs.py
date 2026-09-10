# -*- coding: utf-8 -*-
"""
format_refs.py — 引用条目格式化（论文引用 skill）

按目标格式把 refs.csv 中的一条记录渲染成参考文献文本：
  gbt7714   GB/T 7714-2015 顺序编码制（中文期刊/毕业论文最常见）
  apa       APA 第 7 版
  vancouver Vancouver / ICMJE 编号式（多数生物医学期刊）
  mla       MLA 第 9 版

用法（脚本内调用）：
    from format_refs import format_ref, gbt7714, apa, vancouver, mla
    text = format_ref(ref_dict, style="gbt7714")
"""

import re
from datetime import datetime


# ---------- 作者串处理 ----------

def _split_authors(raw):
    """拆作者串。中文按顿号/逗号；英文 Google Scholar 风格
    'Vaswani, A., Shazeer, N.' -> ['Vaswani, A.', 'Shazeer, N.']。
    尾部 et al. 剔除（截断由格式函数处理）。"""
    if not raw:
        return []
    raw = re.sub(r"(?i)\s*et\s+al\.?$", "", raw.strip())
    parts = re.split(r"[;；]", raw)
    out = []
    for part in parts:
        part = part.strip()
        if not part:
            continue
        if _is_cjk(part):
            out.extend(p.strip() for p in re.split(r"[、，,]", part) if p.strip())
        else:
            # 英文：按 ", " 拆，缩写片段（'A.'、'M.-W.'、'A. N.'）并入前一个作者
            segs = [s.strip() for s in part.split(",")]
            buf = segs[0]
            for s in segs[1:]:
                if not s:
                    continue
                if (re.match(r"^([A-Z]\.\s*)+$", s)          # A. / A. N.
                        or re.match(r"^[A-Z]\.-[A-Z]\.$", s)  # M.-W.
                        or re.match(r"^[A-Z][a-z]*\.$", s)):  # Jr. / Ed.
                    buf = buf + ", " + s
                else:
                    if buf:
                        out.append(buf)
                    buf = s
            if buf:
                out.append(buf)
    return [a.rstrip(".") for a in out if a]


def _is_cjk(s):
    return bool(re.search(r"[\u4e00-\u9fff]", s))


def _first_author_surname(a):
    """取首作者姓（用于 APA 等）。英文 'Smith, J.' -> 'Smith'；'J. Smith' -> 'Smith'"""
    a = a.strip()
    if "," in a:
        return a.split(",")[0].strip()
    words = a.split()
    return words[-1] if words else a


# ---------- 三作者截断（GB/T 7714：前 3 + 等/et al.）----------

def _gb_authors(raw):
    lst = _split_authors(raw)
    if not lst:
        return ""
    head = [_gb_single(a) for a in lst[:3]]
    if len(lst) <= 3:
        return ", ".join(head)
    # 截断符号按文献主语言（首作者语种）判断：中文用"等"，其他用 et al.
    tail = "等" if _is_cjk(lst[0]) else "et al."
    return ", ".join(head) + ", " + tail


def _gb_single(a):
    """英文 'Smith, J.' -> 'SMITH J'；中文原样"""
    a = a.strip()
    if _is_cjk(a):
        return a
    if "," in a:
        surname, given = a.split(",", 1)
        given_initials = "".join(re.findall(r"[A-Z]", given.upper())) or ""
        return f"{surname.strip().upper()} {given_initials}".strip()
    words = a.split()
    if len(words) == 1:
        return words[0].upper()
    surname = words[-1].upper()
    initials = "".join(w[0].upper() for w in words[:-1])
    return f"{surname} {initials}".strip()


def _gb_authors_full(raw):
    lst = _split_authors(raw)
    return ", ".join(_gb_single(a) for a in lst)


# ---------- APA ----------

def _apa_authors(raw):
    lst = _split_authors(raw)
    if not lst:
        return ""
    if len(lst) == 1:
        return _apa_single(lst[0])
    if len(lst) == 2:
        return f"{_apa_single(lst[0])}, & {_apa_single(lst[1])}"
    head = [_apa_single(a) for a in lst[: len(lst) - 1]]
    return f"{', '.join(head)}, & {_apa_single(lst[-1])}"


def _apa_single(a):
    a = a.strip()
    if _is_cjk(a):
        return a
    if "," in a:
        surname, given = a.split(",", 1)
        given = given.strip()
        if given and not given.endswith("."):
            given += "."
        return f"{surname.strip()}, {given}".strip(", ") if given else surname.strip()
    words = a.split()
    if len(words) == 1:
        return a
    return f"{words[-1]}, {''.join(w[0].upper() + '.' for w in words[:-1])}"


# ---------- Vancouver ----------

def _van_authors(raw, max_a=6):
    lst = _split_authors(raw)
    if not lst:
        return ""
    head = [_gb_single(a) for a in lst[:max_a]]
    if len(lst) > max_a:
        head.append("et al.")
    return ", ".join(head)


# ---------- MLA ----------

def _mla_authors(raw):
    """MLA：作者保持 '姓, 名缩写' 形式，末位前加 and。"""
    lst = _split_authors(raw)
    if not lst:
        return ""
    parts = [_apa_single(a) for a in lst]
    if len(parts) == 1:
        return parts[0]
    return ", ".join(parts[:-1]) + ", and " + parts[-1]


def _mla_single_first(a):
    a = a.strip()
    if _is_cjk(a):
        return a
    if "," in a:
        return a
    words = a.split()
    if len(words) <= 1:
        return a
    return f"{words[-1]}, {words[0]}"


def _mla_single_rest(a):
    a = a.strip()
    if _is_cjk(a):
        return a
    if "," in a:
        surname, given = a.split(",", 1)
        return given.strip() + " " + surname.strip()
    return a


# ---------- 主格式化 ----------

def _type_tag(t):
    return {"journal": "J", "conference": "C", "book": "M",
            "thesis": "D", "web": "EB/OL"}.get(t, "J")


def gbt7714(r):
    """GB/T 7714-2015 顺序编码制。
    支持 city（出版地/保存地）：专著[M]/学位论文[D]/会议[C] 输出 出版地: 出版者；
    电子资源[EB/OL] 输出 (发布日期)[引用日期]. URL（引用日期自动取当天）。
    斜体标记：西文期刊名/书名用 *...* 包裹（GB/T 7714 规范：西文刊名斜体、
    中文刊名正体）；format_ref_segments 会据此拆分为斜体 run。"""
    authors = _gb_authors(r.get("authors", ""))
    title = r.get("title", "").strip()
    src = r.get("source", "").strip()
    year = r.get("year", "").strip()
    vol = r.get("volume", "").strip()
    iss = r.get("issue", "").strip()
    pages = r.get("pages", "").strip()
    doi = r.get("doi", "").strip()
    city = r.get("city", "").strip()
    url = r.get("url", "").strip()
    ttype = _type_tag(r.get("type", "journal"))
    # 西文出处名斜体（GB/T 7714：西文刊名/书名斜体；中文用正体）
    src_it = f"*{src}*" if src and not _is_cjk(src) else src
    # 作者段与题名之间：若作者已以“等/et al.”结尾则只留一个空格
    sep = " " if authors.endswith(("等", "et al.")) else ". "
    base = f"{authors}{sep}{title}[{ttype}]"
    if ttype == "EB/OL":
        # [EB/OL]. (发布日期)[引用日期]. 获取和访问路径. DOI.
        pub = year or ""
        today = datetime.now().strftime("%Y-%m-%d")
        base += f". ({pub})[{today}]"
        if url:
            base += f". {url}"
        elif src:
            base += f". {src}"
        base += "."
        if doi:
            base += f" https://doi.org/{doi}."
        return base
    if ttype in ("M", "D", "C"):
        # 专著/学位论文/会议录：出版地: 出版者, 年（city 缺省则省略出版地）
        place = f"{city}: " if city else ""
        base += f". {place}{src_it}, {year}."
        if doi:
            base += f" https://doi.org/{doi}."
        return base
    # 期刊等：期刊名, 年, 卷(期): 页码
    vp = ""
    if vol:
        vp = f", {vol}"
        if iss:
            vp += f"({iss})"
        if pages:
            vp += f": {pages}"
    base += f". {src_it}, {year}{vp}."
    if doi:
        base += f" https://doi.org/{doi}."
    return base


def apa(r):
    """APA 第 7 版。"""
    authors = _apa_authors(r.get("authors", ""))
    year = r.get("year", "").strip()
    title = r.get("title", "").strip()
    src = r.get("source", "").strip()
    vol = r.get("volume", "").strip()
    iss = r.get("issue", "").strip()
    pages = r.get("pages", "").strip()
    doi = r.get("doi", "").strip()
    ttype = r.get("type", "journal")
    if ttype == "book":
        return f"{authors} ({year}). *{title}*. {src}."
    head = f"{authors} ({year}). {title}."
    if src:
        it = f"*{src}*"
        if vol:
            it += f", *{vol}*"
            if iss:
                it += f"({iss})"
            if pages:
                it += f", {pages}"
        head += " " + it + "."
    if doi:
        head += f" https://doi.org/{doi}"
    return head


def vancouver(r):
    """Vancouver / ICMJE 编号式（西文期刊名斜体）。"""
    authors = _van_authors(r.get("authors", ""))
    title = r.get("title", "").strip()
    src = r.get("source", "").strip()
    year = r.get("year", "").strip()
    vol = r.get("volume", "").strip()
    iss = r.get("issue", "").strip()
    pages = r.get("pages", "").strip()
    doi = r.get("doi", "").strip()
    src_it = f"*{src}*" if src and not _is_cjk(src) else src
    seg = []
    if authors:
        seg.append(authors.rstrip(".") + ".")
    if title:
        seg.append(f"{title}.")
    if src_it:
        seg.append(f"{src_it}.")
    tail = f"{year}"
    if vol:
        tail += f";{vol}"
        if iss:
            tail += f"({iss})"
        if pages:
            tail += f":{pages}"
    seg.append(tail + ".")
    if doi:
        seg.append(f"doi: {doi}")
    return " ".join(seg)


def mla(r):
    """MLA 第 9 版（期刊文章，期刊名斜体）。"""
    authors = _mla_authors(r.get("authors", ""))
    title = r.get("title", "").strip()
    src = r.get("source", "").strip()
    vol = r.get("volume", "").strip()
    iss = r.get("issue", "").strip()
    year = r.get("year", "").strip()
    pages = r.get("pages", "").strip()
    if vol and iss:
        src += f", vol. {vol}, no. {iss}"
    src_it = f"*{src}*" if src and not _is_cjk(src) else src
    seg = [f'{authors.rstrip(".")}. "{title}." {src_it}.']
    if pages:
        seg.append(f"pp. {pages}.")
    seg.append(f"{year}.")
    if r.get("doi", "").strip():
        seg.append(f"https://doi.org/{r['doi'].strip()}.")
    return " ".join(x for x in seg if x)


FORMATS = {"gbt7714": gbt7714, "apa": apa, "vancouver": vancouver, "mla": mla}


def format_ref(r, style="gbt7714"):
    fn = FORMATS.get(style.lower(), gbt7714)
    return fn(r)


def format_ref_segments(r, style="gbt7714"):
    """把 format_ref 的纯文本（含 *...* 斜体标记）拆分为
    [(text, italic:bool), ...]，供 insert_refs 生成带斜体的 Word run。
    星号只作标记，不进入输出文本。"""
    text = format_ref(r, style)
    segs = []
    for i, chunk in enumerate(re.split(r"\*([^*]*)\*", text)):
        if chunk == "":
            continue
        segs.append((chunk, i % 2 == 1))
    return segs


if __name__ == "__main__":
    import json
    import sys
    if len(sys.argv) >= 2:
        style = sys.argv[1]
        data = json.loads(sys.stdin.read())
        for r in data:
            print(format_ref(r, style))
    else:
        print("usage: echo '[{\"title\":\"...\",...}]' | python format_refs.py gbt7714")
