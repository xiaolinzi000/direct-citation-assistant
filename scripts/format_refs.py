# -*- coding: utf-8 -*-
"""
format_refs.py — 引用条目格式化（论文引用 skill）

按目标格式把 refs.csv 中的一条记录渲染成参考文献文本：
  gbt7714   GB/T 7714-2015 顺序编码制（中文期刊/毕业论文最常见）
  apa       APA 第 7 版
  vancouver Vancouver / ICMJE 编号式（多数生物医学期刊）
  mla       MLA 第 9 版
  harvard   Harvard 作者-年份制（Cite Them Right 风格，英国体系常用）
  ieee      IEEE 参考文献格式（IEEE Reference Guide；TIM 等 IEEE 期刊）

用法（脚本内调用）：
    from format_refs import format_ref, gbt7714, apa, vancouver, mla, harvard, ieee
    text = format_ref(ref_dict, style="ieee")
"""

import re
from datetime import datetime

# ---------- DOI 精确解析（v1.9：不依赖字符剥除，防重复前缀/猜填）----------

DOI_RE = re.compile(r"\b10\.\d{4,9}/[^\s，。；;<>\"'`]+")


def normalize_doi(raw):
    """从录入值精确提取 DOI。

    兼容多种录入形态：'10.1109/TIM.2020.2978991'、
    'https://doi.org/10.1109/TIM.2020.2978991'、'doi:10.1109/…'、
    'DOIs: 10.1109/…'、'https://dx.doi.org/10.…'；
    自动剥离尾随句点/逗号/括号。提取不到返回 None（视为"无可信 DOI"，
    由上层明确报告，不猜填）。"""
    if not raw or not str(raw).strip():
        return None
    m = DOI_RE.search(str(raw).strip())
    if not m:
        return None
    return m.group(0).rstrip(".,;，。；)】]\"'` ") 


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

def _icmje_single(a):
    """ICMJE/Vancouver 单作者：'Landeta, C.' -> 'Landeta C'；
    'Collet, J.-F.' -> 'Collet JF'（姓保留原大小写，首字母连写、不带句点）。"""
    a = a.strip()
    if _is_cjk(a):
        return a
    if "," in a:
        surname, given = a.split(",", 1)
        initials = "".join(re.findall(r"[A-Z]", given.upper()))
        return f"{surname.strip()} {initials}".strip()
    words = a.split()
    if len(words) == 1:
        return words[0]
    surname = words[-1]
    initials = "".join(w[0].upper() for w in words[:-1])
    return f"{surname} {initials}".strip()


def _van_authors(raw, max_a=6):
    lst = _split_authors(raw)
    if not lst:
        return ""
    head = [_icmje_single(a) for a in lst[:max_a]]
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
    doi = normalize_doi(r.get("doi", ""))
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
    doi = normalize_doi(r.get("doi", ""))
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
    doi = normalize_doi(r.get("doi", ""))
    src_it = f"*{src}*" if src and not _is_cjk(src) else src
    seg = []
    if authors:
        seg.append(authors.rstrip(".") + ".")
    if title:
        seg.append(title if title.endswith(".") else f"{title}.")
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
    doi = normalize_doi(r.get("doi", ""))
    if doi:
        seg.append(f"https://doi.org/{doi}.")
    return " ".join(x for x in seg if x)


# ---------- Harvard（作者-年份制，Cite Them Right）----------

def _harvard_single(a):
    """Harvard 单作者：'Smith, J.' 原样；'J. Smith' -> 'Smith, J.'；中文原样"""
    a = a.strip()
    if _is_cjk(a):
        return a
    if "," in a:
        return a
    words = a.split()
    if len(words) <= 1:
        return a
    return f"{words[-1]}, {''.join(w[0].upper() + '.' for w in words[:-1])}"


def _harvard_authors(raw):
    """Harvard 全部作者列出：A, B and C（末位 and）。"""
    lst = [_harvard_single(a) for a in _split_authors(raw)]
    if not lst:
        return ""
    if len(lst) == 1:
        return lst[0]
    return ", ".join(lst[:-1]) + " and " + lst[-1]


def harvard(r):
    """Harvard（作者-年份制）。
    期刊：Surname, I. (Year) 'Title', Journal, Vol(Issue), pp. pages. doi/URL.
    专著：Surname, I. (Year) Title. City: Publisher.
    网页：Surname, I. (Year) Title. Available at: URL (Accessed: 日期).
    学位论文：Surname, I. (Year) Title. PhD thesis. City: 机构. """
    authors = _harvard_authors(r.get("authors", ""))
    title = r.get("title", "").strip()
    year = r.get("year", "").strip()
    src = r.get("source", "").strip()
    vol = r.get("volume", "").strip()
    iss = r.get("issue", "").strip()
    pages = r.get("pages", "").strip()
    doi = normalize_doi(r.get("doi", ""))
    city = r.get("city", "").strip()
    url = r.get("url", "").strip()
    ttype = r.get("type", "journal")
    head = f"{authors} ({year}). {title}." if authors else f"({year}). {title}."
    if ttype == "book":
        base = f"{head} *{src}*." if src else f"{head}."
        if city:
            base = f"{head} *{src}*. {city}." if src else f"{head} {city}."
        return base
    if ttype == "thesis":
        return f"{head} PhD thesis. {city + '. ' if city else ''}{src}."
    if ttype == "web":
        base = f"{head}"
        if url:
            today = datetime.now().strftime("%Y-%m-%d")
            base += f" Available at: {url} (Accessed: {today})."
        elif src:
            base += f" {src}."
        return base
    # 期刊/会议
    base = f"{head} *{src}*," if src else f"{head}"
    tail = f" {year}"
    if vol:
        tail += f", {vol}"
        if iss:
            tail += f"({iss})"
        if pages:
            tail += f", pp. {pages}"
    base += tail + "."
    if doi:
        base += f" doi: {doi}."
    elif url:
        today = datetime.now().strftime("%Y-%m-%d")
        base += f" Available at: {url} (Accessed: {today})."
    return base


# ---------- IEEE（TIM 等 IEEE 期刊，v1.9 新增）----------

# IEEE 常用期刊全称 → 官方缩写（IEEE Reference Guide 附录 IV）。
# source 字段命中全称时自动替换；未命中保持原样并在校验层提示手填缩写。
IEEE_JOURNAL_ABBREV = {
    "IEEE Transactions on Instrumentation and Measurement":
        "IEEE Trans. Instrum. Meas.",
    "IEEE Transactions on Industrial Electronics":
        "IEEE Trans. Ind. Electron.",
    "IEEE Transactions on Industrial Informatics":
        "IEEE Trans. Ind. Informat.",
    "IEEE Transactions on Power Electronics":
        "IEEE Trans. Power Electron.",
    "IEEE Transactions on Signal Processing":
        "IEEE Trans. Signal Process.",
    "IEEE Transactions on Instrumentation & Measurement":
        "IEEE Trans. Instrum. Meas.",
    "IEEE Sensors Journal": "IEEE Sensors J.",
    "IEEE Transactions on Neural Networks and Learning Systems":
        "IEEE Trans. Neural Netw. Learn. Syst.",
    "IEEE Transactions on Pattern Analysis and Machine Intelligence":
        "IEEE Trans. Pattern Anal. Mach. Intell.",
    "IEEE Transactions on Geoscience and Remote Sensing":
        "IEEE Trans. Geosci. Remote Sens.",
    "IEEE Journal of Solid-State Circuits": "IEEE J. Solid-State Circuits",
    "IEEE Transactions on Antennas and Propagation":
        "IEEE Trans. Antennas Propag.",
    "IEEE Communications Magazine": "IEEE Commun. Mag.",
    "IEEE Access": "IEEE Access",
    "Nature": "Nature",
    "Science": "Science",
    "Proceedings of the IEEE": "Proc. IEEE",
}

IEEE_MONTHS = {
    "jan": "Jan.", "feb": "Feb.", "mar": "Mar.", "apr": "Apr.",
    "may": "May", "jun": "Jun.", "jul": "Jul.", "aug": "Aug.",
    "sep": "Sep.", "oct": "Oct.", "nov": "Nov.", "dec": "Dec.",
}


def abbrev_journal(source):
    """把期刊全称替换为 IEEE 官方缩写；未命中保持原样。"""
    if not source:
        return ""
    return IEEE_JOURNAL_ABBREV.get(source.strip(), source.strip())


def _ieee_month(year_month):
    """'Feb 2025' / '2025-02' / '2025' -> 'Feb. 2025' / 'Feb. 2025' / '2025'"""
    if not year_month:
        return ""
    s = str(year_month).strip()
    # 2025-02 / 2025/02 / 02.2025
    m = re.match(r"^(\d{4})[-/.](\d{1,2})$", s)
    if m:
        mn = IEEE_MONTHS.get(
            ["jan", "feb", "mar", "apr", "may", "jun",
             "jul", "aug", "sep", "oct", "nov", "dec"][int(m.group(2)) - 1])
        return f"{mn} {m.group(1)}"
    # Feb 2025 / February 2025
    m = re.match(r"^([A-Za-z]+)[\s.]*(\d{4})$", s)
    if m:
        mn = IEEE_MONTHS.get(m.group(1).lower()[:3])
        return f"{mn} {m.group(2)}" if mn else s
    return s


def _ieee_single(a):
    """IEEE 单作者：名缩写（每名一个字母带点）在前、姓在后。
    'Smith, J.' -> 'J. Smith'；'Vaswani, A.' -> 'A. Vaswani'；中文原样。"""
    a = a.strip()
    if _is_cjk(a):
        return a
    if "," in a:
        surname, given = a.split(",", 1)
        initials = [w.strip().rstrip(".") for w in
                    re.findall(r"[A-Z][^,\s]*", given.upper())]
        # 复合缩写 J.-F. -> 'J.-F.'；普通 A. -> 'A.'
        initials = " ".join(w if "-" in w else w[0] + "." for w in initials)
        return f"{initials} {surname.strip()}".strip()
    words = a.split()
    if len(words) == 1:
        return a
    surname = words[-1]
    initials = " ".join(w[0].upper() + "." for w in words[:-1])
    return f"{initials} {surname}".strip()


def _ieee_authors(raw, max_a=6):
    """IEEE 作者列表：'J. Smith, K. Jones, and L. Lee'；>6 个作者用 et al.
    （IEEE 规范：最多列 6 位，超出用 et al.）"""
    lst = [_ieee_single(a) for a in _split_authors(raw)]
    if not lst:
        return ""
    if len(lst) == 1:
        return lst[0]
    if len(lst) <= max_a:
        return ", ".join(lst[:-1]) + ", and " + lst[-1]
    return ", ".join(lst[:max_a]) + ", et al."


def ieee(r):
    """IEEE 参考文献格式（IEEE Reference Guide；TIM 等 IEEE 期刊采用）。

    期刊：J. K. Author, "Title," Abbrev. J. Name, vol. x, no. x,
          pp. xxx-xxx, Abbrev. Month, year, doi: 10.xxxx/xxx.
    会议：J. K. Author, "Title," in Proc. Abbrev. Conf. Name, City,
          State, Country, year, pp. xxx-xxx.
    专著：J. K. Author, Title. City, State, Country: Publisher, year, pp.
    学位论文：J. K. Author, "Title," Ph.D. dissertation, Dept., Univ.,
              City, State, year.
    网页：J. K. Author, "Title," Site Name, year. [Online]. Available: URL
    正文编号不带上标（IEEE 正文引用 [1] 为行内方括号）。"""
    authors = _ieee_authors(r.get("authors", ""))
    title = r.get("title", "").strip()
    src = abbrev_journal(r.get("source", "").strip())
    year = r.get("year", "").strip()
    vol = r.get("volume", "").strip()
    iss = r.get("issue", "").strip()
    pages = r.get("pages", "").strip()
    doi = normalize_doi(r.get("doi", ""))
    city = r.get("city", "").strip()
    url = r.get("url", "").strip()
    month = _ieee_month(r.get("month", "") or year)
    ttype = r.get("type", "journal")
    # 作者段（IEEE 不用句号收尾，后接题名）
    head = authors or ""
    ymo = month or year
    if ttype == "book":
        src_it = f"*{src}*" if src and not _is_cjk(src) else src
        base = f'{head}, *{title}*.' if head else f"*{title}*."
        if city:
            base += f" {city}: {src_it}," if src_it else f" {city}."
        elif src_it:
            base += f" {src_it}."
        if year:
            base += f" {year}."
        if pages:
            base += f" pp. {pages}."
        if doi:
            base += f" doi: {doi}."
        return base
    if ttype == "thesis":
        degree = "Ph.D. dissertation"
        base = (f'{head}, "{title}," {degree}' if head
                else (f'"{title}," {degree}' if title else degree))
        if city:
            base += f", {city}"
        if src:
            base += f", {src}"
        if year:
            base += f", {year}"
        base += "."
        if doi:
            base += f" doi: {doi}."
        return base
    if ttype == "web":
        base = (f'{head}, "{title}," {src}' if head
                else (f'"{title}," {src}' if title else src))
        if year:
            base += f", {year}"
        base += ". [Online]. Available: " + (url or "")
        return base
    if ttype == "conference":
        base = (f'{head}, "{title}," in {src}' if head
                else (f'"{title}," in {src}' if title else f"in {src}"))
        if city:
            base += f", {city}"
        if year:
            base += f", {year}"
        if pages:
            base += f", pp. {pages}"
        base += "."
        if doi:
            base += f" doi: {doi}."
        return base
    # 期刊（默认）
    src_it = f"*{src}*" if src and not _is_cjk(src) else src
    base = (f'{head}, "{title}," {src_it}' if head
            else (f'"{title}," {src_it}' if title else src_it))
    if vol:
        base += f", vol. {vol}"
    if iss:
        base += f", no. {iss}"
    if pages:
        base += f", pp. {pages}"
    if ymo:
        base += f", {ymo}"
    base += "."
    if doi:
        base += f" doi: {doi}."
    return base


FORMATS = {"gbt7714": gbt7714, "apa": apa, "vancouver": vancouver,
           "mla": mla, "harvard": harvard, "ieee": ieee}


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
