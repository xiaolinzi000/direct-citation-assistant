# -*- coding: utf-8 -*-
"""
insert_refs.py — 在 Word 文档中插入/更新带超链接的论文引用（论文引用 skill）

功能（对应需求 5、6）：
  1. 正文中的引用点替换为可点击跳转的编号引用（如 [1]，默认上标）；
     点击编号即可跳转到文末对应文献条目（Word 中默认 Ctrl+点击，可在
     文件→选项→高级→「用 Ctrl+单击跟踪超链接」取消勾选后改为单击跳转）。
  2. 文末自动生成参考文献目录：在已有的 References/参考文献 标题后生成；
     没有该标题则自动创建。
  3. 自动重编号：删除中间某处引用后重跑本脚本，其余编号自动连续更新；
     同一篇论文多次引用只占一个编号（文末只列一次）。

引用点（正文中）：
  [?]            —— 按出现顺序自动取 refs.csv 中尚未被引用的下一条
  [CITE:key]     —— 指定引用 refs.csv 中 key 对应的条目
  （脚本上次生成的 [n] 超链接也是引用点，重跑时自动重新编号）

用法：
  python insert_refs.py --docx 文稿.docx --refs refs.csv [--style gbt7714]
  python insert_refs.py --docx 第1章.docx 第2章.docx 第3章.docx ^
      --refs refs.csv            # 分章论文：全文统一编号，文末表生成在最后一个文档
  python insert_refs.py --docx 第1章.docx 第2章.docx --main 第2章.docx
                                 # 指定主文档（参考文献表所在，默认最后一个）
  python insert_refs.py --docx 文稿.docx --dry-run      # 只打印计划不写文件
  常用参数：
    --style gbt7714|apa|vancouver|mla   参考文献格式（默认 gbt7714）
    --citation numbered|author-year     编号制（默认，正文 [1] 上标）/
                                       作者-年份制（正文 (Smith et al., 2023)）
    --heading 参考文献                   自定义标题文本（找不到时创建）
    --no-superscript                    正文编号不用上标
    --bracket ()                        正文编号括号（默认 []）
    --entry-num num|bracket             文末条目编号样式（默认 bracket -> [1]）
    --font-en "Times New Roman"         西文/数字字体（默认 Times New Roman）
    --font-cn "宋体"                     中文字体（默认宋体）
    --no-indent                         关闭条目悬挂缩进（默认缩进 2 字符）
    --hanging-pt 21                     悬挂缩进量 pt（五号 21 / 小四 24 / 四号 28）
    --align both|left|""                条目对齐（默认 both 两端对齐）
    --bold-num                          条目编号加粗（默认不加粗）
    --no-italic-source                  关闭出处（西文期刊名/书名）斜体（默认按规范斜体）
    --main 路径                         多文档时指定主文档（默认最后一个 --docx）
    --backup-dir 路径                    备份目录（默认 docx 同目录 _backup）
    --dry-run                           只打印计划不写文件
"""

import argparse
import io
import os
import re
import shutil
import sys
import zipfile
from datetime import datetime

from lxml import etree

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from format_refs import format_ref_segments
from refs_db import load_refs

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W}
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
P_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
CITE_RE = re.compile(r"\[CITE:([A-Za-z0-9_\-\u4e00-\u9fff.]+)\]", re.IGNORECASE)
ANY_RE = re.compile(r"\[\?\]")
# 旧引用识别：编号制 [1] 或 (1)（bracket () 场景），或 author-year (Author, 2023)
OLD_REF_RE = re.compile(r"^\[\d+\]$|^\(\d+\)$|^\([^()]*\d{4}[^()]*\)$")
ENTRY_START_RE = re.compile(r"^\[(\d+)\]\s")

# DOI 识别：URL 形式（https://doi.org/…）或原始形式（10.xxxx/…）
DOI_URL_RE = re.compile(r"https?://doi\.org/\S+", re.IGNORECASE)
DOI_RAW_RE = re.compile(r"10\.\d{4,9}/[^\s，。；;]+")

TARGET_HEADINGS = ("references", "reference", "bibliography", "参考文献",
                   "参考文献(references)", "references(参考文献)", "文献目录")


def w(tag):
    return f"{{{W}}}{tag}"


# ---------- 文档读取 ----------

def load_docx(path):
    """读 docx：返回 (zip_items, document_root, styles_root)。"""
    with zipfile.ZipFile(path, "r") as z:
        items = {n: z.read(n) for n in z.namelist()}
    doc_root = etree.fromstring(items["word/document.xml"])
    styles_root = None
    if "word/styles.xml" in items:
        try:
            styles_root = etree.fromstring(items["word/styles.xml"])
        except Exception:
            styles_root = None
    return items, doc_root, styles_root


def save_docx(path, items, doc_root):
    """写回 docx（仅替换 document.xml，其余原样保留）。
    Word 独占锁定时会失败，给出可操作的提示。"""
    items["word/document.xml"] = etree.tostring(doc_root, xml_declaration=True,
                                                encoding="UTF-8", standalone=True)
    tmp = path + ".tmp"
    try:
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
            for name, data in items.items():
                z.writestr(name, data)
        os.replace(tmp, path)
    except PermissionError:
        if os.path.exists(tmp):
            try:
                os.remove(tmp)
            except OSError:
                pass
        raise RuntimeError(
            f"无法写入文档（可能正被 Word 打开）：{path}\n"
            "请先关闭该文档的 Word 窗口，再重新运行本脚本。")


def backup_docx(path, backup_dir=None, keep=30):
    """备份 docx 并清理 _backup 目录，只保留最近 keep 份。"""
    target_dir = backup_dir or os.path.join(os.path.dirname(os.path.abspath(path)), "_backup")
    os.makedirs(target_dir, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    bpath = os.path.join(target_dir, f"{ts}__{os.path.basename(path)}")
    try:
        shutil.copy2(path, bpath)
    except PermissionError:
        raise RuntimeError(
            f"无法备份文档（可能正被 Word 打开）：{path}\n"
            "请先关闭该文档的 Word 窗口，再重新运行本脚本。")
    # 清理：同目标文件名前缀，保留最新的 keep 份
    prefix = f"__{os.path.basename(path)}"
    try:
        files = [f for f in os.listdir(target_dir) if f.endswith(prefix)]
        files.sort(reverse=True)
        for old in files[keep:]:
            try:
                os.remove(os.path.join(target_dir, old))
            except OSError:
                pass
    except OSError:
        pass
    return bpath


# ---------- 样式探测 ----------

def find_style_id(styles_root, names):
    """在 styles.xml 中按 name 找 styleId。"""
    if styles_root is None:
        return None
    for s in styles_root.findall(w("style")):
        sid = s.get(w("styleId"))
        nm = s.find(w("name"))
        if nm is not None and nm.get(w("val")) in names:
            return sid
    return None


def run_text(run):
    return "".join(t.text or "" for t in run.findall(w("t")))


def para_text(p):
    """段落完整文本（含超链接内 run）。"""
    return "".join(run_text(r) for r in p.iter(w("r")))


def is_ref_hyperlink(el):
    """判断是否为脚本生成的引用超链接（anchor 以 ref_ 开头、文本形如 [n]）。"""
    if el.tag != w("hyperlink"):
        return False
    anchor = el.get(w("anchor")) or ""
    if not anchor.startswith("ref_"):
        return False
    txt = "".join(run_text(r) for r in el.findall(w("r")))
    return bool(OLD_REF_RE.match(txt.strip()))


# ---------- 正文引用点收集 ----------

def collect_points(body):
    """遍历正文，返回引用点列表：
       {kind: 'old'|'cite'|'any', key: str|None, hyperlink: el|None}
       顺序即文档顺序。"""
    points = []
    for p in body.iter(w("p")):
        # 1) 旧超链接（上一次生成）——直接保留为引用点
        for h in p.findall(w("hyperlink")):
            if is_ref_hyperlink(h):
                anchor = h.get(w("anchor"))
                points.append({"kind": "old", "key": anchor[4:], "hyperlink": h,
                               "paragraph": p})
        # 2) 占位符 —— 循环扫描，处理一个重扫一次
        while True:
            texts = [(r, run_text(r)) for r in p.findall(w("r"))]
            full = "".join(t for _, t in texts)
            m = CITE_RE.search(full)
            if not m:
                m = ANY_RE.search(full)
                kind = "any"
                if not m:
                    break
            else:
                kind = "cite"
            key = m.group(1) if kind == "cite" else None
            anchor_run = _strip_placeholder(p, m.start(), m.end())
            if anchor_run is None:
                raise RuntimeError(f"占位符定位失败：段落「{full[:40]}…」")
            points.append({"kind": kind, "key": key, "anchor": anchor_run,
                           "paragraph": p})
    return points


def _set_run_text(r, text):
    for tt in r.findall(w("t")):
        r.remove(tt)
    if text:
        t = etree.SubElement(r, w("t"))
        t.text = text


def _strip_placeholder(p, start, end):
    """移除段落中 [start,end) 字符区间的占位符文本，在占位符原位插入
    一个空 run 作锚点（后续替换为超链接）。返回锚点 run。"""
    texts = [(r, run_text(r)) for r in p.findall(w("r"))]
    full = "".join(t for _, t in texts)
    if start < 0 or end > len(full):
        return None
    # 定位起止 run
    acc = 0
    i = j = None
    si = sj = 0
    for idx, (r, t) in enumerate(texts):
        r_start, r_end = acc, acc + len(t)
        if i is None and start < r_end:
            i, si = idx, start - r_start
        if j is None and end <= r_end:
            j, sj = idx, end - r_start
            break
        acc = r_end
    if i is None or j is None:
        return None
    anchor = etree.Element(w("r"))  # 空 run 锚点
    if i == j:
        r_i, t_i = texts[i]
        _set_run_text(r_i, t_i[:si] + t_i[sj:])
        r_i.addnext(anchor)
    else:
        r_i, t_i = texts[i]
        r_j, t_j = texts[j]
        _set_run_text(r_i, t_i[:si])
        _set_run_text(r_j, t_j[sj:])
        # 删除 i 与 j 之间的所有 run
        for idx in range(i + 1, j):
            p.remove(texts[idx][0])
        if not t_i[:si]:
            p.remove(r_i)
            r_j.addprevious(anchor)
        else:
            r_i.addnext(anchor)
    return anchor


# ---------- 编号分配 ----------

def assign_numbers(points, refs):
    """给引用点分配编号：key 首次出现定号，重复引用同一 key 复用编号；
       'any' 占位符依次取未引用的条目（按 refs.csv 顺序）。
       返回 (points, key_to_num, used_keys 顺序列表) 或抛错。"""
    key_to_num = {}
    used_keys = []
    any_pool = [r["key"] for r in refs]
    for pt in points:
        key = pt["key"]
        if pt["kind"] == "any":
            cands = [k for k in any_pool if k not in used_keys]
            if not cands:
                raise RuntimeError("正文 [?] 数量多于 refs.csv 中未引用的条目，"
                                   "请先 add_refs.py 补充文献。")
            key = cands[0]
            pt["key"] = key
        if key not in key_to_num:
            key_to_num[key] = len(key_to_num) + 1
            used_keys.append(key)
        pt["num"] = key_to_num[key]
    return points, key_to_num, used_keys


# ---------- 正文替换 ----------

def set_run_fonts(rpr, font_en, font_cn):
    """设置 run 字体：西文（英文/数字/标点）与中文分别指定。
       默认规则：英文与数字用 Times New Roman，中文用宋体。"""
    rf = etree.SubElement(rpr, w("rFonts"))
    if font_en:
        rf.set(w("ascii"), font_en)
        rf.set(w("hAnsi"), font_en)
    if font_cn:
        rf.set(w("eastAsia"), font_cn)


def build_hyperlink(anchor, text, superscript, hyperlink_style_id,
                    font_en="", font_cn=""):
    h = etree.Element(w("hyperlink"))
    h.set(w("anchor"), anchor)
    r = etree.SubElement(h, w("r"))
    rpr = etree.SubElement(r, w("rPr"))
    set_run_fonts(rpr, font_en, font_cn)
    if hyperlink_style_id:
        st = etree.SubElement(rpr, w("rStyle"))
        st.set(w("val"), hyperlink_style_id)
    else:
        # 文档无 Hyperlink 样式时的 fallback：蓝色 + 下划线，保证可点击外观
        color = etree.SubElement(rpr, w("color"))
        color.set(w("val"), "0563C1")
        u = etree.SubElement(rpr, w("u"))
        u.set(w("val"), "single")
    if superscript:
        va = etree.SubElement(rpr, w("vertAlign"))
        va.set(w("val"), "superscript")
    t = etree.SubElement(r, w("t"))
    t.text = text
    return h


def compute_year_suffixes(used_keys, refs_by_key):
    """author-year 制：同一（首作者, 年份）的多篇文献按题名字母序加 a/b/c 后缀
    （GB/T 7714 著者-出版年制规范，如 (Zhang, 2023a)/(Zhang, 2023b)）。
    返回 {key: suffix}；无同年同作者时返回空 dict。"""
    try:
        from format_refs import _split_authors
    except Exception:
        _split_authors = lambda s: []
    groups = {}
    for key in used_keys:
        ref = refs_by_key.get(key, {})
        lst = _split_authors(ref.get("authors", ""))
        first = (lst[0] if lst else "").lower()
        g = (first, ref.get("year", ""))
        groups.setdefault(g, []).append(key)
    suffix = {}
    for g, keys in groups.items():
        if len(keys) <= 1:
            continue
        ordered = sorted(
            keys,
            key=lambda k: ((refs_by_key.get(k, {}) or {}).get("title", "").lower()))
        for i, k in enumerate(ordered, 1):
            suffix[k] = chr(ord("a") + i - 1)
    return suffix


def make_label(pt, bracket, citation_mode, refs_by_key, year_suffix=None):
    """生成正文引用显示文本：
       numbered  ->  [n] 或 (n)（括号由 --bracket 定）
       author-year -> (Author, Year)；多作者英文 (Smith et al., 2023)、中文 (张三等, 2023)
    """
    key = pt["key"]
    year_suffix = year_suffix or {}
    if citation_mode == "author-year":
        ref = refs_by_key.get(key, {})
        year = ref.get("year", "").strip()
        sfx = year_suffix.get(key, "")
        try:
            from format_refs import _is_cjk, _split_authors
        except Exception:
            _is_cjk = lambda s: False
            _split_authors = lambda s: []
        lst = _split_authors(ref.get("authors", ""))
        if lst:
            first = lst[0]
            if "," in first:
                surname = first.split(",")[0].strip()
            else:
                words = first.split()
                surname = words[-1] if words else first
            if len(lst) > 1:
                suffix = "等" if _is_cjk(first) else " et al."
            else:
                suffix = ""
            name = f"{surname}{suffix}"
            return f"({name}, {year}{sfx})" if year else f"({name})"
        return f"({year}{sfx})" if year else "(?)"
    return f"{bracket[0]}{pt['num']}{bracket[1]}"


def insert_hyperlinks(points, bracket, superscript, hyperlink_style_id,
                      citation_mode="numbered", refs_by_key=None,
                      font_en="", font_cn="", year_suffix=None):
    """为占位符类引用点把锚点 run 替换为超链接；更新旧超链接文本。返回替换数量。"""
    refs_by_key = refs_by_key or {}
    new_count = 0
    for pt in points:
        label = make_label(pt, bracket, citation_mode, refs_by_key, year_suffix)
        anchor = f"ref_{pt['key']}"
        if pt["kind"] == "old":
            # 更新已有超链接内 run 文本
            for r in pt["hyperlink"].findall(w("r")):
                for t in r.findall(w("t")):
                    if OLD_REF_RE.match((t.text or "").strip()):
                        t.text = label
            continue
        # 占位符：把超链接插到锚点 run 的位置，移除锚点
        anchor_el = pt["anchor"]
        h = build_hyperlink(anchor, label, superscript, hyperlink_style_id,
                            font_en, font_cn)
        anchor_el.addprevious(h)
        anchor_el.getparent().remove(anchor_el)
        new_count += 1
    return new_count


# ---------- 参考文献表 ----------

def find_or_create_heading(body, heading_text, heading_style_id):
    """找标题段（文本归一化后匹配 TARGET_HEADINGS 或用户指定 heading_text），
       没有则创建。返回标题段元素。"""
    def _norm(s):
        # 去空白/冒号/句点/数字后缀，统一小写，便于宽容匹配
        s = re.sub(r"[\s:：.。·\-–—]+", "", (s or "").strip().casefold())
        return re.sub(r"\d+$", "", s)

    candidates = set()
    if heading_text:
        candidates.add(_norm(heading_text))
    for t in TARGET_HEADINGS:
        candidates.add(_norm(t))
    for p in body.iter(w("p")):
        if _norm(para_text(p)) in candidates:
            return p
    # 创建标题段（文末）
    p = etree.SubElement(body, w("p"))
    if heading_style_id:
        ppr = etree.SubElement(p, w("pPr"))
        ps = etree.SubElement(ppr, w("pStyle"))
        ps.set(w("val"), heading_style_id)
    r = etree.SubElement(p, w("r"))
    t = etree.SubElement(r, w("t"))
    t.text = heading_text or "参考文献"
    return p


def sort_key_ref(ref):
    """author-year 模式的排序键：首作者姓 + 年份。
       中文姓转汉语拼音排序（无 pypinyin 时回退 Unicode 码点）。"""
    authors = ref.get("authors", "")
    try:
        from format_refs import _split_authors, _is_cjk
        lst = _split_authors(authors)
    except Exception:
        lst = []
    first = lst[0] if lst else ""
    surname = first.split(",")[0].strip().lower() if "," in first else first.lower()
    if _is_cjk(surname):
        try:
            from pypinyin import lazy_pinyin
            surname = "".join(lazy_pinyin(surname))
        except Exception:
            pass
    return (surname, ref.get("year", ""))


def build_entry_paragraph(num, key, ref, style, entry_num_style, bookmark_id,
                          citation_mode="numbered", font_en="", font_cn="",
                          indent=True, align="both", bold_num=False,
                          hanging_pt=21.0, italic_source=True,
                          doi_links=None, hyperlink_style_id=None):
    p = etree.Element(w("p"))
    # 段落格式：通行惯例「悬挂缩进 2 字符」（国标未强制版式，可 --no-indent 关闭）。
    # 必须用 twips 单位 w:left/w:hanging（=2字符×字号），leftChars/hangingChars 在
    # Word 中首行无法正确归零（实测首行仍右移约1字符）。2 字符=21pt(五号)/24pt(小四)/
    # 28pt(四号)，默认按五号 21pt，正文字号不同时用 --hanging-pt 覆盖。
    ppr = etree.SubElement(p, w("pPr"))
    if indent:
        ind = etree.SubElement(ppr, w("ind"))
        twips = int(round(hanging_pt * 20))
        ind.set(w("left"), str(twips))
        ind.set(w("hanging"), str(twips))
    if align:
        jc = etree.SubElement(ppr, w("jc"))
        jc.set(w("val"), align)
    # 书签
    bs = etree.SubElement(p, w("bookmarkStart"))
    bs.set(w("id"), str(bookmark_id))
    bs.set(w("name"), f"ref_{key}")
    # 编号（author-year 模式无编号）：编号 + 制表符，续行悬挂缩进后与内容首字符对齐
    if citation_mode == "numbered":
        prefix = f"[{num}]" if entry_num_style == "bracket" else f"{num}."
        r1 = etree.SubElement(p, w("r"))
        r1pr = etree.SubElement(r1, w("rPr"))
        set_run_fonts(r1pr, font_en, font_cn)
        if bold_num:
            b = etree.SubElement(r1pr, w("b"))
        t1 = etree.SubElement(r1, w("t"))
        t1.text = prefix
        rtab = etree.SubElement(p, w("r"))
        tab = etree.SubElement(rtab, w("tab"))  # 制表符：内容推进到悬挂缩进位置
    # 条目文本分段渲染：中英文混排（中文宋体、英文/数字 Times New Roman），
    # 西文期刊名/书名等规范斜体部分拆为独立 run（加 <w:i/>）；
    # DOI 片段（https://doi.org/… 或 10.xxxx/…）做成可点击超链接指向论文页面。
    for text, italic in format_ref_segments(ref, style):
        if italic and not italic_source:
            italic = False
        m = None
        if doi_links:
            m = DOI_URL_RE.search(text)
            if m is None:
                m = DOI_RAW_RE.search(text)
        if m:
            raw = m.group(0)
            url = raw.rstrip(".,;：:。，；")
            drop = raw[len(url):]  # 被剥掉的尾部标点（句子句点等）保留在超链接外
            target = url if url.lower().startswith("https://doi.org/") \
                else "https://doi.org/" + url
            rid = doi_links.get(target)
            head = text[:m.start()]
            tail = drop + text[m.end():]
            if head:
                r0 = etree.SubElement(p, w("r"))
                r0pr = etree.SubElement(r0, w("rPr"))
                set_run_fonts(r0pr, font_en, font_cn)
                if italic:
                    etree.SubElement(r0pr, w("i"))
                t0 = etree.SubElement(r0, w("t"))
                t0.text = head
            if rid:
                h = etree.SubElement(p, w("hyperlink"))
                h.set(f"{{{R_NS}}}id", rid)
                h.set(w("tooltip"), "打开论文页面")
                r1 = etree.SubElement(h, w("r"))
                r1pr = etree.SubElement(r1, w("rPr"))
                set_run_fonts(r1pr, font_en, font_cn)
                if italic:
                    etree.SubElement(r1pr, w("i"))
                if hyperlink_style_id:
                    st = etree.SubElement(r1pr, w("rStyle"))
                    st.set(w("val"), hyperlink_style_id)
                else:
                    # 无 Hyperlink 样式时 fallback：蓝色 + 下划线，保证可点击外观
                    color = etree.SubElement(r1pr, w("color"))
                    color.set(w("val"), "0563C1")
                    u = etree.SubElement(r1pr, w("u"))
                    u.set(w("val"), "single")
                t1 = etree.SubElement(r1, w("t"))
                t1.text = url
            if tail:
                r2 = etree.SubElement(p, w("r"))
                r2pr = etree.SubElement(r2, w("rPr"))
                set_run_fonts(r2pr, font_en, font_cn)
                if italic:
                    etree.SubElement(r2pr, w("i"))
                t2 = etree.SubElement(r2, w("t"))
                t2.text = tail
            continue
        r2 = etree.SubElement(p, w("r"))
        r2pr = etree.SubElement(r2, w("rPr"))
        set_run_fonts(r2pr, font_en, font_cn)
        if italic:
            etree.SubElement(r2pr, w("i"))
        t2 = etree.SubElement(r2, w("t"))
        t2.text = text
    # 书签结束
    be = etree.SubElement(p, w("bookmarkEnd"))
    be.set(w("id"), str(bookmark_id))
    return p


def doi_links_for_entries(items, refs, used_keys):
    """为文末条目中的 DOI 建立可点击超链接：
    document.xml.rels 增加 External 关系（Target=https://doi.org/xxx）。
    已有同 Target 的 hyperlink 关系复用（重跑幂等，不新增）。
    返回 {doi_url: rId} 映射；无 DOI 时返回 None。"""
    urls = []
    key_to_ref = {r["key"]: r for r in refs}
    for key in used_keys:
        ref = key_to_ref.get(key) or {}
        doi = (ref.get("doi") or "").strip()
        if doi:
            urls.append("https://doi.org/" + doi.lstrip("https://doi.org/").lstrip("/"))
    if not urls:
        return None
    rels_name = "word/_rels/document.xml.rels"
    rels_root = None
    if rels_name in items:
        try:
            rels_root = etree.fromstring(items[rels_name])
        except Exception:
            rels_root = None
    if rels_root is None:
        rels_root = etree.Element(f"{{{P_REL}}}Relationships")
    max_id = 0
    existing = {}
    for rel in rels_root:
        m = re.match(r"rId(\d+)", rel.get("Id") or "")
        if m:
            max_id = max(max_id, int(m.group(1)))
        if rel.get("TargetMode") == "External" and \
                (rel.get("Type") or "").endswith("/hyperlink"):
            t = rel.get("Target") or ""
            if t.startswith("https://doi.org/"):
                existing[t] = rel.get("Id")
    mapping = {}
    for url in urls:
        if url in existing:
            mapping[url] = existing[url]
            continue
        max_id += 1
        rid = f"rId{max_id}"
        rel = etree.SubElement(rels_root, f"{{{P_REL}}}Relationship")
        rel.set("Id", rid)
        rel.set("Type", ("http://schemas.openxmlformats.org/officeDocument/"
                         "2006/relationships/hyperlink"))
        rel.set("Target", url)
        rel.set("TargetMode", "External")
        existing[url] = rid
        mapping[url] = rid
    items[rels_name] = etree.tostring(rels_root, xml_declaration=True,
                                      encoding="UTF-8", standalone=True)
    return mapping


def rebuild_bibliography(body, heading, used_keys, refs, style,
                         entry_num_style, bookmark_id_base,
                         citation_mode="numbered", font_en="", font_cn="",
                         indent=True, align="both", bold_num=False,
                         hanging_pt=21.0, italic_source=True,
                         items=None, hyperlink_style_id=None,
                         year_suffix=None):
    """在标题段之后重建条目段：删除旧的（含 ref_ 书签的段），插入新条目。
       numbered   按出现顺序编号；
       author-year 按作者字母序排列、不编号。
       items 传入 zip 条目字典时，条目中的 DOI 会被渲染为可点击超链接。"""
    key_to_ref = {r["key"]: r for r in refs}
    # 收集标题后需要删除的旧条目段。
    # 判别规则：带 ref_ 书签的段必然是脚本生成的条目；无书签但以 [n] 开头
    # 的段，仅当它紧跟在条目段之后（兼容历史上无书签的旧条目）才视为条目，
    # 否则视为正文/附录段并停止清理——防止把参考文献标题后的正文误删。
    old_entries = []
    after = False
    prev_entry = False
    for child in body:
        if after:
            if child.tag == w("p"):
                if child is heading:
                    continue
                has_ref_bookmark = any(
                    (bs.get(w("name")) or "").startswith("ref_")
                    for bs in child.iter(w("bookmarkStart")))
                txt = para_text(child).strip()
                if has_ref_bookmark:
                    old_entries.append(child)
                    prev_entry = True
                elif ENTRY_START_RE.match(txt) and prev_entry:
                    old_entries.append(child)
                    prev_entry = True
                elif txt:
                    # 非空非条目段（正文/附录/其他内容）即停止清理
                    break
                else:
                    prev_entry = False  # 空段（空行分隔）跳过，继续扫描
        if child is heading:
            after = True
            prev_entry = False
    for e in old_entries:
        body.remove(e)
    # 插入新条目（标题段之后）
    if citation_mode == "author-year":
        used_keys = sorted(used_keys,
                           key=lambda k: sort_key_ref(key_to_ref.get(k, {})))
    doi_links = None
    if items is not None:
        doi_links = doi_links_for_entries(items, refs, used_keys)
    year_suffix = year_suffix or {}
    anchor_el = heading
    bid = bookmark_id_base
    for num, key in enumerate(used_keys, 1):
        ref = key_to_ref.get(key)
        if ref is None:
            raise RuntimeError(f"引用表中找不到 key={key}")
        if year_suffix.get(key):
            ref = dict(ref, year=(ref.get("year") or "") + year_suffix[key])
        p = build_entry_paragraph(num, key, ref, style, entry_num_style, bid,
                                  citation_mode, font_en, font_cn, indent,
                                  align, bold_num, hanging_pt, italic_source,
                                  doi_links, hyperlink_style_id)
        anchor_el.addnext(p)
        anchor_el = p
        bid += 1
    return len(used_keys)


# ---------- 跨文档链接 ----------

def add_cross_doc_links(items, doc_root, main_path, this_path):
    """把本文档中的引用超链接指向主文档（跨文档跳转）：
    document.xml.rels 增加 External 关系（Target=主文档相对路径），
    hyperlink 增加 r:id；已有 r:id 的跳过（重跑幂等）。返回新增链接数。
    用于分章论文：各章文档的 [n] 点击后打开主文档并跳转到对应文献条目。"""
    rels_name = "word/_rels/document.xml.rels"
    rels_root = None
    if rels_name in items:
        try:
            rels_root = etree.fromstring(items[rels_name])
        except Exception:
            rels_root = None
    if rels_root is None:
        rels_root = etree.Element(f"{{{P_REL}}}Relationships")
    max_id = 0
    for rel in rels_root:
        m = re.match(r"rId(\d+)", rel.get("Id") or "")
        if m:
            max_id = max(max_id, int(m.group(1)))
    rel_path = os.path.relpath(main_path,
                               os.path.dirname(os.path.abspath(this_path)))
    rel_path = rel_path.replace("\\", "/")
    count = 0
    for h in doc_root.iter(w("hyperlink")):
        anchor = h.get(w("anchor")) or ""
        if not anchor.startswith("ref_"):
            continue
        if h.get(f"{{{R_NS}}}id"):
            continue
        max_id += 1
        rid = f"rId{max_id}"
        h.set(f"{{{R_NS}}}id", rid)
        rel = etree.SubElement(rels_root, f"{{{P_REL}}}Relationship")
        rel.set("Id", rid)
        rel.set("Type", ("http://schemas.openxmlformats.org/officeDocument/"
                         "2006/relationships/hyperlink"))
        rel.set("Target", rel_path)
        rel.set("TargetMode", "External")
        count += 1
    items[rels_name] = etree.tostring(rels_root, xml_declaration=True,
                                      encoding="UTF-8", standalone=True)
    return count


# ---------- 主流程 ----------

def main():
    ap = argparse.ArgumentParser(
        description="Word 文档插入/更新带超链接的论文引用（支持多文档分章统一编号）")
    ap.add_argument("--docx", required=True, nargs="+",
                    help="Word 文档路径（可多个：分章论文全文统一编号，如 --docx 第1章.docx 第2章.docx）")
    ap.add_argument("--main", default=None,
                    help="主文档路径（参考文献表所在，默认最后一个 --docx）")
    ap.add_argument("--refs", default=None, help="refs.csv 路径（默认 docx 同目录）")
    ap.add_argument("--style", default="gbt7714",
                    choices=["gbt7714", "apa", "vancouver", "mla"])
    ap.add_argument("--heading", default="", help="参考文献标题文本（找不到时创建）")
    ap.add_argument("--no-superscript", action="store_true", help="正文编号不用上标")
    ap.add_argument("--bracket", default="[]", help="正文编号括号，如 [] 或 ()")
    ap.add_argument("--entry-num", default="bracket", choices=["bracket", "num"],
                    help="文末条目编号样式")
    ap.add_argument("--backup-dir", default=None, help="备份目录")
    ap.add_argument("--dry-run", action="store_true", help="只打印计划")
    ap.add_argument("--citation", default="numbered",
                    choices=["numbered", "author-year"],
                    help="正文引用样式：numbered 编号制 / author-year 作者-年份制（APA/MLA 类）")
    ap.add_argument("--font-en", default="Times New Roman",
                    help="西文/数字字体（默认 Times New Roman）")
    ap.add_argument("--font-cn", default="宋体",
                    help="中文字体（默认宋体）")
    ap.add_argument("--no-indent", action="store_true",
                    help="关闭参考文献条目悬挂缩进（默认按通行惯例缩进 2 字符）")
    ap.add_argument("--align", default="both", choices=["both", "left", ""],
                    help="参考文献条目对齐：both 两端对齐（默认，通行惯例）/ left 左对齐 / 空 继承文档")
    ap.add_argument("--bold-num", action="store_true",
                    help="文献编号加粗（默认不加粗，通行惯例与条目同格式）")
    ap.add_argument("--hanging-pt", type=float, default=21.0,
                    help="悬挂缩进量 pt（2 字符×正文字号：五号 21 / 小四 24 / 四号 28，默认 21）")
    ap.add_argument("--no-italic-source", action="store_true",
                    help="关闭出处（西文期刊名/书名）斜体（默认按 GB/T 7714 规范斜体）")
    ap.add_argument("--mapping", action="store_true",
                    help="输出「正文引用 ↔ 文献」对照表（逐处核对引用是否真实支撑该句）")
    args = ap.parse_args()

    docx_list = [os.path.abspath(d) for d in args.docx]
    for d in docx_list:
        if not os.path.exists(d):
            print(f"找不到文档：{d}")
            sys.exit(1)
        if not d.lower().endswith(".docx"):
            print(f"只支持 .docx 文档（当前：{d}）。")
            print("若为 .doc 格式，请先在 Word 中「另存为」.docx 后再运行。")
            sys.exit(1)
    if args.main:
        main_abs = os.path.abspath(args.main)
        if main_abs not in docx_list:
            print(f"--main 不在 --docx 列表中：{args.main}")
            sys.exit(1)
        main_idx = docx_list.index(main_abs)
    else:
        main_idx = len(docx_list) - 1

    refs_path = args.refs or os.path.join(os.path.dirname(docx_list[0]), "refs.csv")
    if not os.path.exists(refs_path):
        print(f"找不到引用表：{refs_path}")
        print("提示：工作区结构为 <项目>/引用目录/refs.csv，请用 --refs 显式指定；")
        print("      或先运行 new_project.py 创建项目、add_refs.py 添加文献。")
        sys.exit(1)
    refs = load_refs(refs_path)
    if not refs:
        print(f"引用表为空：{refs_path}，先用 add_refs.py 添加文献。")
        sys.exit(1)
    print(f"引用表：{refs_path}（{len(refs)} 条）｜文档 {len(docx_list)} 份"
          f"（主文档：{os.path.basename(docx_list[main_idx])}）")

    # 1. 读入全部文档并收集引用点
    docs = []
    for d in docx_list:
        items, doc_root, styles_root = load_docx(d)
        body = doc_root.find(w("body"))
        if body is None:
            print(f"文档结构异常（无 body）：{d}")
            sys.exit(1)
        points = collect_points(body)
        docs.append({"path": d, "items": items, "root": doc_root,
                     "body": body, "styles_root": styles_root, "points": points})
    if not any(x["points"] for x in docs):
        print("正文中未找到引用点（[?] / [CITE:key] / 上次生成的引用超链接）。")
        print("在需要引用的句子末尾放 [?] 或 [CITE:key] 后重跑。")
        sys.exit(1)

    # 2. 全局编号（跨文档连续）
    flat = [pt for x in docs for pt in x["points"]]
    flat, key_to_num, used_keys = assign_numbers(flat, refs)
    idx = 0
    for x in docs:
        n = len(x["points"])
        x["points"] = flat[idx:idx + n]
        idx += n

    # 3. 更新/插入正文引用（样式取自主文档）
    bracket = args.bracket if len(args.bracket) == 2 else "[]"
    superscript = not args.no_superscript
    refs_by_key = {r["key"]: r for r in refs}
    year_suffix = {}
    if args.citation == "author-year":
        year_suffix = compute_year_suffixes(used_keys, refs_by_key)
    main_doc = docs[main_idx]
    heading_style_id = find_style_id(main_doc["styles_root"],
                                     ["heading 1", "Heading 1", "标题 1", "1"])
    hyperlink_style_id = find_style_id(main_doc["styles_root"],
                                       ["Hyperlink", "超链接"])
    total_new = 0
    for x in docs:
        total_new += insert_hyperlinks(x["points"], bracket, superscript,
                                       hyperlink_style_id, args.citation,
                                       refs_by_key, args.font_en, args.font_cn,
                                       year_suffix)

    # 4. 主文档参考文献表
    heading = find_or_create_heading(main_doc["body"], args.heading,
                                     heading_style_id)
    bookmark_id_base = 1
    for bs in main_doc["root"].iter(w("bookmarkStart")):
        try:
            bookmark_id_base = max(bookmark_id_base,
                                   int(bs.get(w("id"), 1)) + 1)
        except ValueError:
            pass
    entry_count = rebuild_bibliography(
        main_doc["body"], heading, used_keys, refs, args.style,
        args.entry_num, bookmark_id_base, args.citation,
        args.font_en, args.font_cn, not args.no_indent, args.align,
        args.bold_num, args.hanging_pt, not args.no_italic_source,
        items=main_doc["items"], hyperlink_style_id=hyperlink_style_id,
        year_suffix=year_suffix)

    # 5. 非主文档：引用超链接指向主文档（跨文档跳转）
    cross_count = 0
    for i, x in enumerate(docs):
        if i == main_idx:
            continue
        cross_count += add_cross_doc_links(x["items"], x["root"],
                                           main_doc["path"], x["path"])

    total_points = len(flat)
    print(f"引用点：{total_points} 处（新增 {total_new}、保留更新 {total_points - total_new}）"
          f"｜参考文献条目：{entry_count} 条"
          f"｜跨文档链接：{cross_count}")
    if args.citation == "numbered":
        print("编号方案：", ", ".join(f"{k}->{n}" for k, n in key_to_num.items()))
    else:
        print("引用样式：作者-年份制（--citation author-year）")

    # 5b. 引用对应性对照表（核对每处引用是否真实支撑所在句子）
    if args.mapping:
        print("\n== 正文引用 ↔ 文献对照表（逐处核对引用是否对得上） ==")
        for pt in points:
            key = pt["key"]
            ref = refs_by_key.get(key, {})
            num = pt.get("num")
            para = para_text(pt["paragraph"]).strip()
            if len(para) > 90:
                para = para[:90] + "…"
            tag = f"[{num}]" if num else "—"
            print(f"{tag} 段落：「{para}」")
            print(f"    ↳ key={key} → {(ref.get('title') or '')[:60]}"
                  f"（{(ref.get('year') or '')}）")
            note = (ref.get("note") or "").strip()
            if note:
                print(f"       note：{note[:100]}")
        print()

    if args.dry_run:
        print("--dry-run：未写文件。")
        return

    # 6. 备份并保存全部文档
    for x in docs:
        bpath = backup_docx(x["path"], args.backup_dir)
        save_docx(x["path"], x["items"], x["root"])
        print(f"已写入：{x['path']}（原文件备份：{bpath}）")

    # 7. 回读验证
    all_nums = []
    for i, x in enumerate(docs):
        items2, doc_root2, _ = load_docx(x["path"])
        body2 = doc_root2.find(w("body"))
        hyps = [el for el in body2.iter(w("hyperlink"))
                if is_ref_hyperlink(el)]
        txt_all = para_text(body2)
        leftovers = CITE_RE.findall(txt_all) + ANY_RE.findall(txt_all)
        tag = "（主文档）" if i == main_idx else ""
        print(f"回读验证[{os.path.basename(x['path'])}{tag}]："
              f"正文引用 {len(hyps)} 个｜残留占位符 {len(leftovers)} 个")
        if leftovers:
            print("警告：仍有占位符残留：", leftovers)
        if args.citation == "numbered":
            for h in hyps:
                m = re.search(r"\d+", "".join(run_text(r) for r in h.findall(w("r"))))
                if m:
                    all_nums.append(int(m.group()))
    if args.citation == "numbered":
        if set(all_nums) != set(range(1, entry_count + 1)):
            print(f"警告：编号不连续！全文正文编号 {sorted(all_nums)}，条目 {entry_count}，"
                  f"缺失 {set(range(1, entry_count + 1)) - set(all_nums)}")
        else:
            print("编号连续性验证通过（全文跨文档）。")
    bookmarks = [bs.get(w("name")) for bs in main_doc["root"].iter(w("bookmarkStart"))
                 if (bs.get(w("name")) or "").startswith("ref_")]
    print(f"文末书签 {len(bookmarks)} 个")
    if len(set(bookmarks)) != len(bookmarks):
        print("警告：书签名重复！")


if __name__ == "__main__":
    main()
