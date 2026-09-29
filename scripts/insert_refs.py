# -*- coding: utf-8 -*-
"""
insert_refs.py — 在 Word 文档中插入/更新带超链接的论文引用（论文引用 skill）

功能（对应需求 5、6）：
  1. 正文中的引用点替换为可点击跳转的编号引用（如 [1]，默认上标）；
     点击编号即可跳转到文末对应文献条目（Word 中默认 Ctrl+点击，可在
     文件→选项→高级→「用 Ctrl+单击跟踪超链接」取消勾选后改为单击跳转）。
  2. 文末自动生成参考文献目录：在已有的 References/参考文献 标题后生成；
     没有该标题则自动创建。定位改进：取「最后一个」匹配标题，自动跳过
     TOC（目录）区域/样式，避免把条目误插进目录。
  3. 自动重编号：删除中间某处引用后重跑本脚本，其余编号自动连续更新；
     同一篇论文多次引用只占一个编号（文末只列一次）。

v1.8 新增（句级落点 + 严格引用，默认行为）：
  4. 句级切分：脚本按句切分段落并定位每个引用点所在句子；
     占位符若误放在句号之后，自动归位到句末标点之前（编号紧跟引文）。
  5. 堆叠告警：插入后自动扫描「单句/单段堆叠编号」（默认阈值 单句 3、
     单段 5），提醒把编号拆分到各自支撑的句子，不在段末打包。
  6. 编号状态与增量更新：运行状态（key→编号）落盘到
     <主文档同目录>/_refs_state/，每次运行输出与上次的「编号映射 diff」
     （新增/移除/编号变化）；--freeze 锁定编号（key 恒为原编号，
     只做增量增删），重排不再是每次全量重编号。
  7. 变更审计报告：--audit <路径>.md 输出新增/移除/移位引用句、
     编号映射与堆叠告警，配合 _backup/ 一键回滚。

v1.9 新增（验收门槛硬化，未通过即报告失败）：
  8. 链接完整性检查：逐条确认正文引用超链接指向文末对应条目，
     文内编号与条目一一对应，无孤立编号/悬空链接；AI 改完文档后
     重跑或独立运行 verify_links.py 再检查一次。未通过默认退出码 1。
  9. DOI 保留与检查：有可靠 DOI 一律保留并渲染可点击链接；来源表缺
     DOI 或 DOI 无法解析时输出明确报告（不静默删、不猜填）；--strict-doi
     时未通过直接失败。DOI 前缀解析改为精确正则（format_refs.normalize_doi）。
  10. IEEE/TIM 格式：--style ieee（IEEE Reference Guide；正文编号不上标、
      条目按 IEEE 规范：作者名缩写在前、期刊缩写、vol./no./pp.、doi:）。
  11. 逐句证据核验记录：refs.csv 新增 evidence 字段（论断|出处|核验深度|
      支持程度），verify_support.py --evidence-csv 可把核验结果写回；
      无法确认的标「待核实」，不自动挪引文。

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
  python insert_refs.py --docx 文稿.docx --freeze       # 冻结编号：key 恒为原编号
  python insert_refs.py --docx 文稿.docx --audit 变更.md  # 输出变更审计报告
  常用参数：
    --style gbt7714|apa|vancouver|mla|harvard   参考文献格式（默认 gbt7714）
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
    --freeze                            冻结编号（key 恒为原编号，仅增量增删）
    --no-reuse-gaps                     freeze 下新 key 不复用空号（用最大号+1）
    --stack-sentence N                  单句堆叠告警阈值（默认 3）
    --stack-paragraph N                 单段堆叠告警阈值（默认 5）
    --no-stack-warning                  关闭堆叠告警
    --audit 路径                         输出变更审计报告（Markdown）
    --no-state                          不读写编号状态（无 diff/无 freeze）
    --state 路径                         自定义状态文件路径
"""

import argparse
import copy
import io
import json
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


# ---------- 句级切分（v1.8：句级落点 + 堆叠统计）----------

# 常见缩写点（"." 前是该词时不当作句末标点）
_ABBREV = {"e.g", "i.g", "i.e", "et", "al", "dr", "mr", "mrs", "ms", "vs",
           "fig", "figs", "eq", "eqs", "no", "vol", "pp", "ca", "approx",
           "inc", "ltd", "co", "dept", "univ", "st", "ave", "jr", "sr",
           "esp", "cf", "ibid", "eds", "ed", "trans", "p", "pp", "sec",
           "ch", "chap", "ref", "refs"}
# 句末标点
_SENT_END = ".。！？!?"
# 切分点后紧跟并归入前一句的闭合符号
_CLOSERS = "）】」』”’>)]}\u300b\u3009\uff62"


def split_sentences(text):
    """把文本按句末标点切分为 [(start, end, sentence_text)]。

    规则（保守，避免误切）：
      - 小数点 / 版本号（数字.数字）不切；
      - 常见缩写点（e.g.、et al.、fig. 等）不切；
      - 省略号（连续 3+ 个点）不切；
      - 切分点后紧跟的闭合引号/括号归入前一句。
    返回值用于：定位引用点所在句、单句/单段堆叠统计、句级映射报告。
    """
    sents = []
    n = len(text)
    start = 0
    i = 0
    while i < n:
        c = text[i]
        if c not in _SENT_END:
            i += 1
            continue
        if c == ".":
            prev = text[i - 1] if i > 0 else ""
            nxt = text[i + 1] if i + 1 < n else ""
            if prev.isdigit() and nxt.isdigit():
                i += 1
                continue  # 小数点/版本号
            if i + 1 < n and text[i + 1] == ".":
                j = i
                while j < n and text[j] == ".":
                    j += 1
                i = j
                continue  # 省略号（连续点）
            # 缩写检查：向前取连续词（允许词内含点，如 e.g.）
            tok_start = i
            while (tok_start > 0 and
                   (text[tok_start - 1].isalnum() or
                    text[tok_start - 1] in "-'.")):
                tok_start -= 1
            tok = text[tok_start:i].lower()
            if tok in _ABBREV:
                i += 1
                continue
        # 切分点后紧跟闭合符号则并入前一句
        j = i + 1
        while j < n and text[j] in _CLOSERS:
            j += 1
        sents.append((start, j, text[start:j].strip()))
        i = j
        start = j
    if start < n:
        sents.append((start, n, text[start:n].strip()))
    return sents


def locate_sentence(sents, offset):
    """返回 offset 所在句的 (句索引, 句子文本)；找不到返回 (None, "")。"""
    for idx, (s, e, txt) in enumerate(sents):
        if s <= offset < e:
            return idx, txt
    return None, ""


def _snap_anchor_before_punct(anchor):
    """句内落点校正：占位符若误放在句末标点之后（如「句子。[CITE]」），
    把句末标点剥离到占位符之后——编号紧跟引文、置于句末标点之前（规范）。

    规则（保守）：
      - 跳过 _strip_placeholder 留下的空 run 找前一个有文字的 run；
      - 仅当该 run 以句末标点结尾且与锚点之间无其他文字时调整；
      - 调整方式：标点从原 run 剥离，新建同格式 run 放在锚点之后。"""
    prev = anchor.getprevious()
    while prev is not None and run_text(prev) == "":
        prev = prev.getprevious()
    if prev is None:
        return
    t = run_text(prev)
    if not t or t[-1] not in _SENT_END:
        return
    # 检查 prev 与锚点之间只允许空 run（不允许其他文字）
    cur = prev
    while cur is not None and cur.getnext() is not anchor:
        cur = cur.getnext()
        if cur is None:
            return
        if run_text(cur) != "":
            return
    # 剥离尾标点：prev 去掉标点，标点作为新 run 放到锚点之后
    head, tail = t[:-1], t[-1]
    if head:
        _set_run_text(prev, head)
    else:
        _set_run_text(prev, "")
    p = anchor.getparent()
    punct_run = etree.Element(w("r"))
    rpr = prev.find(w("rPr"))
    if rpr is not None:
        punct_run.append(copy.deepcopy(rpr))
    tt = etree.SubElement(punct_run, w("t"))
    tt.text = tail
    anchor.addnext(punct_run)


def _ref_marks_in_para(p):
    """段内引用标记列表 [(offset, label)]：脚本生成的引用超链接文本
    （[n] / (Author, 2023)）与残留占位符（[CITE:key] / [?]）。"""
    marks = []
    acc = 0
    for child in p:
        if child.tag == w("hyperlink"):
            if is_ref_hyperlink(child):
                txt = "".join(run_text(r) for r in child.findall(w("r")))
                marks.append((acc, txt))
                acc += len(txt)
        elif child.tag == w("r"):
            txt = run_text(child)
            m = CITE_RE.search(txt) or ANY_RE.search(txt)
            if m:
                marks.append((acc + m.start(), m.group(0)))
            acc += len(txt)
    return marks


def ref_offset_in_para(p, key):
    """返回 key 对应引用标记在段文本中的字符偏移；找不到返回 None。"""
    acc = 0
    for child in p:
        if child.tag == w("hyperlink"):
            anchor = child.get(w("anchor")) or ""
            txt = "".join(run_text(r) for r in child.findall(w("r")))
            if anchor == f"ref_{key}":
                return acc
            acc += len(txt)
        elif child.tag == w("r"):
            txt = run_text(child)
            m = CITE_RE.search(txt)
            kk = m.group(1) if m else None
            if kk and kk.lower() == key.lower():
                return acc + m.start()
            acc += len(txt)
    return None


def scan_stacking(body, sent_th=3, para_th=5, citation_mode="numbered"):
    """扫描单句/单段堆叠编号，返回告警列表：
    [(段号, 类型('句'/'段'), 数量, 阈值, 编号列表, 文本片段)]"""
    warnings = []
    p_idx = 0
    for p in body.iter(w("p")):
        p_idx += 1
        full = para_text(p)
        if not full.strip():
            continue
        marks = _ref_marks_in_para(p)
        if not marks:
            continue
        labels = [lab for _, lab in marks]
        nums = []
        for lab in labels:
            m = re.search(r"\[(\d+)\]", lab)
            if m:
                nums.append(int(m.group(1)))
        if citation_mode == "numbered":
            total = len(nums)
            if total >= para_th:
                warnings.append((p_idx, "段", total, para_th, nums, full[:80]))
            sents = split_sentences(full)
            for sidx, (s, e, _t) in enumerate(sents):
                sn = [nm for nm, (so, _l) in zip(nums, marks)
                      if s <= so < e]
                if len(sn) >= sent_th:
                    warnings.append((p_idx, f"句{sidx + 1}", len(sn), sent_th,
                                     sn, full[s:e][:80]))
        else:
            # author-year：按引用标记数量统计
            if len(marks) >= para_th:
                warnings.append((p_idx, "段", len(marks), para_th,
                                 labels, full[:80]))
    return warnings


# ---------- TOC 区域检测（v1.8：避免把条目误插进目录）----------

def _para_has_toc_field(p):
    """段落含 TOC 域（fldSimple instr=TOC 或 instrText 含 TOC）。"""
    for it in p.iter(w("instrText")):
        if "TOC" in (it.text or "").upper():
            return True
    for fs in p.iter(w("fldSimple")):
        if "TOC" in (fs.get(w("instr")) or "").upper():
            return True
    return False


def _para_style_name(p, styles_root):
    """段落样式名（pStyle val -> style name）；无样式返回 ''。"""
    if styles_root is None:
        return ""
    ppr = p.find(w("pPr"))
    if ppr is None:
        return ""
    ps = ppr.find(w("pStyle"))
    if ps is None:
        return ""
    val = ps.get(w("val")) or ""
    if not val:
        return ""
    for s in styles_root.findall(w("style")):
        if s.get(w("styleId")) == val:
            nm = s.find(w("name"))
            if nm is not None and nm.get(w("val")):
                return nm.get(w("val"))
            break
    return val


def _para_in_toc_region(p, styles_root):
    """判断段落是否位于目录（TOC）区域：
    (a) 自身含 TOC 域；(b) 位于 fldSimple TOC 之内；(c) 样式名含 toc。"""
    cur = p
    while cur is not None:
        if cur.tag == w("fldSimple"):
            if "TOC" in (cur.get(w("instr")) or "").upper():
                return True
        cur = cur.getparent()
    if _para_has_toc_field(p):
        return True
    style = _para_style_name(p, styles_root)
    if style and "toc" in style.lower():
        return True
    return False


def toc_domain_set(body):
    """返回位于 TOC 复杂域（fldChar begin instrText=TOC … end）之间的段落集合。
    覆盖目录内容段（separate 与 end 之间的超链接行），修复
    「带目录文档把条目误插进 TOC」的踩坑场景。"""
    inside = False
    dom = set()
    for p in body.iter(w("p")):
        instr = [(it.text or "") for it in p.iter(w("instrText"))]
        has_toc_instr = any("TOC" in t.upper() for t in instr)
        flds = list(p.iter(w("fldChar")))
        has_begin = any(f.get(w("fldCharType")) == "begin" for f in flds)
        has_end = any(f.get(w("fldCharType")) == "end" for f in flds)
        if inside:
            dom.add(p)
        if has_begin and has_toc_instr:
            inside = True
        if has_end:
            inside = False
    return dom


def make_toc_filter(body, styles_root):
    """返回 TOC 区域判定函数 is_toc(p)（域内容段 + fldSimple + 样式名）。"""
    dom = toc_domain_set(body)

    def is_toc(p):
        if p in dom:
            return True
        return _para_in_toc_region(p, styles_root)
    return is_toc


# ---------- 编号状态（v1.8：diff / freeze）----------

STATE_DIR_NAME = "_refs_state"
STATE_VERSION = 2


def state_path(main_path, state_override=None):
    if state_override:
        return state_override
    d = os.path.join(os.path.dirname(os.path.abspath(main_path)), STATE_DIR_NAME)
    return os.path.join(d, os.path.splitext(os.path.basename(main_path))[0] + ".json")


def load_state(path):
    """读取状态文件。返回 (key_num dict, key_order list)；不存在/损坏返回 (None, None)。"""
    try:
        with io.open(path, "r", encoding="utf-8") as f:
            d = json.load(f)
        return d.get("key_num") or {}, d.get("key_order") or []
    except Exception:
        return None, None


def save_state(path, key_num, key_order):
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    data = {"version": STATE_VERSION,
            "saved_at": datetime.now().isoformat(timespec="seconds"),
            "key_num": key_num, "key_order": key_order}
    tmp = path + ".tmp"
    with io.open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=1)
    os.replace(tmp, path)


def diff_key_nums(old_map, new_map):
    """比较新旧编号映射，返回 (added, removed, changed, stable)：
      added   [(key, None, 新号)]         —— 本次新增引用
      removed [(key, 旧号, None)]         —— 本次移除引用
      changed [(key, 旧号, 新号)]         —— 编号发生变化（移位）
      stable  [key]                       —— 编号不变"""
    old, new = old_map or {}, new_map or {}
    added = [(k, None, new[k]) for k in new if k not in old]
    removed = [(k, old[k], None) for k in old if k not in new]
    changed = [(k, old[k], new[k]) for k in old
               if k in new and old[k] != new[k]]
    stable = [k for k in old if k in new and old[k] == new[k]]
    return added, removed, changed, stable


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

def assign_numbers(points, refs, freeze_key_num=None, reuse_gaps=True):
    """给引用点分配编号。

    非冻结（默认）：key 首次出现定号，重复引用同一 key 复用编号；
      'any' 占位符依次取未引用的条目（按 refs.csv 顺序）。
    冻结（--freeze，freeze_key_num 非空）：已冻结 key 恒为原编号；
      新出现的 key 分配「当前未占用的最小编号」（--no-reuse-gaps 时取最大号+1）；
      从正文消失的 key 编号释放（下次冻结可复用）。
    返回 (points, key_to_num, used_keys 顺序列表) 或抛错。"""
    key_to_num = {}
    used_keys = []
    any_pool = [r["key"] for r in refs]
    if freeze_key_num is None:
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
    # 冻结分支：已冻结 key 恒为原编号；新 key 分配「当前未被正文占用的
    # 最小编号」（--no-reuse-gaps 时取最大号+1）。已从正文消失的旧 key
    # 编号随之释放（不再占用空号）。
    freeze_key_num = freeze_key_num or {}
    maxn = max(freeze_key_num.values()) if freeze_key_num else 0
    # 第一遍：确定每个 point 的 key（[?] 自动分配），并收集正文实际占用的冻结号
    occupied = set()
    for pt in points:
        key = pt["key"]
        if pt["kind"] == "any":
            cands = [k for k in any_pool if k not in used_keys]
            if not cands:
                raise RuntimeError("正文 [?] 数量多于 refs.csv 中未引用的条目，"
                                   "请先 add_refs.py 补充文献。")
            key = cands[0]
            pt["key"] = key
        if key in freeze_key_num:
            occupied.add(freeze_key_num[key])
        if key not in used_keys:
            used_keys.append(key)
    # 第二遍：定号
    for pt in points:
        key = pt["key"]
        if key in freeze_key_num:
            key_to_num[key] = freeze_key_num[key]
        else:
            if reuse_gaps:
                n = 1
                while n in occupied or n in key_to_num.values():
                    n += 1
            else:
                maxn += 1
                n = maxn
            key_to_num[key] = n
            occupied.add(n)
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
        # 占位符：句内落点校正（误放句号后则归位到标点前），
        # 然后把超链接插到锚点 run 的位置，移除锚点
        anchor_el = pt["anchor"]
        _snap_anchor_before_punct(anchor_el)
        h = build_hyperlink(anchor, label, superscript, hyperlink_style_id,
                            font_en, font_cn)
        anchor_el.addprevious(h)
        anchor_el.getparent().remove(anchor_el)
        new_count += 1
    return new_count


# ---------- 参考文献表 ----------

def find_or_create_heading(body, heading_text, heading_style_id,
                           styles_root=None):
    """找标题段（文本归一化后匹配 TARGET_HEADINGS 或用户指定 heading_text），
       没有则创建。返回标题段元素。

    v1.8 定位改进：
      - 匹配全部候选后取【最后一个】匹配段（更接近文末，避免被正文中
        早出现的“References”字样误导）；
      - 自动跳过目录（TOC）区域：含 TOC 域 / 位于 fldSimple TOC 内 /
        样式名含 toc 的段不作为标题候选——修复“带目录文档把条目误插进 TOC”。
    """
    def _norm(s):
        # 去空白/冒号/句点/数字后缀，统一小写，便于宽容匹配
        s = re.sub(r"[\s:：.。·\-–—]+", "", (s or "").strip().casefold())
        return re.sub(r"\d+$", "", s)

    candidates = set()
    if heading_text:
        candidates.add(_norm(heading_text))
    for t in TARGET_HEADINGS:
        candidates.add(_norm(t))
    is_toc = make_toc_filter(body, styles_root)
    found = []
    for p in body.iter(w("p")):
        if _norm(para_text(p)) not in candidates:
            continue
        if is_toc(p):
            continue  # 目录里的“参考文献”条目，不是真实标题
        found.append(p)
    if found:
        return found[-1]  # 取最后一个（更接近文末）
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
                         year_suffix=None, num_map=None):
    """在标题段之后重建条目段：删除旧的（含 ref_ 书签的段），插入新条目。
       numbered   按出现顺序编号（num_map 提供时按 key 的冻结编号显示）；
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
        disp_num = (num_map or {}).get(key, num)
        p = build_entry_paragraph(disp_num, key, ref, style, entry_num_style, bid,
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


# ---------- 转占位符模式（交给 AI 修改前的保护） ----------

def _norm_heading_text(s):
    """标题归一化（与 find_or_create_heading 一致的去空白/标点/数字后缀）。"""
    s = re.sub(r"[\s:：.。·\-–—]+", "", (s or "").strip().casefold())
    return re.sub(r"\d+$", "", s)


def to_placeholders_mode(docx_list, backup_dir=None):
    """把正文引用超链接还原为 [CITE:key] 占位符纯文本，并删除文末参考文献表。

    用途：把生成好的文档交给 AI / 他人自由修改正文之前先跑一次。
    占位符 [CITE:key] 是普通文本，任何文字处理（AI 重写、python-docx
    重存、复制粘贴等）都不会破坏引用标记；对方改完后重新运行
    insert_refs.py（正常参数）即可恢复超链接、编号与文末表——彻底避免
    "AI 改完超链接就没了"的问题。

    返回转换的引用超链接数量。"""
    total = 0
    for d in docx_list:
        items, doc_root, styles_root = load_docx(d)
        body = doc_root.find(w("body"))
        # 1) 引用超链接 -> [CITE:key] 文本 run（保留原 run 的 rPr 格式）
        for h in list(body.iter(w("hyperlink"))):
            if not is_ref_hyperlink(h):
                continue
            key = (h.get(w("anchor")) or "")[4:]
            parent = h.getparent()
            rpr = None
            for r in h.findall(w("r")):
                rp = r.find(w("rPr"))
                if rp is not None:
                    rpr = rp
                    break
            new_r = etree.SubElement(parent, w("r"))
            if rpr is not None:
                new_r.append(copy.deepcopy(rpr))
            t = etree.SubElement(new_r, w("t"))
            t.text = f"[CITE:{key}]"
            parent.replace(h, new_r)
            total += 1
        # 2) 删除文末参考文献表（标题段之后的条目段：带 ref_ 书签的段
        #    + 紧跟其后的 [n] 开头段；遇到其他非空段即停止）
        #    标题定位与 insert 一致：跳过 TOC 区域，取最后一个匹配段
        target_heads = {_norm_heading_text(t) for t in TARGET_HEADINGS}
        is_toc = make_toc_filter(body, styles_root)
        head = None
        for p in body.iter(w("p")):
            if _norm_heading_text(para_text(p)) not in target_heads:
                continue
            if is_toc(p):
                continue
            head = p  # 取最后一个
        removed = 0
        if head is not None:
            after = False
            prev_entry = False
            for child in list(body):
                if after:
                    if child.tag == w("p"):
                        has_ref_bookmark = any(
                            (bs.get(w("name")) or "").startswith("ref_")
                            for bs in child.iter(w("bookmarkStart")))
                        txt = para_text(child).strip()
                        if has_ref_bookmark or (ENTRY_START_RE.match(txt) and prev_entry):
                            body.remove(child)
                            removed += 1
                            prev_entry = True
                        elif txt:
                            break
                        else:
                            prev_entry = False
                elif child is head:
                    after = True
                    prev_entry = False
        # 3) 备份并保存
        bpath = backup_docx(d, backup_dir)
        save_docx(d, items, doc_root)
        print(f"已转换：{d}（文末条目移除 {removed} 条；原文件备份：{bpath}）")
    return total


def write_audit_report(path, docs, main_idx, key_to_num, used_keys,
                       old_map, old_order, refs_by_key, citation_mode,
                       args, bpaths, stack_warns, num_diff):
    """生成变更审计报告（--audit）：新增/移除/移位引用句 + 编号映射 diff
    + 堆叠告警 + 备份清单，配合 _backup/ 一键回滚。"""
    lines = ["# 引用变更审计报告", "",
             f"- 生成时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
             f"- 文档：{'、'.join(os.path.basename(x['path']) for x in docs)}"
             f"（主文档：{os.path.basename(docs[main_idx]['path'])}）",
             f"- 引用表：{args.refs or '(docx 同目录)'}",
             f"- 样式：{args.style} / {args.citation}"
             f"{' / 冻结编号(--freeze)' if args.freeze else ''}",
             f"- 引用点：{sum(len(x['points']) for x in docs)} 处",
             f"- 文末条目：{len(used_keys)} 条", ""]
    # 1) 编号映射 diff
    lines.append("## 1. 编号映射 diff（vs 上次）")
    if num_diff:
        added, removed, changed, stable = num_diff
        if added or removed or changed:
            lines.append("| key | 旧编号 | 新编号 | 变更 |")
            lines.append("|---|---|---|---|")
            for k, o, n in added:
                lines.append(f"| {k} | — | {n} | 新增引用 |")
            for k, o, n in removed:
                lines.append(f"| {k} | {o} | — | 移除引用 |")
            for k, o, n in changed:
                lines.append(f"| {k} | {o} | {n} | 编号移位 |")
        else:
            lines.append("（无变化）")
    else:
        lines.append("（无上次状态可对比：首次运行 / --no-state）")
    lines.append("")
    # 2) 引用点清单（句级）
    lines.append("## 2. 引用点清单（句级落点）")
    lines.append("| 文档 | 段 | 句 | 编号 | key | 引用句 |")
    lines.append("|---|---|---|---|---|---|")
    for di, x in enumerate(docs):
        tag = "主" if di == main_idx else str(di + 1)
        pidx = 0
        for p in x["body"].iter(w("p")):
            pidx += 1
            pts = [pt for pt in x["points"] if pt["paragraph"] is p]
            if not pts:
                continue
            full = para_text(p)
            sents = split_sentences(full)
            for pt in pts:
                key = pt["key"]
                num = pt.get("num")
                off = ref_offset_in_para(p, key)
                si, stxt = locate_sentence(sents, off) \
                    if off is not None else (None, "")
                sent_disp = (stxt if si is not None else full)[:80]
                sent_disp = sent_disp.replace("|", "｜").replace("\n", " ")
                lines.append(f"| {tag} | {pidx} | {si + 1 if si is not None else '—'} | "
                             f"{num if num is not None else '—'} | {key} | {sent_disp} |")
    lines.append("")
    # 3) 堆叠告警
    lines.append("## 3. 堆叠告警")
    if stack_warns:
        for docname, (pidx, typ, cnt, th, nums, frag) in stack_warns:
            frag = frag.replace("|", "｜")
            lines.append(f"- {docname} 段{pidx} {typ}：{cnt} 个编号（阈值 {th}）"
                         f"{nums}「{frag}」")
    else:
        lines.append("（无）")
    lines.append("")
    # 4) 文末条目
    lines.append("## 4. 文末参考文献条目")
    lines.append("| 编号 | key | 文献 |")
    lines.append("|---|---|---|")
    key_to_ref = refs_by_key
    for num, key in enumerate(sorted(used_keys,
                                     key=lambda k: key_to_num.get(k, 0)), 1):
        ref = key_to_ref.get(key, {})
        title = (ref.get("title") or "")[:70].replace("|", "｜")
        lines.append(f"| {key_to_num.get(key, num)} | {key} | {title} |")
    lines.append("")
    # 5) 备份与回滚
    lines.append("## 5. 备份与一键回滚")
    lines.append("本次修改前的原文件已备份到 `_backup/`：")
    for path, bpath in bpaths:
        lines.append(f"- `{os.path.basename(path)}` → `{bpath}`")
    lines.append("")
    lines.append("回滚：把对应备份文件复制回原路径覆盖即可（或在 Word 中关闭文档后操作）。")
    lines.append("")
    with io.open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")


# ---------- 链接完整性检查（v1.9：逐条验收，未通过即报告失败）----------

def verify_docx_links(docx_paths, main_idx, citation_mode="numbered",
                      key_to_num=None):
    """逐个确认文档引用体系完整：

      1. 悬空链接：每个正文引用超链接（anchor=ref_key）在主文档文末
         都有对应书签（ref_key）；
      2. 孤立条目：文末每个书签都有正文引用（没有"表里有、正文没引"）；
      3. 编号一一对应（编号制）：正文出现的编号集合与文末条目编号集合
         完全一致——无孤立编号、无悬空编号、无重复书签；
      4. 编号↔条目一致（编号制，需提供 key_to_num）：正文 [n] 的 key
         与文末 [n] 条目（按 key_to_num）一致，防止张冠李戴。

    返回 (ok, issues)；issues 为逐条描述列表。"""
    issues = []
    body_anchors = {}    # key -> [num or label]
    body_nums = set()
    for di, path in enumerate(docx_paths):
        try:
            _items, doc_root, _styles = load_docx(path)
        except Exception as e:
            issues.append(f"无法读取文档 {os.path.basename(path)}：{type(e).__name__}")
            continue
        body = doc_root.find(w("body"))
        if body is None:
            issues.append(f"{os.path.basename(path)}：文档无 body（结构异常）")
            continue
        for h in body.iter(w("hyperlink")):
            if not is_ref_hyperlink(h):
                continue
            anchor = h.get(w("anchor")) or ""
            key = anchor[4:] if anchor.startswith("ref_") else anchor
            txt = "".join(run_text(r) for r in h.findall(w("r"))).strip()
            m = re.search(r"\[(\d+)\]", txt)
            num = int(m.group(1)) if m else txt
            body_anchors.setdefault(key, []).append(num)
            if isinstance(num, int):
                body_nums.add(num)
    # 主文档文末条目：书签 + 条目编号
    main_items, main_root, _ms = load_docx(docx_paths[main_idx])
    main_body = main_root.find(w("body"))
    entry_bookmarks = {}
    for bs in main_root.iter(w("bookmarkStart")):
        nm = bs.get(w("name")) or ""
        if nm.startswith("ref_"):
            entry_bookmarks[nm[4:]] = True
    entry_nums = set()
    if main_body is not None:
        for p in main_body.iter(w("p")):
            has_ref = any((bs.get(w("name")) or "").startswith("ref_")
                          for bs in p.iter(w("bookmarkStart")))
            if not has_ref:
                continue
            txt = para_text(p).strip()
            m = re.match(r"\[?(\d+)\]?", txt)
            if m:
                entry_nums.add(int(m.group(1)))
    # 1) 悬空链接：正文 anchor 无对应文末书签
    for key, nums in sorted(body_anchors.items()):
        if key not in entry_bookmarks:
            issues.append(f"悬空链接：正文引用 key={key}（{nums}）在文末没有对应条目书签")
    # 2) 孤立条目：文末书签无正文引用
    for key in sorted(entry_bookmarks):
        if key not in body_anchors:
            issues.append(f"孤立条目：文末书签 ref_{key} 在正文中没有被引用")
    # 3) 编号一一对应（编号制）
    if citation_mode == "numbered":
        if body_nums != entry_nums:
            only_body = sorted(body_nums - entry_nums)
            only_entry = sorted(entry_nums - body_nums)
            if only_body:
                issues.append(f"孤立编号：正文出现但文末无条目 {only_body}")
            if only_entry:
                issues.append(f"悬空编号：文末有条目但正文无引用 {only_entry}")
        # 4) 编号↔条目一致
        if key_to_num:
            for key, nums in sorted(body_anchors.items()):
                want = key_to_num.get(key)
                if want is None:
                    continue
                for n in nums:
                    if isinstance(n, int) and n != want:
                        issues.append(
                            f"编号↔条目不一致：正文 key={key} 显示 [{n}]，"
                            f"但编号方案为 [{want}]（文末条目 {want} 对应 key={key}）")
    return not issues, issues


def missing_doi_entries(used_keys, refs_by_key):
    """列出缺 DOI / DOI 无法解析的条目（v1.9：不静默删、不猜填）。
    返回 [(key, title, 原始 doi 值)]；有值但无法解析的单独标注可疑。"""
    from format_refs import normalize_doi
    missing = []
    suspicious = []
    for key in used_keys:
        ref = refs_by_key.get(key, {})
        raw = (ref.get("doi") or "").strip()
        if not raw:
            missing.append((key, (ref.get("title") or "")[:60], ""))
        elif normalize_doi(raw) is None:
            suspicious.append((key, (ref.get("title") or "")[:60], raw))
    return missing, suspicious


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
                    choices=["gbt7714", "apa", "vancouver", "mla", "harvard",
                             "ieee"],
                    help="格式体系：gbt7714 / apa / vancouver / mla / harvard / "
                         "ieee（IEEE 期刊如 TIM 采用：正文编号不上标、条目按 "
                         "IEEE Reference Guide）")
    ap.add_argument("--heading", default="", help="参考文献标题文本（找不到时创建）")
    ap.add_argument("--no-superscript", action="store_true", help="正文编号不用上标")
    ap.add_argument("--superscript", action="store_true",
                    help="正文编号强制用上标（ieee 样式默认不上标，用此参数覆盖）")
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
    ap.add_argument("--to-placeholders", action="store_true",
                    help="把正文引用超链接转为 [CITE:key] 占位符纯文本并移除文末参考文献表："
                         "生成纯文本草稿交给 AI/他人修改正文，占位符是普通文本不会被破坏；"
                         "对方改完后重跑本脚本（正常参数）即恢复超链接、编号与文末表。"
                         "此模式不需要 --refs。")
    ap.add_argument("--freeze", action="store_true",
                    help="冻结编号：以 refs.csv 的 key 为准保持编号不变（collet2020 恒为 1），"
                         "只做增量增删——新增 key 取未占用的最小编号（--no-reuse-gaps 时取最大号+1），"
                         "从正文消失的 key 释放编号；适合交稿后维护、需要编号稳定对稿的场景。")
    ap.add_argument("--no-reuse-gaps", action="store_true",
                    help="--freeze 下新 key 不复用释放的空号（用当前最大编号+1）")
    ap.add_argument("--stack-sentence", type=int, default=3,
                    help="单句内堆叠编号告警阈值（默认 3）")
    ap.add_argument("--stack-paragraph", type=int, default=5,
                    help="单段内堆叠编号告警阈值（默认 5）")
    ap.add_argument("--no-stack-warning", action="store_true",
                    help="关闭句级/段级堆叠编号告警")
    ap.add_argument("--audit", default="",
                    help="输出变更审计报告到 Markdown 文件：新增/移除/移位引用句 + 编号映射 diff"
                         " + 堆叠告警 + 备份清单（配合 _backup/ 一键回滚）")
    ap.add_argument("--no-state", action="store_true",
                    help="不读写编号状态文件（无编号 diff、--freeze 不可用）")
    ap.add_argument("--state", default="",
                    help="自定义编号状态文件路径（默认 <主文档同目录>/_refs_state/<主文档名>.json）")
    ap.add_argument("--strict-doi", action="store_true",
                    help="文末条目存在缺 DOI / DOI 无法解析时直接失败退出"
                         "（默认只输出缺失清单，不静默删、不猜填）")
    ap.add_argument("--warn-only", action="store_true",
                    help="回读/链接完整性检查未通过时仅警告，不退出非零"
                         "（默认：未通过检查就报告失败）")
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
    # --to-placeholders：不需要 refs.csv，直接转占位符
    if args.to_placeholders:
        n = to_placeholders_mode(docx_list, args.backup_dir)
        print(f"\n已把 {n} 处引用超链接转为 [CITE:key] 占位符纯文本，文末参考文献表已移除。")
        print("现在可把这份文档交给 AI/他人自由修改正文（占位符是普通文本，不会被破坏）；")
        print("改完后重新运行本脚本（正常参数，需带 --refs）即恢复超链接、编号和文末表。")
        return
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

    # 2. 编号状态（v1.8：diff / freeze）
    state_fp = None
    old_map, old_order = None, None
    if not args.no_state:
        state_fp = state_path(docx_list[main_idx], args.state or None)
        old_map, old_order = load_state(state_fp)
    freeze_key_num = old_map if args.freeze else None
    if args.freeze and args.citation != "numbered":
        print("注意：--freeze 仅对编号制有效（author-year 无编号），本次仅输出引用集合 diff。")
        freeze_key_num = None

    # 3. 全局编号（跨文档连续；--freeze 时以 key 冻结编号，仅增量增删）
    flat = [pt for x in docs for pt in x["points"]]
    flat, key_to_num, used_keys = assign_numbers(
        flat, refs, freeze_key_num,
        reuse_gaps=not args.no_reuse_gaps)
    idx = 0
    for x in docs:
        n = len(x["points"])
        x["points"] = flat[idx:idx + n]
        idx += n
    # freeze 下文末表按编号升序排列（正文编号仍按 key 定号）
    if args.freeze and args.citation == "numbered":
        used_keys = sorted(used_keys, key=lambda k: key_to_num[k])

    # 4. 更新/插入正文引用（样式取自主文档）
    bracket = args.bracket if len(args.bracket) == 2 else "[]"
    # ieee 样式按 IEEE 规范正文编号为行内方括号（不上标）；
    # 其他样式默认上标；--superscript / --no-superscript 可显式覆盖。
    superscript = (args.superscript
                   or (not args.no_superscript and args.style != "ieee"))
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

    # 5. 主文档参考文献表
    heading = find_or_create_heading(main_doc["body"], args.heading,
                                     heading_style_id,
                                     styles_root=main_doc["styles_root"])
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
        year_suffix=year_suffix,
        num_map=key_to_num if args.freeze and args.citation == "numbered" else None)

    # 5b. DOI 保留与检查（v1.9）：缺 DOI / DOI 可疑 → 明确报告，不静默
    missing_doi, suspicious_doi = missing_doi_entries(used_keys, refs_by_key)
    if missing_doi or suspicious_doi:
        print("\n⚠ DOI 检查：以下被引用的条目缺 DOI 或 DOI 无法解析（不静默删、不猜填）：")
        for key, title, _ in missing_doi:
            print(f"  [缺] {key}：{title}")
        for key, title, raw in suspicious_doi:
            print(f"  [疑] {key}：{title}（录入值无法解析：{raw[:60]}）")
        print("  处理：用 fetch_doi.py 自动补查（--apply 写回），或人工到出版社页面补录；")
        print("        verify_support.py / 文末条目会保留缺失状态，不会自动编造 DOI。")
        if args.strict_doi and (missing_doi or suspicious_doi):
            print("--strict-doi：存在缺 DOI / 可疑 DOI，按失败处理。")
            sys.exit(1)

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

    # 编号方案 + 与上次状态的 diff（v1.8：增量更新可视化）
    num_diff = None
    if args.citation == "numbered":
        print("编号方案：", ", ".join(f"{k}->{n}" for k, n in key_to_num.items()))
        if args.freeze:
            print("模式：--freeze 冻结（key 编号固定，仅增量增删）")
        if old_map is not None and old_map:
            added, removed, changed, stable = diff_key_nums(old_map, key_to_num)
            num_diff = (added, removed, changed, stable)
            print(f"编号 diff（vs 上次）：新增 {len(added)} ｜移除 {len(removed)} ｜"
                  f"变化 {len(changed)} ｜不变 {len(stable)}")
            for k, _o, n in added:
                print(f"  + {k} -> {n}（新增引用）")
            for k, o, _n in removed:
                print(f"  - {k}（原编号 {o}，本次移除）")
            for k, o, n in changed:
                print(f"  ~ {k}: {o} -> {n}（编号移位）")
    else:
        print("引用样式：作者-年份制（--citation author-year）")
        if old_order is not None:
            new_order = []
            for pt in flat:
                if pt["key"] not in new_order:
                    new_order.append(pt["key"])
            added = [k for k in new_order if k not in old_order]
            removed = [k for k in old_order if k not in new_order]
            num_diff = ([(k, None, None) for k in added],
                        [(k, None, None) for k in removed], [], [])
            if added or removed:
                print(f"引用集合 diff（vs 上次）：新增 {added} ｜移除 {removed}")

    # 5b. 引用对应性对照表（句级：每处编号对应哪篇文献、所在句，核对是否对得上）
    if args.mapping:
        print("\n== 正文引用 ↔ 文献对照表（句级，逐处核对引用是否对得上） ==")
        for di, x in enumerate(docs):
            tag = "（主文档）" if di == main_idx else ""
            pidx = 0
            for p in x["body"].iter(w("p")):
                pidx += 1
                pts = [pt for pt in x["points"] if pt["paragraph"] is p]
                if not pts:
                    continue
                full = para_text(p)
                sents = split_sentences(full)
                for pt in pts:
                    key = pt["key"]
                    ref = refs_by_key.get(key, {})
                    num = pt.get("num")
                    off = ref_offset_in_para(p, key)
                    si, stxt = locate_sentence(sents, off) \
                        if off is not None else (None, "")
                    if si is None:
                        sent_disp = full[:100] + ("…" if len(full) > 100 else "")
                    else:
                        sent_disp = stxt[:120] + ("…" if len(stxt) > 120 else "")
                    tag2 = f"{tag} 段{pidx}" + (f" 句{si + 1}" if si is not None else "")
                    label = f"[{num}]" if num else "—"
                    print(f"{label} {tag2}：「{sent_disp}」")
                    print(f"    ↳ key={key} → {(ref.get('title') or '')[:60]}"
                          f"（{(ref.get('year') or '')}）")
                    note = (ref.get("note") or "").strip()
                    if note:
                        print(f"       note：{note[:100]}")
                    ev = (ref.get("evidence") or "").strip()
                    if ev:
                        print(f"       evidence：{ev[:120]}")
        print()

    # 5c. 堆叠告警（句级落点检查：单句/单段堆叠编号提示拆分）
    stack_warns = []
    if not args.no_stack_warning:
        for di, x in enumerate(docs):
            warns = scan_stacking(x["body"], args.stack_sentence,
                                  args.stack_paragraph, args.citation)
            if warns:
                tag = "（主文档）" if di == main_idx else ""
                print(f"\n⚠ 堆叠告警[{os.path.basename(x['path'])}{tag}]"
                      f"——请把编号拆分到各自支撑的句子，避免段末打包：")
                for pidx, typ, cnt, th, nums, frag in warns:
                    print(f"  [段{pidx} {typ}] {cnt} 个编号（阈值 {th}）"
                          f"{nums}：「{frag}」")
                stack_warns.extend((os.path.basename(x["path"]), w)
                                   for w in warns)

    if args.dry_run:
        print("--dry-run：未写文件。")
        return

    # 6. 备份并保存全部文档
    bpaths = []
    for x in docs:
        bpath = backup_docx(x["path"], args.backup_dir)
        save_docx(x["path"], x["items"], x["root"])
        bpaths.append((x["path"], bpath))
        print(f"已写入：{x['path']}（原文件备份：{bpath}）")

    # 7. 更新编号状态（供下次 diff / --freeze 使用）
    if state_fp:
        order = []
        for pt in flat:
            if pt["key"] not in order:
                order.append(pt["key"])
        save_state(state_fp, key_to_num, order)
        print(f"编号状态已更新：{state_fp}")

    # 8. 变更审计报告（--audit）
    if args.audit:
        write_audit_report(args.audit, docs, main_idx, key_to_num, used_keys,
                           old_map, old_order, refs_by_key, args.citation,
                           args, bpaths, stack_warns, num_diff)
        print(f"变更审计报告：{args.audit}")

    # 9. 回读验证（v1.9：编号一致性 + 链接完整性逐条检查，未通过即报告失败）
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
    fail = []
    if args.citation == "numbered":
        expected = (set(key_to_num.values()) if args.freeze
                    else set(range(1, entry_count + 1)))
        if set(all_nums) != expected:
            fail.append(f"编号不一致：全文正文编号 {sorted(set(all_nums))}，"
                        f"期望 {sorted(expected)}，"
                        f"缺失 {sorted(expected - set(all_nums))}，"
                        f"多余 {sorted(set(all_nums) - expected)}")
        else:
            print("编号一致性验证通过（全文跨文档）。")
    bookmarks = [bs.get(w("name")) for bs in main_doc["root"].iter(w("bookmarkStart"))
                 if (bs.get(w("name")) or "").startswith("ref_")]
    print(f"文末书签 {len(bookmarks)} 个")
    if len(set(bookmarks)) != len(bookmarks):
        fail.append(f"书签名重复：{sorted(set(b for b in bookmarks if bookmarks.count(b) > 1))}")
    # v1.9：链接完整性逐条检查（悬空链接 / 孤立条目 / 孤立编号 / 编号↔条目一致）
    link_ok, link_issues = verify_docx_links(
        docx_list, main_idx, args.citation,
        key_to_num if args.citation == "numbered" else None)
    fail.extend(link_issues)
    if fail:
        print("\n❌ 回读验证未通过（未通过检查 = 报告失败）：")
        for iss in fail:
            print(f"  - {iss}")
        if args.warn_only:
            print("--warn-only：以上仅警告，本次不按失败处理。")
        else:
            print("请修复后重跑；若确认无误可加 --warn-only 放行。")
            sys.exit(1)
    else:
        print("链接完整性检查通过：正文引用↔文末条目逐条对应，"
              "无孤立编号/悬空链接/悬空条目。")


if __name__ == "__main__":
    main()
