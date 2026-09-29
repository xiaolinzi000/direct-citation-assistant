# -*- coding: utf-8 -*-
"""
verify_support.py — 内容级核验：引用句是否真的被该文献摘要支持（论文引用 skill）

verify_refs.py 只核「书目信息真实、与权威库一致」；本脚本再进一步核
「正文这句话，是否真的被被引文献的内容（摘要）支撑」——防张冠李戴。

原理（确定性代码，非提示词判断）：
  1. 从 docx 提取每处引用所在的句子（句级切分，与 insert_refs 同一套逻辑）；
  2. 拉取该文献摘要：有 DOI 先查 Crossref（JATS XML 摘要）；
     取不到再查 OpenAlex（abstract_inverted_index 还原）；
  3. 引用句与摘要做 token 覆盖率：
       coverage = |引用句 token ∩ 摘要 token| / |引用句 token|
     覆盖率越高，该句被该文献摘要支持的可信度越高。
  4. 判定：
       ✅ 强支持   coverage ≥ 0.25
       ⚠️ 弱支持   coverage < 0.25（很可能张冠李戴 / 引错文，人工复核）
       🔶 无摘要   Crossref/OpenAlex 均无摘要（无法机器核验，人工核对）

用法：
  python verify_support.py --docx <项目>/论文/文稿.docx --refs <项目>/引用目录/refs.csv
  python verify_support.py --docx 第1章.docx 第2章.docx --refs refs.csv --report 支持核对.md
  python verify_support.py --docx 文稿.docx --refs refs.csv --min-coverage 0.3
                          # 自定义弱支持阈值（默认 0.25）
  python verify_support.py --docx 文稿.docx --refs refs.csv --evidence-csv
                          # v1.9：把逐句核验记录写回 refs.csv 的 evidence 列
                          # （论断|出处|核验深度|支持程度；无法确认标「待核实」，
                          #  不自动挪引文）；写入前自动备份 refs.csv

注意：
  - 覆盖率是启发式指标：摘要用词与正文用词不同时可能误报「弱」，
    误报时不代表引用错，按报告逐条人工判断即可；
  - 无摘要条目只能人工打开原文核对；
  - 联网行为统一走 api_client（退避 + 缓存，重跑不再占额度）。
"""

import argparse
import io
import os
import re
import sys
import urllib.parse
from collections import Counter

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from api_client import fetch_json
from refs_db import backup_refs, load_refs, save_refs
from insert_refs import (ANY_RE, CITE_RE, is_ref_hyperlink, load_docx,
                         locate_sentence, para_text, ref_offset_in_para,
                         run_text, split_sentences)

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def w(tag):
    return f"{{{W}}}{tag}"

# 英文常用停用词（摘要/正文共有的高频虚词不计入覆盖率）
STOP_EN = set("""a an and are as at be but by for from has have i in is it its
of on or that the this to was we were with our their you your will can may
not no nor than then so such which who whom while when where study studies
paper research result results method methods conclusion conclusions doi
article journal""".split())
# 中文常见字（去空格标点后逐字过滤）
STOP_CN = set("的了是在与及和等有被把将我们你们它们这那其并对而或但且还也又都只")


def norm_text(s):
    """只保留中文汉字与英文单词字符。"""
    return re.sub(r"[^0-9A-Za-z\u4e00-\u9fff]+", "", (s or "").lower())


def token_set(text, use_stop=True):
    """token 集合：英文按单词；中文按双字（bigram）避免单字覆盖率虚高。
    返回 (set, 计数)。"""
    t = norm_text(text)
    cjk = re.findall(r"[\u4e00-\u9fff]+", t)
    eng = re.findall(r"[a-z0-9]+", t)
    toks = []
    for chunk in cjk:
        if len(chunk) == 1:
            if use_stop and chunk not in STOP_CN:
                toks.append(chunk)
        else:
            for i in range(len(chunk) - 1):
                bigram = chunk[i:i + 2]
                if use_stop and bigram[0] in STOP_CN and bigram[1] in STOP_CN:
                    continue
                toks.append(bigram)
    for w in eng:
        if use_stop and w in STOP_EN:
            continue
        toks.append(w)
    return set(toks), len(toks)


def restore_inverted(inv):
    """OpenAlex abstract_inverted_index -> 正常文本。"""
    words = {}
    for w, poss in (inv or {}).items():
        for p in poss:
            words[p] = w
    return " ".join(words[i] for i in sorted(words)) if words else ""


def strip_jats(abstract):
    """Crossref abstract 常为 JATS XML 片段，剥掉标签还原纯文本。"""
    return re.sub(r"<[^>]+>", " ", abstract or "").strip()


def get_abstract(ref, use_cache=True):
    """拉取文献摘要：Crossref（JATS）→ OpenAlex（倒排索引）。返回 (text, 来源)。"""
    doi = (ref.get("doi") or "").strip()
    if doi:
        try:
            url = ("https://api.crossref.org/works/"
                   + urllib.parse.quote(doi))
            data = fetch_json(url, cache_key="support:crossref:" + doi,
                              use_cache=use_cache)
            ab = (data.get("message") or {}).get("abstract")
            if ab:
                text = strip_jats(ab)
                if len(text) > 40:
                    return text, "Crossref"
        except RuntimeError:
            pass
    # OpenAlex：优先按 DOI，失败按题名搜索
    for mode, q in (("doi", "doi:" + doi) if doi else ("title", None),):
        if q is None:
            continue
        try:
            url = ("https://api.openalex.org/works/"
                   + urllib.parse.quote(q)
                   + "?select=title,abstract_inverted_index")
            data = fetch_json(url, cache_key="support:openalex:" + q,
                              use_cache=use_cache)
            inv = (data or {}).get("abstract_inverted_index")
            text = restore_inverted(inv)
            if len(text) > 40:
                return text, "OpenAlex"
        except RuntimeError:
            pass
    try:
        url = ("https://api.openalex.org/works?search="
               + urllib.parse.quote((ref.get("title") or ""))
               + "&per-page=1&select=title,abstract_inverted_index")
        data = fetch_json(url, cache_key="support:openalex:title:" + (ref.get("title") or ""),
                          use_cache=use_cache)
        res = data.get("results") or []
        if res:
            inv = res[0].get("abstract_inverted_index")
            text = restore_inverted(inv)
            if len(text) > 40:
                return text, "OpenAlex"
    except RuntimeError:
        pass
    return None, None


def extract_citations(docx_list):
    """从各 docx 提取引用点：(doc, pidx, sidx, key, sentence)。"""
    rows = []
    for d in docx_list:
        items, root, _ = load_docx(d)
        body = root.find(w("body"))
        pidx = 0
        for p in body.iter(w("p")):
            pidx += 1
            full = para_text(p)
            if not full.strip():
                continue
            sents = split_sentences(full)
            marks = []
            for h in p.iter(w("hyperlink")):
                if is_ref_hyperlink(h):
                    key = (h.get(w("anchor")) or "")[4:]
                    off = ref_offset_in_para(p, key)
                    if off is not None:
                        marks.append((off, key))
            for m in re.finditer(r"\[CITE:([^\]\[]+)\]", full, re.IGNORECASE):
                marks.append((m.start(), m.group(1).strip()))
            for off, key in marks:
                si, stxt = locate_sentence(sents, off)
                if si is None:
                    si, stxt = 0, full[:120]
                rows.append((os.path.basename(d), pidx, si + 1, key,
                             stxt[:300]))
    return rows


def support_score(sentence, abstract):
    """引用句 vs 摘要的 token 覆盖率。返回 (coverage, 命中数, 引用句token数)。"""
    st, n_st = token_set(sentence)
    at, _ = token_set(abstract)
    if not st:
        return 0.0, 0, n_st
    inter = st & at
    return len(inter) / len(st), len(inter), n_st


def cross_language(sentence, abstract):
    """跨语言检测：引用句与摘要语言不一致（如中文句对英文摘要）时，
    自动比对不可用，应标注人工核对而非简单判弱。"""
    has_cjk_s = bool(re.search(r"[\u4e00-\u9fff]", sentence))
    has_cjk_a = bool(re.search(r"[\u4e00-\u9fff]", abstract))
    return has_cjk_s != has_cjk_a


def main():
    ap = argparse.ArgumentParser(
        description="内容级核验：引用句 × 文献摘要相似度（防张冠李戴）")
    ap.add_argument("--docx", required=True, nargs="+", help="Word 文档路径（可多份）")
    ap.add_argument("--refs", default="", help="refs.csv 路径（默认 docx 同目录）")
    ap.add_argument("--min-coverage", type=float, default=0.25,
                    help="弱支持判定阈值（默认 0.25；coverage 低于该值输出 ⚠️）")
    ap.add_argument("--report", default="", help="导出核验报告到 Markdown 文件")
    ap.add_argument("--evidence-csv", action="store_true",
                    help="v1.9：把逐句核验记录写回 refs.csv 的 evidence 列"
                         "（论断|出处|核验深度|支持程度；无法确认标「待核实」）")
    ap.add_argument("--refresh-cache", action="store_true",
                    help="忽略缓存强制实时拉摘要")
    args = ap.parse_args()

    docx_list = [os.path.abspath(d) for d in args.docx]
    for d in docx_list:
        if not os.path.exists(d):
            print(f"找不到文档：{d}")
            sys.exit(1)
    refs_path = args.refs or os.path.join(os.path.dirname(docx_list[0]), "refs.csv")
    if not os.path.exists(refs_path):
        print(f"找不到引用表：{refs_path}")
        sys.exit(1)
    refs = load_refs(refs_path)
    refs_by_key = {r["key"]: r for r in refs}
    if not refs:
        print(f"引用表为空：{refs_path}")
        sys.exit(1)

    print("== 内容级核验：引用句 ↔ 文献摘要 ==")
    rows = extract_citations(docx_list)
    if not rows:
        print("正文未找到可核验的引用（[CITE:key] 或脚本生成的引用超链接）。")
        print("提示：先运行 insert_refs.py 生成引用后再核验。")
        return

    results = []
    weak = 0
    noabs = 0
    for doc, pidx, sidx, key, sent in rows:
        ref = refs_by_key.get(key)
        if ref is None:
            print(f"⚠ [{doc} 段{pidx} 句{sidx}] key={key} 不在 refs.csv（引用表缺这条？）")
            results.append((doc, pidx, sidx, key, sent, "❓ 表中无此 key", "", "", ""))
            continue
        ab, src = get_abstract(ref, use_cache=not args.refresh_cache)
        if not ab:
            status = "🔶 无摘要"
            cov = hit = nst = ""
            noabs += 1
        elif cross_language(sent, ab):
            # 中英跨语言：token 覆盖率必然近 0，无法自动比对，标注人工核对
            status = "🌐 跨语言（人工核对）"
            cov = "—"
            hit = nst = ""
        else:
            cov, hit, nst = support_score(sent, ab)
            status = "✅ 强支持" if cov >= args.min_coverage else "⚠️ 弱支持"
            if cov < args.min_coverage:
                weak += 1
            cov = f"{cov:.2f}"
        title = (ref.get("title") or "")[:60]
        print(f"[{key}] {status}（{doc} 段{pidx} 句{sidx}）cover={cov} 命中={hit}/{nst}"
              f" 摘要={src or '—'}")
        print(f"    句：「{sent[:90]}」")
        print(f"    ↳ {title}（{ref.get('year') or ''}）")
        results.append((doc, pidx, sidx, key, sent, status, cov, hit, src))

    total = len(results)
    strong = sum(1 for r in results if r[5] == "✅ 强支持")
    xlang = sum(1 for r in results if r[5].startswith("🌐"))
    print(f"\n汇总：引用处 {total} ｜ 强支持 {strong} ｜ 弱支持 {weak} ｜ "
          f"跨语言待人工 {xlang} ｜ 无摘要 {noabs}")
    if weak:
        print("⚠ 存在弱支持引用：该句很可能未被该文献摘要支撑（张冠李戴/引错文），"
              "请打开原文逐条人工复核，必要时更换文献或删引用。")
    if xlang:
        print("🌐 跨语言引用（中文句对英文摘要等）无法自动比对覆盖率，请人工核对。")
    if noabs:
        print("🔶 无摘要条目无法机器核验，请人工打开原文核对。")

    if args.report:
        lines = ["# 引用内容支持度核验报告", "",
                 f"- 文档：{'、'.join(os.path.basename(d) for d in docx_list)}",
                 f"- 引用表：`{refs_path}`",
                 f"- 弱支持阈值：coverage < {args.min_coverage}",
                 "",
                 f"汇总：引用处 {total} ｜ 强支持 {strong} ｜ 弱支持 {weak} ｜ "
                 f"跨语言待人工 {xlang} ｜ 无摘要 {noabs}",
                 "",
                 "| 文档 | 段 | 句 | key | 状态 | 覆盖率 | 引用句 | 文献 |",
                 "|---|---|---|---|---|---|---|---|"]
        for doc, pidx, sidx, key, sent, status, cov, hit, src in results:
            sent = sent.replace("|", "｜")[:60]
            title = (refs_by_key.get(key) or {}).get("title", "")[:50].replace("|", "｜")
            lines.append(f"| {doc} | {pidx} | {sidx} | {key} | {status} | "
                         f"{cov} | {sent} | {title} |")
        with io.open(args.report, "w", encoding="utf-8") as f:
            f.write("\n".join(lines) + "\n")
        print(f"\n报告已导出：{args.report}")

    if args.evidence_csv:
        # v1.9：逐句证据核验记录写回 refs.csv 的 evidence 列
        # 格式：论断|出处（段/句）|核验深度|支持程度；无法确认标「待核实」，
        # 不自动挪引文。写入前自动备份。
        by_key = {}
        for doc, pidx, sidx, key, sent, status, cov, hit, src in results:
            by_key.setdefault(key, []).append((doc, pidx, sidx, sent, status, src))
        bpath = backup_refs(refs_path)
        changed = 0
        pending = 0
        for r in refs:
            k = r["key"]
            if k not in by_key:
                continue
            parts = []
            for doc, pidx, sidx, sent, status, src in by_key[k]:
                depth = "摘要" if src else "无法核验（无摘要）"
                sup = {"✅ 强支持": "强", "⚠️ 弱支持": "弱"}.get(status, "待核实")
                if sup == "待核实":
                    pending += 1
                claim = sent.replace("|", "｜").strip()[:80]
                parts.append(f"{claim}|{doc} 段{pidx} 句{sidx}|{depth}|{sup}")
            ev = "；".join(parts)
            if (r.get("evidence") or "").strip() != ev:
                r["evidence"] = ev
                changed += 1
        if changed:
            save_refs(refs, refs_path)
            print(f"\n已写回 evidence 列：{changed} 条（备份于 {bpath}）。")
        else:
            print("\nevidence 列无需更新（与现有记录一致）。")
        if pending:
            print(f"⚠ 其中 {pending} 处支持程度为「待核实」：请人工打开原文核对后"
                  "补正 evidence；不得自动挪引文。")


if __name__ == "__main__":
    main()
