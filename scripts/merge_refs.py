# -*- coding: utf-8 -*-
"""
merge_refs.py — 多章节论文合并：统一全局编号 + 合并文末参考文献表（论文引用 skill）

场景：论文分章写作，各章各自插入过引用（每章内部编号 1,2,3…），最后合成
一份完整 docx 时编号会冲突、文末表重复。本脚本一键合并：

  1. 读入各章 docx，移除各章自带文末参考文献表（标题 + 条目段）；
  2. 按章节顺序拼接正文（保留最后一章页面设置），书签 ID 全局重排，
     移除跨文档跳转关系（合并后为同一文档，用书签内部跳转）；
  3. 全文全局重新编号（key 首次出现定号，重复引用复用编号）；
  4. 在文末生成一份合并后的参考文献表；
  5. 输出「章节旧编号 → 全局新编号」映射表（逐处对照验收）；
  6. 备份 + 回读验证。

用法：
  python merge_refs.py --docx 第1章.docx 第2章.docx 第3章.docx ^
      --refs <项目>/引用目录/refs.csv --out 合并稿.docx

常用参数：--style gbt7714|apa|vancouver|mla|harvard、--citation numbered|author-year、
  --entry-num num|bracket、--heading 标题、--no-superscript、--mapping、
  --dry-run、--backup-dir 路径。

注意：各章文档请用 insert_refs.py 生成引用后再合并（脚本识别引用超链接与
[CITE:key] 占位符）；纯文本手打编号 [1] 无法识别，需先转占位符。
"""

import argparse
import io
import os
import re
import sys
import zipfile

from lxml import etree

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from refs_db import load_refs
from insert_refs import (ANY_RE, CITE_RE, ENTRY_START_RE, TARGET_HEADINGS,
                         assign_numbers, backup_docx, collect_points,
                         compute_year_suffixes, find_or_create_heading,
                         find_style_id, insert_hyperlinks, is_ref_hyperlink,
                         load_docx, make_toc_filter, para_text,
                         rebuild_bibliography, run_text, save_docx,
                         split_sentences, _norm_heading_text, w)

R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def remove_bibliography(body, styles_root):
    """移除文末参考文献表（最后一个 References 标题及其后的条目段）。
    返回移除的条目段数。"""
    is_toc = make_toc_filter(body, styles_root)
    heads = [p for p in body.iter(w("p"))
             if _norm_heading_text(para_text(p)) in
             {_norm_heading_text(t) for t in TARGET_HEADINGS}
             and not is_toc(p)]
    if not heads:
        return 0
    head = heads[-1]
    removed = 0
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
    body.remove(head)
    return removed


def collect_old_numbers(body):
    """收集章节内现有引用的旧显示编号：[key -> [旧编号, 段文本片段]]。"""
    out = {}
    pidx = 0
    for p in body.iter(w("p")):
        pidx += 1
        for h in p.findall(w("hyperlink")):
            if not is_ref_hyperlink(h):
                continue
            key = (h.get(w("anchor")) or "")[4:]
            txt = "".join(run_text(r) for r in h.findall(w("r"))).strip()
            m = re.search(r"\d+", txt)
            out.setdefault(key, []).append((int(m.group()) if m else "?",
                                            pidx, para_text(p)[:60]))
    return out


def renumber_bookmark_pairs(body):
    """书签 start/end ID 全局重排：start 按出现顺序编连续 ID（同名复用），
    end 按出现顺序编连续 ID。合并前各章书签 ID 可能重复，重排避免冲突。"""
    next_id = 1
    name_id = {}
    for bs in body.iter(w("bookmarkStart")):
        name = bs.get(w("name")) or ""
        if name and name in name_id:
            bs.set(w("id"), name_id[name])
        else:
            name_id[name] = str(next_id)
            bs.set(w("id"), str(next_id))
            next_id += 1
    n2 = 1
    for be in body.iter(w("bookmarkEnd")):
        be.set(w("id"), str(n2))
        n2 += 1


def merge_bodies(target_body, src_body):
    """把 src body 的子元素（除 sectPr）依次追加到 target_body。"""
    for child in list(src_body):
        if child.tag == w("sectPr"):
            continue
        target_body.append(child)


def main():
    ap = argparse.ArgumentParser(
        description="多章节论文合并：统一全局编号 + 合并文末参考文献表")
    ap.add_argument("--docx", required=True, nargs="+",
                    help="章节文档路径（按顺序合并；需先各自 insert_refs 生成引用）")
    ap.add_argument("--out", default="",
                    help="输出合并稿路径（默认 <第一章目录>/合并稿.docx）")
    ap.add_argument("--refs", default="", help="refs.csv 路径（默认 docx 同目录）")
    ap.add_argument("--style", default="gbt7714",
                    choices=["gbt7714", "apa", "vancouver", "mla", "harvard"])
    ap.add_argument("--citation", default="numbered",
                    choices=["numbered", "author-year"])
    ap.add_argument("--heading", default="", help="参考文献标题文本")
    ap.add_argument("--entry-num", default="bracket", choices=["bracket", "num"])
    ap.add_argument("--no-superscript", action="store_true")
    ap.add_argument("--bracket", default="[]")
    ap.add_argument("--font-en", default="Times New Roman")
    ap.add_argument("--font-cn", default="宋体")
    ap.add_argument("--no-indent", action="store_true")
    ap.add_argument("--hanging-pt", type=float, default=21.0)
    ap.add_argument("--align", default="both", choices=["both", "left", ""])
    ap.add_argument("--bold-num", action="store_true")
    ap.add_argument("--no-italic-source", action="store_true")
    ap.add_argument("--backup-dir", default="")
    ap.add_argument("--mapping", action="store_true")
    ap.add_argument("--dry-run", action="store_true")
    args = ap.parse_args()

    docx_list = [os.path.abspath(d) for d in args.docx]
    for d in docx_list:
        if not os.path.exists(d):
            print(f"找不到文档：{d}")
            sys.exit(1)
    out_path = os.path.abspath(args.out or os.path.join(
        os.path.dirname(docx_list[0]), "合并稿.docx"))
    refs_path = args.refs or os.path.join(os.path.dirname(docx_list[0]), "refs.csv")
    if not os.path.exists(refs_path):
        print(f"找不到引用表：{refs_path}")
        sys.exit(1)
    refs = load_refs(refs_path)
    if not refs:
        print(f"引用表为空：{refs_path}")
        sys.exit(1)
    print(f"章节 {len(docx_list)} 份 → 输出：{out_path}")
    print(f"引用表：{refs_path}（{len(refs)} 条）")

    # 1. 读入各章、移除文末表、收集旧编号
    merged_items, merged_root, merged_styles = load_docx(docx_list[0])
    merged_body = merged_root.find(w("body"))
    old_by_chapter = []
    for i, d in enumerate(docx_list):
        items, root, styles_root = load_docx(d)
        body = root.find(w("body"))
        removed = remove_bibliography(body, styles_root)
        old_map = collect_old_numbers(body)
        old_by_chapter.append((os.path.basename(d), old_map))
        if i == 0:
            continue  # 第一章即目标文档
        merge_bodies(merged_body, body)
        print(f"已并入：{d}（移除文末条目 {removed} 段，识别旧引用 "
              f"{sum(len(v) for v in old_map.values())} 处）")

    # 2. 清理：移除跨文档跳转 r:id、重排书签 ID
    for h in merged_body.iter(w("hyperlink")):
        if is_ref_hyperlink(h) and h.get(f"{{{R_NS}}}id"):
            del h.attrib[f"{{{R_NS}}}id"]
    renumber_bookmark_pairs(merged_body)

    # 3. 全局编号 + 重建引用
    points = collect_points(merged_body)
    if not points:
        print("正文中未找到引用点（[CITE:key] / 引用超链接）。")
        print("提示：各章请先跑 insert_refs.py 生成引用（或转占位符）后再合并。")
        sys.exit(1)
    flat, key_to_num, used_keys = assign_numbers(points, refs)
    refs_by_key = {r["key"]: r for r in refs}
    year_suffix = {}
    if args.citation == "author-year":
        year_suffix = compute_year_suffixes(used_keys, refs_by_key)
    bracket = args.bracket if len(args.bracket) == 2 else "[]"
    hyperlink_style_id = find_style_id(merged_styles, ["Hyperlink", "超链接"])
    new_count = insert_hyperlinks(points, bracket, not args.no_superscript,
                                  hyperlink_style_id, args.citation,
                                  refs_by_key, args.font_en, args.font_cn,
                                  year_suffix)
    heading_style_id = find_style_id(merged_styles, ["heading 1", "Heading 1",
                                                     "标题 1", "1"])
    heading = find_or_create_heading(merged_body, args.heading,
                                     heading_style_id, merged_styles)
    bid_base = 1
    for bs in merged_root.iter(w("bookmarkStart")):
        try:
            bid_base = max(bid_base, int(bs.get(w("id"), 1)) + 1)
        except ValueError:
            pass
    entry_count = rebuild_bibliography(
        merged_body, heading, used_keys, refs, args.style,
        args.entry_num, bid_base, args.citation,
        args.font_en, args.font_cn, not args.no_indent, args.align,
        args.bold_num, args.hanging_pt, not args.no_italic_source,
        items=merged_items, hyperlink_style_id=hyperlink_style_id,
        year_suffix=year_suffix)

    # 4. 章节旧编号 → 全局新编号 映射表
    print("\n== 章节旧编号 → 全局新编号 映射表 ==")
    for chap, old_map in old_by_chapter:
        for key, items in old_map.items():
            newn = key_to_num.get(key)
            for oldn, pidx, frag in items:
                title = (refs_by_key.get(key, {}) or {}).get("title", "")[:40]
                print(f"  {chap} 段{pidx} [{oldn}] -> [{newn}]"
                      f"（{key}：{title}）")
    print(f"\n引用点：{len(flat)} 处（新增 {new_count}）｜文末条目：{entry_count} 条")

    # 5. mapping（可选）
    if args.mapping:
        print("\n== 合并后正文引用 ↔ 文献对照表 ==")
        pidx = 0
        for p in merged_body.iter(w("p")):
            pidx += 1
            pts = [pt for pt in points if pt["paragraph"] is p]
            if not pts:
                continue
            for pt in pts:
                key = pt["key"]
                ref = refs_by_key.get(key, {})
                num = pt.get("num")
                full = para_text(p)
                print(f"[{num}] 段{pidx}：「{full[:80]}」")
                print(f"    ↳ {key} → {(ref.get('title') or '')[:60]}"
                      f"（{(ref.get('year') or '')}）")

    if args.dry_run:
        print("--dry-run：未写文件。")
        return

    # 6. 保存 + 回读验证
    if args.backup_dir:
        bpath = backup_docx(out_path, args.backup_dir) if os.path.exists(out_path) else None
    else:
        bpath = backup_docx(out_path) if os.path.exists(out_path) else None
    save_docx(out_path, merged_items, merged_root)
    print(f"已写入：{out_path}" + (f"（原文件备份：{bpath}）" if bpath else ""))

    items2, root2, _ = load_docx(out_path)
    body2 = root2.find(w("body"))
    hyps = [el for el in body2.iter(w("hyperlink")) if is_ref_hyperlink(el)]
    txt_all = para_text(body2)
    leftovers = CITE_RE.findall(txt_all) + ANY_RE.findall(txt_all)
    all_nums = []
    if args.citation == "numbered":
        for h in hyps:
            m = re.search(r"\d+", "".join(run_text(r) for r in h.findall(w("r"))))
            if m:
                all_nums.append(int(m.group()))
    print(f"回读验证：正文引用 {len(hyps)} 个｜残留占位符 {len(leftovers)} 个")
    if leftovers:
        print("警告：仍有占位符残留：", leftovers)
    if args.citation == "numbered":
        if set(all_nums) != set(range(1, entry_count + 1)):
            print(f"警告：编号不一致！正文 {sorted(all_nums)} vs 条目 {entry_count}")
        else:
            print("编号一致性验证通过。")


if __name__ == "__main__":
    main()
