# -*- coding: utf-8 -*-
"""
add_refs.py — 把收集到的引用信息加入 refs.csv（论文引用 skill）

用法：
  # 参数模式（推荐，可多次 --ref 添加多条，或单条字段）
  python add_refs.py --refs refs.csv ^
      --title "Attention is all you need" --authors "Vaswani, A., Shazeer, N., Parmar, N., et al." ^
      --source "Advances in Neural Information Processing Systems" --year 2017 ^
      --volume 30 --doi 10.5555/3295222.3295349 --type journal

  # 单条多字段也可分次
  python add_refs.py --refs refs.csv --title "..." ...

  # 交互模式（跟着提示填）
  python add_refs.py --refs refs.csv --interactive

  # 查看现有条目
  python add_refs.py --refs refs.csv --list

说明：Google Scholar 复制引用时直接取它的作者/标题/期刊/年/卷/期/页/DOI，
      作者串保留原样（如 "Smith, J., Jones, K."），格式化脚本会自动转换。
"""

import argparse
import io
import os
import re
import sys
from datetime import datetime

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from refs_db import (REFS_HEADER, VALID_TYPES, backup_refs, load_refs,
                     make_key, save_refs)


def _ask(prompt, default=""):
    if default:
        prompt = f"{prompt} [{default}]"
    v = input(prompt + ": ").strip()
    return v if v else default


def interactive(refs):
    print("—— 新增引用条目（直接回车跳过可选字段，key 留空则自动生成）——")
    r = {k: "" for k in REFS_HEADER}
    r["key"] = _ask("key（唯一标识，如 smith2023；留空自动生成）")
    r["type"] = _ask("类型", "journal")
    if r["type"] not in VALID_TYPES:
        print(f"  ⚠ 类型不合法（可选 {', '.join(VALID_TYPES)}），已重置为 journal")
        r["type"] = "journal"
    r["title"] = _ask("题名")
    r["authors"] = _ask("作者（Google Scholar 原样）")
    r["source"] = _ask("期刊/会议/出版社")
    r["year"] = _ask("年份")
    r["volume"] = _ask("卷")
    r["issue"] = _ask("期")
    r["pages"] = _ask("页码")
    r["doi"] = _ask("DOI")
    r["url"] = _ask("来源链接")
    r["note"] = _ask("支撑句子/用途说明")
    if not r["title"]:
        print("未填题名，放弃添加。")
        return False
    if not r["key"]:
        r["key"] = make_key(r["title"], r["year"] or None, [x["key"] for x in refs])
    keys = {x["key"] for x in refs}
    if r["key"] in keys:
        print(f"key 已存在：{r['key']}（已跳过，可用 --force 覆盖）")
        return False
    refs.append(r)
    return True


def main():
    ap = argparse.ArgumentParser(description="把引用信息加入 refs.csv")
    ap.add_argument("--refs", default="refs.csv", help="引用信息表路径")
    ap.add_argument("--interactive", action="store_true", help="交互模式逐项录入")
    ap.add_argument("--list", action="store_true", help="列出当前条目")
    ap.add_argument("--force", action="store_true", help="key 冲突时覆盖")
    for k in REFS_HEADER:
        ap.add_argument(f"--{k}", default="", help=f"字段 {k}")
    args = ap.parse_args()

    refs = load_refs(args.refs)
    if args.list:
        for i, r in enumerate(refs, 1):
            print(f"{i}. [{r['key']}] {r['title']} ({r['year']})")
        return

    changed = False
    if args.interactive:
        while interactive(refs):
            changed = True
            again = input("继续添加？(y/N): ").strip().lower()
            if again != "y":
                break
    else:
        r = {k: getattr(args, k, "") for k in REFS_HEADER}
        if not r["title"]:
            print("缺少 --title，无法添加。用 --interactive 交互录入，或 --help 看字段。")
            sys.exit(1)
        if r["type"] and r["type"] not in VALID_TYPES:
            print(f"类型不合法：{r['type']}（可选 {', '.join(VALID_TYPES)}）")
            sys.exit(1)
        if not r["key"]:
            r["key"] = make_key(r["title"], r["year"] or None, [x["key"] for x in refs])
        keys = {x["key"] for x in refs}
        if r["key"] in keys and not args.force:
            print(f"key 已存在：{r['key']}，未添加（--force 覆盖）")
            return
        refs = [x for x in refs if not (args.force and x["key"] == r["key"])]
        refs.append(r)
        changed = True

    if changed:
        b = backup_refs(args.refs)
        save_refs(refs, args.refs)
        print(f"已写入 {args.refs}（当前 {len(refs)} 条）")
        if b:
            print(f"原表已备份：{b}")
    else:
        print("无变更。")


if __name__ == "__main__":
    main()
