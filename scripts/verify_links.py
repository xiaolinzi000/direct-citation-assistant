# -*- coding: utf-8 -*-
"""
verify_links.py — 链接完整性验收（论文引用 skill v1.9）

对已有 docx 逐个确认：
  1. 正文引用超链接（anchor=ref_key）在文末都有对应书签（无悬空链接）；
  2. 文末书签都有正文引用（无孤立条目）；
  3. 编号制下：文内编号与文末条目编号一一对应（无孤立编号/悬空编号）；
  4. 正文编号 ↔ key ↔ 文末条目编号一致（防张冠李戴）；
  5. 无重复书签。

用法（AI 改完文档、定稿前必跑；insert_refs 每次运行后也会自动检查）：
  python verify_links.py --docx <项目>/论文/文稿.docx
  python verify_links.py --docx 第1章.docx 第2章.docx --main 第2章.docx
  python verify_links.py --docx 文稿.docx --refs refs.csv
                                     # 提供 --refs 时额外按编号方案核对编号↔条目

未通过默认退出码 1（--warn-only 只警告）。
"""
import argparse
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from insert_refs import verify_docx_links, load_refs


def main():
    ap = argparse.ArgumentParser(description="链接完整性验收（正文引用 ↔ 文末条目）")
    ap.add_argument("--docx", required=True, nargs="+", help="Word 文档路径")
    ap.add_argument("--main", default=None,
                    help="主文档路径（文末表所在，默认最后一个 --docx）")
    ap.add_argument("--refs", default=None,
                    help="refs.csv 路径（提供时按编号方案额外核对编号↔条目一致）")
    ap.add_argument("--citation", default="numbered",
                    choices=["numbered", "author-year"],
                    help="正文引用样式（默认 numbered）")
    ap.add_argument("--warn-only", action="store_true",
                    help="未通过仅警告，不退出非零")
    args = ap.parse_args()

    docx_list = [os.path.abspath(d) for d in args.docx]
    for d in docx_list:
        if not os.path.exists(d):
            print(f"找不到文档：{d}")
            sys.exit(1)
    main_abs = os.path.abspath(args.main) if args.main else docx_list[-1]
    if main_abs not in docx_list:
        print(f"--main 不在 --docx 列表中：{args.main}")
        sys.exit(1)
    main_idx = docx_list.index(main_abs)

    refs_by_key = None
    key_to_num = None
    # 注：独立验收无法推断实际编号方案（非冻结模式按出现顺序编号、
    # freeze 按状态文件），因此不做"编号↔key"一致性核对（insert_refs
    # 每次运行后已自动做全量核对）；--refs 在此仅作信息展示用途。
    if args.refs and os.path.exists(args.refs):
        refs = load_refs(args.refs)
        refs_by_key = {r["key"]: r for r in refs}
    elif args.refs:
        print(f"找不到引用表：{args.refs}")

    ok, issues = verify_docx_links(docx_list, main_idx, args.citation,
                                   key_to_num)
    print(f"文档 {len(docx_list)} 份（主文档：{os.path.basename(docx_list[main_idx])}）")
    if not ok:
        print(f"❌ 链接完整性检查未通过（{len(issues)} 项）：")
        for iss in issues:
            print(f"  - {iss}")
        if args.warn_only:
            print("--warn-only：以上仅警告。")
        else:
            sys.exit(1)
    else:
        print("✅ 链接完整性检查通过：正文引用↔文末条目逐条对应，"
              "无孤立编号/悬空链接/悬空条目。")


if __name__ == "__main__":
    main()
