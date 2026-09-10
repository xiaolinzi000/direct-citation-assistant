# -*- coding: utf-8 -*-
"""
new_project.py — 创建论文引用项目工作区（论文引用 skill）

用户使用本技能时：
  1. 先问清/确认工作文件夹；用户已指定 → 直接用。
  2. 未指定 → 运行本脚本，自动在 C 盘之外（按 D、E、F… 顺序找第一个
     可写盘符，找不到则退回用户主目录）创建项目文件夹，内含三个子文件夹：
       论文/      —— 存放 Word 论文文稿
       引用目录/   —— 存放引用信息表 refs.csv（引用目录列表）
       文献/      —— 存放对应的论文 PDF（能下载的尽量下载）
  3. 同时生成 refs.csv 模板（带表头，Excel 可直接打开）。

用法：
  python new_project.py --name "我的毕业论文_开题报告"
  python new_project.py --name "论文题目" --base "E:\文件1号"
  python new_project.py                        # 名称缺省用日期

说明：项目名称建议用「主题+阶段」（如 毕业论文_开题报告 / 期刊名_投稿），
      便于后期多个项目区分。创建后把 Word 文稿放进 论文/，引用信息写进
      引用目录/refs.csv，下载的论文 PDF 放 文献/（下载后可用 rename_papers.py
      按题目重命名归档，--in 指向 文献/ 的下载暂存处）。
"""

import argparse
import ctypes
import io
import os
import string
import sys
from datetime import datetime

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from refs_db import REFS_HEADER, save_refs

SUBDIRS = {"论文": "存放 Word 论文文稿",
           "引用目录": "存放引用信息表 refs.csv（引用目录列表）",
           "文献": "存放对应的论文 PDF（能下载的尽量下载）"}


def find_non_c_drive():
    """返回 C 盘之外第一个可写的盘符根目录（如 E:\\），找不到返回 None。"""
    drives = []
    bitmask = ctypes.windll.kernel32.GetLogicalDrives()
    for i, letter in enumerate(string.ascii_uppercase):
        if bitmask & (1 << i):
            drives.append(f"{letter}:\\")
    for d in drives:
        if d.upper().startswith("C:"):
            continue
        try:
            test = os.path.join(d, ".write_test_tmp")
            with open(test, "w") as f:
                f.write("ok")
            os.remove(test)
            return d
        except Exception:
            continue
    return None


def main():
    ap = argparse.ArgumentParser(description="创建论文引用项目工作区")
    ap.add_argument("--name", default="", help="项目名称（建议 主题_阶段）")
    ap.add_argument("--base", default="", help="创建位置（默认自动选 C 盘外第一个可写盘）")
    args = ap.parse_args()

    name = args.name.strip()
    if not name:
        name = f"论文引用_{datetime.now().strftime('%Y%m%d')}"

    if args.base.strip():
        base = args.base.strip().rstrip("\\/")
    else:
        d = find_non_c_drive()
        base = d.rstrip("\\/") if d else os.path.expanduser("~")
        if d:
            print(f"自动选择工作盘：{d}（C 盘之外第一个可写盘）")
        else:
            print("未找到 C 盘外的可写盘，退回用户主目录。")

    root = os.path.join(base, name)
    if os.path.exists(root) and os.listdir(root):
        print(f"⚠ 目标文件夹已存在且非空：{root}")
        print("  为避免覆盖已有内容，请换一个 --name，或确认后手动整理。")
        sys.exit(1)

    os.makedirs(root, exist_ok=True)
    for sub, desc in SUBDIRS.items():
        os.makedirs(os.path.join(root, sub), exist_ok=True)
    refs_path = os.path.join(root, "引用目录", "refs.csv")
    save_refs([], refs_path)

    # 简短使用说明
    note_path = os.path.join(root, "使用说明.txt")
    with io.open(note_path, "w", encoding="utf-8") as f:
        f.write(f"项目：{name}\n\n")
        f.write("三个子文件夹：\n")
        for sub, desc in SUBDIRS.items():
            f.write(f"  {sub}/ —— {desc}\n")
        f.write("\n使用流程：\n")
        f.write("1. Word 文稿放入 论文/；\n")
        f.write("2. 从 Google Scholar/知网等复制题录，用 add_refs.py 写入 引用目录/refs.csv；\n")
        f.write("3. 论文 PDF 下载到 文献/（可用 rename_papers.py 按题目重命名归档）；\n")
        f.write("4. 文稿中放 [?] 或 [CITE:key] 占位符，运行 insert_refs.py 生成引用和参考文献目录。\n")
        f.write(f"\n（技能脚本位置：{os.path.dirname(os.path.abspath(__file__))}）\n")

    print(f"\n✅ 项目工作区已创建：{root}")
    for sub, desc in SUBDIRS.items():
        print(f"   - {sub}/（{desc}）")
    print(f"   - 引用目录/refs.csv（模板已生成，含表头）")
    print(f"   - 使用说明.txt")


if __name__ == "__main__":
    main()
