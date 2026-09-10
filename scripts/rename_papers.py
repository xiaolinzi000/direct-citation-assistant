# -*- coding: utf-8 -*-
"""
rename_papers.py — 下载论文按题目重命名归档（论文引用 skill）

把浏览器下载的论文 PDF 按 refs.csv 中的题名重命名，统一归档到 papers/：
  1) 自动匹配：扫描下载目录中的 PDF，与引用表题名做归一化比对，
     唯一命中即重命名并归档；命中多条/未命中则列出待处理。
  2) 手动指定：--by-key + --file 明确把某个 PDF 归到引用表的某条目。

用法：
  python rename_papers.py --refs refs.csv --in 下载目录 --out papers
  python rename_papers.py --refs refs.csv --by-key smith2023 --file 下载的.pdf
  python rename_papers.py --refs refs.csv --list           # 列出待匹配 PDF

说明：下载论文建议在浏览器中完成（Google Scholar/期刊页有下载按钮），
      下载完成后运行本脚本即可按题目命名归档。
"""

import argparse
import os
import re
import shutil
import sys
import unicodedata

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from refs_db import load_refs


def norm(s):
    """归一化：去空格/标点/大小写/全半角，便于比对。"""
    s = unicodedata.normalize("NFKC", s or "")
    s = s.lower()
    s = re.sub(r"[\s\-_\.:：,，、\(\)（）\[\]【】/\\'\"“”‘’]+", "", s)
    return s


def find_in_dir(directory, exts=(".pdf",)):
    if not os.path.isdir(directory):
        return []
    return [os.path.join(directory, f) for f in os.listdir(directory)
            if f.lower().endswith(exts) and not f.startswith("~")]


def match_refs(filepath, refs):
    """返回与文件匹配的 key 列表：先按文件名，再按 PDF 首页文本（标题出现）。"""
    base = os.path.splitext(os.path.basename(filepath))[0]
    n_base = norm(base)
    hits = []
    for r in refs:
        n_title = norm(r.get("title", ""))
        if not n_title:
            continue
        # 文件名包含题名主体，或题名包含文件名主体
        if n_title in n_base or n_base in n_title:
            hits.append(r["key"])
    if hits:
        return hits
    # 读 PDF 前几页文本再匹配
    try:
        import pypdf
        with open(filepath, "rb") as f:
            reader = pypdf.PdfReader(f)
            n_pages = min(len(reader.pages), 3)
            text = " ".join((reader.pages[i].extract_text() or "")
                            for i in range(n_pages))
        n_text = norm(text)
        for r in refs:
            n_title = norm(r.get("title", ""))
            if n_title and len(n_title) >= 15 and n_title in n_text:
                hits.append(r["key"])
    except Exception:
        pass
    return hits


def safe_name(title):
    title = title.replace(":", "：").replace("/", "／")
    return re.sub(r'[\\:*?"<>|]+', "_", title).strip()


def move_unique(src, dst):
    """移动文件；目标已存在时自动追加序号 _1/_2，避免覆盖。返回实际目标路径。"""
    if not os.path.exists(dst):
        shutil.move(src, dst)
        return dst
    base, ext = os.path.splitext(dst)
    n = 1
    while os.path.exists(f"{base}_{n}{ext}"):
        n += 1
    dst2 = f"{base}_{n}{ext}"
    shutil.move(src, dst2)
    return dst2


def main():
    ap = argparse.ArgumentParser(description="论文 PDF 按题目重命名归档")
    ap.add_argument("--refs", default="refs.csv")
    ap.add_argument("--in", dest="in_dir", default="papers_download",
                    help="下载目录（自动匹配时扫描这里）")
    ap.add_argument("--out", default="papers", help="归档目录")
    ap.add_argument("--by-key", default=None, help="指定引用表 key")
    ap.add_argument("--file", default=None, help="指定 PDF 文件路径（配 --by-key）")
    ap.add_argument("--list", action="store_true", help="只列出待匹配 PDF")
    args = ap.parse_args()

    refs = load_refs(args.refs)
    if not refs:
        print(f"引用表为空或不存在：{args.refs}")
        sys.exit(1)
    by_key = {r["key"]: r for r in refs}
    os.makedirs(args.out, exist_ok=True)

    if args.by_key:
        if args.by_key not in by_key:
            print(f"引用表无此 key：{args.by_key}")
            sys.exit(1)
        if not args.file:
            print("请用 --file 指定要归档的 PDF。")
            sys.exit(1)
        r = by_key[args.by_key]
        dst = os.path.join(args.out, safe_name(r["title"]) + ".pdf")
        dst = move_unique(args.file, dst)
        print(f"已归档：{args.file}\n  -> {dst}")
        return

    files = find_in_dir(args.in_dir)
    if not files:
        print(f"下载目录没有 PDF：{args.in_dir}")
        sys.exit(0)

    todo = []
    for f in files:
        hits = match_refs(f, refs)
        if len(hits) == 1:
            r = by_key[hits[0]]
            dst = os.path.join(args.out, safe_name(r["title"]) + ".pdf")
            dst = move_unique(f, dst)
            print(f"✓ {os.path.basename(f)}\n   -> {dst}")
        elif len(hits) > 1:
            print(f"? 多条候选：{os.path.basename(f)} -> {hits}")
            todo.append((f, hits))
        else:
            print(f"? 未匹配：{os.path.basename(f)}")
            todo.append((f, []))
    if todo and not args.list:
        print("\n未归档文件可用 --by-key <key> --file <路径> 手动指定。")
    print("完成。")


if __name__ == "__main__":
    main()
