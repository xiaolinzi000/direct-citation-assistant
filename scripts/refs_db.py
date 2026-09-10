# -*- coding: utf-8 -*-
"""
refs_db.py — 引用信息库读写模块（论文引用 skill）

引用信息统一存放在 refs.csv（UTF-8 with BOM，Excel 可直接打开）：
  列：
    key       唯一标识（如 smith2023），用于正文占位符 [CITE:key]
    type      文献类型：journal(期刊) / conference(会议) / book(专著) /
              thesis(学位论文) / web(网页)
    title     题名
    authors   作者列表。中文用顿号/逗号分隔；英文建议 Google Scholar
              原样（如 "Smith, J., Jones, K."），格式化时会自动转 GB/T 7714 写法
    source    期刊名 / 会议名 / 出版社 / 网站名
    year      年份
    volume    卷
    issue     期
    pages     页码
    city      出版地/保存地（专著、学位论文、会议录的 GB/T 7714 格式需要；
              学位论文填保存地如"北京"，专著填出版地如"北京"）
    doi       DOI（不带 https://doi.org/）
    url       来源链接（Google Scholar / 原文页；网页[EB/OL] 类型必须填）
    note      支撑的句子/用途说明（写论文时填写）
"""

import csv
import io
import os
import re
import sys
from datetime import datetime

REFS_HEADER = ["key", "type", "title", "authors", "source", "year",
               "volume", "issue", "pages", "city", "doi", "url", "note"]

VALID_TYPES = ("journal", "conference", "book", "thesis", "web")


def default_refs_path(docx_path=None):
    """默认引用信息表路径：docx 同目录 refs.csv，否则当前目录 refs.csv"""
    base = os.path.dirname(os.path.abspath(docx_path)) if docx_path else os.getcwd()
    return os.path.join(base, "refs.csv")


def load_refs(path):
    """读取 refs.csv，返回 dict[refs] 列表。文件不存在返回 []。"""
    if not os.path.exists(path):
        return []
    with io.open(path, "r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        refs = []
        for row in reader:
            d = {k: (row.get(k) or "").strip() for k in REFS_HEADER}
            if d["key"]:
                refs.append(d)
        return refs


def save_refs(refs, path):
    """写入 refs.csv（UTF-8 BOM）。"""
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    with io.open(path, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=REFS_HEADER)
        writer.writeheader()
        for r in refs:
            writer.writerow({k: r.get(k, "") for k in REFS_HEADER})


def backup_refs(path, backup_dir=None):
    """备份 refs.csv（可选，写入前调用）。返回备份路径或 None。"""
    if not os.path.exists(path):
        return None
    target_dir = backup_dir or os.path.join(os.path.dirname(os.path.abspath(path)), "_backup")
    os.makedirs(target_dir, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    bpath = os.path.join(target_dir, f"{ts}__refs.csv")
    with io.open(path, "rb") as src, io.open(bpath, "wb") as dst:
        dst.write(src.read())
    return bpath


def make_key(title, year=None, existing=None):
    """根据题名自动生成 key：ref+年份。"""
    existing = existing or []
    base = "ref" + str(year) if year else "ref"
    key = base
    n = 1
    while key in existing:
        n += 1
        key = f"{base}_{n}"
    return key


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "--list":
        p = sys.argv[2] if len(sys.argv) > 2 else "refs.csv"
        for i, r in enumerate(load_refs(p), 1):
            print(f"{i}. [{r['key']}] {r['title']} ({r['year']})")
    else:
        print(__doc__)
