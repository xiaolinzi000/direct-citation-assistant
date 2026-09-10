# -*- coding: utf-8 -*-
"""author-year 模式回归测试：正文 (作者, 年) 引用、文末字母序不编号、重跑幂等。
用法：先 make_demo.py 生成素材，再跑本脚本。"""
import os
import re
import shutil
import subprocess
import sys
import zipfile

from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
HERE = os.path.dirname(os.path.abspath(__file__))
SCRIPTS = os.path.dirname(HERE)
DOCX = os.path.join(HERE, "文稿.docx")
REFS = os.path.join(HERE, "refs.csv")
INSERT = os.path.join(SCRIPTS, "insert_refs.py")


def w(tag):
    return f"{{{W}}}{tag}"


def run_insert(extra=None):
    cmd = [sys.executable, INSERT, "--docx", DOCX, "--refs", REFS,
           "--citation", "author-year"]
    if extra:
        cmd += extra
    r = subprocess.run(cmd, capture_output=True, text=True,
                       encoding="utf-8", errors="replace")
    return r.stdout + r.stderr


def para_text(p):
    return "".join(t.text or "" for t in p.iter(w("t")))


def summary():
    with zipfile.ZipFile(DOCX) as z:
        root = etree.fromstring(z.read("word/document.xml"))
    body = root.find(w("body"))
    hyps = []
    for h in body.iter(w("hyperlink")):
        if not (h.get(w("anchor")) or ""):
            continue  # 跳过文末条目的 DOI 外部超链接
        txt = para_text(h).strip()
        if txt:
            hyps.append((h.get(w("anchor")) or "", txt))
    # 文末条目：含书签 ref_ 的段（书签名即 key），按出现顺序
    entry_keys = []
    for p in body.iter(w("p")):
        for bs in p.iter(w("bookmarkStart")):
            nm = (bs.get(w("name")) or "")
            if nm.startswith("ref_") and nm[4:] not in entry_keys:
                entry_keys.append(nm[4:])
    body_txt = "".join(para_text(p) for p in body.iter(w("p")))
    return hyps, entry_keys, body_txt


# 1) 第一次运行
print("== author-year 首次运行 ==")
print(run_insert().strip().splitlines()[-2:])
hyps, entry_keys, _ = summary()
print("正文引用:", [t for _, t in hyps])
print("文末条目顺序:", entry_keys)

# 断言 1：正文引用为 (Author, Year) 形式
for _, t in hyps:
    assert re.match(r"^\([^()]+,\s*\d{4}\)$", t), f"引用格式异常: {t}"
# 断言 2：文末按作者字母序 devlin < vaswani（zhang2023 未被正文引用，不出现）
expect = ["devlin2019", "vaswani2017"]
assert entry_keys == expect, f"文末顺序异常: {entry_keys}（期望 {expect}）"
# 断言 3：文末条目不编号（无 [n] 前缀）
with zipfile.ZipFile(DOCX) as z:
    root = etree.fromstring(z.read("word/document.xml"))
for p in root.iter(w("p")):
    if any((bs.get(w("name")) or "").startswith("ref_")
           for bs in p.iter(w("bookmarkStart"))):
        txt = para_text(p).strip()
        assert not txt.startswith("["), f"条目不应编号: {txt[:40]}"
print("✅ 断言1-3 通过（引用格式 / 字母序 / 不编号）")

# 2) 重跑幂等
print("\n== 重跑（幂等检查）==")
print(run_insert().strip().splitlines()[-2:])
hyps2, entry_keys2, body_txt2 = summary()
assert len(hyps2) == len(hyps), "重跑后引用数量变化！"
assert len(entry_keys2) == len(entry_keys), "重跑后条目数量变化！"
assert "[?]" not in body_txt2 and "[CITE:" not in body_txt2, "占位符残留！"
print(f"重跑后：正文引用 {len(hyps2)} 个 / 条目 {len(entry_keys2)} 条，无占位符残留")
print("✅ author-year 幂等测试通过")
