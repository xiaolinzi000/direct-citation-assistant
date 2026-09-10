# -*- coding: utf-8 -*-
"""场景测试：删除中间引用后自动重编号 + 新增引用后编号重排。
用法：先 make_demo.py 生成干净素材，再 insert_refs.py 一次，最后跑本脚本。"""
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
DOCX = os.path.join(HERE, "文稿2.docx")
REFS = os.path.join(HERE, "refs.csv")
INSERT = os.path.join(SCRIPTS, "insert_refs.py")


def w(tag):
    return f"{{{W}}}{tag}"


def run_insert():
    r = subprocess.run([sys.executable, INSERT, "--docx", DOCX, "--refs", REFS],
                       capture_output=True, text=True, encoding="utf-8", errors="replace")
    return r.stdout + r.stderr


def edit_docx(fn):
    """用 fn(items, root) 修改 docx 并保存。"""
    with zipfile.ZipFile(DOCX, "r") as z:
        items = {n: z.read(n) for n in z.namelist()}
    root = etree.fromstring(items["word/document.xml"])
    fn(items, root)
    items["word/document.xml"] = etree.tostring(root, xml_declaration=True,
                                                encoding="UTF-8", standalone=True)
    with zipfile.ZipFile(DOCX, "w", zipfile.ZIP_DEFLATED) as z:
        for n, d in items.items():
            z.writestr(n, d)


def para_text(p):
    return "".join(t.text or "" for t in p.iter(w("t")))


def summary():
    with zipfile.ZipFile(DOCX) as z:
        root = etree.fromstring(z.read("word/document.xml"))
    body = root.find(w("body"))
    refs = []
    for p in body.iter(w("p")):
        txt = para_text(p)
        m = re.search(r"\[(\d+)\]", txt)
        if m and any((bs.get(w("name")) or "").startswith("ref_")
                     for bs in p.iter(w("bookmarkStart"))):
            refs.append(m.group(1))
    hyps = [h.get(w("anchor")) for h in body.iter(w("hyperlink"))
            if (h.get(w("anchor")) or "").startswith("ref_")]
    body_txt = "".join(para_text(p) for p in body.iter(w("p")))
    return hyps, refs, body_txt


shutil.copy(os.path.join(HERE, "文稿.docx"), DOCX)
print("== 初始（3 处引用，编号 1,1,2）==")
print(run_insert().strip().splitlines()[-3:])
h, b, t = summary()
print(f"正文引用: {h} | 文末条目: {b}")

# 场景 A：删除中间引用（段落 3 的 [2]，即 devlin2019）
def delete_middle(items, root):
    body = root.find(w("body"))
    for p in list(body.iter(w("p"))):
        for h in list(p.findall(w("hyperlink"))):
            anchor = h.get(w("anchor")) or ""
            if anchor == "ref_devlin2019":
                p.remove(h)
edit_docx(delete_middle)
print("\n== 场景A：删除中间引用 devlin2019 后重跑 ==")
print(run_insert().strip().splitlines()[-3:])
h, b, t = summary()
print(f"正文引用: {h} | 文末条目: {b}")

# 场景 B：末尾新增一处 [?]（应自动取 zhang2023 编号 2）
def add_new_placeholder(items, root):
    body = root.find(w("body"))
    for p in body.iter(w("p")):
        if para_text(p).startswith("本文在前人工作"):
            r = etree.SubElement(p, w("r"))
            tt = etree.SubElement(r, w("t"))
            tt.text = "[?]"
edit_docx(add_new_placeholder)
print("\n== 场景B：末尾新增 [?] 后重跑（应按表序取 devlin2019 编为 2）==")
print(run_insert().strip().splitlines()[-3:])
h, b, t = summary()
print(f"正文引用: {h} | 文末条目: {b}")
assert "ref_devlin2019" in h, "devlin2019 未出现！"
assert h.count("ref_devlin2019") == 1 and h.count("ref_vaswani2017") == 2
print("\n✅ 自动重编号场景测试全部通过")
