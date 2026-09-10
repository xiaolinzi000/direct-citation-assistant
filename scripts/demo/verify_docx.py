# -*- coding: utf-8 -*-
"""深度验证：解包 docx，检查超链接/书签/条目结构是否完整正确"""
import sys
import zipfile
from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W}


def w(tag):
    return f"{{{W}}}{tag}"


def run_text(r):
    return "".join(t.text or "" for t in r.iter(w("t")))


def para_text(p):
    return "".join(run_text(r) for r in p.iter(w("r")))


path = sys.argv[1]
with zipfile.ZipFile(path) as z:
    root = etree.fromstring(z.read("word/document.xml"))

body = root.find(w("body"))
print("===== 正文段落（含超链接文本）=====")
for p in body.iter(w("p")):
    txt = para_text(p)
    if txt.strip():
        print("|", txt[:120])

print("\n===== 超链接 =====")
for h in body.iter(w("hyperlink")):
    anchor = h.get(w("anchor"))
    print(f"  anchor={anchor}  文本={para_text(h)}")

print("\n===== 书签 =====")
for bs in root.iter(w("bookmarkStart")):
    name = bs.get(w("name")) or ""
    if name.startswith("ref_"):
        print(f"  {name}  (id={bs.get(w('id'))})")

# 验证锚点跳转完整性：每个正文超链接 anchor 都有对应书签
anchors = [h.get(w("anchor")) for h in body.iter(w("hyperlink"))]
names = {bs.get(w("name")) for bs in root.iter(w("bookmarkStart"))}
missing = [a for a in anchors if a not in names]
print(f"\n===== 完整性 =====")
print(f"超链接 {len(anchors)} 个 -> 书签 {len(names)} 个；无对应书签的：{missing or '无'}")
print("OK" if not missing else "FAIL")
