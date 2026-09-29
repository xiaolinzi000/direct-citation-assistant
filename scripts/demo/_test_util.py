# -*- coding: utf-8 -*-
"""测试公共工具：构造最小 docx + 读取解析（论文引用 skill 回归测试用）"""
import os
import zipfile

from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W}

CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      '<Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" '
      'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '</Types>')

RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
        'Target="word/document.xml"/>'
        '</Relationships>')


def w(tag):
    return f"{{{W}}}{tag}"


def make_docx(path, body_inner_xml):
    """构造最小 docx。body_inner_xml 为 <w:body> 内部 XML 字符串（不含 body 标签）。"""
    document = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:body>' + body_inner_xml + '</w:body></w:document>')
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/document.xml", document)


def esc(s):
    """XML 文本转义（构造字符串型 body 时用）。"""
    return (s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))


def make_para(text, placeholder=None, heading=False):
    """构造 <w:p>。placeholder 为占位符文本（如 [CITE:key] / [?]）。
    heading=True 时带 Heading1 样式。"""
    ppr = ('<w:pPr><w:pStyle w:val="Heading1"/></w:pPr>' if heading else "")
    runs = []
    if text:
        runs.append(f'<w:r><w:t>{esc(text)}</w:t></w:r>')
    if placeholder:
        runs.append(f'<w:r><w:t>{esc(placeholder)}</w:t></w:r>')
    return f"<w:p>{ppr}" + "".join(runs) + "</w:p>"


def read_docx(path):
    """读取 docx 返回 document.xml 的 etree。"""
    with zipfile.ZipFile(path) as z:
        return etree.fromstring(z.read("word/document.xml"))


def para_text(p):
    return "".join(t.text or "" for t in p.iter(w("t")))


def paras(root):
    """返回 body 下全部段落元素。"""
    body = root.find(w("body"))
    return [p for p in body.iter(w("p"))]


def ref_anchors(root):
    """全部引用超链接 anchor 名。"""
    body = root.find(w("body"))
    return [h.get(w("anchor")) for h in body.iter(w("hyperlink"))
            if (h.get(w("anchor")) or "").startswith("ref_")]


def entry_numbers(root):
    """文末条目编号（带 ref_ 书签的段，取 [n] 或 n. 前缀）。"""
    import re
    body = root.find(w("body"))
    out = []
    for p in body.iter(w("p")):
        has_ref = any((bs.get(w("name")) or "").startswith("ref_")
                      for bs in p.iter(w("bookmarkStart")))
        if has_ref:
            m = re.match(r"\[?(\d+)\]?", para_text(p).strip())
            if m:
                out.append(int(m.group(1)))
    return out
