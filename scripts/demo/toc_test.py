# -*- coding: utf-8 -*-
"""回归测试（v1.8 改进 3）：文末定位 —— 跳过 TOC 区域 + 取最后一个标题。

场景（已踩过的坑）：带目录的文档里 TOC 内容含「References」行，
旧逻辑按第一个匹配标题定位，把条目误插进目录；现在：
  - TOC 复杂域（fldChar begin instrText=TOC … end）内的段落不算标题；
  - 样式名含 toc 的段不算标题；
  - 多个候选时取【最后一个】匹配段（更接近文末）。
"""
import os
import subprocess
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
SCRIPTS = os.path.dirname(HERE)
sys.path.insert(0, SCRIPTS)

from demo._test_util import make_docx, read_docx, para_text, esc, w

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
INSERT = os.path.join(SCRIPTS, "insert_refs.py")

TOC_BEGIN = ('<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r>'
             '<w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u </w:instrText></w:r>'
             '<w:r><w:fldChar w:fldCharType="separate"/></w:r></w:p>')
TOC_END = '<w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>'


def has_entry_after(root, p_el):
    """p 段之后是否紧跟脚本生成的条目段（带 ref_ 书签）。"""
    body = root.find(w("body"))
    after = False
    for child in body:
        if child is p_el:
            after = True
            continue
        if not after:
            continue
        if child.tag != w("p"):
            continue
        if any((bs.get(w("name")) or "").startswith("ref_")
               for bs in child.iter(w("bookmarkStart"))):
            return True
        if para_text(child).strip():
            return False
    return False


if __name__ == "__main__":
    tmp = tempfile.mkdtemp(prefix="toc_test_")
    try:
        refs = os.path.join(tmp, "refs.csv")
        with open(refs, "w", encoding="utf-8-sig", newline="") as f:
            f.write("key,type,title,authors,source,year,volume,issue,pages,doi,url,note\n")
            f.write("k1,journal,Some title,Smith J.,Journal A,2020,1,1,1-2,10.1/a,,\n")
        docx = os.path.join(tmp, "目录稿.docx")
        body = (
            # 正文
            '<w:p><w:r><w:t>正文内容</w:t></w:r><w:r><w:t>[CITE:k1]</w:t></w:r></w:p>'
            # TOC 域：内容行里有 References（旧逻辑会误插这里）
            + TOC_BEGIN
            + '<w:p><w:hyperlink w:anchor="_Toc1"><w:r><w:t>References</w:t></w:r></w:hyperlink></w:p>'
            + TOC_END
            # 正文中间出现一个普通 References 字样段落（候选 1，靠前）
            + '<w:p><w:r><w:t>References</w:t></w:r></w:p>'
            # 文末真实标题（候选 2，最后一个）
            + '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr>'
            + '<w:r><w:t>References</w:t></w:r></w:p>'
        )
        make_docx(docx, body)
        r = subprocess.run([sys.executable, INSERT, "--docx", docx,
                            "--refs", refs],
                           capture_output=True, text=True, encoding="utf-8",
                           errors="replace")
        out = r.stdout + r.stderr
        assert "编号一致性验证通过" in out, out[-800:]
        root = read_docx(docx)
        body_root = root.find(w("body"))
        refs_paras = [p for p in body_root.iter(w("p"))
                      if para_text(p).strip() == "References"]
        assert len(refs_paras) >= 3, [para_text(p) for p in refs_paras]
        # TOC 内容行（第 1 个 References）之后不应有条目
        assert not has_entry_after(root, refs_paras[0]), "条目被插进 TOC！"
        # 正文中间 References（第 2 个）之后不应有条目（取最后一个）
        assert not has_entry_after(root, refs_paras[1]), "误用非末尾标题！"
        # 文末真实标题（最后一个）之后应有条目
        assert has_entry_after(root, refs_paras[-1]), "条目未插到文末标题后！"
        print("  ✓ TOC 内容行未插入条目")
        print("  ✓ 取最后一个 References 标题（非正文中间字样）")
        print("  ✓ 条目正确插在文末真实标题之后")
        print("\n✅ TOC 定位回归测试通过")
    finally:
        import shutil
        shutil.rmtree(tmp, ignore_errors=True)
