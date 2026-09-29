# -*- coding: utf-8 -*-
"""回归测试（v1.8 改进 7）：重音/撇号/XML 特殊字符贯穿 format→insert 全链路。

覆盖 Künzler、Dall'Angelo、O'Connor、Müller 等重音/撇号字符，以及
& < > 等 XML 特殊字符：此前需手工转义，本次固化为回归测试防止复发。
"""
import io
import os
import re
import subprocess
import sys
import zipfile

from lxml import etree

HERE = os.path.dirname(os.path.abspath(__file__))
SCRIPTS = os.path.dirname(HERE)
sys.path.insert(0, SCRIPTS)

from format_refs import _split_authors, format_ref, format_ref_segments
from insert_refs import load_docx, para_text
from demo._test_util import esc, make_docx, make_para, read_docx, ref_anchors

REFS = [
    {"key": "kunzler2020", "type": "journal",
     "title": "Künzler's method & its variants: a study of Dall'Angelo",
     "authors": "Künzler, P., Dall'Angelo, R., O'Connor, J.",
     "source": "Journal für Praktische Chemie", "year": "2020",
     "volume": "12", "issue": "3", "pages": "45-60",
     "doi": "10.1000/xyz.123", "url": "", "note": "特殊字符回归"},
    {"key": "muller2021", "type": "journal",
     "title": "Müller <und> Söhne: über die Wirkung",
     "authors": "Müller, M., Böhm, K.",
     "source": "Zeitschrift für Forschung", "year": "2021",
     "volume": "8", "issue": "", "pages": "1-10",
     "doi": "", "url": "", "note": "德文特殊字符"},
]


def write_refs(path):
    import csv
    with io.open(path, "w", encoding="utf-8-sig", newline="") as f:
        wtr = csv.DictWriter(f, fieldnames=list(REFS[0].keys()))
        wtr.writeheader()
        for r in REFS:
            wtr.writerow(r)


def xml_roundtrip(segs):
    """模拟 insert_refs 建 run：序列化后重新解析，断言文本无损。"""
    root = etree.fromstring(
        '<w:root xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '</w:root>')
    for text, italic in segs:
        r = etree.SubElement(root, "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r")
        t = etree.SubElement(r, "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t")
        t.text = text
    raw = etree.tostring(root)
    root2 = etree.fromstring(raw)
    got = "".join(tt.text or "" for tt in root2.iter(
        "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t"))
    return got


def check_format_layer():
    print("== 1) format 层：特殊字符不丢失 ==")
    ok = True
    # 每条 ref 只检查它自身应包含的字符。
    # 注意：GB/T 7714 按规范把西文作者姓大写（KÜNZLER、O'CONNOR），
    # 字符存在性用 casefold 比较；& < 为精确匹配。
    need_by_key = {
        "kunzler2020": (("Künzler", True), ("Dall'Angelo", True),
                        ("O'Connor", True), ("&", False)),
        "muller2021": (("Müller", True), ("Böhm", True), ("<", False),
                       ("für", True)),
    }
    for style in ("gbt7714", "apa", "vancouver", "mla", "harvard"):
        for r in REFS:
            text = format_ref(r, style)
            for needle, casefold in need_by_key[r["key"]]:
                hit = (needle.casefold() in text.casefold()) if casefold \
                    else (needle in text)
                if not hit:
                    # 各格式对 & < 的呈现可能不同（如 Vancouver 只输出 doi），
                    # & 与 < 的强校验放在 gbt7714（题名原样输出）
                    if needle in ("&", "<") and style != "gbt7714":
                        continue
                    print(f"  ✗ {style} 缺 {needle!r}：{text[:80]}")
                    ok = False
            segs = format_ref_segments(r, style)
            round_txt = xml_roundtrip(segs)
            plain = re.sub(r"[*]", "", format_ref(r, style))
            if round_txt != plain:
                print(f"  ✗ {style} XML 回读不一致\n    原：{plain}\n    回：{round_txt}")
                ok = False
    # 作者拆分：撇号不切断作者（_split_authors 会去除作者串尾点）
    lst = _split_authors("Künzler, P., Dall'Angelo, R., O'Connor, J.")
    assert len(lst) == 3, f"作者拆分错误：{lst}"
    assert "Dall'Angelo, R" in lst, f"撇号作者被破坏：{lst}"
    assert "O'Connor, J" in lst, f"撇号作者被破坏：{lst}"
    print(f"  ✓ 作者拆分保留撇号：{lst}")
    print("  ✓ format 层 5 种格式特殊字符全部保留" if ok else "  ✗ format 层有失败")
    return ok


def check_insert_layer(tmp):
    print("== 2) insert 端到端：docx 中特殊字符完整 ==")
    refs_path = os.path.join(tmp, "refs.csv")
    docx_path = os.path.join(tmp, "文稿.docx")
    write_refs(refs_path)
    body = make_para("本文参考了", "[CITE:kunzler2020]") + \
        make_para("以及", "[CITE:muller2021]")
    make_docx(docx_path, body)
    r = subprocess.run([sys.executable,
                        os.path.join(SCRIPTS, "insert_refs.py"),
                        "--docx", docx_path, "--refs", refs_path],
                       capture_output=True, text=True, encoding="utf-8",
                       errors="replace")
    out = r.stdout + r.stderr
    assert "编号一致性验证通过" in out, out[-800:]
    root = read_docx(docx_path)
    full = "".join(para_text(p) for p in root.iter(
        "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p"))
    # 特殊字符存在性（GB 大写作者用 casefold）；& 与 < 精确匹配
    for needle, casefold in (("Künzler", True), ("Dall'Angelo", True),
                             ("O'Connor", True), ("Müller", True),
                             ("für", True), ("Zeitschrift", True),
                             ("über", True), ("&", False), ("<", False)):
        hit = (needle.casefold() in full.casefold()) if casefold \
            else (needle in full)
        assert hit, f"docx 中缺失 {needle!r}\n全文：{full[:600]}"
    anchors = ref_anchors(root)
    assert set(anchors) == {"ref_kunzler2020", "ref_muller2021"}, anchors
    print(f"  ✓ 正文引用 {len(anchors)} 处，文末条目含全部特殊字符")
    print("  ✓ insert 端到端通过")
    return True


if __name__ == "__main__":
    import tempfile
    ok = check_format_layer()
    with tempfile.TemporaryDirectory() as tmp:
        ok = check_insert_layer(tmp) and ok
    print("\n✅ 特殊字符回归测试通过" if ok else "\n❌ 特殊字符回归测试失败")
    sys.exit(0 if ok else 1)
