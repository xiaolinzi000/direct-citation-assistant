# -*- coding: utf-8 -*-
"""回归测试（v1.9）：验收门槛硬化 —— DOI 精确解析 / IEEE-TIM 格式 /
链接完整性逐条检查 / evidence 字段展示。

覆盖：
  A. format_refs.normalize_doi：6 类前缀统一解析，非法值返回 None
  B. format_refs.ieee：期刊（TIM 缩写、作者缩写+et al.、无 month 年份回退）、
     会议 / 专著 / 学位论文 / 网页条目
  C. insert_refs --style ieee 端到端：正文编号不上标、文末条目含期刊缩写
     与 doi: 前缀
  D. verify_docx_links：悬空链接 / 孤立条目 / 孤立编号 逐条检出
  E. verify_links.py 独立验收：正常文档通过、缺陷文档退出非零
"""
import os
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
SCRIPTS = os.path.dirname(HERE)
sys.path.insert(0, SCRIPTS)

from demo._test_util import (entry_numbers, make_docx, make_para,
                             read_docx, ref_anchors, para_text, w)
from format_refs import FORMATS, ieee, normalize_doi
from insert_refs import verify_docx_links, load_docx

REFS_SRC = os.path.join(HERE, "refs.csv")
INSERT = os.path.join(SCRIPTS, "insert_refs.py")
VERIFY_LINKS = os.path.join(SCRIPTS, "verify_links.py")

JOURNAL_REF = {
    "key": "haag2012", "type": "journal",
    "title": "A novel approach to precision measurement in industrial environments",
    "authors": "Haag, M., Uhlemann, F., Schrepf, P., Roth, H., Zhang, Y., Li, W., Wang, Q.",
    "source": "IEEE Transactions on Instrumentation and Measurement",
    "year": "2012", "volume": "61", "issue": "5", "pages": "1350-1358",
    "doi": "10.1109/TIM.2012.2186121",
}


def run(cmd, expect_ok=True):
    r = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8",
                       errors="replace")
    if expect_ok:
        assert r.returncode == 0, (r.stdout + r.stderr)[-1500:]
    return r.stdout + r.stderr


def hyperlink_has_superscript(docx_path):
    """检查正文引用超链接是否带 <w:vertAlign superscript/>。"""
    root = read_docx(docx_path)
    for h in root.iter(w("hyperlink")):
        if (h.get(w("anchor")) or "").startswith("ref_"):
            for r in h.findall(w("r")):
                rpr = r.find(w("rPr"))
                if rpr is not None:
                    va = rpr.find(w("vertAlign"))
                    if va is not None and va.get(w("val")) == "superscript":
                        return True
    return False


def body_xml(path):
    with zipfile.ZipFile(path) as z:
        return z.read("word/document.xml").decode("utf-8")


if __name__ == "__main__":
    tmp = tempfile.mkdtemp(prefix="v19_test_")
    try:
        # ---------- A. normalize_doi 精确解析 ----------
        assert normalize_doi("10.1109/TIM.2023.1234567") == "10.1109/TIM.2023.1234567"
        assert normalize_doi("https://doi.org/10.1109/TIM.2023.1234567") == \
            "10.1109/TIM.2023.1234567"
        assert normalize_doi("doi: 10.1109/TIM.2023.1234567") == \
            "10.1109/TIM.2023.1234567"
        assert normalize_doi("dx.doi.org/10.1109/TIM.2023.1234567") == \
            "10.1109/TIM.2023.1234567"
        assert normalize_doi("10.1109/TIM.2023.1234567.") == \
            "10.1109/TIM.2023.1234567"   # 尾随句点剥除
        assert normalize_doi("not a doi") is None
        assert normalize_doi("") is None
        print("A. normalize_doi 精确解析 ✅（7 断言）")

        # ---------- B. IEEE 渲染 ----------
        assert "ieee" in FORMATS
        s = ieee(JOURNAL_REF)
        assert "IEEE Trans. Instrum. Meas." in s, s   # 期刊官方缩写
        assert "M. Haag, F. Uhlemann, P. Schrepf" in s, s  # 名缩写在前
        assert "Q. Wang, et al." in s or "et al." in s, s  # 第 7 作者起 et al.
        assert "vol. 61, no. 5, pp. 1350-1358" in s, s
        assert "2012." in s, s   # 无 month 时回退年份
        assert "doi: 10.1109/TIM.2012.2186121" in s, s
        conf = ieee({"key": "k", "type": "conference", "title": "Conf paper",
                     "authors": "Zhang, Y., Li, W.",
                     "source": "Proc. IEEE Int. Conf. Meas.", "year": "2021",
                     "pages": "10-15", "city": "Beijing"})
        assert "Proc. IEEE Int. Conf. Meas." in conf and "Beijing" in conf, conf
        book = ieee({"key": "b", "type": "book", "title": "Sensors and signal processing",
                     "authors": "Brown, R.", "source": "Springer", "year": "2019",
                     "city": "New York"})
        assert "Springer" in book and "New York" in book, book
        web = ieee({"key": "w", "type": "web", "title": "IEEE Reference Guide",
                    "authors": "", "source": "IEEE Author Center",
                    "year": "2025", "url": "https://ieeeauthorcenter.ieee.org"})
        assert "IEEE Author Center" in web, web
        print("B. IEEE 渲染 ✅（期刊/会议/专著/网页 + 作者缩写 + 期刊缩写）")

        # ---------- C. insert_refs --style ieee 端到端 ----------
        refs = os.path.join(tmp, "refs.csv")
        shutil.copy(REFS_SRC, refs)
        # 追加两条带 DOI 的 IEEE 期刊条目（覆盖 demo/refs.csv 无 DOI 的情况）
        import csv
        with open(refs, "a", encoding="utf-8", newline="") as f:
            wr = csv.writer(f)
            wr.writerow(["haag2012", "journal",
                         "A novel approach to precision measurement in industrial environments",
                         "Haag, M., Uhlemann, F., Schrepf, P., Roth, H., Zhang, Y., Li, W., Wang, Q.",
                         "IEEE Transactions on Instrumentation and Measurement", "2012",
                         "61", "5", "1350-1358", "", "10.1109/TIM.2012.2186121",
                         "", "", ""])
            wr.writerow(["collet2020", "journal",
                         "Sensory feedback in industrial automation",
                         "Collet, J.-F., Berger, T.",
                         "IEEE Transactions on Industrial Informatics", "2020",
                         "16", "3", "2010-2018", "", "10.1109/TII.2019.2940011",
                         "", "", ""])
        doc = os.path.join(tmp, "ieee稿.docx")
        make_docx(doc, make_para("这是正文第一句", "[CITE:haag2012]") +
                  make_para("这是正文第二句", "[CITE:collet2020]"))
        out = run([sys.executable, INSERT, "--docx", doc, "--refs", refs,
                   "--style", "ieee"])
        assert "链接完整性检查通过" in out, out[-800:]
        assert not hyperlink_has_superscript(doc), "ieee 样式正文编号不应上标"
        xml = body_xml(doc)
        assert "doi: " in xml, "文末条目应渲染 doi: 前缀"
        assert "10.1109/TIM.2012.2186121" in xml, "文末条目应渲染 DOI"
        assert 'w:tooltip="打开论文页面"' in xml, "DOI 应渲染为可点击超链接"
        assert "IEEE Trans. Instrum. Meas." in xml, "期刊应渲染为 IEEE 官方缩写"
        print("C. insert --style ieee 端到端 ✅（不上标 + DOI 链接 + 期刊缩写 + 检查通过）")

        # ---------- D. verify_docx_links 缺陷检出 ----------
        # D1 悬空链接：正文 anchor 无文末书签
        d1 = os.path.join(tmp, "悬空.docx")
        make_docx(d1,
                  '<w:p><w:hyperlink w:anchor="ref_nonexist">'
                  '<w:r><w:t>[1]</w:t></w:r></w:hyperlink></w:p>' +
                  make_para("参考文献") + make_para("无书签条目"))
        ok1, iss1 = verify_docx_links([d1], 0, "numbered", None)
        assert not ok1 and any("悬空链接" in i for i in iss1), iss1
        # D2 孤立条目：文末书签无正文引用
        d2 = os.path.join(tmp, "孤立条目.docx")
        make_docx(d2, make_para("正文无引用") + make_para("参考文献") +
                  '<w:p><w:bookmarkStart w:id="0" w:name="ref_smith2020"/>'
                  '<w:r><w:t>[1] 孤立条目</w:t></w:r></w:p>')
        ok2, iss2 = verify_docx_links([d2], 0, "numbered", None)
        assert not ok2 and any("孤立条目" in i for i in iss2), iss2
        # D3 孤立编号：正文 [9] 文末 [1]
        d3 = os.path.join(tmp, "孤立编号.docx")
        make_docx(d3,
                  '<w:p><w:hyperlink w:anchor="ref_a"><w:r><w:t>[9]</w:t></w:r>'
                  '</w:hyperlink></w:p>' +
                  '<w:p><w:bookmarkStart w:id="0" w:name="ref_a"/>'
                  '<w:r><w:t>[1] 条目A</w:t></w:r></w:p>')
        ok3, iss3 = verify_docx_links([d3], 0, "numbered", None)
        assert not ok3 and any("孤立编号" in i for i in iss3), iss3
        print("D. verify_docx_links 缺陷检出 ✅（悬空/孤立条目/孤立编号）")

        # ---------- E. verify_links.py 独立验收 ----------
        # 正常文档：通过（exit 0）
        out = run([sys.executable, VERIFY_LINKS, "--docx", doc])
        assert "✅ 链接完整性检查通过" in out, out[-600:]
        # 缺陷文档：退出非零
        r = subprocess.run([sys.executable, VERIFY_LINKS, "--docx", d1],
                           capture_output=True, text=True, encoding="utf-8",
                           errors="replace")
        assert r.returncode != 0, r.stdout
        assert "悬空链接" in r.stdout, r.stdout
        print("E. verify_links.py 独立验收 ✅（通过 exit0 / 缺陷 exit≠0）")

        print("\n🎉 v1.9 回归测试全部通过（DOI 解析 / IEEE 格式 / "
              "链接完整性 / 独立验收）")
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
