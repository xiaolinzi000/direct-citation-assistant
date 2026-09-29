# -*- coding: utf-8 -*-
"""回归测试（v1.8 改进 1，最高优先级）：句级落点 + 严格引用。

覆盖：
  1. split_sentences 句级切分器（中文/英文/小数点/缩写/省略号/引号归属）；
  2. 句内落点校正：占位符误放在句号之后 → 自动归位到句末标点之前；
  3. 堆叠告警：单句/单段编号堆叠 → 输出提示（阈值 3/5 默认）。
"""
import os
import re
import subprocess
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
SCRIPTS = os.path.dirname(HERE)
sys.path.insert(0, SCRIPTS)

from insert_refs import split_sentences
from demo._test_util import make_docx, make_para, read_docx, para_text, esc

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def w(tag):
    return f"{{{W}}}{tag}"


def test_splitter():
    print("== 1) split_sentences 句级切分器 ==")
    # 中文三句
    s = split_sentences("句子一。句子二！句子三？尾句")
    assert len(s) == 4 and s[0][2] == "句子一。" and s[-1][2] == "尾句", s
    # 小数/版本号不切：整句仍是 1 段
    s = split_sentences("版本 1.2.3 已发布。")
    assert len(s) == 1 and s[0][2] == "版本 1.2.3 已发布。", s
    # 缩写不切
    s = split_sentences("This was shown by Smith et al. in 2020. Done.")
    assert len(s) == 2 and "et al." in s[0][2], s
    # 省略号不切（英文连续点；中文省略号不是句末标点天然不切）
    s = split_sentences("Wait... let me think. Done.")
    assert len(s) == 2 and s[0][2].startswith("Wait..."), s
    s = split_sentences("内容……然后。继续。")
    assert len(s) == 2, s
    # 闭合引号归属
    s = split_sentences("他说：“这是重点。”然后继续。")
    assert len(s) == 2, s
    assert s[0][2].endswith("。”"), s
    print(f"  ✓ 6 组切分断言通过（示例：{split_sentences('版本 1.2.3 已发布。')[0][2]}）")
    return True


def test_sentence_placement(tmp):
    print("== 2) 句内落点校正：句号后的占位符自动归位 ==")
    refs = os.path.join(tmp, "refs.csv")
    with open(refs, "w", encoding="utf-8-sig", newline="") as f:
        f.write("key,type,title,authors,source,year,volume,issue,pages,doi,url,note\n")
        f.write("k1,journal,Title one,Smith J.,Journal A,2020,1,1,1-2,10.1/a,,\n")
        f.write("k2,journal,Title two,Jones K.,Journal B,2021,2,2,3-4,10.2/b,,\n")
    docx = os.path.join(tmp, "文.docx")
    # 第一句末尾句号后误放 [CITE:k1]；第二句正常 [CITE:k2]
    body = ('<w:p><w:r><w:t>第一句内容。</w:t></w:r>'
            '<w:r><w:t>[CITE:k1]</w:t></w:r>'
            '<w:r><w:t> 第二句内容</w:t></w:r>'
            '<w:r><w:t>[CITE:k2]</w:t></w:r>'
            '<w:r><w:t>。</w:t></w:r></w:p>')
    make_docx(docx, body)
    r = subprocess.run([sys.executable,
                        os.path.join(SCRIPTS, "insert_refs.py"),
                        "--docx", docx, "--refs", refs, "--mapping"],
                       capture_output=True, text=True, encoding="utf-8",
                       errors="replace")
    out = r.stdout + r.stderr
    assert "编号一致性验证通过" in out, out[-800:]
    root = read_docx(docx)
    p = next(iter(root.iter(w("p"))))
    txt = para_text(p)
    # [1] 应位于句号前：第一句内容[1]。 第二句内容[2]。
    assert txt.index("内容") < txt.index("[1]") < txt.index("。") < \
        txt.index("第二句"), txt
    assert txt.index("第二句内容") < txt.index("[2]") < txt.rindex("。"), txt
    print(f"  ✓ 归位成功：{txt}")
    return True


def test_stacking_warning(tmp):
    print("== 3) 堆叠告警：单段 6 个编号 → 输出提示 ==")
    refs = os.path.join(tmp, "refs.csv")
    lines = ["key,type,title,authors,source,year,volume,issue,pages,doi,url,note"]
    for i in range(1, 7):
        lines.append(f"k{i},journal,T{i},A{i}.,J{i},2020,1,1,1-2,,,")
    with open(refs, "w", encoding="utf-8-sig", newline="") as f:
        f.write("\n".join(lines) + "\n")
    docx = os.path.join(tmp, "堆叠.docx")
    runs = "".join(f'<w:r><w:t>[CITE:k{i}]</w:t></w:r>' for i in range(1, 7))
    body = ("<w:p><w:r><w:t>一句话里堆了六个编号。</w:t></w:r>" + runs + "</w:p>")
    make_docx(docx, body)
    r = subprocess.run([sys.executable,
                        os.path.join(SCRIPTS, "insert_refs.py"),
                        "--docx", docx, "--refs", refs],
                       capture_output=True, text=True, encoding="utf-8",
                       errors="replace")
    out = r.stdout + r.stderr
    assert "堆叠告警" in out, out[-800:]
    assert "句1" in out, out[-800:]
    print("  ✓ 堆叠告警已触发（句 1 内 6 个编号）")
    return True


if __name__ == "__main__":
    ok = test_splitter()
    with tempfile.TemporaryDirectory() as tmp:
        ok = test_sentence_placement(tmp) and ok
        ok = test_stacking_warning(tmp) and ok
    print("\n✅ 句级落点回归测试通过" if ok else "\n❌ 句级落点回归测试失败")
    sys.exit(0 if ok else 1)
