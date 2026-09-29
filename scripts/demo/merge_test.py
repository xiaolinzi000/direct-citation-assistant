# -*- coding: utf-8 -*-
"""回归测试（v1.8 改进 8）：多章节合并 —— 统一全局编号 + 合并文末表 + 映射表。

场景：章1、章2 各自 insert_refs 生成引用（各自编号 1,2…），合并后：
  - 全文一个 References 标题、一份文末表；
  - 全局重编号（key 首次出现定号）；
  - 输出「章节旧编号 → 全局新编号」映射表。
"""
import os
import re
import shutil
import subprocess
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
SCRIPTS = os.path.dirname(HERE)
sys.path.insert(0, SCRIPTS)

from demo._test_util import (entry_numbers, make_docx, make_para,
                             read_docx, ref_anchors, para_text, w)

REFS_SRC = os.path.join(HERE, "refs.csv")
INSERT = os.path.join(SCRIPTS, "insert_refs.py")
MERGE = os.path.join(SCRIPTS, "merge_refs.py")


def run(cmd):
    r = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8",
                       errors="replace")
    assert r.returncode == 0, (r.stdout + r.stderr)[-1500:]
    return r.stdout + r.stderr


if __name__ == "__main__":
    tmp = tempfile.mkdtemp(prefix="merge_test_")
    try:
        refs = os.path.join(tmp, "refs.csv")
        shutil.copy(REFS_SRC, refs)
        ch1 = os.path.join(tmp, "第1章.docx")
        ch2 = os.path.join(tmp, "第2章.docx")
        # 章1：v 与 d
        make_docx(ch1, make_para("第一章内容一", "[CITE:vaswani2017]") +
                  make_para("第一章内容二", "[CITE:devlin2019]"))
        run([sys.executable, INSERT, "--docx", ch1, "--refs", refs])
        # 章2：z 与 v（重复引用 v）
        make_docx(ch2, make_para("第二章内容一", "[CITE:zhang2023]") +
                  make_para("第二章内容二", "[CITE:vaswani2017]"))
        run([sys.executable, INSERT, "--docx", ch2, "--refs", refs])
        out_1 = run([sys.executable, INSERT, "--docx", ch1, "--refs", refs,
                     "--dry-run"])
        assert "vaswani2017->1" in out_1 and "devlin2019->2" in out_1, out_1
        out_2 = run([sys.executable, INSERT, "--docx", ch2, "--refs", refs,
                     "--dry-run"])
        assert "zhang2023->1" in out_2 and "vaswani2017->2" in out_2, out_2
        print("== 1) 各章单独编号 ==\n  ✓ 章1: v=1 d=2 ｜ 章2: z=1 v=2")

        # 合并
        merged = os.path.join(tmp, "合并稿.docx")
        out = run([sys.executable, MERGE, "--docx", ch1, ch2,
                   "--refs", refs, "--out", merged])
        # 映射表应出现章节 → 全局
        assert "第1章" in out and "第2章" in out, out[-1000:]
        assert re.search(r"第2章.*\[1\] -> \[3\]", out), out[-1000:]
        assert re.search(r"第2章.*\[2\] -> \[1\]", out), out[-1000:]
        print("== 2) 合并 + 映射表 ==\n  ✓ 第2章 [1]->[3]（zhang2023）、[2]->[1]（vaswani2017 复用）")

        root = read_docx(merged)
        anchors = ref_anchors(root)
        assert len(anchors) == 4, anchors
        assert anchors.count("ref_vaswani2017") == 2
        assert anchors.count("ref_devlin2019") == 1
        assert anchors.count("ref_zhang2023") == 1
        assert entry_numbers(root) == [1, 2, 3], entry_numbers(root)
        # 只有一个参考文献标题（脚本默认创建中文「参考文献」）
        body = root.find(w("body"))
        heads = [p for p in body.iter(w("p"))
                 if para_text(p).strip() in ("References", "参考文献")]
        assert len(heads) == 1, f"参考文献标题数量：{len(heads)}"
        print("  ✓ 正文引用 4 处（v×2、d、z）｜文末条目 3 条 ｜ References 标题唯一")
        print("\n✅ 多章节合并回归测试通过")
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
