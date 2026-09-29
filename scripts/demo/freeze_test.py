# -*- coding: utf-8 -*-
"""回归测试（v1.8 改进 2）：编号冻结 + 增量更新 + 映射 diff。

场景（key 编号恒定的核心诉求，对应「collet2020 恒为 1」）：
  1. 首次运行：按出现顺序编号，状态落盘；
  2. 同文档普通重跑：幂等，diff 无变化；
  3. 删 v 增 z 后 --freeze：已冻结 key 编号不变（devlin2019 恒为 2），
     新 key 复用释放的空号（zhang2023 -> 1）；
  4. --freeze 幂等：再次运行无变化。
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

from demo._test_util import entry_numbers, make_docx, make_para, read_docx

REFS_SRC = os.path.join(HERE, "refs.csv")
INSERT = os.path.join(SCRIPTS, "insert_refs.py")


def run_insert(docx, refs, *extra):
    r = subprocess.run([sys.executable, INSERT, "--docx", docx,
                        "--refs", refs] + list(extra),
                       capture_output=True, text=True, encoding="utf-8",
                       errors="replace")
    return r.stdout + r.stderr


def main_nums(out):
    """从运行输出提取「编号方案：k->n」映射。"""
    m = re.search(r"编号方案：(.+)", out)
    if not m:
        return {}
    return {k: int(n) for k, n in
            re.findall(r"([A-Za-z0-9_\-\u4e00-\u9fff.]+)->(\d+)", m.group(1))}


if __name__ == "__main__":
    tmp = tempfile.mkdtemp(prefix="freeze_test_")
    try:
        refs = os.path.join(tmp, "refs.csv")
        shutil.copy(REFS_SRC, refs)
        docx = os.path.join(tmp, "文稿.docx")
        state_p = os.path.join(tmp, "_refs_state", "文稿.json")

        # 1) 首次：v 先出现 → 1，d → 2
        make_docx(docx, make_para("A", "[CITE:vaswani2017]") +
                  make_para("B", "[CITE:devlin2019]"))
        out = run_insert(docx, refs)
        m1 = main_nums(out)
        assert m1 == {"vaswani2017": 1, "devlin2019": 2}, m1
        assert os.path.exists(state_p), "状态文件未生成"
        print("== 1) 首次运行 ==\n  ✓", m1, "｜ 状态已落盘")

        # 2) 同文档普通重跑：幂等（无 diff 变化）
        out = run_insert(docx, refs)
        m2 = main_nums(out)
        assert m2 == m1, m2
        assert "新增 0" in out and "移除 0" in out and "不变 2" in out, out[-500:]
        print("== 2) 同文档普通重跑 ==\n  ✓ 幂等：新增 0 / 移除 0 / 变化 0 / 不变 2")

        # 3) 删 v、增 z 后 --freeze：d 恒为 2，z 复用空号 1
        make_docx(docx, make_para("B", "[CITE:devlin2019]") +
                  make_para("Z", "[CITE:zhang2023]"))
        out = run_insert(docx, refs, "--freeze")
        m3 = main_nums(out)
        assert m3 == {"devlin2019": 2, "zhang2023": 1}, m3
        assert "冻结" in out and "移除" in out and "新增" in out, out[-500:]
        root = read_docx(docx)
        txt = "".join(p.text or "" for p in root.iter(
            "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t"))
        assert "[1]" in txt and "[2]" in txt, txt
        assert entry_numbers(root) == [1, 2], entry_numbers(root)
        print("== 3) --freeze 冻结（删 v 增 z） ==\n"
              "  ✓ devlin2019 恒为 2、zhang2023 复用空号 1｜diff 含移除 v / 新增 z")

        # 4) --freeze 幂等
        out = run_insert(docx, refs, "--freeze")
        assert "新增 0" in out and "移除 0" in out and "不变 2" in out, out[-500:]
        print("== 4) --freeze 幂等 ==\n  ✓ 无变化")
        print("\n✅ 编号冻结/增量更新回归测试通过")
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
