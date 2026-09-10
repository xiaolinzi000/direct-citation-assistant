# -*- coding: utf-8 -*-
"""
verify_refs_test.py — verify_refs.py 离线回归测试

覆盖（不依赖网络）：
  1. 字段完整性：缺 title/authors/year/note 均能提示
  2. --offline 模式：返回 ⬜ 字段检查，不做联网判定
  3. 题名归一化/相似度函数：norm 与 ratio 行为
联网真实性判定（Crossref/OpenAlex）由人工在真实网络下抽查（demo refs.csv：
vaswani2017 / devlin2019 应为 ✅，虚构的 zhang2023 应为 ❌）。

用法：python verify_refs_test.py
"""
import io
import os
import sys

_here = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, _here)
sys.path.insert(0, os.path.join(_here, ".."))
import verify_refs as vr

PASS = 0


def check(name, cond, extra=""):
    global PASS
    if cond:
        PASS += 1
        print(f"  ✅ {name}")
    else:
        print(f"  ❌ {name} {extra}")
        sys.exit(1)


def build_refs(tmp_path, rows):
    import csv
    fields = ["key", "type", "title", "authors", "source", "year",
              "volume", "issue", "pages", "city", "doi", "url", "note"]
    with io.open(tmp_path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=fields)
        w.writeheader()
        for r in rows:
            row = {k: "" for k in fields}
            row.update(r)
            w.writerow(row)


def test_field_checks():
    print("1) 字段完整性检查")
    good = {"key": "a", "type": "journal", "title": "Some paper",
            "authors": "Smith, J.", "year": "2020", "note": "支撑句"}
    status, detail, issues = vr.verify(good, offline=True)
    check("完整条目无字段问题", status == "⬜ 字段检查" and not issues,
          f"issues={issues}")

    bad = {"key": "b", "type": "journal", "title": "", "authors": "",
           "year": "", "note": ""}
    status, detail, issues = vr.verify(bad, offline=True)
    need = {"缺题名", "缺作者", "缺年份", "note 未填（建议记录本条目支撑的句子，便于核对是否对得上）"}
    check("缺字段条目全部提示", need.issubset(set(issues)),
          f"issues={issues}")

    no_note = {"key": "c", "type": "journal", "title": "X", "authors": "Y",
               "year": "2021", "note": ""}
    status, detail, issues = vr.verify(no_note, offline=True)
    check("note 缺失单独提示", any("note" in i for i in issues), f"{issues}")


def test_norm_ratio():
    print("2) 归一化与相似度")
    check("norm 小写去符号", vr.norm("Attention Is All You Need!")
          == "attentionisallyouneed")
    check("norm 中文保留", "注意力机制" in vr.norm("注意力机制研究"))
    check("完全一致 ratio=1", vr.ratio("a b", "a b") == 1.0)
    check("相似标题 ratio>0.65", vr.ratio(
        "BERT: Pre-training of Deep Bidirectional Transformers for Language "
        "Understanding",
        "BERT: Pre-training of deep bidirectional transformers") >= 0.65)
    check("无关标题 ratio<0.5", vr.ratio("abc def", "xyz qwe") < 0.5)


def test_offline_mode():
    print("3) --offline 模式不联网")
    tmp = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                       "_vr_test_refs.csv")
    build_refs(tmp, [
        {"key": "a", "type": "journal", "title": "Some paper",
         "authors": "Smith, J.", "year": "2020", "note": "支撑句"},
    ])
    old_stdout = sys.stdout
    sys.stdout = io.StringIO()
    try:
        vr.main(["--refs", tmp, "--offline"])
    except SystemExit:
        pass
    finally:
        out = sys.stdout.getvalue()
        sys.stdout = old_stdout
    os.remove(tmp)
    check("离线输出字段检查", "⬜ 字段检查" in out and "汇总：1 条" in out,
          out[:200])


if __name__ == "__main__":
    test_field_checks()
    test_norm_ratio()
    test_offline_mode()
    print(f"\n🎉 verify_refs 离线回归全部通过（{PASS} 项断言）")
