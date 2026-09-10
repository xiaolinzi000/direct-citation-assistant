# -*- coding: utf-8 -*-
"""生成 demo 测试素材：refs.csv + 文稿.docx（供脚本自测与用户参考）"""
import io
import os
import csv

from docx import Document
from docx.shared import Pt

HERE = os.path.dirname(os.path.abspath(__file__))
REFS = [
    {"key": "vaswani2017", "type": "journal",
     "title": "Attention is all you need",
     "authors": "Vaswani, A., Shazeer, N., Parmar, N., Uszkoreit, J., Jones, L., Gomez, A. N., Kaiser, L., Polosukhin, I.",
     "source": "Advances in Neural Information Processing Systems",
     "year": "2017", "volume": "30", "issue": "", "pages": "",
     "doi": "10.5555/3295222.3295349", "url": "", "note": "Transformer 原始出处"},
    {"key": "devlin2019", "type": "journal",
     "title": "BERT: Pre-training of deep bidirectional transformers for language understanding",
     "authors": "Devlin, J., Chang, M.-W., Lee, K., Toutanova, K.",
     "source": "Proceedings of NAACL-HLT", "year": "2019", "volume": "",
     "issue": "", "pages": "4171-4186",
     "doi": "10.18653/v1/N19-1423", "url": "", "note": "预训练模型代表工作"},
    {"key": "zhang2023", "type": "journal",
     "title": "大语言模型研究进展综述",
     "authors": "张三, 李四, 王五",
     "source": "计算机学报", "year": "2023", "volume": "46",
     "issue": "5", "pages": "1000-1020",
     "doi": "10.7544/issn1000-1239.2023.xxxxx", "url": "", "note": "中文综述示例"},
]


def write_refs(path):
    with io.open(path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=list(REFS[0].keys()))
        w.writeheader()
        for r in REFS:
            w.writerow(r)


def build_docx(path):
    doc = Document()
    doc.add_heading("测试文稿：预训练语言模型研究", level=1)
    p1 = doc.add_paragraph()
    p1.add_run("深度学习在自然语言处理领域取得了显著进展")
    r = p1.add_run("[?]")
    r.font.superscript = True  # 占位符本身不做上标，仅示意
    p1.add_run("，这一结论在多项研究中得到验证。")

    p2 = doc.add_paragraph()
    p2.add_run("注意力机制是 Transformer 的核心组件")
    r2 = p2.add_run("[CITE:vaswani2017]")
    p2.add_run("，其重要性已被广泛认可。")

    p3 = doc.add_paragraph()
    p3.add_run("近期研究进一步表明")
    p3.add_run("[?")
    p3.add_run("]预训练模型可以显著提升下游任务表现")
    p3.add_run("。")

    p4 = doc.add_paragraph()
    p4.add_run("本文在前人工作基础上开展研究，提出了一种改进的训练策略。")

    doc.save(path)


if __name__ == "__main__":
    refs_p = os.path.join(HERE, "refs.csv")
    docx_p = os.path.join(HERE, "文稿.docx")
    write_refs(refs_p)
    build_docx(docx_p)
    print("已生成：", refs_p)
    print("已生成：", docx_p)
