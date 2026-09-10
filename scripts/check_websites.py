# -*- coding: utf-8 -*-
"""
check_websites.py — 文献检索网站可用性检查（代码执行，非提示词判断）

对技能内置的文献检索网站清单逐一发起 HTTP 检查（并发、超时控制），
把结果以表格形式输出，并用 --update 把最新状态写回 references/websites.md。

状态判定（确定性规则）：
  ✅ 可达      HTTP 2xx/3xx（含重定向）
  ⚠️ 异常      HTTP 4xx/5xx（记录状态码）
  ❌ 不可达    DNS/连接/超时/SSL 失败（记录原因）

用法：
  python check_websites.py                 # 只输出检查报告
  python check_websites.py --update        # 检查并更新 references/websites.md
  python check_websites.py --timeout 15    # 自定义超时（秒）
"""

import argparse
import io
import os
import sys
import threading
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from urllib import error, request

# 网站清单：名称、URL、用途、备注（单一权威来源；references/websites.md 由此生成）
SITES = [
    ("谷歌学术 Google Scholar", "https://scholar.google.com",
     "国际文献检索、引用次数、题录", "大陆直连受限，可用镜像/代理"),
    ("百度学术", "https://xueshu.baidu.com",
     "中文文献聚合检索、引用信息", ""),
    ("知网 CNKI", "https://www.cnki.net",
     "中文期刊/学位论文", "学位论文下载需机构权限"),
    ("万方数据", "https://www.wanfangdata.com.cn",
     "中文文献、学位论文", ""),
    ("维普", "https://www.cqvip.com",
     "中文期刊", ""),
    ("PubMed", "https://pubmed.ncbi.nlm.nih.gov",
     "生物医学文献", "NLM 官方免费检索"),
    ("Semantic Scholar", "https://www.semanticscholar.org",
     "免费语义学术搜索、引用网络", "API 免费"),
    ("Crossref", "https://search.crossref.org",
     "DOI 题录核验", "api.crossref.org 可取权威元数据"),
    ("arXiv", "https://arxiv.org",
     "预印本", "计算机/物理等"),
    ("OpenAlex", "https://openalex.org",
     "开放学术数据、引用关系", "API 免费"),
    ("Web of Science", "https://www.webofscience.com",
     "权威引文索引", "订阅制"),
    ("Scopus", "https://www.scopus.com",
     "权威引文索引", "订阅制"),
    ("ScienceDirect", "https://www.sciencedirect.com",
     "Elsevier 期刊全文", "部分开放获取"),
    ("IEEE Xplore", "https://ieeexplore.ieee.org",
     "电子/计算机领域", "订阅制"),
    ("SpringerLink", "https://link.springer.com",
     "Springer 期刊/图书", "部分开放获取"),
    ("Wiley Online Library", "https://onlinelibrary.wiley.com",
     "Wiley 期刊", "订阅制"),
    ("Nature", "https://www.nature.com",
     "Nature 系列期刊", ""),
    ("Science", "https://www.science.org",
     "Science 期刊", ""),
]

UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/126.0 Safari/537.36")


def check_one(url, timeout):
    """检查单个 URL，返回 (status, detail)。
    status: ok / warn / down；detail 为描述文字。"""
    req = request.Request(url, headers={
        "User-Agent": UA,
        "Range": "bytes=0-1024",          # 只取首段，避免大页面下载
        "Accept": "text/html,application/xhtml+xml,*/*;q=0.8",
    })
    try:
        with request.urlopen(req, timeout=timeout) as resp:
            code = resp.getcode()
            if 200 <= code < 400:
                return "ok", f"HTTP {code}"
            if code in (403, 407):
                return "warn", f"HTTP {code}（反爬拦截，浏览器通常可访问）"
            return "warn", f"HTTP {code}"
    except error.HTTPError as e:
        if e.code in (301, 302, 303, 307, 308):
            # 重定向（urllib 不自动跟随 307/308）仍视为可达
            return "ok", f"HTTP {e.code}（重定向）"
        if e.code in (403, 407):
            return "warn", f"HTTP {e.code}（反爬拦截，浏览器通常可访问）"
        return "warn", f"HTTP {e.code}"
    except error.URLError as e:
        reason = getattr(e, "reason", e)
        return "down", f"{type(reason).__name__}: {reason}"
    except Exception as e:
        return "down", f"{type(e).__name__}: {e}"


def run_checks(timeout):
    results = []
    with ThreadPoolExecutor(max_workers=8) as ex:
        futs = {ex.submit(check_one, url, timeout): name
                for name, url, _, _ in SITES}
        for fut in as_completed(futs):
            name = futs[fut]
            status, detail = fut.result()
            results.append((name, status, detail))
    # 按原清单顺序输出
    order = {name: i for i, (name, _, _, _) in enumerate(SITES)}
    results.sort(key=lambda r: order[r[0]])
    return results


def status_icon(status):
    return {"ok": "✅ 可达", "warn": "⚠️ 异常", "down": "❌ 不可达"}[status]


def render_md(results, check_time):
    lines = ["# 文献检索网站清单", "",
             "> 本清单供「论文引用 skill」检索文献时选用。",
             "> 状态由 `scripts/check_websites.py` 自动检查更新（代码执行，非人工判断）。",
             f"> 最后检查时间：{check_time}", "",
             "| 网站 | URL | 用途 | 状态 | 备注 |",
             "|---|---|---|---|---|"]
    by_name = {n: (s, d) for n, s, d in results}
    for name, url, purpose, note in SITES:
        status, detail = by_name.get(name, ("?", "?"))
        cell = status_icon(status)
        if status != "ok":
            cell += f"（{detail}）"
        lines.append(f"| {name} | {url} | {purpose} | {cell} | {note} |")
    lines.append("")
    return "\n".join(lines)


def main():
    ap = argparse.ArgumentParser(description="文献检索网站可用性检查")
    ap.add_argument("--update", action="store_true",
                    help="检查后更新 references/websites.md")
    ap.add_argument("--timeout", type=int, default=12, help="单站超时秒数")
    args = ap.parse_args()

    t0 = time.time()
    results = run_checks(args.timeout)
    check_time = datetime.now().strftime("%Y-%m-%d %H:%M")

    ok_n = sum(1 for _, s, _ in results if s == "ok")
    warn_n = sum(1 for _, s, _ in results if s == "warn")
    down_n = sum(1 for _, s, _ in results if s == "down")
    print(f"检查完成：{len(results)} 站，✅ {ok_n} ｜ ⚠️ {warn_n} ｜ ❌ {down_n}"
          f"（耗时 {time.time() - t0:.1f}s）")
    print()
    for name, status, detail in results:
        print(f"  {status_icon(status):<6} {name}  {detail}")

    if args.update:
        md = render_md(results, check_time)
        here = os.path.dirname(os.path.abspath(__file__))
        ref_dir = os.path.join(here, "..", "references")
        os.makedirs(ref_dir, exist_ok=True)
        path = os.path.join(ref_dir, "websites.md")
        with io.open(path, "w", encoding="utf-8") as f:
            f.write(md)
        print(f"\n已更新：{path}")


if __name__ == "__main__":
    main()
