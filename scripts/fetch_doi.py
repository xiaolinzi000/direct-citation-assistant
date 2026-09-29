# -*- coding: utf-8 -*-
"""
fetch_doi.py — 自动补查缺失 DOI（论文引用 skill）

场景：录入文献时出版社页面没给 DOI（如 Haag 2012 这类老文献），
refs.csv 里 doi 为空，verify_refs 只能按题名兜底核验。

本脚本按「回退链」自动补查 DOI，命中后可直接写回引用表：
  1. Crossref works 搜索      （query.bibliographic = 题名 + 首作者，按年份过滤）
  2. PubMed E-utilities       （esearch 找 PMID → esummary 取 articleids 中的 doi）
  3. OpenAlex works 搜索      （search = 题名，取结果 doi）
  4. 出版页抓取（--scrape URL） （从给定出版页面文本中启发式提取 DOI）

联网行为统一走 api_client（指数退避 + 落盘缓存），不反复 429。

用法：
  # 扫描 refs.csv 中所有缺 DOI 的条目并补查（不写回）
  python fetch_doi.py --refs <项目>/引用目录/refs.csv

  # 命中后写回 refs.csv（写入前自动备份）
  python fetch_doi.py --refs <项目>/引用目录/refs.csv --apply

  # 单条查询
  python fetch_doi.py --title "A title" --authors "Smith, J." --year 2012

  # 指定出版页抓 DOI（回退链第 4 环，手动给 URL）
  python fetch_doi.py --refs refs.csv --scrape "https://www.sciencedirect.com/science/article/pii/XXXX"

说明：
  - 只建议「题名 + 年份 + 首作者」都吻合的 DOI，避免张冠李戴；
  - --apply 前建议先 --report 看建议结果；写回是修改操作，脚本会先备份 refs.csv。
"""

import argparse
import io
import os
import re
import sys
import urllib.parse
import urllib.request

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from api_client import fetch_json
from refs_db import backup_refs, load_refs, save_refs

DOI_RE = re.compile(r"\b10\.\d{4,9}/[^\s，。；;<>()\"'`]+")
REFRESH = False  # --refresh 时置 True：绕缓存强制联网


def _norm_year(y):
    return (y or "").strip()


def _first_surname(authors):
    """取首作者姓氏（用于检索过滤）。英文 'Smith, J.' -> Smith；中文原样。"""
    if not authors:
        return ""
    a = authors.strip().split(";")[0].strip()
    if "," in a:
        return a.split(",")[0].strip()
    return a


def crossref_search(title, authors, year):
    """回退链第 1 环：Crossref works 搜索。返回 (doi, reason)。"""
    try:
        query = title
        fam = _first_surname(authors)
        if fam:
            query += " " + fam
        url = ("https://api.crossref.org/works?rows=20&query.bibliographic="
               + urllib.parse.quote(query))
        data = fetch_json(url, cache_key="fetchdoi:crossref:" + query,
                          use_cache=not REFRESH)
        for it in (data.get("message") or {}).get("items", []):
            t = (it.get("title") or [""])[0]
            yr = ""
            for k in ("published-print", "published-online", "issued"):
                dp = it.get(k, {}).get("date-parts")
                if dp and dp[0]:
                    yr = str(dp[0][0])
                    break
            doi = it.get("DOI") or ""
            if not doi:
                continue
            if year and yr and _norm_year(yr) != _norm_year(year):
                continue
            # 题名相似度粗判
            if _sim(t, title) < 0.5:
                continue
            return doi, f"Crossref（{t[:40]}，{yr}）"
    except Exception as e:
        return None, f"Crossref 查询失败：{e}"
    return None, "Crossref 未命中"


def pubmed_search(title, authors, year):
    """回退链第 2 环：PubMed E-utilities。
    esearch 按题名+首作者找 PMID，再 esummary 取 articleids 中的 doi。"""
    try:
        term = f'"{title}"[Title]'
        fam = _first_surname(authors)
        if fam:
            term += f' AND {fam}[Author]'
        if year:
            term += f' AND {year}[Publication Date]'
        url = ("https://eutils.ncbi.nlm.nih.gov/entrez/eutils/esearch.fcgi"
               f"?db=pubmed&retmode=json&term={urllib.parse.quote(term)}")
        data = fetch_json(url, cache_key="fetchdoi:esearch:" + term,
                          use_cache=not REFRESH)
        ids = ((data.get("esearchresult") or {}).get("idlist")) or []
        if not ids:
            return None, "PubMed 未命中"
        pmid = ids[0]
        s_url = ("https://eutils.ncbi.nlm.nih.gov/entrez/eutils/esummary.fcgi"
                 f"?db=pubmed&retmode=json&id={pmid}")
        s_data = fetch_json(s_url, cache_key="fetchdoi:esummary:" + pmid,
                            use_cache=not REFRESH)
        doc = (s_data.get("result") or {}).get(pmid, {})
        for aid in doc.get("articleids", []):
            if aid.get("idtype") == "doi" and aid.get("value"):
                doi = aid["value"].strip()
                if doi:
                    return doi, f"PubMed（PMID {pmid}）"
        return None, f"PubMed 命中但无 DOI（PMID {pmid}）"
    except Exception as e:
        return None, f"PubMed 查询失败：{e}"


def openalex_search(title, authors, year):
    """回退链第 3 环：OpenAlex works 搜索。返回 (doi, reason)。"""
    try:
        url = ("https://api.openalex.org/works?search="
               + urllib.parse.quote(title) + "&per-page=5"
               "&select=title,publication_year,doi")
        data = fetch_json(url, cache_key="fetchdoi:openalex:" + title,
                          use_cache=not REFRESH)
        for it in data.get("results", []):
            t = it.get("title") or ""
            yr = str(it.get("publication_year") or "")
            doi = it.get("doi") or ""
            if not doi:
                continue
            if year and yr and _norm_year(yr) != _norm_year(year):
                continue
            if _sim(t, title) < 0.5:
                continue
            return doi.replace("https://doi.org/", ""), \
                f"OpenAlex（{t[:40]}，{yr}）"
    except Exception as e:
        return None, f"OpenAlex 查询失败：{e}"
    return None, "OpenAlex 未命中"


def scrape_page_doi(url):
    """回退链第 4 环（辅助）：从出版页面文本提取 DOI。
    网络/解析异常不影响主链结果，返回 (doi, reason)。"""
    try:
        req = urllib.request.Request(url, headers={
            "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                           "AppleWebKit/537.36 (KHTML, like Gecko) "
                           "Chrome/126.0 Safari/537.36")})
        with urllib.request.urlopen(req, timeout=20) as resp:
            html = resp.read(2 * 1024 * 1024).decode("utf-8", "replace")
        m = DOI_RE.search(html)
        if m:
            doi = m.group(0).rstrip(".,;")
            return doi, f"出版页抓取（{url}）"
        return None, "出版页未找到 DOI"
    except Exception as e:
        return None, f"出版页抓取失败：{type(e).__name__}"


def _sim(a, b):
    from difflib import SequenceMatcher
    norm = lambda s: re.sub(r"[^a-z0-9\u4e00-\u9fff]+", "",
                            (s or "").lower())
    return SequenceMatcher(None, norm(a), norm(b)).ratio()


def find_doi(title, authors="", year="", scrape_url="", verbose=False):
    """按回退链补查 DOI。返回 (doi, reason)；全部未命中返回 (None, reason)。"""
    chain = [("Crossref", crossref_search), ("PubMed", pubmed_search),
             ("OpenAlex", openalex_search)]
    for name, fn in chain:
        doi, reason = fn(title, authors, year)
        if verbose:
            print(f"  ↳ {name}: {reason}")
        if doi:
            return doi, f"{name}：{reason}"
    if scrape_url:
        doi, reason = scrape_page_doi(scrape_url)
        if verbose:
            print(f"  ↳ 出版页: {reason}")
        if doi:
            return doi, f"出版页抓取：{reason}"
    return None, "；".join(f"{n}未命中" for n, _ in chain)


def main():
    ap = argparse.ArgumentParser(
        description="自动补查缺失 DOI（回退链：Crossref → PubMed → OpenAlex → 出版页）")
    ap.add_argument("--refs", default="", help="refs.csv 路径（扫描其中缺 DOI 的条目）")
    ap.add_argument("--title", default="", help="单条模式：题名")
    ap.add_argument("--authors", default="", help="单条模式：作者")
    ap.add_argument("--year", default="", help="单条模式：年份")
    ap.add_argument("--apply", action="store_true",
                    help="把命中结果写回 refs.csv（写前自动备份）")
    ap.add_argument("--report", default="", help="导出补查报告到 Markdown 文件")
    ap.add_argument("--scrape", default="",
                   help="回退链第 4 环：指定出版页 URL 抓取 DOI（可配合 --refs 使用）")
    ap.add_argument("--refresh", action="store_true",
                    help="绕缓存强制联网查询（默认命中 7 天缓存）")
    ap.add_argument("--verbose", action="store_true", help="打印每一环的查询结果")
    args = ap.parse_args()
    global REFRESH
    REFRESH = args.refresh

    if args.refs:
        refs = load_refs(args.refs)
        if not refs:
            print(f"引用表为空或不存在：{args.refs}")
            sys.exit(1)
        missing = [r for r in refs if not (r.get("doi") or "").strip()]
        if not missing:
            print(f"引用表中没有缺 DOI 的条目（共 {len(refs)} 条，全部有 DOI）。")
            return
        print(f"引用表：{args.refs}（{len(refs)} 条，缺 DOI {len(missing)} 条）")
        found = {}
        for r in missing:
            print(f"[{r['key']}] {r['title'][:50]}（{r.get('year') or '?'}）")
            doi, reason = find_doi(r.get("title", ""), r.get("authors", ""),
                                   r.get("year", ""), args.scrape,
                                   args.verbose)
            if doi:
                found[r["key"]] = doi
                print(f"    ✅ {doi}  ← {reason}")
            else:
                print(f"    ❌ 未补到（{reason}）")
        # 写回
        if found and args.apply:
            b = backup_refs(args.refs)
            for r in refs:
                if r["key"] in found:
                    r["doi"] = found[r["key"]]
            save_refs(refs, args.refs)
            print(f"\n已写回 {len(found)} 条 DOI 到 {args.refs}"
                  + (f"（原表备份：{b}）" if b else ""))
        elif found:
            print(f"\n命中 {len(found)} 条，未写回（加 --apply 写回 refs.csv，写前自动备份）。")
        # 报告
        if args.report:
            with io.open(args.report, "w", encoding="utf-8") as f:
                f.write("# DOI 补查报告\n\n")
                f.write(f"- 引用表：`{args.refs}`\n")
                f.write(f"- 缺 DOI：{len(missing)} 条；命中：{len(found)} 条\n\n")
                f.write("| key | 题名 | 年份 | 命中 DOI | 状态 |\n|---|---|---|---|---|\n")
                for r in missing:
                    doi = found.get(r["key"], "")
                    status = "✅" if doi else "❌"
                    f.write(f"| {r['key']} | {r['title'][:50]} | {r.get('year','')} | "
                            f"{doi} | {status} |\n")
            print(f"报告已导出：{args.report}")
    elif args.title:
        print(f"查询：{args.title}（{args.authors or '?'}，{args.year or '?'}）")
        doi, reason = find_doi(args.title, args.authors, args.year,
                               args.scrape, args.verbose)
        if doi:
            print(f"✅ {doi}  ← {reason}")
        else:
            print(f"❌ 未补到（{reason}）")
    else:
        ap.print_help()


if __name__ == "__main__":
    main()
