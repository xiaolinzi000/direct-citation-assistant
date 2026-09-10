# -*- coding: utf-8 -*-
"""
verify_refs.py — 引用真实性与字段完整性校验（论文引用 skill）

核心目标：保证参考文献【真实存在】且与录入信息对得上，防止编造引用。
校验方式（确定性代码 + 权威数据库，非提示词判断）：
  1. 字段完整性：title/authors/year 必填；note 建议填写（记录支撑的句子）
  2. 真实性（联网校验）：
     - 有 DOI  → 查 Crossref API（api.crossref.org/works/{doi}），核对题名/年份
     - 无 DOI  → 查 OpenAlex API 按题名搜索，比对最匹配结果的题名相似度
       （DOI 在 Crossref 查不到时，自动用 OpenAlex 按题名兜底）
  3. 判定结果：
     ✅ 已核实       权威数据库命中且信息一致
     ⚠️ 信息有出入    命中但题名/年份不一致（录入可能有误，需人工核对）
     ❌ 未找到       数据库查不到 → 极可能是编造或信息严重有误，必须人工核对
     🔶 无法联网     API 不可达（不判假，标注待人工核）

用法：
  python verify_refs.py --refs refs.csv                # 完整校验（联网）
  python verify_refs.py --refs refs.csv --offline      # 仅字段完整性（不联网）
  python verify_refs.py --refs refs.csv --report 核对报告.md   # 同时导出 Markdown 报告
"""

import argparse
import io
import json
import os
import re
import sys
import urllib.error
import urllib.parse
import urllib.request
from difflib import SequenceMatcher
from datetime import datetime

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from refs_db import load_refs

UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/126.0 Safari/537.36 "
      "direct-citation-assistant/1.4")


def norm(s):
    return re.sub(r"[^a-z0-9\u4e00-\u9fff]+", "", (s or "").lower())


def ratio(a, b):
    return SequenceMatcher(None, norm(a), norm(b)).ratio()


def http_json(url, timeout=20, retries=2):
    """GET JSON，429/5xx 时等待后重试（OpenAlex/Crossref 对高频请求限流）。"""
    import time
    for attempt in range(retries):
        try:
            req = urllib.request.Request(url, headers={
                "User-Agent": UA, "Accept": "application/json"})
            with urllib.request.urlopen(req, timeout=timeout) as resp:
                return json.loads(resp.read().decode("utf-8"))
        except urllib.error.HTTPError as e:
            if e.code in (429, 500, 502, 503, 504) and attempt < retries - 1:
                time.sleep(1.2 * (attempt + 1))
                continue
            raise
    raise RuntimeError("unreachable")


def check_crossref(doi):
    """按 DOI 查 Crossref。
    返回 (found, title, year, container, volume, issue, page, error)。"""
    try:
        data = http_json(
            "https://api.crossref.org/works/" + urllib.parse.quote(doi))
        msg = data.get("message", {})
        title = (msg.get("title") or [""])[0]
        year = ""
        for k in ("published-print", "published-online", "issued"):
            dp = msg.get(k, {}).get("date-parts")
            if dp and dp[0]:
                year = str(dp[0][0])
                break
        container = (msg.get("container-title") or [""])[0]
        volume = str(msg.get("volume") or "")
        issue = str(msg.get("issue") or "")
        page = str(msg.get("page") or "")
        return True, title, year, container, volume, issue, page, None
    except urllib.error.HTTPError as e:
        if e.code == 404:
            return False, "", "", "", "", "", "", None
        return False, "", "", "", "", "", "", f"HTTP {e.code}"
    except Exception as e:
        return False, "", "", "", "", "", "", f"{type(e).__name__}: {e}"


def search_openalex(title):
    """按题名搜 OpenAlex。返回 (found, match_title, year, error)。"""
    import time
    try:
        time.sleep(0.3)  # 对 OpenAlex 保持礼貌间隔，降低 429 概率
        url = ("https://api.openalex.org/works?search="
               + urllib.parse.quote(title)
               + "&per-page=1&select=title,publication_year,biblio")
        data = http_json(url)
        res = data.get("results") or []
        if not res:
            return False, "", "", None
        t = res[0].get("title") or ""
        y = str(res[0].get("publication_year") or "")
        return True, t, y, None
    except Exception as e:
        return False, "", "", f"{type(e).__name__}: {e}"


def _norm_num(s):
    """归一化卷/期/页数值：去空白、en-dash 统一为连字符、去尾部句点。"""
    return (s or "").strip().replace("–", "-").replace(" ", "").rstrip(".,;：:。，；")


def verify(ref, offline):
    """校验单条引用，返回 (status, detail, field_issues)。"""
    title = (ref.get("title") or "").strip()
    issues = []
    if not title:
        issues.append("缺题名")
    if not (ref.get("authors") or "").strip():
        issues.append("缺作者")
    if not (ref.get("year") or "").strip():
        issues.append("缺年份")
    note = (ref.get("note") or "").strip()
    if not note:
        issues.append("note 未填（建议记录本条目支撑的句子，便于核对是否对得上）")
    elif len(note) < 6:
        issues.append("note 过短（<6字），请写明该文献支撑的具体句子/观点（如“支持句3的结论”）")

    if offline:
        return "⬜ 字段检查", "（--offline 未联网）", issues

    doi = (ref.get("doi") or "").strip()
    if doi:
        found, c_title, c_year, c_src, c_vol, c_iss, c_page, err = \
            check_crossref(doi)
        if err:
            return "🔶 无法联网", f"Crossref: {err}", issues
        if found:
            mism = []
            if ratio(c_title, title) < 0.65:
                mism.append(f"题名不符（库「{c_title[:45]}」vs 录入「{title[:30]}」）")
            for fld, cv in (("year", c_year), ("volume", c_vol),
                            ("issue", c_iss), ("pages", c_page)):
                v = _norm_num(ref.get(fld))
                c = _norm_num(cv)
                if v and c and c != v:
                    mism.append(f"{fld}不符（库 {c} vs 录入 {v}）")
            if mism:
                return "⚠️ 信息有出入", "；".join(mism), issues
            return "✅ 已核实", \
                f"Crossref DOI 命中：{c_title[:45]}（{c_year}）", issues
        # DOI 404 → OpenAlex 按题名兜底（只核对题名：OpenAlex 搜索结果可能
        # 匹配到不同版本/预印本，年份核对会误报，如 Attention Is All You Need
        # 命中 2025 重印版；年份核对仅在 Crossref DOI 精确命中时进行）
        f2, o_title, o_year, err2 = search_openalex(title)
        if err2:
            return "🔶 无法联网", f"DOI 404 且 OpenAlex: {err2}", issues
        if f2 and ratio(o_title, title) >= 0.6:
            return "✅ 已核实", \
                f"DOI 404（非 Crossref 收录），OpenAlex 题名命中：{o_title[:45]}（{o_year}）", issues
        return "❌ 未找到", \
            f"DOI 查不到且题名也搜不到（{title[:35]}）——可能是编造或信息有误", issues
    # 无 DOI → OpenAlex 题名搜索（同上：只核对题名，不核对年份）
    f, o_title, o_year, err = search_openalex(title)
    if err:
        return "🔶 无法联网", f"OpenAlex: {err}", issues
    if f and ratio(o_title, title) >= 0.6:
        return "✅ 已核实", f"OpenAlex 题名命中：{o_title[:45]}（{o_year}）", issues
    return "❌ 未找到", \
        f"题名搜不到（{title[:35]}）——可能是编造或信息有误", issues


def main(argv=None):
    ap = argparse.ArgumentParser(
        description="引用真实性与字段完整性校验（Crossref DOI + OpenAlex 题名）")
    ap.add_argument("--refs", default="refs.csv", help="引用信息表路径")
    ap.add_argument("--offline", action="store_true",
                    help="仅检查字段完整性，不做联网真实性校验")
    ap.add_argument("--report", default="", help="导出核对报告到 Markdown 文件")
    args = ap.parse_args(argv)

    refs = load_refs(args.refs)
    if not refs:
        print(f"引用表为空或不存在：{args.refs}")
        sys.exit(1)

    rows = []
    for i, r in enumerate(refs, 1):
        status, detail, issues = verify(r, args.offline)
        rows.append((i, r, status, detail, issues))
        mark = "⚠" if status.startswith(("⚠", "❌")) else " "
        flag = " ".join(issues)
        print(f"{mark} [{r['key']}] {status}  {detail}")
        if flag:
            print(f"   字段问题：{flag}")

    ok_n = sum(1 for _, _, s, _, _ in rows if s == "✅ 已核实")
    warn_n = sum(1 for _, _, s, _, _ in rows if s.startswith(("⚠", "🔶")))
    bad_n = sum(1 for _, _, s, _, _ in rows if s == "❌ 未找到")
    print(f"\n汇总：{len(rows)} 条 ｜ ✅ {ok_n} ｜ ⚠/🔶 {warn_n} ｜ ❌ {bad_n}")
    if bad_n:
        print("⚠ 存在「❌ 未找到」条目：极可能是编造或不存在的文献，"
              "请务必人工核实后删除或更正，不要直接写入论文！")

    if args.report:
        path = args.report
        lines = [f"# 引用真实性核对报告", "",
                 f"- 引用表：`{args.refs}`",
                 f"- 校验时间：{datetime.now().strftime('%Y-%m-%d %H:%M')}",
                 f"- 方式：{'仅字段检查（--offline）' if args.offline else 'Crossref DOI + OpenAlex 题名'}",
                 "",
                 f"汇总：{len(rows)} 条 ｜ ✅ {ok_n} ｜ ⚠/🔶 {warn_n} ｜ ❌ {bad_n}",
                 "",
                 "| # | key | 状态 | 题名 | 详情 | 字段问题 |",
                 "|---|---|---|---|---|---|"]
        for i, r, status, detail, issues in rows:
            title = (r.get("title") or "").replace("|", "｜")[:60]
            flag = "；".join(issues).replace("|", "｜") or "—"
            lines.append(f"| {i} | {r.get('key','')} | {status} | {title} | "
                         f"{detail.replace('|','｜')} | {flag} |")
        with io.open(path, "w", encoding="utf-8") as f:
            f.write("\n".join(lines) + "\n")
        print(f"\n报告已导出：{path}")


if __name__ == "__main__":
    main()
