# -*- coding: utf-8 -*-
"""
api_client.py — 论文引用 skill 的联网公共模块

统一处理对外 API 请求的三大问题：
  1. 限流退避：429/5xx 等可重试错误按「指数退避 + 随机抖动」重试，
     避免反复 429 把可用额度耗尽（Crossref / OpenAlex / PubMed 均高频限流）。
  2. 落盘缓存：同参数请求的 JSON 响应缓存到 <scripts>/_api_cache/，
     默认 7 天有效，重跑校验不再重复占用 API 额度（结果以文件 hash 为键）。
  3. 统一 UA 与超时：可配置，默认带浏览器 UA 规避部分 403。

用法：
    from api_client import fetch_json, clear_cache, CACHE_DIR
    data = fetch_json("https://api.crossref.org/works/10.xxxx/yyy",
                      cache_key="crossref:10.xxxx/yyy")
"""

import hashlib
import json
import os
import random
import sys
import time
import urllib.error
import urllib.request

UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/126.0 Safari/537.36 "
      "direct-citation-assistant/1.8")

CACHE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "_api_cache")
DEFAULT_TTL = 7 * 24 * 3600  # 缓存 7 天
RETRYABLE = frozenset((429, 500, 502, 503, 504))


def cache_path(cache_key):
    """缓存文件路径：以 cache_key 的 sha1 命名，避免路径/特殊字符问题。"""
    h = hashlib.sha1(cache_key.encode("utf-8", "replace")).hexdigest()[:16]
    return os.path.join(CACHE_DIR, h + ".json")


def read_cache(cache_key, ttl=DEFAULT_TTL):
    """读取未过期缓存；无/过期/损坏返回 None。"""
    p = cache_path(cache_key)
    if not os.path.exists(p):
        return None
    try:
        age = time.time() - os.path.getmtime(p)
        if age > ttl:
            return None
        with open(p, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return None


def write_cache(cache_key, data):
    """写入缓存（失败静默：缓存只是优化，不影响主流程）。"""
    try:
        os.makedirs(CACHE_DIR, exist_ok=True)
        with open(cache_path(cache_key), "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False)
    except Exception:
        pass


def clear_cache():
    """清空整个 API 缓存目录（用于调试/强制刷新）。"""
    n = 0
    if os.path.isdir(CACHE_DIR):
        for fn in os.listdir(CACHE_DIR):
            try:
                os.remove(os.path.join(CACHE_DIR, fn))
                n += 1
            except OSError:
                pass
    return n


def fetch_json(url, cache_key=None, timeout=20, retries=4, use_cache=True,
               poll=None):
    """GET JSON：指数退避重试 + 落盘缓存。

    参数：
      url       目标 API 地址
      cache_key 缓存键（建议含 API 名与查询串）；None 表示不缓存
      timeout   单次请求超时（秒）
      retries   最大尝试次数（含首次）
      use_cache 是否读/写缓存（False 强制实时请求）
      poll      可选的 {interval: 秒, max_wait: 秒}：请求成功但返回体需要
                轮询（如 PubMed 某些异步接口），此参数暂为扩展保留
    返回：
      解析后的 JSON（dict/list）。
    抛错：
      RuntimeError —— 网络不可达 / 重试耗尽 / 非可重试 HTTP 错误。
    """
    if use_cache and cache_key:
        hit = read_cache(cache_key)
        if hit is not None:
            return hit
    last_err = None
    for attempt in range(retries):
        try:
            req = urllib.request.Request(url, headers={
                "User-Agent": UA, "Accept": "application/json"})
            with urllib.request.urlopen(req, timeout=timeout) as resp:
                raw = resp.read().decode("utf-8", "replace")
                data = json.loads(raw)
                if use_cache and cache_key:
                    write_cache(cache_key, data)
                return data
        except urllib.error.HTTPError as e:
            if e.code in RETRYABLE and attempt < retries - 1:
                last_err = f"HTTP {e.code}"
                time.sleep(_backoff(attempt))
                continue
            if e.code == 404:
                # 404 明确不存在，不缓存错误但直接上抛，由调用方区分
                raise RuntimeError(f"HTTP 404: {url}")
            raise RuntimeError(f"HTTP {e.code}: {url}")
        except (urllib.error.URLError, TimeoutError, OSError) as e:
            if attempt < retries - 1:
                last_err = f"{type(e).__name__}"
                time.sleep(_backoff(attempt))
                continue
            raise RuntimeError(f"网络错误 {last_err or type(e).__name__}: {url}")
        except ValueError as e:
            raise RuntimeError(f"响应非 JSON（{e}）: {url}")
    raise RuntimeError(f"重试耗尽（{last_err}）: {url}")


def _backoff(attempt):
    """指数退避：1.5 * 2^attempt 秒 + 0~0.5s 抖动，上限 20s。"""
    base = 1.5 * (2 ** attempt)
    return min(base + random.uniform(0, 0.5), 20.0)


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "--clear-cache":
        print(f"已清理缓存 {clear_cache()} 个文件（{CACHE_DIR}）")
    else:
        print(__doc__)
