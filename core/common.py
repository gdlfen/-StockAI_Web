# -*- coding: utf-8 -*-
"""
通用工具层：HTTP 重试、磁盘缓存、日志回调、数值清洗。

云端（Streamlit Community Cloud / 容器）特点：
- 文件系统可写但**不保证持久**，因此缓存与产物都放在可配置的 data_dir 下；
- 无 GUI、无浏览器、无本机代理，所有网络访问走直连 + 重试。
"""
from __future__ import annotations

import hashlib
import json
import math
import os
import re
import time
from datetime import datetime, timedelta
from typing import Any, Callable, Dict, Iterable, List, Optional

import pandas as pd
import requests

LogFn = Callable[[str], None]

DEFAULT_UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
              "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")


# ======================================================================
# 日志
# ======================================================================
def _noop(_msg: str) -> None:  # pragma: no cover
    pass


def make_logger(callback: Optional[Callable[[str], None]] = None) -> LogFn:
    """统一日志出口：既打印到 stdout（云端日志可见），也可回调到任务队列。"""
    def _log(msg: str) -> None:
        text = f"[{datetime.now().strftime('%H:%M:%S')}] {msg}"
        try:
            print(text, flush=True)
        except Exception:
            pass
        if callback is not None:
            try:
                callback(text)
            except Exception:
                pass

    return _log


# ======================================================================
# HTTP
# ======================================================================
def http_get(url: str, *, params: Optional[dict] = None, headers: Optional[dict] = None,
             timeout: int = 20, encoding: Optional[str] = None, retries: int = 3,
             label: str = "", log: Optional[LogFn] = None, allow_proxy_env: bool = False) -> Optional[requests.Response]:
    """带重试与退避的 GET。默认忽略环境变量里的代理（云端常有残留代理配置）。"""
    hdrs = {"User-Agent": DEFAULT_UA}
    if headers:
        hdrs.update(headers)
    last: Optional[Exception] = None
    session = requests.Session()
    if not allow_proxy_env:
        session.trust_env = False
    for i in range(1, max(1, retries) + 1):
        try:
            r = session.get(url, params=params, headers=hdrs, timeout=timeout)
            if encoding:
                r.encoding = encoding
            return r
        except Exception as e:  # noqa: BLE001
            last = e
            if i < retries:
                time.sleep(min(0.8 * i, 3.0))
    if label:
        (log or _noop)(f"⚠️ {label} 请求失败（已重试{retries}次）: {type(last).__name__}: {str(last)[:120]}")
    return None


def http_post(url: str, *, data: Optional[dict] = None, headers: Optional[dict] = None,
              timeout: int = 20, retries: int = 3, label: str = "",
              log: Optional[LogFn] = None) -> Optional[requests.Response]:
    hdrs = {"User-Agent": DEFAULT_UA}
    if headers:
        hdrs.update(headers)
    last: Optional[Exception] = None
    session = requests.Session()
    session.trust_env = False
    for i in range(1, max(1, retries) + 1):
        try:
            return session.post(url, data=data, headers=hdrs, timeout=timeout)
        except Exception as e:  # noqa: BLE001
            last = e
            if i < retries:
                time.sleep(min(0.8 * i, 3.0))
    if label:
        (log or _noop)(f"⚠️ {label} 请求失败（已重试{retries}次）: {type(last).__name__}: {str(last)[:120]}")
    return None


# ======================================================================
# 磁盘缓存（JSON，带 TTL）
# ======================================================================
class DiskCache:
    """简单可靠的磁盘缓存。云端重启后失效也没关系，只是慢一点。"""

    def __init__(self, cache_dir: str, ttl_hours: float = 12.0):
        self.cache_dir = cache_dir
        self.ttl = timedelta(hours=ttl_hours)
        os.makedirs(self.cache_dir, exist_ok=True)

    def _path(self, key: str) -> str:
        safe = hashlib.md5(key.encode("utf-8")).hexdigest()
        return os.path.join(self.cache_dir, f"{safe}.json")

    def get(self, key: str) -> Optional[Any]:
        path = self._path(key)
        try:
            if not os.path.exists(path):
                return None
            if datetime.now() - datetime.fromtimestamp(os.path.getmtime(path)) > self.ttl:
                return None
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return None

    def set(self, key: str, value: Any) -> None:
        try:
            with open(self._path(key), "w", encoding="utf-8") as f:
                json.dump(value, f, ensure_ascii=False)
        except Exception:
            pass

    def cached(self, key: str, producer: Callable[[], Any], *, log: Optional[LogFn] = None,
               label: str = "") -> Any:
        hit = self.get(key)
        if hit is not None:
            return hit
        val = producer()
        if val is not None:
            self.set(key, val)
        return val


# ======================================================================
# 数值/文本清洗
# ======================================================================
def safe_float(x: Any) -> Optional[float]:
    """与桌面版行为一致：'12.3%' -> 12.3，'1,234' -> 1234.0，非法值 -> None。"""
    if x is None:
        return None
    if isinstance(x, (int, float)):
        try:
            if isinstance(x, float) and math.isnan(x):
                return None
        except Exception:
            pass
        return float(x)
    s = str(x).strip()
    if s in ("", "-", "--", "N/A", "NA", "None", "null", "nan", "False"):
        return None
    s = s.replace("%", "").replace(",", "").replace("，", "").strip()
    m = re.search(r"-?\d+(\.\d+)?", s)
    try:
        return float(m.group()) if m else None
    except Exception:
        return None


def to_yuan(val: Any, unit_hint: str = "") -> Optional[float]:
    """把 '1.02亿' / '3350.57万' / '1234' 统一换算成元。"""
    if val is None:
        return None
    if isinstance(val, (int, float)):
        try:
            if isinstance(val, float) and math.isnan(val):
                return None
        except Exception:
            pass
        return float(val)
    s = str(val).strip().replace(",", "").replace("，", "")
    if s in ("", "-", "--", "False", "None", "nan"):
        return None
    mult = 1.0
    if "万亿" in s:
        mult = 1e12
    elif "亿" in s:
        mult = 1e8
    elif "万" in s:
        mult = 1e4
    elif "千" in s:
        mult = 1e3
    elif unit_hint == "亿":
        mult = 1e8
    elif unit_hint == "万":
        mult = 1e4
    num = safe_float(s)
    return None if num is None else num * mult


def norm_text(x: Any) -> str:
    if x is None:
        return ""
    try:
        if isinstance(x, float) and math.isnan(x):
            return ""
    except Exception:
        pass
    return str(x).strip().replace("\n", "").replace("\r", "").replace("\u200b", "")


def norm_item(name: Any) -> str:
    """科目名归一：去空格、全角括号转半角、帐->账（与原程序一致）。"""
    s = norm_text(name)
    s = s.replace(" ", "").replace("　", "")
    s = s.replace("（", "(").replace("）", ")")
    s = s.replace("帐", "账")
    return s


def now_ts() -> str:
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def ensure_dir(path: str) -> str:
    os.makedirs(path, exist_ok=True)
    return path


def df_to_records(df: Optional[pd.DataFrame], limit: int = 200) -> List[Dict[str, Any]]:
    """把 DataFrame 安全转成可 JSON 序列化的记录列表（供前端展示）。"""
    if df is None or not isinstance(df, pd.DataFrame) or df.empty:
        return []
    out: List[Dict[str, Any]] = []
    for _, row in df.head(limit).iterrows():
        rec = {}
        for k, v in row.items():
            if pd.isna(v):
                rec[str(k)] = ""
            elif isinstance(v, (int, float)):
                rec[str(k)] = v
            else:
                rec[str(k)] = str(v)
        out.append(rec)
    return out


def chunked(items: Iterable[Any], size: int) -> Iterable[List[Any]]:
    buf: List[Any] = []
    for it in items:
        buf.append(it)
        if len(buf) >= size:
            yield buf
            buf = []
    if buf:
        yield buf
