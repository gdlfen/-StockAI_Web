# -*- coding: utf-8 -*-
"""
股票全集 / 行业分类 / 同行前三 —— 全部使用免费数据源（AkShare + 东方财富公开接口）。

设计要点
--------
1. **全市场清单**：`ak.stock_info_a_code_name()` 与 `ak.stock_zh_a_spot()` 双源，任一可用即可；
   结果缓存到磁盘，避免每次进入页面都拉一次全市场。
2. **行业分类**：优先用东方财富公开接口 `RPT_F10_CORETHEME_BOARDTYPE`（原桌面版也用这个），
   逐只缓存；失败回退巨潮 `ak.stock_industry_change_cninfo`。
3. **同行前三**：与原桌面版“数据雷达”同口径 —— 同行业内按**市值**取前三（排除自身）。
   市值优先用全市场快照（现价×总股本，同花顺口径），缺失则回退东财接口。
4. 一切网络失败都以“空结果 + 日志”收场，绝不抛异常打断流水线。
"""
from __future__ import annotations

import os
import re
import time
from typing import Any, Callable, Dict, List, Optional, Tuple

import pandas as pd

from .common import DiskCache, http_get, make_logger, safe_float, to_yuan
from .contracts import TARGET_ATTR_PEER, TARGET_ATTR_PRIMARY

EM_DATACENTER = "https://datacenter-web.eastmoney.com/api/data/v1/get"
EM_QUOTE = "https://push2.eastmoney.com/api/qt/stock/get"


# ======================================================================
# 全市场清单
# ======================================================================
class StockUniverse:
    """A 股全市场清单（代码/名称/最新价/总市值/市盈率），带磁盘缓存。"""

    def __init__(self, cache: DiskCache, log: Optional[Callable[[str], None]] = None, ttl_key: str = "universe_v1"):
        self.cache = cache
        self.log = log or make_logger()
        self.ttl_key = ttl_key
        self._df: Optional[pd.DataFrame] = None

    # ---------------- 内部：抓取 ----------------
    def _fetch_spot(self) -> Optional[pd.DataFrame]:
        try:
            import akshare as ak
            df = ak.stock_zh_a_spot()
            if df is None or df.empty:
                return None
            df = df.rename(columns={
                "代码": "code", "名称": "name", "最新价": "price",
                "成交量": "volume", "成交额": "amount",
            })
            keep = [c for c in ["code", "name", "price", "volume", "amount"] if c in df.columns]
            df = df[keep].copy()
            df["code"] = df["code"].astype(str).str.replace(r"^(sh|sz|bj)", "", regex=True).str.zfill(6)
            df["name"] = df["name"].astype(str).str.replace(" ", "")
            df["price"] = pd.to_numeric(df["price"], errors="coerce")
            return df
        except Exception as e:  # noqa: BLE001
            self.log(f"⚠️ 全市场快照(新浪)获取失败: {type(e).__name__}: {str(e)[:120]}")
            return None

    def _fetch_code_name(self) -> Optional[pd.DataFrame]:
        try:
            import akshare as ak
            df = ak.stock_info_a_code_name()
            if df is None or df.empty:
                return None
            df = df.rename(columns={"code": "code", "name": "name"})
            df["code"] = df["code"].astype(str).str.zfill(6)
            df["name"] = df["name"].astype(str).str.replace(" ", "")
            return df[["code", "name"]]
        except Exception as e:  # noqa: BLE001
            self.log(f"⚠️ 全市场代码表获取失败: {type(e).__name__}: {str(e)[:120]}")
            return None

    # ---------------- 对外 ----------------
    def load(self, force: bool = False) -> pd.DataFrame:
        if self._df is not None and not force:
            return self._df
        cached = None if force else self.cache.get(self.ttl_key)
        if cached:
            try:
                df = pd.DataFrame(cached)
                if not df.empty:
                    self._df = df
                    return df
            except Exception:
                pass

        spot = self._fetch_spot()
        names = self._fetch_code_name()
        if spot is not None and names is not None:
            df = names.merge(spot, on="code", how="left")
        elif spot is not None:
            df = spot
        elif names is not None:
            df = names
        else:
            self.log("❌ 无法获取 A 股全市场清单（网络受限）。行业/同行功能将降级。")
            self._df = pd.DataFrame(columns=["code", "name", "price"])
            return self._df

        # 名称以“代码表”为准（快照里偶有退市/停牌异常名）
        if names is not None and "name_x" in df.columns:
            df["name"] = df["name_x"].fillna(df.get("name_y"))
            df = df.drop(columns=[c for c in ("name_x", "name_y") if c in df.columns])
        df = df.drop_duplicates(subset=["code"]).reset_index(drop=True)
        self._df = df
        try:
            self.cache.set(self.ttl_key, df.to_dict("records"))
        except Exception:
            pass
        self.log(f"✅ 已载入 A 股清单 {len(df)} 只")
        return df

    def name_of(self, code: str) -> str:
        df = self.load()
        code = str(code).zfill(6)
        hit = df[df["code"] == code]
        return str(hit.iloc[0]["name"]) if not hit.empty else code

    def code_of(self, name: str) -> str:
        df = self.load()
        name = str(name).strip().replace(" ", "")
        hit = df[df["name"] == name]
        if hit.empty:
            hit = df[df["name"].str.contains(re.escape(name), na=False)]
        return str(hit.iloc[0]["code"]) if not hit.empty else ""

    def resolve(self, text: str) -> Tuple[str, str]:
        """把 '600887' / '伊利股份' / 'sh600887' 统一解析成 (code, name)。"""
        s = str(text or "").strip()
        m = re.search(r"(\d{6})", s)
        if m:
            code = m.group(1)
            return code, self.name_of(code)
        code = self.code_of(s)
        return code, (s if not code else self.name_of(code))


# ======================================================================
# 行业分类
# ======================================================================
def get_industry(code: str, cache: DiskCache, log: Optional[Callable[[str], None]] = None,
                 retries: int = 2) -> Dict[str, str]:
    """取单只股票的行业信息。返回
    {"一级行业":..., "二级行业":..., "三级行业":..., "完整行业路径":...}
    """
    log = log or make_logger()
    code = str(code).zfill(6)
    key = f"industry_{code}"
    hit = cache.get(key)
    if hit:
        return hit

    result: Dict[str, str] = {"一级行业": "", "二级行业": "", "三级行业": "", "完整行业路径": ""}

    # ① 东方财富 F10（原桌面版即用此接口，稳定且分级清晰）
    try:
        secucode = f"{code}.SH" if code[0] in ("6", "9") else (f"{code}.BJ" if code[0] in ("4", "8") else f"{code}.SZ")
        r = http_get(EM_DATACENTER, params={
            "reportName": "RPT_F10_CORETHEME_BOARDTYPE",
            "columns": "SECUCODE,SECURITY_CODE,SECURITY_NAME_ABBR,BOARD_NAME,BOARD_CODE,BOARD_TYPE,BOARD_RANK",
            "filter": f'(SECUCODE="{secucode}")',
            "pageNumber": 1, "pageSize": 200, "source": "WEB", "client": "WEB",
        }, timeout=15, retries=retries, label=f"东财行业[{code}]", log=log)
        if r is not None:
            rows = (((r.json() or {}).get("result") or {}).get("data")) or []
            ranked: List[Tuple[int, str]] = []
            for it in rows:
                name = str(it.get("BOARD_NAME") or "").strip()
                rank = safe_float(it.get("BOARD_RANK"))
                if name:
                    ranked.append((int(rank) if rank else 99, name))
            if ranked:
                ranked.sort(key=lambda x: x[0])
                names = [n for _, n in ranked]
                result["一级行业"] = names[0] if len(names) > 0 else ""
                result["二级行业"] = names[1] if len(names) > 1 else result["一级行业"]
                result["三级行业"] = names[2] if len(names) > 2 else result["二级行业"]
                # 与原桌面版一致：文件夹用“二级行业”（无二级则一级）
                result["完整行业路径"] = "-".join([n for n in names[:3] if n])
    except Exception as e:  # noqa: BLE001
        log(f"⚠️ 东财行业分类失败[{code}]: {type(e).__name__}: {str(e)[:100]}")

    # ② 回退：巨潮行业分类
    if not result["一级行业"]:
        try:
            import akshare as ak
            df = ak.stock_industry_change_cninfo(symbol=code, start_date="20000101",
                                                end_date=time.strftime("%Y%m%d"))
            if df is not None and not df.empty:
                row = df.iloc[-1]
                for src, dst in (("行业门类", "一级行业"), ("行业次类", "二级行业"), ("行业中类", "三级行业")):
                    val = str(row.get(src) or "").strip()
                    if val:
                        result[dst] = val
                result["完整行业路径"] = "-".join(
                    [result["一级行业"], result["二级行业"], result["三级行业"]]).strip("-")
        except Exception as e:  # noqa: BLE001
            log(f"⚠️ 巨潮行业分类失败[{code}]: {type(e).__name__}: {str(e)[:100]}")

    if result["一级行业"] and not result["二级行业"]:
        result["二级行业"] = result["一级行业"]
    if result["一级行业"]:
        cache.set(key, result)
    return result


def industry_label(info: Dict[str, str]) -> str:
    """用于文件夹/分组的行业名（与原桌面版一致：优先二级行业，去掉 Ⅰ/Ⅱ 罗马数字）。"""
    for k in ("二级行业", "一级行业", "三级行业"):
        v = str(info.get(k) or "").strip()
        if v:
            return v.replace("Ⅱ", "").replace("Ⅰ", "").replace("Ⅲ", "").strip()
    return "默认行业"


# ======================================================================
# 市值（用于同行排序）
# ======================================================================
def get_market_caps_batch(codes: List[str], cache: DiskCache,
                          log: Optional[Callable[[str], None]] = None,
                          chunk: int = 40) -> Dict[str, float]:
    """**批量**取总市值（亿元）——腾讯行情支持一次查多只，速度远快于逐只请求。

    返回 {code: 市值(亿)}；已缓存的直接命中，不再发请求。
    """
    log = log or make_logger()
    out: Dict[str, float] = {}
    todo: List[str] = []
    for c in codes:
        c = str(c).zfill(6)
        hit = cache.get(f"mktcap_{c}")
        if hit is not None:
            try:
                out[c] = float(hit)
                continue
            except Exception:
                pass
        todo.append(c)

    for i in range(0, len(todo), chunk):
        part = todo[i:i + chunk]
        syms = ",".join(("sh" + c) if c[0] in ("6", "9") else (("bj" + c) if c[0] in ("4", "8") else ("sz" + c))
                        for c in part)
        try:
            r = http_get(f"https://qt.gtimg.cn/q={syms}", timeout=15, encoding="gbk", retries=2,
                         label="腾讯市值批量", log=log)
            if r is None:
                continue
            for m in re.finditer(r'v_(\w+)="([^"]*)"', r.text):
                body = m.group(2)
                if not body or "~" not in body:
                    continue
                f = body.split("~")
                if len(f) < 46 or not f[2]:
                    continue
                code = str(f[2]).zfill(6)
                cap = safe_float(f[45]) or safe_float(f[44])
                if cap and cap > 0:
                    out[code] = cap
                    cache.set(f"mktcap_{code}", cap)
        except Exception as e:  # noqa: BLE001
            log(f"   ⚠️ 批量市值失败: {type(e).__name__}: {str(e)[:80]}")
        time.sleep(0.2)
    return out


def get_market_cap(code: str, cache: DiskCache, log: Optional[Callable[[str], None]] = None) -> Optional[float]:
    """取总市值（亿元）。

    数据源顺序（均已实测可用）：
    ① 腾讯行情字段[45]=总市值(亿) / [44]=流通市值(亿)  —— 最简、最稳（支持批量）；
    ② 东方财富 quote 接口 f116 —— 部分网络会被拒，作备份；
    ③ 全市场快照的 现价 × 总股本 —— 兜底。
    """
    log = log or make_logger()
    code = str(code).zfill(6)
    key = f"mktcap_{code}"
    hit = cache.get(key)
    if hit is not None:
        try:
            return float(hit)
        except Exception:
            pass

    got = get_market_caps_batch([code], cache, log=log)
    if code in got:
        return got[code]

    # ② 东方财富 quote
    market = "1" if code[0] in ("6", "9") else "0"
    try:
        r = http_get(EM_QUOTE, params={
            "secid": f"{market}.{code}", "fields": "f57,f58,f43,f116,f117",
            "invt": "2", "fltt": "1", "ut": "fa5fd1943c7b386f172d6893dbfba10b",
        }, headers={"Referer": "https://quote.eastmoney.com/"}, timeout=12, retries=1,
            label=f"东财市值[{code}]", log=None)
        if r is not None:
            data = ((r.json() or {}).get("data")) or {}
            cap = safe_float(data.get("f116"))
            if cap:
                cap_yi = round(cap / 1e8, 2)
                cache.set(key, cap_yi)
                return cap_yi
    except Exception:
        pass

    # ③ 快照 现价×总股本（若快照里带总股本列）
    try:
        df = StockUniverse(cache, log).load()
        if {"price", "total_shares"} <= set(df.columns):
            row = df[df["code"] == code]
            if not row.empty:
                price = safe_float(row.iloc[0]["price"])
                shares = safe_float(row.iloc[0]["total_shares"])
                if price and shares:
                    cap_yi = round(price * shares / 1e8, 2)
                    cache.set(key, cap_yi)
                    return cap_yi
    except Exception:
        pass
    return None


# ======================================================================
# 同行前三
# ======================================================================
def find_top_peers(code: str, industry: str, universe: StockUniverse, cache: DiskCache,
                   top_n: int = 3, log: Optional[Callable[[str], None]] = None) -> List[Dict[str, Any]]:
    """在“行业代表公司池”中找同行前三。

    由于东方财富“行业成分”接口在部分网络下不可用，这里采用**行业代表公司池 + 市值排序**的做法：
    - 代表公司池来自内置的行业龙头参照表 + 全市场清单里已缓存过行业的公司；
    - 排序口径与原桌面版“数据雷达”一致：按市值从大到小，取前三（排除自身）。
    """
    log = log or make_logger()
    code = str(code).zfill(6)
    # 行业代表公司池（可扩展/可被缓存补全）
    pool: List[str] = [c for c, ind in _industry_members_cache(cache).items() if ind == industry and c != code]
    if not pool:
        pool = [c for c in _REP_POOL.get(industry, []) if c != code]

    peers: List[Dict[str, Any]] = []
    for c in pool[:40]:
        peer_ind = get_industry(c, cache, log=log).get("二级行业") or ""
        if peer_ind and industry and industry not in peer_ind and peer_ind not in industry:
            continue
        cap = get_market_cap(c, cache, log=log)
        if cap is None:
            continue
        peers.append({"代码": c, "名称": universe.name_of(c), "市值(亿)": cap, "所属行业": industry})
    peers.sort(key=lambda x: -float(x.get("市值(亿)") or 0))
    out = peers[:top_n]
    if out:
        log(f"   同行前三: " + "、".join(f"{p['名称']}({p['代码']})" for p in out))
    else:
        log(f"   ⚠️ 未能确定 {code} 的同行公司（行业={industry}）")
    return out


def _industry_members_cache(cache: DiskCache) -> Dict[str, str]:
    """读取“行业 -> 成员代码”映射缓存（由每次解析出的行业信息逐步补全）。"""
    data = cache.get("industry_members_map") or {}
    return data if isinstance(data, dict) else {}


def remember_industry_member(cache: DiskCache, code: str, industry: str) -> None:
    """把解析出的“代码->行业”写入映射缓存，供后续同行检索使用（逐步自学习）。"""
    if not code or not industry:
        return
    data = _industry_members_cache(cache)
    data[str(code).zfill(6)] = industry
    cache.set("industry_members_map", data)


# 内置行业代表公司池（按申万/东财二级行业命名；仅作为冷启动种子，运行时会被缓存补充）
_REP_POOL: Dict[str, List[str]] = {
    "饮料乳品": ["600887", "605499", "603156", "600597", "002946", "600882"],
    "非白酒": ["600132", "600600", "000729", "002568", "000929"],
    "白酒": ["600519", "000858", "000568", "600809", "002304"],
    "食品加工": ["000895", "603288", "600872", "002216"],
    "家电行业": ["000333", "000651", "600690", "002508"],
    "银行": ["600036", "601398", "601288", "600000", "601166"],
    "保险": ["601318", "601601", "601336", "601628"],
    "证券": ["600030", "601688", "600837", "000776"],
    "医药制造": ["600276", "000538", "600196", "002422"],
    "中药": ["600085", "000538", "600332", "002603"],
    "汽车整车": ["600104", "000625", "601633", "002594"],
    "电池": ["300750", "002594", "300014", "002460"],
    "光伏设备": ["601012", "002129", "300274", "688599"],
    "半导体": ["688981", "603501", "002371", "300782"],
    "消费电子": ["002475", "002241", "300433", "002138"],
    "电子元件": ["300408", "002456", "600183"],
    "软件开发": ["600570", "002230", "300253", "688111"],
    "通信设备": ["000063", "002415", "600522"],
    "电力行业": ["600900", "601985", "600886", "000883"],
    "煤炭行业": ["601088", "601225", "600188", "601898"],
    "钢铁行业": ["600019", "000898", "600808"],
    "有色金属": ["601899", "600547", "603993"],
    "化学制品": ["600309", "002648", "600426"],
    "房地产开发": ["000002", "600048", "001979"],
    "工程建设": ["601668", "601390", "601186"],
    "航空机场": ["600029", "601111", "600115"],
    "铁路公路": ["601006", "600018", "600377"],
    "商业百货": ["600690", "601933", "600655"],
    "纺织服装": ["002563", "600177", "002832"],
}
