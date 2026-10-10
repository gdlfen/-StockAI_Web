# -*- coding: utf-8 -*-
"""员工人数补充数据（用于恢复排雷 C1/C2 与桌面版一致的判定）。

背景
----
桌面版的排雷 C1「员工与营收背离」与 C2「人均薪酬异常」依赖**员工人数**，
该数据在桌面版来自年报 PDF 的「员工情况」章节。免费公开数据源（AkShare /
Baostock / TuShare / 巨潮 / 东财）都**没有**可用的“按公司查历年员工人数”接口，
因此云端版默认把这两条标注为「数据不足」而不是按 0 计算（按 0 会得出
`员工降幅=100%`、`营收增幅=1.1e13%` 这类垃圾值并误报预警）。

本模块让用户**手工补充**员工人数；一旦补齐，C1/C2 自动恢复为与桌面版
逐字一致的算式与阈值判定（见 ``core.fraud.execute_18_conditions_analysis``）。

支持的输入格式（自动识别）
--------------------------
1) **长表**（推荐）：三列 ``代码 | 年份 | 员工人数``
2) **宽表**：``代码 | 2021 | 2022 | 2023 ...``（首列为代码，其余列为各年人数）
3) 列名宽松：代码列可叫 ``代码/股票代码/证券代码/code``；年份列可叫
   ``年份/年度/报告期/year``；人数列可叫 ``员工人数/员工总数/人数/职工人数/employee``；
   人数还允许写成 ``12,345``、``12345人``、``1.23万`` 等。

文件格式：``.xlsx`` / ``.xls`` / ``.csv``（CSV 自动尝试 utf-8-sig / gbk）。
"""
from __future__ import annotations

import io
import os
import re
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd

EMPLOYEE_FILE_CANDIDATES = [
    "员工人数补充.csv",
    "员工人数补充.xlsx",
    "员工人数.csv",
    "员工人数.xlsx",
    "employee_counts.csv",
    "employee_counts.xlsx",
]

_CODE_COLS = ("代码", "股票代码", "证券代码", "证券代码(代码)", "code", "symbol", "股票代码(代码)")
_YEAR_COLS = ("年份", "年度", "报告期", "会计年度", "year", "报表年度")
_NUM_COLS = ("员工人数", "员工总数", "人数", "职工人数", "在职员工人数", "员工数量",
             "employee", "employees", "staff", "count")


# ----------------------------------------------------------------------
# 解析工具
# ----------------------------------------------------------------------
def parse_count(v: Any) -> Optional[int]:
    """把 ``12345`` / ``'12,345'`` / ``'12345人'`` / ``'1.23万'`` 等解析成整数人数。"""
    if v is None:
        return None
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        try:
            f = float(v)
        except Exception:
            return None
        if f != f or f <= 0:          # NaN / 非正
            return None
        return int(round(f))
    s = str(v).strip().replace(",", "").replace("，", "").replace(" ", "")
    if not s or s.lower() in ("nan", "none", "-", "—", "null"):
        return None
    m = re.search(r"(-?\d+(?:\.\d+)?)", s)
    if not m:
        return None
    try:
        f = float(m.group(1))
    except Exception:
        return None
    if f <= 0:
        return None
    if "万" in s:
        f *= 10000
    return int(round(f))


def _norm_code(v: Any) -> str:
    s = str(v).strip()
    m = re.search(r"(\d{1,6})", s.split(".")[0] if s.endswith(".0") else s)
    return m.group(1).zfill(6) if m else ""


def _norm_year(v: Any) -> Optional[int]:
    s = str(v).strip()
    m = re.search(r"(19|20)\d{2}", s)
    if not m:
        return None
    y = int(m.group(0))
    return y if 1990 <= y <= 2100 else None


# ----------------------------------------------------------------------
# 读取
# ----------------------------------------------------------------------
def _read_table(data: bytes, filename: str) -> List[pd.DataFrame]:
    """把上传的字节读成若干 DataFrame（xlsx 每个 sheet 一个）。"""
    name = (filename or "").lower()
    frames: List[pd.DataFrame] = []
    if name.endswith((".xlsx", ".xls")):
        xl = pd.ExcelFile(io.BytesIO(data))
        for sh in xl.sheet_names:
            df = xl.parse(sh, dtype=object)
            df.attrs["__sheet__"] = sh
            frames.append(df)
        return frames
    last_err: Optional[Exception] = None
    for enc in ("utf-8-sig", "gbk", "utf-8"):
        try:
            return [pd.read_csv(io.BytesIO(data), encoding=enc, dtype=object)]
        except Exception as e:  # noqa: BLE001
            last_err = e
    raise ValueError(f"CSV 解析失败（已尝试 utf-8-sig/gbk/utf-8）：{last_err}")


def _find_col(df: pd.DataFrame, names: Tuple[str, ...]) -> Optional[str]:
    cols = [str(c).strip() for c in df.columns]
    for want in names:
        for c in cols:
            if c.lower() == want.lower():
                return c
    for want in names:
        for c in cols:
            if want.lower() in c.lower():
                return c
    return None


def parse_employee_bytes(data: bytes, filename: str = "") -> Tuple[Dict[str, Dict[int, int]], List[str]]:
    """解析上传内容 → ``({代码: {年份: 人数}}, 跳过说明)``。

    同时支持长表与宽表；无法识别的行会被记入第二个返回值（用于界面提示）。
    任何格式错误都只记入提示、不抛异常。
    """
    out: Dict[str, Dict[int, int]] = {}
    notes: List[str] = []
    if not data or not data.strip(b"\x00\r\n\t "):
        return {}, ["文件内容为空"]
    try:
        frames = _read_table(data, filename)
    except Exception as e:  # noqa: BLE001
        return {}, [f"文件解析失败：{type(e).__name__}: {str(e)[:160]}"]

    for df in frames:
        sheet = df.attrs.get("__sheet__", "")
        if df is None or df.empty:
            continue
        df = df.dropna(how="all")
        if df.empty:
            continue
        code_col = _find_col(df, _CODE_COLS)
        year_col = _find_col(df, _YEAR_COLS)
        num_col = _find_col(df, _NUM_COLS)

        # ---- 形态 A：长表（代码 | 年份 | 人数）----
        if code_col and year_col and num_col:
            for i, r in df.iterrows():
                code = _norm_code(r.get(code_col))
                year = _norm_year(r.get(year_col))
                n = parse_count(r.get(num_col))
                if not code or not year or not n:
                    if code or year or r.get(num_col):
                        notes.append(f"[{sheet or filename}] 第{i + 2}行无法识别："
                                     f"代码={r.get(code_col)!r} 年份={r.get(year_col)!r} "
                                     f"人数={r.get(num_col)!r}")
                    continue
                out.setdefault(code, {})[year] = n
            continue

        # ---- 形态 B：宽表（首列代码，其余列名是年份）----
        if code_col:
            year_map: Dict[str, int] = {}
            for c in df.columns:
                if str(c) == str(code_col):
                    continue
                y = _norm_year(c)
                if y:
                    year_map[str(c)] = y
            if not year_map:
                notes.append(f"[{sheet or filename}] 未找到年份列，已跳过（列名：{list(df.columns)[:8]}）")
                continue
            for i, r in df.iterrows():
                code = _norm_code(r.get(code_col))
                if not code:
                    continue
                for c, y in year_map.items():
                    n = parse_count(r.get(c))
                    if n:
                        out.setdefault(code, {})[y] = n
            continue

        notes.append(f"[{sheet or filename}] 既没有“代码”列也没有年份列，已跳过")

    # 去掉只解析出空字典的代码
    out = {c: v for c, v in out.items() if v}
    return out, notes


def load_employee_file(path: str) -> Tuple[Dict[str, Dict[int, int]], List[str]]:
    """从磁盘路径读取（用于读取项目内随仓库一起上传的补充文件）。"""
    if not os.path.isfile(path):
        return {}, []
    with open(path, "rb") as fh:
        return parse_employee_bytes(fh.read(), os.path.basename(path))


def load_supplement(data_dir: str, project_root: str = "") -> Tuple[Dict[str, Dict[int, int]], str]:
    """按优先级查找员工人数补充数据。

    1. ``<会话目录>/员工人数补充.*``（用户在页面上传后落盘，本次会话有效）
    2. ``<项目根>/员工人数补充.*``（随 GitHub 仓库一起上传 → **跨重启持久有效**）

    返回 ``(数据, 来源说明)``。
    """
    for base, tag in ((data_dir, "本次会话上传"), (project_root, "项目内文件（随仓库上传，持久）")):
        if not base:
            continue
        for fn in EMPLOYEE_FILE_CANDIDATES:
            p = os.path.join(base, fn)
            if os.path.isfile(p):
                try:
                    d, _notes = load_employee_file(p)
                    if d:
                        return d, f"{tag}：{fn}（{len(d)} 家公司）"
                except Exception:  # noqa: BLE001
                    continue
    return {}, ""


def save_supplement(data_dir: str, data: Dict[str, Dict[int, int]]) -> str:
    """把补充数据落盘到会话目录（CSV，便于用户下载/复用/上传到仓库）。"""
    os.makedirs(data_dir, exist_ok=True)
    p = os.path.join(data_dir, "员工人数补充.csv")
    rows = [{"代码": c, "年份": y, "员工人数": n}
            for c in sorted(data) for y, n in sorted(data[c].items())]
    pd.DataFrame(rows, columns=["代码", "年份", "员工人数"]).to_csv(
        p, index=False, encoding="utf-8-sig")
    return p


def to_csv_bytes(data: Dict[str, Dict[int, int]]) -> bytes:
    rows = [{"代码": c, "年份": y, "员工人数": n}
            for c in sorted(data) for y, n in sorted(data[c].items())]
    return pd.DataFrame(rows, columns=["代码", "年份", "员工人数"]).to_csv(
        index=False, encoding="utf-8-sig").encode("utf-8-sig")


def template_bytes() -> bytes:
    """下载用模板（长表 + 示例行）。"""
    demo = {
        "600887": {2021: 56416, 2022: 60489, 2023: 62698, 2024: 63119, 2025: 64000},
        "605499": {2021: 8210, 2022: 9180, 2023: 10240, 2024: 11350, 2025: 12000},
    }
    return to_csv_bytes(demo)


# ----------------------------------------------------------------------
# 合并进排雷的公司数据
# ----------------------------------------------------------------------
def merge_into_company_data(comp_data: Dict[str, Any],
                            supplement: Dict[str, Dict[int, int]],
                            log=None) -> int:
    """把补充的员工人数写入 ``comp_data``（结构同 ``fraud.extract_company_data``）。

    ``comp_data`` 的键形如 ``"600887 伊利股份"``，年度的键是字符串 ``"2022"``。
    返回实际写入的条数。
    """
    if not supplement:
        return 0
    filled = 0
    for key, years_map in comp_data.items():
        code = _norm_code(str(key).split()[0] if str(key).strip() else "")
        if not code or code not in supplement:
            continue
        for ykey, metrics in (years_map or {}).items():
            y = _norm_year(ykey)
            if not y or not isinstance(metrics, dict):
                continue
            n = supplement[code].get(y)
            if n:
                metrics["员工人数"] = float(n)
                filled += 1
    if filled and log is not None:
        with_emp = sum(1 for v in comp_data.values()
                       if any((m or {}).get("员工人数") for m in v.values() if isinstance(m, dict)))
        log(f"   ✅ 已补充员工人数：写入 {filled} 条年度记录（覆盖 {with_emp} 家公司），"
            f"C1/C2 将按桌面版口径判定")
    return filled
