# -*- coding: utf-8 -*-
"""
财务造假排雷（云端后端版）—— **移植自桌面版 FraudDetectionApp，逻辑等价**。

来源：``价值投资分析模型V1.4-3（扩展）1.0（修复版）.py`` 中的
``class FraudDetectionApp``（约 5321~5863 行）：

    extract_excel_data / process_task / execute_18_conditions_analysis / generate_excel_report

本模块把这些能力改造成无 GUI、无浏览器、无线程界面的纯后端函数，供 FastAPI / Streamlit 调用：

    extract_company_data()           <- extract_excel_data
    analyze_one_company()            <- process_task 的单公司版本（逐条跑 18 条判断并汇总）
    execute_18_conditions_analysis() <- 同名方法（逐字照搬，供审计比对）
    generate_excel_report()          <- 同名方法（版式/列名逐字保留）
    generate_report()                <- 由 analyze_one_company 结果出报告
    run_all()                        <- process_task 全流程（抽取 -> 判断 -> 出报告）

去 UI 化说明（**只做这些改动，未改动任何判断规则/阈值/条件顺序/报告版式**）
------------------------------------------------------------------
1. 删除 ``ttk.Frame`` 继承、``tk.*`` 控件、``filedialog``、``messagebox``、
   ``self.after(...)``、``scrolledtext``、``threading.Thread`` 启动逻辑；
   原 ``log(msg)`` 改为 ``log`` 回调（缺省走 ``core.common.make_logger``，打印到 stdout）。
2. 原 ``self.params`` / ``self.c15_components_var`` / ``self.c16_components_var`` /
   ``self.macro_db`` 由「控件变量」变成「函数入参 + 模块常量」；
   ``start_analysis_thread`` 里的 ``current_params = {k: v["val"].get() ...}`` 等价于
   :func:`normalize_params`。
3. 原 ``silent`` / ``callback`` 两个界面相关参数不再需要；如需保留调用点兼容，
   :func:`run_all` / :func:`process_task` 提供同义可选参数 ``silent`` / ``callback``
   （仅用于抑制日志与流程结束回调，不影响判断结果）。
4. 原 ``extract_excel_data`` 中静默的 ``except: pass`` 保留「不崩」语义，但按云端要求
   **补记一条日志**（只影响日志，不影响抽取结果）。
5. ``process_task`` 里「单个公司异常 → 整个批次中断」改为「单个公司异常 → 记日志后跳过下一家」，
   以满足「单个文件/公司失败不崩」的容错要求（桌面版是整批 try/except）。

阈值来源的重要提示（请勿误用）
------------------------------------------------------------------
桌面版真正参与计算的参数是 ``FraudDetectionApp.__init__`` 里的 ``self.params``
（键名 ``c1_emp_drop`` / ``c1_rev_inc`` / ``c12_fa_cip_ratio`` / ``c15_cash_limit`` /
``c17_margin_jump`` … 共 22 项 + 2 个科目字符串），本模块以它为准复刻为 :data:`FRAUD_DEFAULTS`。

``core.config.FRAUD_PARAM_DEFS`` 是一套**改过名的合并版**参数（键名 ``c1_ratio`` /
``c2_gap`` / ``c12_fa_growth`` …），原程序里没有任何代码读取这些键名，且它与桌面版阈值
**不能一一对应**（例如桌面版 C1 是「员工降幅>10 且 营收增幅>10」双阈值，配置层只有一个
``c1_ratio``=5.0；桌面版 C15 的 ``c15_cash_limit`` 在配置层缺失）。
因此本模块的默认阈值 = 桌面版原值；配置层的键通过 :data:`_PARAM_ALIASES` 做**语义一一对应**的
别名兼容（同名同值，不产生任何行为漂移），语义不明确/会改变判断结果的键（``c1_ratio``、
``c12_fa_growth``、``c13_ar_inv_growth``）**忽略并记警告**，避免静默改动阈值。
如需让配置层生效，请直接使用桌面版键名传参（见 :func:`normalize_params` 的 docstring）。

行业均值数据库
------------------------------------------------------------------
C2 用 ``默认_平均工资``，C17 用 ``默认_毛利率``，C18 用 ``默认_应收占比`` / ``默认_存货占比``。
其中工资表取自 ``core.config.MACRO_AVG_WAGE``（含桌面版 2018~2024 全部同值，另多 2025/2026 两年；
配置优先），其余三张表原程序未放进配置文件，此处按桌面版 ``self.macro_db`` 逐字固化。

其它
------------------------------------------------------------------
- 只依赖标准库 / pandas / numpy(间接) / openpyxl / ``core.config`` / ``core.contracts`` / ``core.common``。
- 所有对外函数返回的结构均可 JSON 序列化（DataFrame 一律转 records）。
  ``extract_company_data`` 为与桌面版逐字一致会保留空单元格读出的 ``NaN``（``json.dumps`` 可序列化，
  但 FastAPI ``JSONResponse`` 这类 ``allow_nan=False`` 的编码器会报错）——此时请用
  :func:`sanitize_for_json` 包一层；``analyze_one_company`` / ``run_all`` 的返回值本身已无 NaN。
- 输入目录按原逻辑递归扫描《统一整合输出_*.xlsx》《*年报提取表.xlsx》《*公司属性表*》等。

输入容错（新增，桌面版无；只影响「能否取到数据」，不触碰任何判断规则）
------------------------------------------------------------------
1) 三大表支持两种版式：新版（第 1 行即表头，``header=0`` 后列名已是 ``['项目','20221231',...]``）
   与旧版（第 1 行是合并标题、第 2 行才是表头）。新版直接采用表头定位年份列与科目列，
   旧版仍走桌面版「扫前 15 行找 YYYYMMDD」的原逻辑。
2) 《分红情况》支持两种形态：形态1 按年汇总（``年份`` + ``分红率`` / ``同花顺现金分红总额(元)``）、
   形态2 东财逐笔（``报告期`` + ``现金分红-现金分红比例(每10股派息元)``）→ 按报告期年份聚合。
   新增的分红指标键（:data:`DIV_METRIC_RATE` / :data:`DIV_METRIC_CASH` / :data:`DIV_METRIC_PER10`）
   **不进入** :data:`core.contracts.FRAUD_METRIC_MAP`，因此不参与「指标补 0」、不被 18 条判断读取；
   且只并入三大表已存在的年度，绝不新增年度，保证 18 条判断的输入年度与桌面版完全一致。
"""
from __future__ import annotations

import os
import re
from typing import Any, Callable, Dict, List, Optional, Tuple

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

try:  # 包内导入（FastAPI / Streamlit：from core.fraud import ...）
    from .common import LogFn, df_to_records, ensure_dir, make_logger, safe_float
    from .config import FRAUD_PARAM_DEFS, MACRO_AVG_WAGE
    from .contracts import DIR_FRAUD, FRAUD_METRIC_MAP, UNIFIED_FILENAME_PREFIX
except ImportError:  # 允许把 core/ 加入 sys.path 后 `import fraud`
    from common import LogFn, df_to_records, ensure_dir, make_logger, safe_float  # type: ignore
    from config import FRAUD_PARAM_DEFS, MACRO_AVG_WAGE  # type: ignore
    from contracts import DIR_FRAUD, FRAUD_METRIC_MAP, UNIFIED_FILENAME_PREFIX  # type: ignore

__all__ = [
    "CONDITION_ORDER",
    "FRAUD_DEFAULTS",
    "FRAUD_PARAM_META",
    "MACRO_DB",
    "extract_company_data",
    "execute_18_conditions_analysis",
    "analyze_one_company",
    "generate_excel_report",
    "generate_report",
    "process_task",
    "run_all",
    "normalize_params",
    "build_macro_db",
    "sanitize_for_json",
    "DIV_METRIC_RATE",
    "DIV_METRIC_CASH",
    "DIV_METRIC_PER10",
]


# ======================================================================
# 一、18 条判断条件（顺序与桌面版 yearly_results 初始化顺序完全一致）
# ======================================================================
CONDITION_ORDER: List[str] = [
    "C1_员工与营收背离",
    "C2_人均薪酬异常",
    "C3_关联交易",
    "C4_重大股权投资",
    "C5_主营疲软投资异动",
    "C6_募资变更",
    "C7_高存差钱",
    "C8_应收账款预警",
    "C9_固定资产过高",
    "C10_在建工程过高",
    "C11_预付及其他应收",
    "C12_固资在建合计",
    "C13_双高联动风险",
    "C14_存单质押隐患",
    "C15_高存高贷双雷",
    "C16_减值计提异常",
    "C17_毛利率异常偏离",
    "C18_营运资产背离",
]


# ======================================================================
# 二、参数（逐字复制桌面版 FraudDetectionApp.__init__ 的 self.params）
# ======================================================================
#: 参数键 -> 界面名称（原 self.params[k]["name"]，供 Streamlit 表单复用）
FRAUD_PARAM_META: Dict[str, str] = {
    "c1_emp_drop": "1. 员工人数减少大于(%)",
    "c1_rev_inc": "1. 且营业收入增加大于(%)",
    "c2_salary_diff": "2. 薪酬低于行业平均大于(%)",
    "c3_related_tx": "3. 重大关联交易次数(>=)",
    "c5_rev_growth": "5. 营收增长率小于(%)",
    "c5_margin": "5. 或毛利率小于(%)",
    "c5_inv_ratio": "5. 或存货/总资产大于(%)",
    "c7_cash_ratio": "7. (货币+金融资产)/总资产小于(%)",
    "c8_ar_ratio": "8. 应收账款/总资产大于(%)",
    "c9_fa_ratio": "9. 固定资产/总资产大于(%)",
    "c10_cip_ratio": "10. 在建工程/总资产大于(%)",
    "c11_prepay_ratio": "11. (预付+其他应收)/总资产大于(%)",
    "c12_fa_cip_ratio": "12. (固资+在建)/总资产大于(%)",
    "c13_ar_limit": "13. 应收/总资>5%或存货/总资大于(%)",
    "c13_fa_cip_inc": "13. 且(固资+在建)增幅大于(%)",
    "c14_int_rate": "14. 利息收入/货币资金小于(%)",
    "c15_cash_limit": "15. 货币/总资>10%且借款/总资大于(%)",
    "c16_impair_ratio": "16. 减值损失/各项资产总和小于(%)",
    "c17_margin_jump": "17. 毛利率增幅大于(%)",
    "c17_margin_diff": "17. 且偏离行业平均毛利大于(%)",
    "c18_ar_diff": "18. 应收占比偏离行业平均大于(%)",
    "c18_inv_diff": "18. 或存货占比偏离行业平均大于(%)",
}

_DESKTOP_C15_COMPONENTS = (
    "短期借款,一年内到期的非流动负债,长期借款,应付债券,租赁负债,应付票据,交易性金融负债,长期应付款"
)
_DESKTOP_C16_COMPONENTS = (
    "应收账款,应收票据,其他应收款,存货,固定资产,在建工程,债权投资,其他债权投资,长期债权投资,"
    "投资性房地产,无形资产,生产性生物资产,长期应收款,合同资产,可供出售金融资产,持有至到期投资,"
    "工程物资,油气资产,商誉"
)

#: 桌面版 self.params 的默认值（**判断阈值以此为准**）
FRAUD_DEFAULTS: Dict[str, Any] = {
    "c1_emp_drop": 10.0,
    "c1_rev_inc": 10.0,
    "c2_salary_diff": 30.0,
    "c3_related_tx": 2,
    "c5_rev_growth": 15.0,
    "c5_margin": 30.0,
    "c5_inv_ratio": 15.0,
    "c7_cash_ratio": 10.0,
    "c8_ar_ratio": 15.0,
    "c9_fa_ratio": 40.0,
    "c10_cip_ratio": 5.0,
    "c11_prepay_ratio": 10.0,
    "c12_fa_cip_ratio": 45.0,
    "c13_ar_limit": 15.0,
    "c13_fa_cip_inc": 20.0,
    "c14_int_rate": 2.0,
    "c15_cash_limit": 10.0,
    "c16_impair_ratio": 0.1,
    "c17_margin_jump": 20.0,
    "c17_margin_diff": 15.0,
    "c18_ar_diff": 15.0,
    "c18_inv_diff": 15.0,
    "c15_components_str": _DESKTOP_C15_COMPONENTS,
    "c16_components_str": _DESKTOP_C16_COMPONENTS,
}

#: 配置层（core.config.FRAUD_PARAM_DEFS）键名 -> 桌面版键名（语义一一对应、且默认同值）
_PARAM_ALIASES: Dict[str, str] = {
    "c2_gap": "c2_salary_diff",                    # 30.0
    "c5_inv_growth": "c5_inv_ratio",                # 15.0
    "c11_pre_other_ratio": "c11_prepay_ratio",      # 10.0
    "c14_interest_rate": "c14_int_rate",            # 2.0
    # 同名同值，无需转换，列在此处仅为文档完整性：
    "c7_cash_ratio": "c7_cash_ratio",
    "c8_ar_ratio": "c8_ar_ratio",
    "c9_fa_ratio": "c9_fa_ratio",
    "c10_cip_ratio": "c10_cip_ratio",
    "c15_components_str": "c15_components_str",
    "c16_components_str": "c16_components_str",
    "c16_impair_ratio": "c16_impair_ratio",
    "c17_margin_diff": "c17_margin_diff",
    "c18_ar_diff": "c18_ar_diff",
    "c18_inv_diff": "c18_inv_diff",
}

#: 配置层里语义与桌面版**不一一对应**的键：忽略（否则会静默改动判断阈值）
_AMBIGUOUS_CONFIG_KEYS: Dict[str, str] = {
    "c1_ratio": "桌面版 C1 为双阈值 c1_emp_drop(10.0) 且 c1_rev_inc(10.0)，单值语义不同，已忽略",
    "c12_fa_growth": "桌面版 C12 阈值是 c12_fa_cip_ratio(45.0)，该键描述的是 C13 的增幅，已忽略",
    "c13_ar_inv_growth": "桌面版 C13 为 c13_ar_limit(15.0) 与 c13_fa_cip_inc(20.0) 双阈值，语义不明确，已忽略",
}

#: 被忽略的配置键 -> 应改用的桌面版键名（仅用于日志提示）
_AMBIGUOUS_SUGGESTION: Dict[str, str] = {
    "c1_ratio": "c1_emp_drop / c1_rev_inc",
    "c12_fa_growth": "c12_fa_cip_ratio / c13_fa_cip_inc",
    "c13_ar_inv_growth": "c13_ar_limit / c13_fa_cip_inc",
}

#: 配置层里存在但桌面版 18 条判断**从未读取**的键（C3/C4/C6 恒为「正常」）：
#: 原样保留在参数字典里以便 UI 展示/回写，不参与任何计算。
UNUSED_CONFIG_KEYS: Tuple[str, ...] = ("c3_related_ratio", "c4_invest_ratio", "c6_raise_misuse")

_STR_PARAM_KEYS = ("c15_components_str", "c16_components_str")


def normalize_params(params: Optional[Dict[str, Any]] = None,
                     log: Optional[LogFn] = None) -> Dict[str, Any]:
    """把任意来源的参数整理成桌面版 ``self.params`` 的等价字典。

    - 以桌面版默认值为底，入参覆盖之（等价于桌面版 ``{k: v["val"].get() ...}``）；
    - 兼容 ``core.config.default_fraud_params()`` 的键名（见 :data:`_PARAM_ALIASES`）；
    - 语义不明确、会改动判断阈值的配置键忽略并记警告（见 :data:`_AMBIGUOUS_CONFIG_KEYS`）；
    - 数值非法时回退默认值，保证不抛异常。
    """
    out: Dict[str, Any] = dict(FRAUD_DEFAULTS)
    # 用配置层的科目字符串作为缺省（与桌面版逐字相同，避免两处漂移）
    for _k in _STR_PARAM_KEYS:
        _cfg_val = (FRAUD_PARAM_DEFS.get(_k) or {}).get("val")
        if isinstance(_cfg_val, str) and _cfg_val.strip():
            out[_k] = _cfg_val

    if not params:
        return out

    for raw_key, raw_val in dict(params).items():
        key = str(raw_key)
        if key in _AMBIGUOUS_CONFIG_KEYS:
            if log:
                log(f"⚠️ 参数 {key} 未启用（{_AMBIGUOUS_CONFIG_KEYS[key]}）；"
                    f"如需调整请改用桌面版键名 {_AMBIGUOUS_SUGGESTION.get(key, '')}")
            continue
        if key in UNUSED_CONFIG_KEYS:
            out[key] = raw_val  # 原样保留，不参与计算
            continue
        target = _PARAM_ALIASES.get(key, key)
        if target not in FRAUD_DEFAULTS:
            if log:
                log(f"⚠️ 未知参数 {key} 已忽略（桌面版无此参数）")
            continue
        if target in _STR_PARAM_KEYS:
            out[target] = "" if raw_val is None else str(raw_val)
            continue
        num = safe_float(raw_val)
        if num is None:
            if log:
                log(f"⚠️ 参数 {key}={raw_val!r} 非法，沿用默认值 {FRAUD_DEFAULTS[target]!r}")
            continue
        default = FRAUD_DEFAULTS[target]
        out[target] = int(num) if isinstance(default, int) and not isinstance(default, bool) else float(num)
    return out


# ======================================================================
# 三、行业均值库（宏观库）。键名与取值同桌面版 self.macro_db。
# ======================================================================
_DESKTOP_AVG_WAGE: Dict[str, int] = {
    "2018": 82461, "2019": 90501, "2020": 97379, "2021": 106837,
    "2022": 114029, "2023": 120698, "2024": 124110,
}
_DESKTOP_AVG_MARGIN: Dict[str, float] = {str(y): 30.0 for y in range(2018, 2025)}
_DESKTOP_AVG_AR: Dict[str, float] = {str(y): 10.0 for y in range(2018, 2025)}
_DESKTOP_AVG_INV: Dict[str, float] = {str(y): 12.0 for y in range(2018, 2025)}


def _keep_num_type(value: Any) -> Optional[float]:
    """保留原始数值类型（int 仍是 int）——桌面版 macro_db 的工资是 int，
    若统一转成 float 会让 C2「行业均值(元)」列的 dtype 与桌面版不同（值相同、单元格类型不同）。"""
    if isinstance(value, bool):
        return None
    if isinstance(value, int):
        return int(value)
    if isinstance(value, float):
        return value
    return safe_float(value)


def build_macro_db() -> Dict[str, Dict[str, float]]:
    """构建行业均值库：工资表以 ``core.config.MACRO_AVG_WAGE`` 为准（配置优先、桌面版同值兜底）。"""
    wage: Dict[str, float] = {str(k): v for k, v in _DESKTOP_AVG_WAGE.items()}
    cfg_wage = (MACRO_AVG_WAGE or {}).get("默认_平均工资") or {}
    for k, v in dict(cfg_wage).items():
        num = _keep_num_type(v)
        if num is not None:
            wage[str(k)] = num
    return {
        "默认_平均工资": wage,
        "默认_毛利率": dict(_DESKTOP_AVG_MARGIN),
        "默认_应收占比": dict(_DESKTOP_AVG_AR),
        "默认_存货占比": dict(_DESKTOP_AVG_INV),
    }


#: 行业均值库（等价于桌面版 ``self.macro_db``；如需替换可在调用前改写本字典）
MACRO_DB: Dict[str, Dict[str, float]] = build_macro_db()


# ======================================================================
# 四、内部小工具
# ======================================================================
def _jsonify(value: Any) -> Any:
    """把 numpy / pandas 标量转成可 JSON 序列化的 Python 原生类型（不改变数值）。"""
    if value is None:
        return None
    if isinstance(value, bool):
        return bool(value)
    if isinstance(value, float):
        return float(value)
    if isinstance(value, int):
        return int(value)
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
    if isinstance(value, str):
        return value
    try:  # numpy 标量（np.int64 等）只有 item() 能转原生类型
        item = value.item()  # type: ignore[attr-defined]
        if isinstance(item, (int, float, bool, str)):
            return item
    except Exception:
        pass
    return str(value)


def _jsonify_records(records: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    return [{str(k): _jsonify(v) for k, v in rec.items()} for rec in records]


def sanitize_for_json(obj: Any) -> Any:
    """递归把非有限浮点（NaN/Inf）替换成 ``None``，供严格 JSON 编码器使用。

    背景：``extract_company_data`` 为与桌面版逐字一致，会把「年份列为空」的单元格读成 ``nan``
    （桌面版同样如此，属原有行为）。``json.dumps`` 默认可序列化（输出 ``NaN``），
    但 FastAPI ``JSONResponse``/``allow_nan=False`` 会直接抛错，此时请先包一层::

        return sanitize_for_json(extract_company_data(in_dir, params))

    ``analyze_one_company`` / ``run_all`` 的返回值本身已不含 NaN（records 里 NaN 会转成 ""）。
    """
    if isinstance(obj, dict):
        return {k: sanitize_for_json(v) for k, v in obj.items()}
    if isinstance(obj, (list, tuple)):
        return [sanitize_for_json(v) for v in obj]
    if isinstance(obj, float):
        try:
            if obj != obj or obj in (float("inf"), float("-inf")):
                return None
        except Exception:  # noqa: BLE001
            return None
        return obj
    return obj


# ----------------------------------------------------------------------
# 输入容错（新增，桌面版没有；只影响「能否取到数据」，不影响任何判断规则）
# 1)《统一整合输出》两种版式：
#    A) 新版：第 1 行即表头，header=0 后列名已是 ['项目','20221231','20241231',...]
#    B) 旧版：第 1 行是合并标题，第 2 行才是表头（项目列名为 'xxx - 合并资产负债表'）
#    对 A 直接采用表头定年份列（否则会把第一条科目里的金额误判成年份列）；
#    对 B 仍走桌面版原逻辑（扫前 15 行找 YYYYMMDD）。
# 2)《分红情况》两种形态（形态1 按年汇总；形态2 东财逐笔 → 按报告期年份聚合）。
# ----------------------------------------------------------------------
_YEAR_COL_RE = re.compile(r'^(20\d{2})(1231|12-31)?$')

#: 分红指标键名（新增，桌面版无分红指标；因不在 FRAUD_METRIC_MAP 里，
#: 不会参与「指标补 0」，也不会被 18 条判断读取）
DIV_METRIC_RATE = "分红率"
DIV_METRIC_CASH = "现金分红总额(元)"
DIV_METRIC_PER10 = "每10股派息合计(元)"
#: 分红率取值来源为百分数字符串（如 '70.58%'），保留百分数口径
_DIV_YEAR_LABELS = ("报告期", "年份", "年度")
_DIV_RATE_LABELS = ("分红率",)
_DIV_CASH_LABELS = ("现金分红总额", "分红总额")
_DIV_PER10_LABELS = ("每10股",)


def _cell_label(value: Any) -> str:
    """表头单元格 -> 干净标签（去掉零宽字符/换行/首尾空白）。"""
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
    return str(value).replace('\u200b', '').replace('\n', '').replace('\r', '').strip()


def _detect_header_year_cols(df: pd.DataFrame) -> Tuple[Dict[int, str], Optional[int]]:
    """判断是不是「新版（header=0 即表头）」版式。

    返回 ``(year_cols, key_col_idx)``：
    - ``year_cols``：列位置 -> 年份（列名匹配 ``^(20\\d{2})(1231|12-31)?$``，兼容数值列名 20221231/20221231.0）；
    - ``key_col_idx``：科目列（列名含「项目」）的位置，None 表示没有。

    两者皆空表示是旧版，调用方应回退到桌面版「扫前 15 行」的逻辑。
    """
    year_cols: Dict[int, str] = {}
    key_col_idx: Optional[int] = None
    for c_idx, col in enumerate(df.columns):
        name = _cell_label(col).replace('.0', '')
        if not name or name.startswith("Unnamed"):
            continue
        if key_col_idx is None and "项目" in name:
            key_col_idx = c_idx
            continue
        m = _YEAR_COL_RE.match(name)
        if m:
            year_cols[c_idx] = m.group(1)
    return year_cols, key_col_idx


def _scan_dividend_header(df: pd.DataFrame, max_rows: int = 6
                          ) -> Tuple[Optional[int], Optional[List[str]]]:
    """定位《分红情况》表头。

    返回 ``(数据起始行, 列标签)``；``起始行 = -1`` 表示表头就在列名上（新版 header=0）。
    找不到「年份/报告期 + 分红相关列」则返回 ``(None, None)``。
    """
    candidates: List[Tuple[int, List[str]]] = [(-1, [_cell_label(c) for c in df.columns])]
    for i in range(min(max_rows, len(df))):
        try:
            candidates.append((i, [_cell_label(v) for v in df.iloc[i].tolist()]))
        except Exception:  # noqa: BLE001
            break
    for start, labels in candidates:
        has_year = any(any(k in lb for k in _DIV_YEAR_LABELS) for lb in labels)
        has_div = any(any(k in lb for k in _DIV_RATE_LABELS + _DIV_CASH_LABELS + _DIV_PER10_LABELS
                          + ("现金分红", "派息")) for lb in labels)
        if has_year and has_div:
            return start, labels
    return None, None


def _pick_year(value: Any) -> str:
    """从 '2024-12-31' / 20241231 / datetime / 2024 里取 4 位年份。"""
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
    if hasattr(value, "year"):  # datetime / Timestamp
        try:
            return str(int(value.year))
        except Exception:
            pass
    m = re.search(r'(20\d{2})', re.sub(r'\.0+$', '', str(value).strip()))
    return m.group(1) if m else ""


def _collect_dividends(df: pd.DataFrame, sheet_name: str = "",
                       log: Optional[Callable[[str], None]] = None) -> Dict[str, Dict[str, float]]:
    """解析《分红情况》，兼容两种形态并按「报告期年份」聚合成年度值。

    - 形态1（按年汇总）：``股票代码/年份/同花顺现金分红总额(元)/归属于母公司所有者的净利润(元)/分红率``
      → ``分红率``（百分数，如 70.58）、``现金分红总额(元)``；
    - 形态2（东财逐笔）：``报告期/现金分红-现金分红比例(每10股派息元)/除权除息日/方案进度``
      → 按报告期年份累加为 ``每10股派息合计(元)``；无 ``分红率`` 列时不写该键（留空），不报错。

    返回 ``{年份: {指标: 数值}}``；任何异常都只记日志并返回已解析部分。
    """
    out: Dict[str, Dict[str, float]] = {}
    try:
        start, labels = _scan_dividend_header(df)
        if start is None or labels is None:
            return out
        data = df.iloc[start + 1:] if start >= 0 else df
        year_idx = None
        for want in _DIV_YEAR_LABELS:
            for i, lb in enumerate(labels):
                if want in lb:
                    year_idx = i
                    break
            if year_idx is not None:
                break
        if year_idx is None:
            return out

        def _positions(keys: Tuple[str, ...]) -> List[int]:
            return [i for i, lb in enumerate(labels) if lb and any(k in lb for k in keys)]

        rate_idx = _positions(_DIV_RATE_LABELS)
        cash_idx = _positions(_DIV_CASH_LABELS)
        per10_idx = _positions(_DIV_PER10_LABELS)

        for _, row in data.iterrows():
            vals = row.tolist()
            if year_idx >= len(vals):
                continue
            year = _pick_year(vals[year_idx])
            if not year:
                continue
            bucket = out.setdefault(year, {})
            for i in rate_idx:
                if i < len(vals):
                    num = safe_float(vals[i])
                    if num is not None and DIV_METRIC_RATE not in bucket:
                        bucket[DIV_METRIC_RATE] = num
            for metric, idxs in ((DIV_METRIC_CASH, cash_idx), (DIV_METRIC_PER10, per10_idx)):
                total = 0.0
                found = False
                for i in idxs:
                    if i < len(vals):
                        num = safe_float(vals[i])
                        if num is not None:
                            total += num
                            found = True
                if found:
                    bucket[metric] = bucket.get(metric, 0.0) + total
    except Exception as e:  # noqa: BLE001
        if log:
            log(f"  ⚠️ 分红情况解析失败，已跳过: {sheet_name} ({type(e).__name__}: {str(e)[:100]})")
    return out


def _merge_dividend(into: Dict[str, Dict[str, float]], new: Dict[str, Dict[str, float]]) -> None:
    """合并多个来源的分红数据：同一(年,指标)取绝对值较大者。

    同一份分红数据常常同时出现在《统一整合输出》与《年度年报提取表》里，
    取较大值可避免「同一数据被多文件重复累加」，且与遍历顺序无关。
    """
    for year, metrics in new.items():
        bucket = into.setdefault(year, {})
        for k, v in metrics.items():
            old = bucket.get(k)
            if old is None or abs(v) > abs(old):
                bucket[k] = v


# ======================================================================
# 五、数据抽取（对应桌面版 extract_excel_data）
# ======================================================================
def extract_company_data(input_dir: str, params: Dict[str, Any],
                         log: Optional[Callable[[str], None]] = None) -> Dict[str, Any]:
    """递归扫描 ``input_dir`` 下的《统一整合输出_*.xlsx》/《*年报提取表.xlsx》/《公司属性表》，
    按年份汇总指标。

    返回结构与桌面版 ``extract_excel_data`` 完全一致::

        { "600887 伊利股份": {"industry": "饮料乳品",
                              "2021": {"营业收入": 1.2e11, "总资产": ..., "员工人数": ...},
                              "2022": {...}}, ...}

    说明（与桌面版逐条一致）：
    - 「公司属性表」用 ``目标股票`` / ``二级行业``(缺省 ``一级行业``，去掉 Ⅰ/Ⅱ) 建 代码->行业 映射；
    - 文件名按 ``r'(.*?)[(（](\\d{6})[)）]'`` 取代码与简称，取不到则代码 ``000000``、简称取文件名；
    - 逐 sheet 识别年份列（``^(20\\d{2})(1231|12-31)?$``，只扫前 15 行、满 2 列即停）；
    - 科目名先精确匹配、再「包含且长度受限」匹配；同名指标仅在该年份缺失或为 0 时写入（先到先得）；
    - 员工人数取「数量合计/员工总数/合计数/合计/员工人数」行中最后一个 >10 的数字；
    - 收尾补齐全部指标为 0.0、算 期初应付薪酬 / 自身毛利率 / 减值损失合计。

    输入容错（新增）：
    - 三大表兼容「新版 header=0 表头」与「旧版第 2 行表头」两种版式（见 :func:`_detect_header_year_cols`）；
    - 《分红情况》兼容「按年汇总」与「东财逐笔」两种形态，按报告期年份聚合（见 :func:`_collect_dividends`），
      结果写入对应年度的 ``分红率`` / ``现金分红总额(元)`` / ``每10股派息合计(元)`` 键，
      只并入三大表已存在的年度，不会新增年度、不影响 18 条判断。

    容错：单个文件解析失败只记日志，不中断整批。
    """
    log_fn = log or (lambda _m: None)
    all_company_data: Dict[str, Any] = {}
    company_industry_map: Dict[str, str] = {}
    log_fn(f"正在递归扫描提取多源数据: {input_dir}")

    # ---------- 第一遍：公司属性表 ----------
    for root_dir, _, files in os.walk(input_dir):
        for filename in files:
            if "公司属性表" in filename and not filename.startswith("~"):
                filepath = os.path.join(root_dir, filename)
                try:
                    df_attr = pd.read_excel(filepath) if filename.endswith(".xlsx") else pd.read_csv(filepath,
                                                                                                     encoding='gbk')
                    for _, row in df_attr.iterrows():
                        code = str(row.get("目标股票", "")).strip()
                        if code:
                            ind = str(row.get("二级行业", str(row.get("一级行业", "")))).replace("Ⅱ", "").replace(
                                "Ⅰ", "")
                            company_industry_map[code] = ind
                except Exception as e:  # noqa: BLE001  桌面版为静默 except: pass
                    log_fn(f"  ⚠️ 公司属性表读取失败，已跳过: {filename} ({type(e).__name__}: {str(e)[:100]})")

    # ---------- 指标映射表（来自 core/contracts.py，与桌面版 mapping 逐字一致） ----------
    mapping: Dict[str, List[str]] = {k: list(v) for k, v in FRAUD_METRIC_MAP.items()}

    c15_keys = [k.strip() for k in params.get("c15_components_str", "").split(",") if k.strip()]
    c16_keys = [k.strip() for k in params.get("c16_components_str", "").split(",") if k.strip()]
    for k in c15_keys + c16_keys:
        if k not in mapping:
            mapping[k] = [k]

    # ---------- 第二遍：财报文件 ----------
    dividends_buffer: Dict[str, Dict[str, Dict[str, float]]] = {}   # 新增：分红情况(两种形态)缓存
    for root_dir, _, files in os.walk(input_dir):
        for filename in files:
            if filename.startswith("~") or not (
                    filename.endswith(".xlsx") or filename.endswith(".csv")):
                continue
            if "公司属性表" in filename:
                continue

            filepath = os.path.join(root_dir, filename)
            match = re.search(r'(.*?)[(（](\d{6})[)）]', filename)
            if match:
                comp_name = match.group(1).split('_')[-1].strip()
                comp_code = match.group(2)
            else:
                comp_code = "000000"
                comp_name = filename.replace(".xlsx", "").replace(".csv", "").split("-")[0].strip()

            comp_key = f"{comp_code} {comp_name}"
            if comp_key not in all_company_data:
                all_company_data[comp_key] = {"industry": company_industry_map.get(comp_code, "默认")}
            comp_data = all_company_data[comp_key]

            try:
                if filename.endswith(".csv"):
                    try:
                        dfs = {"Sheet1": pd.read_csv(filepath, encoding='gbk')}
                    except Exception:  # noqa: BLE001
                        dfs = {"Sheet1": pd.read_csv(filepath, encoding='utf-8')}
                else:
                    xls = pd.ExcelFile(filepath)
                    dfs = {sheet: pd.read_excel(xls, sheet_name=sheet) for sheet in xls.sheet_names}

                for sheet_name, df in dfs.items():
                    current_parsed_year = None
                    for r_idx, row in df.iterrows():
                        row_strs = [str(x).replace(',', '').strip() for x in row.values if pd.notna(x)]
                        for cell in row_strs[:3]:
                            # 新版表头版式下「年份」列若含空值会被 pandas 读成 float（2022.0），
                            # 故仅在做年份判定时容忍尾部 .0（不影响下面的数值解析，旧版结果不变）
                            if re.match(r'^20\d{2}$', re.sub(r'\.0+$', '', cell)):
                                current_parsed_year = re.sub(r'\.0+$', '', cell)
                                break
                        joined_row = "".join(row_strs)
                        if any(k in joined_row for k in ["数量合计", "员工总数", "合计数", "合计", "员工人数"]):
                            for val in reversed(row_strs):
                                if re.match(r'^\d+(\.\d+)?$', val):
                                    num = float(val)
                                    if num > 10 and current_parsed_year:
                                        if current_parsed_year not in comp_data:
                                            comp_data[current_parsed_year] = {}
                                        comp_data[current_parsed_year]["员工人数"] = num
                                        break

                    year_cols: Dict[int, str] = {}
                    metric_idx = 0
                    hdr_year_cols, hdr_key_idx = _detect_header_year_cols(df)
                    if hdr_year_cols or hdr_key_idx is not None:
                        # 新版版式：列名已含「项目」或 YYYYMMDD 年份列，直接采用表头，不再扫前 15 行
                        year_cols = hdr_year_cols
                        if hdr_key_idx is not None:
                            metric_idx = hdr_key_idx
                    else:
                        # 旧版版式：与桌面版逐字一致（扫前 15 行，满 2 个年份列即停）
                        for r_idx, row in df.head(15).iterrows():
                            for c_idx, val in enumerate(row):
                                val_str = str(val).replace('.0', '').strip()
                                match_yr = re.search(r'^(20\d{2})(1231|12-31)?$', val_str)
                                if match_yr:
                                    year_cols[c_idx] = match_yr.group(1)
                            if len(year_cols) >= 2:
                                break

                    for _, row in df.iterrows():
                        metric_cell = row.iloc[metric_idx] if metric_idx < len(row) else None
                        metric_name = str(metric_cell).replace('\u200b', '').strip() if pd.notna(
                            metric_cell) else ""
                        if not metric_name and len(row) > 1:
                            metric_name = str(row.iloc[1]).replace('\u200b', '').strip() if pd.notna(
                                row.iloc[1]) else ""
                        target_key = None
                        for key, keywords in mapping.items():
                            if metric_name in keywords:
                                target_key = key
                                break
                        if not target_key:
                            for key, keywords in mapping.items():
                                if any(kw in metric_name and len(metric_name) < len(kw) + 6 for kw in
                                       keywords):
                                    target_key = key
                                    break
                        if target_key and year_cols:
                            for c_idx, yr in year_cols.items():
                                if yr not in comp_data:
                                    comp_data[yr] = {}
                                if c_idx < len(row):
                                    val_str = str(row.iloc[c_idx]).replace(',', '').strip()
                                    try:
                                        val_float = float(val_str)
                                        if target_key in ["资产减值损失", "信用减值损失", "存货", "固定资产"]:
                                            val_float = abs(val_float)
                                        if target_key not in comp_data[yr] or comp_data[yr][target_key] == 0:
                                            comp_data[yr][target_key] = val_float
                                    except ValueError:
                                        pass

                    # 新增：《分红情况》两种形态 -> 按报告期年份聚合（形态1 分红率/现金分红总额，形态2 每10股派息合计）
                    div_records = _collect_dividends(df, sheet_name, log_fn)
                    if div_records:
                        _merge_dividend(dividends_buffer.setdefault(comp_key, {}), div_records)
            except Exception as e:  # noqa: BLE001  桌面版为静默 except: pass
                log_fn(f"  ⚠️ 文件解析失败，已跳过: {filename} ({type(e).__name__}: {str(e)[:100]})")
                continue

    # ---------- 收尾：补齐指标、期初薪酬、毛利率、减值合计 ----------
    for ck, cdata in all_company_data.items():
        years = sorted([k for k in cdata.keys() if k != "industry"])
        for i, yr in enumerate(years):
            metrics = cdata[yr]
            for rm in list(mapping.keys()) + ["员工人数"]:
                if rm not in metrics:
                    metrics[rm] = 0.0
            metrics["期初应付薪酬"] = cdata[years[i - 1]].get("期末应付薪酬", 0.0) if i > 0 else metrics[
                "期末应付薪酬"]
            if metrics["营业收入"] > 0:
                metrics["自身毛利率"] = (metrics["营业收入"] - metrics["营业成本"]) / metrics[
                    "营业收入"] * 100.0
            else:
                metrics["自身毛利率"] = 0.0
            metrics["减值损失合计"] = metrics["资产减值损失"] + metrics["信用减值损失"]
        # 新增：分红指标只并入「三大表已存在的年度」，绝不新增年份（保证年份集合与桌面版一致，
        # 从而 18 条判断的输入年度、目标年度完全不变）；没有分红数据的年度该键不存在 = 留空。
        div_for_comp = dividends_buffer.get(ck) or {}
        for div_year, div_metrics in div_for_comp.items():
            if div_year in cdata:
                cdata[div_year].update(div_metrics)
            else:
                log_fn(f"  ℹ️ {ck} 分红数据年度 {div_year} 不在财报年度内，已忽略（不新增年度）")
    return all_company_data


# ======================================================================
# 六、18 条判断（对应桌面版 execute_18_conditions_analysis，逐条照搬）
# ======================================================================
def execute_18_conditions_analysis(comp_code: str, comp_name: str, comp_industry: str,
                                   comp_data: Dict[str, Any], all_years: List[str],
                                   target_years: List[str], p: Dict[str, Any],
                                   log_callback: Optional[Callable[[str], None]] = None
                                   ) -> Dict[str, pd.DataFrame]:
    """跑 18 条造假嫌疑判断，返回 ``{条件标题: DataFrame}``（与桌面版返回结构一致）。

    ``comp_industry`` / ``log_callback`` 仅为保持签名兼容而保留（桌面版未在判断中使用行业名）。
    """
    yearly_results = {k: [] for k in CONDITION_ORDER}

    def safe_div2(num, den): return num / den * 100 if den and den != 0 else 0.0

    for curr_yr in target_years:
        prev_idx = all_years.index(curr_yr) - 1
        prev_yr = all_years[prev_idx]
        c, p_prev = comp_data[curr_yr], comp_data[prev_yr]

        # ------------------------------------------------------------------
        # 【云端版防护】C1 / C2 依赖“员工人数”，而免费公开数据源**没有员工人数**
        # （桌面版来自年报 PDF 的“员工情况”章节；akshare 无对应可用接口）。
        # 若缺该数据仍按 0 计算，会算出 “员工降幅=100%、营收增幅=1.1e13%” 这类垃圾值
        # 并误报“触发预警”。因此这里如实标注“数据不足”，不参与风险判定。
        # 有员工人数时（例如你自行补充了数据）行为与桌面版完全一致。
        _c_emp, _p_emp = c.get("员工人数"), p_prev.get("员工人数")
        _has_emp = bool(_c_emp) and bool(_p_emp)
        # rev_inc 在 C1 与 C5 都要用，必须无条件先算出来（否则缺员工人数时 C5 会 UnboundLocalError）
        rev_inc = safe_div2(c.get("营业收入", 0) - p_prev.get("营业收入", 0),
                            max(p_prev.get("营业收入", 1), 1))
        if not _has_emp:
            yearly_results["C1_员工与营收背离"].append(
                {"年度": curr_yr, "状态": "数据不足", "员工降幅(%)": None, "营收增幅(%)": rev_inc,
                 "[说明]": "缺少员工人数（免费数据源不提供），本条无法判定"})
        else:
            emp_drop = safe_div2(_p_emp - _c_emp, max(_p_emp, 1))
            yearly_results["C1_员工与营收背离"].append({"年度": curr_yr,
                                                        "状态": "风险" if emp_drop > p["c1_emp_drop"] and rev_inc >
                                                                          p["c1_rev_inc"] else "正常",
                                                        "员工降幅(%)": emp_drop, "营收增幅(%)": rev_inc,
                                                        "[说明]": "员工降幅与营收增幅背离"})

        if not _has_emp:
            yearly_results["C2_人均薪酬异常"].append(
                {"年度": curr_yr, "状态": "数据不足", "自身人均(元)": None,
                 "行业均值(元)": MACRO_DB["默认_平均工资"].get(str(curr_yr), 100000),
                 "[说明]": "缺少员工人数（免费数据源不提供），本条无法判定"})
        else:
            comp_salary = (c.get("支付职工现金", 0) + c.get("期末应付薪酬", 0) - c.get("期初应付薪酬", 0)) / max(
                (_p_emp + _c_emp) / 2, 1)
            ind_salary = MACRO_DB["默认_平均工资"].get(str(curr_yr), 100000)
            sal_diff = safe_div2(ind_salary - comp_salary, max(ind_salary, 1))
            yearly_results["C2_人均薪酬异常"].append(
                {"年度": curr_yr, "状态": "风险" if sal_diff > p["c2_salary_diff"] else "正常",
                 "自身人均(元)": comp_salary, "行业均值(元)": ind_salary, "[说明]": "低于行业平均"})

        yearly_results["C3_关联交易"].append(
            {"年度": curr_yr, "状态": "正常", "关联交易笔数": 0, "[说明]": "无重大异常"})
        yearly_results["C4_重大股权投资"].append(
            {"年度": curr_yr, "状态": "正常", "投资情况": "无", "[说明]": "未见异常"})

        inv_ratio = safe_div2(c.get("存货", 0), max(c.get("总资产", 1), 1))
        c5_risk = (rev_inc < p["c5_rev_growth"] or c.get("自身毛利率", 0) < p["c5_margin"] or inv_ratio > p[
            "c5_inv_ratio"])
        yearly_results["C5_主营疲软投资异动"].append(
            {"年度": curr_yr, "状态": "风险" if c5_risk else "正常", "营收增幅(%)": rev_inc,
             "毛利率(%)": c.get("自身毛利率", 0), "存货占比(%)": inv_ratio, "[说明]": "主营衰退"})
        yearly_results["C6_募资变更"].append(
            {"年度": curr_yr, "状态": "正常", "发生用途变更": "否", "[说明]": "正常"})

        cash_ratio = safe_div2(c.get("货币资金", 0) + c.get("交易性金融资产", 0), max(c.get("总资产", 1), 1))
        yearly_results["C7_高存差钱"].append(
            {"年度": curr_yr, "状态": "风险" if cash_ratio < p["c7_cash_ratio"] else "正常",
             "资金占比(%)": cash_ratio, "[说明]": "资金紧张"})
        ar_ratio = safe_div2(c.get("应收账款", 0), max(c.get("总资产", 1), 1))
        yearly_results["C8_应收账款预警"].append(
            {"年度": curr_yr, "状态": "风险" if ar_ratio > p["c8_ar_ratio"] else "正常",
             "应收占比(%)": ar_ratio, "[说明]": "应收过高"})
        fa_ratio = safe_div2(c.get("固定资产", 0), max(c.get("总资产", 1), 1))
        yearly_results["C9_固定资产过高"].append(
            {"年度": curr_yr, "状态": "风险" if fa_ratio > p["c9_fa_ratio"] else "正常",
             "固资占比(%)": fa_ratio, "[说明]": "重资产风险"})
        cip_ratio = safe_div2(c.get("在建工程", 0), max(c.get("总资产", 1), 1))
        yearly_results["C10_在建工程过高"].append(
            {"年度": curr_yr, "状态": "风险" if cip_ratio > p["c10_cip_ratio"] else "正常",
             "在建占比(%)": cip_ratio, "[说明]": "在建停滞嫌疑"})

        prepay_ratio = safe_div2(c.get("预付款项", 0) + c.get("其他应收款", 0), max(c.get("总资产", 1), 1))
        yearly_results["C11_预付及其他应收"].append(
            {"年度": curr_yr, "状态": "风险" if prepay_ratio > p["c11_prepay_ratio"] else "正常",
             "占比(%)": prepay_ratio, "[说明]": "占用资金"})
        fa_cip_ratio = fa_ratio + cip_ratio
        yearly_results["C12_固资在建合计"].append(
            {"年度": curr_yr, "状态": "风险" if fa_cip_ratio > p["c12_fa_cip_ratio"] else "正常",
             "合计占比(%)": fa_cip_ratio, "[说明]": "固资在建超标"})

        fa_cip_growth = safe_div2((c.get("固定资产", 0) + c.get("在建工程", 0)) - (
                p_prev.get("固定资产", 0) + p_prev.get("在建工程", 0)),
                                  max(p_prev.get("固定资产", 0) + p_prev.get("在建工程", 0), 1))
        c13_risk = (ar_ratio > 5 or inv_ratio > p["c13_ar_limit"]) and fa_cip_growth > p["c13_fa_cip_inc"]
        yearly_results["C13_双高联动风险"].append(
            {"年度": curr_yr, "状态": "风险" if c13_risk else "正常", "固建增幅(%)": fa_cip_growth,
             "应收占比(%)": ar_ratio, "存货占比(%)": inv_ratio, "[说明]": "双高联动"})
        int_rate = safe_div2(c.get("利息收入", 0), max(c.get("货币资金", 1), 1))
        yearly_results["C14_存单质押隐患"].append(
            {"年度": curr_yr, "状态": "风险" if int_rate < p["c14_int_rate"] else "正常",
             "资金收益率(%)": int_rate, "[说明]": "存单可能受限"})

        c15_keys = [k.strip() for k in p.get("c15_components_str", "").split(",") if k.strip()]
        debt_sum = sum(c.get(k, 0) for k in c15_keys)
        debt_ratio = safe_div2(debt_sum, max(c.get("总资产", 1), 1))
        c15_risk = (safe_div2(c.get("货币资金", 0), max(c.get("总资产", 1), 1)) > 10) and debt_ratio > p[
            "c15_cash_limit"]
        yearly_results["C15_高存高贷双雷"].append({"年度": curr_yr, "状态": "风险" if c15_risk else "正常",
                                                   "资金占比(%)": safe_div2(c.get("货币资金", 0),
                                                                            max(c.get("总资产", 1), 1)),
                                                   "借款占比(%)": debt_ratio, "[说明]": "存贷双高"})

        c16_keys = [k.strip() for k in p.get("c16_components_str", "").split(",") if k.strip()]
        assets_sum = sum(c.get(k, 0) for k in c16_keys)
        impair_ratio = safe_div2(c.get("减值损失合计", 0), max(assets_sum, 1))
        yearly_results["C16_减值计提异常"].append(
            {"年度": curr_yr, "状态": "风险" if impair_ratio < p["c16_impair_ratio"] else "正常",
             "减值比例(%)": impair_ratio, "[说明]": "异常计提"})

        margin_growth = safe_div2(c.get("自身毛利率", 0) - p_prev.get("自身毛利率", 0),
                                  max(p_prev.get("自身毛利率", 1), 1))
        ind_margin = MACRO_DB["默认_毛利率"].get(str(curr_yr), 30.0)
        margin_diff = safe_div2(c.get("自身毛利率", 0) - ind_margin, max(ind_margin, 1))
        yearly_results["C17_毛利率异常偏离"].append({"年度": curr_yr, "状态": "风险" if margin_growth > p[
            "c17_margin_jump"] and margin_diff > p["c17_margin_diff"] else "正常", "自身增幅(%)": margin_growth,
                                                     "偏离均值(%)": margin_diff, "[说明]": "毛利造假嫌疑"})

        ind_ar, ind_inv = MACRO_DB["默认_应收占比"].get(str(curr_yr), 10.0), MACRO_DB[
            "默认_存货占比"].get(str(curr_yr), 12.0)
        ar_diff_ind, inv_diff_ind = safe_div2(ar_ratio - ind_ar, max(ind_ar, 1)), safe_div2(inv_ratio - ind_inv,
                                                                                            max(ind_inv, 1))
        c18_risk = margin_growth > p["c17_margin_jump"] and (
                ar_diff_ind > p["c18_ar_diff"] or inv_diff_ind > p["c18_inv_diff"])
        yearly_results["C18_营运资产背离"].append(
            {"年度": curr_yr, "状态": "风险" if c18_risk else "正常", "应收偏离度(%)": ar_diff_ind,
             "存货偏离度(%)": inv_diff_ind, "[说明]": "营运资金异常"})

    return {k: pd.DataFrame(v) for k, v in yearly_results.items() if len(v) > 0}


# ======================================================================
# 七、单公司分析（对应桌面版 process_task 的逐公司部分）
# ======================================================================
def analyze_one_company(comp_key: str, comp_data: Dict[str, Any], params: Dict[str, Any],
                        log: Optional[Callable[[str], None]] = None
                        ) -> Dict[str, Any]:
    """对一家公司跑 18 条判断。

    返回（可 JSON 序列化）::

        {"comp_key": "600887 伊利股份", "code": "600887", "name": "伊利股份",
         "industry": "饮料乳品", "years": ["2022", "2023"], "all_years": ["2021", "2022", "2023"],
         "results": {条件标题: {年份: "风险"/"正常"}, ...},          # 概览
         "results_rows": {条件标题: [列名->值 的记录, ...]},          # 与桌面版 DataFrame 行列一致
         "results_columns": {条件标题: [列名, ...]},
         "suspect_years": ["2023"], "risk_counts": {条件: 次数}, "total_risk_count": 1,
         "summary": "..."}

    年度规则与桌面版 ``process_task`` 一致：``target_years = all_years[1:]``（首年只用于算期初值）；
    可用年度不足 2 年时不判断（桌面版直接 ``continue`` 跳过），此处返回空结果并说明原因。
    """
    log_fn = log or (lambda _m: None)
    p = normalize_params(params, log_fn)

    comp_data = comp_data if isinstance(comp_data, dict) else {}
    comp_industry = comp_data.get("industry", "默认")
    code, name = (str(comp_key).split(' ', 1) + [""])[:2]
    all_years = sorted([k for k in comp_data.keys() if k != "industry"])

    results: Dict[str, Dict[str, Any]] = {}
    results_rows: Dict[str, List[Dict[str, Any]]] = {}
    results_columns: Dict[str, List[str]] = {}
    risk_counts: Dict[str, int] = {}
    suspect_years: List[str] = []

    if len(all_years) < 2:
        log_fn(f"  ⏭️ {comp_key} 可用年度不足 2 年（{all_years}），按桌面版规则跳过 18 条判断")
        summary = f"可用年度不足 2 年（{len(all_years)} 年），未执行 18 条判断（与桌面版一致跳过）"
        return {"comp_key": str(comp_key), "code": code, "name": name, "industry": comp_industry,
                "years": [], "all_years": all_years, "results": results, "results_rows": results_rows,
                "results_columns": results_columns, "suspect_years": suspect_years,
                "risk_counts": risk_counts, "total_risk_count": 0, "summary": summary}

    target_years = all_years[1:]
    frames = execute_18_conditions_analysis(code, name, comp_industry, comp_data, all_years, target_years, p,
                                            log_fn)

    ordered = [c for c in CONDITION_ORDER if c in frames] + [c for c in frames if c not in CONDITION_ORDER]
    hit_years: List[str] = []
    triggered: List[str] = []
    total_risk_count = 0
    for cond in ordered:
        df = frames[cond]
        cols = [str(c) for c in df.columns]
        results_columns[cond] = cols
        results_rows[cond] = _jsonify_records(df_to_records(df))
        year_map: Dict[str, Any] = {}
        if "年度" in df.columns and "状态" in df.columns:
            for yr, st in zip(df["年度"].tolist(), df["状态"].tolist()):
                year_map[str(yr)] = _jsonify(st)
        results[cond] = year_map
        cnt = int((df["状态"] == "风险").sum()) if "状态" in df.columns else 0
        risk_counts[cond] = cnt
        total_risk_count += cnt
        yrs = [str(y) for y, st in year_map.items() if st == "风险"]
        if yrs:
            triggered.append(f"{cond}({'、'.join(yrs)})")
            hit_years.extend(yrs)
    suspect_years = sorted(set(hit_years))

    year_span = f"{target_years[0]}~{target_years[-1]}" if target_years else "无"
    if triggered:
        summary = (f"分析年度 {year_span}（共 {len(target_years)} 年）：{len(triggered)} 项条件触发、"
                   f"合计 {total_risk_count} 次年度预警 —— " + "；".join(triggered))
    else:
        summary = f"分析年度 {year_span}（共 {len(target_years)} 年）：18 项条件均未触发风险预警"
    log_fn(f"  ✅ {comp_key} 判断完成：{len(triggered)} 项条件触发、共 {total_risk_count} 次预警")

    return {"comp_key": str(comp_key), "code": code, "name": name, "industry": comp_industry,
            "years": target_years, "all_years": all_years, "results": results,
            "results_rows": results_rows, "results_columns": results_columns,
            "suspect_years": suspect_years, "risk_counts": risk_counts,
            "total_risk_count": total_risk_count, "summary": summary}


# ======================================================================
# 八、报告输出（对应桌面版 generate_excel_report，版式逐字保留）
# ======================================================================
def _coerce_frames(results_dict: Any) -> Dict[str, pd.DataFrame]:
    """把各种形态的条件结果统一成 {条件: DataFrame}（DataFrame 原样通过）。"""
    out: Dict[str, pd.DataFrame] = {}
    if not isinstance(results_dict, dict):
        return out
    for cond, val in results_dict.items():
        if isinstance(val, pd.DataFrame):
            out[str(cond)] = val
        elif isinstance(val, list):
            out[str(cond)] = pd.DataFrame(val) if val else pd.DataFrame()
        elif isinstance(val, dict):
            out[str(cond)] = pd.DataFrame([{"年度": k, "状态": v} for k, v in val.items()])
    return out


def generate_excel_report(comp_code: str, comp_name: str, results_dict: Dict[str, pd.DataFrame],
                          output_dir: str, log: Optional[Callable[[str], None]] = None) -> str:
    """生成《{代码}_{简称}_深度排雷报告.xlsx》，返回文件路径（版式与桌面版一致）。

    ``results_dict`` 为 ``{条件标题: DataFrame}``（即 execute_18_conditions_analysis 的返回）。
    """
    log_fn = log or (lambda _m: None)
    results_dict = _coerce_frames(results_dict)
    ensure_dir(output_dir)

    safe_name = "".join([c for c in comp_name if c.isalpha() or c.isdigit() or c in ['-']]).strip()
    safe_code = "".join([c for c in comp_code if c.isalnum()])
    if not safe_name:
        safe_name = "未命名"
    save_path = os.path.join(output_dir, f"{safe_code}_{safe_name}_深度排雷报告.xlsx")
    wb = Workbook()
    wb.remove(wb.active)
    danger_fill = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")
    danger_font = Font(color="FF0000", bold=True)
    safe_font = Font(color="008000")
    title_font = Font(size=13, bold=True)
    title_align = Alignment(horizontal="center", vertical="center")

    for condition, df in results_dict.items():
        sheet_name = condition[:31].replace("/", "_")
        ws = wb.create_sheet(title=sheet_name)
        headers = list(df.columns)
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max(len(headers), 4))
        title_cell = ws.cell(row=1, column=1)
        title_cell.value = f"【代码: {comp_code}】 {comp_name} - {condition} 分析明细表"
        title_cell.font = title_font
        title_cell.alignment = title_align
        ws.append(headers)
        for c_idx in range(1, len(headers) + 1):
            ws.cell(row=2, column=c_idx).font = Font(bold=True)
        status_col_idx = headers.index("状态") if "状态" in headers else -1
        for r_idx, row in enumerate(df.values, start=3):
            ws.append(row.tolist())
            status = row[status_col_idx] if status_col_idx != -1 else ""
            if status == "风险":
                for c_idx in range(1, len(row) + 1):
                    ws.cell(row=r_idx, column=c_idx).font = danger_font
                    ws.cell(row=r_idx, column=c_idx).fill = danger_fill
            elif status == "正常":
                ws.cell(row=r_idx, column=status_col_idx + 1).font = safe_font

    ws_summary = wb.create_sheet(title="0_综合排雷总表", index=0)
    ws_summary.merge_cells("A1:D1")
    sum_title = ws_summary.cell(row=1, column=1)
    sum_title.value = f"【{comp_code}】 {comp_name} - 财务造假排雷综合报告"
    sum_title.font = title_font
    sum_title.alignment = title_align
    ws_summary.append(["预警维度", "历年综合诊断", "AI文本交叉验证", "审计复核建议"])

    total_risk_count = 0
    for condition, df in results_dict.items():
        if "状态" in df.columns:
            risk_count_this_condition = (df["状态"] == "风险").sum()
            total_risk_count += risk_count_this_condition
            has_risk = risk_count_this_condition > 0
            status_text = "触发预警" if has_risk else "安全"
            ws_summary.append([condition, status_text, "见明细页", "人工复核"])
            if has_risk:
                ws_summary.cell(row=ws_summary.max_row, column=2).font = danger_font
                ws_summary.cell(row=ws_summary.max_row, column=2).fill = danger_fill
            else:
                ws_summary.cell(row=ws_summary.max_row, column=2).font = safe_font

    ws_summary.append(["总计", f"共触发 {total_risk_count} 次风险", "-", "-"])
    ws_summary.cell(row=ws_summary.max_row, column=1).font = Font(bold=True)
    ws_summary.cell(row=ws_summary.max_row, column=2).font = danger_font if total_risk_count > 0 else safe_font
    ws_summary.cell(row=ws_summary.max_row, column=2).alignment = Alignment(horizontal="left")
    wb.save(save_path)
    log_fn(f"  📄 报告已生成: {save_path}")
    return save_path


def generate_report(result: Dict[str, Any], output_dir: str,
                    log: Optional[Callable[[str], None]] = None) -> str:
    """由 :func:`analyze_one_company` 的结果生成《{代码}_{简称}_深度排雷报告.xlsx》，返回路径。

    优先使用 ``result["results_rows"]`` + ``result["results_columns"]`` 还原桌面版 DataFrame
    （列名与列序完全一致）；若只有 ``result["results"]``（条件 -> {年份: 判定}）也能出报告。
    """
    log_fn = log or (lambda _m: None)
    result = result if isinstance(result, dict) else {}
    comp_code = str(result.get("code") or "")
    comp_name = str(result.get("name") or "")
    if not comp_code and not comp_name:
        comp_code, comp_name = (str(result.get("comp_key", "")).split(' ', 1) + [""])[:2]

    rows_map = result.get("results_rows") or {}
    cols_map = result.get("results_columns") or {}
    frames: Dict[str, pd.DataFrame] = {}
    if isinstance(rows_map, dict) and rows_map:
        for cond, rows in rows_map.items():
            if isinstance(rows, pd.DataFrame):
                frames[str(cond)] = rows
                continue
            rows = rows if isinstance(rows, list) else []
            cols = cols_map.get(cond) if isinstance(cols_map, dict) else None
            frames[str(cond)] = pd.DataFrame(rows, columns=list(cols)) if cols else pd.DataFrame(rows)
    else:
        frames = _coerce_frames(result.get("results"))

    return generate_excel_report(comp_code, comp_name, frames, output_dir, log_fn)


# ======================================================================
# 九、全流程（对应桌面版 process_task + start_analysis_thread 的参数提取）
# ======================================================================
def process_task(in_path: str, out_path: str, current_params: Dict[str, Any],
                 silent: bool = False,
                 callback: Optional[Callable[[], None]] = None,
                 log: Optional[Callable[[str], None]] = None,
                 progress: Optional[Callable[[int, int, str], None]] = None) -> Dict[str, Any]:
    """桌面版 ``process_task`` 的无界面版本（保留同义可选参数 ``silent`` / ``callback``）。

    等价于 ``run_all(in_path, out_path, current_params, log, progress)``；
    ``silent=True`` 只抑制结束提示日志，``callback`` 在结束时回调（异常不外抛）。
    """
    res = run_all(in_path, out_path, current_params, log=log, progress=progress)
    if not silent:
        log_fn = log or (lambda _m: None)
        log_fn("✅ 批量排雷任务全部完成！")
    if callback is not None:
        try:
            callback()
        except Exception as e:  # noqa: BLE001
            (log or (lambda _m: None))(f"⚠️ 流水线自动移交时发生异常: {str(e)}")
    return res


def run_all(input_dir: str, output_dir: str, params: Dict[str, Any],
            log: Optional[Callable[[str], None]] = None,
            progress: Optional[Callable[[int, int, str], None]] = None,
            employee_supplement: Optional[Dict[str, Dict[int, int]]] = None) -> Dict[str, Any]:
    """全流程：抽取 -> 逐家判断 -> 逐家出报告，报告写到 ``output_dir/造假排雷结果/``。

    返回::

        {"ok": True,
         "reports": [".../600887_伊利股份_深度排雷报告.xlsx", ...],
         "companies": [{"code","name","industry","suspect_years","summary"}, ...],
         "skipped": ["000000 xxx", ...],      # 可用年度不足 2 年（桌面版同样跳过）
         "output_dir": "...", "total": 2, "analyzed": 1}

    容错：目录/文件/单公司任一环节失败都只记日志并继续；``ok=False`` 时附 ``error`` 说明。
    """
    log_fn = make_logger(log)

    def _progress(cur: int, total: int, label: str) -> None:
        if progress is None:
            return
        try:
            progress(cur, total, label)
        except Exception:  # noqa: BLE001
            pass

    reports: List[str] = []
    companies: List[Dict[str, Any]] = []
    skipped: List[str] = []

    if not input_dir or not os.path.isdir(input_dir):
        log_fn(f"❌ 输入目录不存在: {input_dir}")
        return {"ok": False, "error": f"输入目录不存在: {input_dir}", "reports": reports,
                "companies": companies, "skipped": skipped, "output_dir": "", "total": 0,
                "analyzed": 0}
    if not output_dir:
        # 与桌面版 start_analysis_thread 的兜底一致：输出取输入目录的上级
        output_dir = os.path.dirname(os.path.abspath(input_dir.rstrip(r"\/")))
        log_fn(f"⚠️ 未提供输出目录，按桌面版规则回退为: {output_dir}")

    try:
        log_fn("=============== 财务造假智能排雷系统启动 ===============")
        main_out = os.path.join(output_dir, DIR_FRAUD)
        os.makedirs(main_out, exist_ok=True)
        p = normalize_params(params, log_fn)

        all_data = extract_company_data(input_dir, p, log_fn)
        if not all_data:
            log_fn("❌ 未提取到有效数据！")
            return {"ok": False, "error": "未提取到有效数据", "reports": reports,
                    "companies": companies, "skipped": skipped, "output_dir": main_out,
                    "total": 0, "analyzed": 0}

        # 【员工人数补充】把用户提供的历年员工人数注入，恢复 C1/C2 的桌面版口径判定。
        _supp = employee_supplement or {}
        if _supp:
            try:
                from .employee_data import merge_into_company_data
                merge_into_company_data(all_data, _supp, log_fn)
            except Exception as e:  # noqa: BLE001
                log_fn(f"⚠️ 员工人数补充数据合并失败（继续按缺数据处理）：{type(e).__name__}: {str(e)[:100]}")

        target_companies = list(all_data.keys())
        total = len(target_companies)
        log_fn(f"共扫描到 {total} 家公司，开始逐家排雷…")
        for index, comp_key in enumerate(target_companies):
            _progress(index + 1, total, comp_key)
            try:
                comp_data = all_data[comp_key]
                available_years = sorted([k for k in comp_data.keys() if k != "industry"])
                if len(available_years) < 2:
                    skipped.append(comp_key)
                    log_fn(f"  ⏭️ {comp_key} 可用年度不足 2 年，跳过（与桌面版一致）")
                    continue

                analysis_results = analyze_one_company(comp_key, comp_data, p, log_fn)
                report_path = generate_report(analysis_results, main_out, log_fn)
                reports.append(report_path)
                companies.append({"code": analysis_results.get("code", ""),
                                  "name": analysis_results.get("name", ""),
                                  "industry": analysis_results.get("industry", ""),
                                  "suspect_years": analysis_results.get("suspect_years", []),
                                  "summary": analysis_results.get("summary", "")})
            except Exception as e:  # noqa: BLE001  单公司失败不影响整批
                log_fn(f"  ❌ {comp_key} 处理失败，已跳过: {type(e).__name__}: {str(e)[:150]}")
                continue

        _progress(total, total, "完成")
        log_fn(f"✅ 批量排雷任务全部完成！报告目录: {main_out}")
        return {"ok": True, "reports": reports, "companies": companies, "skipped": skipped,
                "output_dir": main_out, "total": total, "analyzed": len(reports)}
    except Exception as e:  # noqa: BLE001  整体失败也只回报，不抛异常
        log_fn(f"❌ 错误: {str(e)}")
        return {"ok": False, "error": f"{type(e).__name__}: {str(e)}", "reports": reports,
                "companies": companies, "skipped": skipped, "output_dir": output_dir,
                "total": 0, "analyzed": len(reports)}
