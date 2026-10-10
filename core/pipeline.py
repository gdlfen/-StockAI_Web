# -*- coding: utf-8 -*-
"""
流水线编排层：把各模块串成与原桌面版一致的顺序，并支持单步运行。

顺序（已按需求调整）：
  1) 海选公司（免费数据）
  2) 年报数据搜索汇总（替代原“下载年报 + 年报提取”，产出结构不变）
  3) 造假排雷（18 项）
  4) 16 维度同行对比
  5) 企业 AI 深度分析（原“AI 企业深度”，保留；已取消 AI 文本分析）
  6) 好价估值

每一步的产物都落在同一个 data_dir 下，目录名与原程序完全一致
（海选情况/报表提取完善/造假排雷结果/16维度对比结果/企业分析/好价分析）。
"""
from __future__ import annotations

import os
import re
import traceback
from dataclasses import dataclass, field
from typing import Any, Callable, Dict, List, Optional

from .common import DiskCache, make_logger
from . import config as cfg_mod
from . import contracts as C

STAGES: List[Dict[str, str]] = [
    {"key": "haixuan", "name": "1. 海选公司", "desc": "免费数据多条件选股（替代问财浏览器抓取）"},
    {"key": "report", "name": "2. 年报数据搜索汇总", "desc": "AkShare 汇总三大表/分红/员工（替代下载+PDF提取）"},
    {"key": "fraud", "name": "3. 造假排雷", "desc": "18 项财务造假嫌疑判断"},
    {"key": "dim16", "name": "4. 16维度分析", "desc": "同行业 16 维度对比打分"},
    {"key": "enterprise", "name": "5. 企业AI深度分析", "desc": "联网检索 + LLM 生成企业深度报告"},
    {"key": "valuation", "name": "6. 好价估值", "desc": "现价/PE(TTM)/TTM股息率 → 好价格判断"},
]
STAGE_KEYS = [s["key"] for s in STAGES]


@dataclass
class PipelineContext:
    data_dir: str                 # 所有产物的根目录（云端建议 tempfile/app 目录）
    cache_dir: str                # 磁盘缓存目录
    config: Dict[str, Any] = field(default_factory=dict)
    log: Callable[[str], None] = field(default_factory=lambda: make_logger())
    progress: Optional[Callable[[int, int, str], None]] = None

    @property
    def cache(self) -> DiskCache:
        if not hasattr(self, "_cache") or self._cache is None:  # type: ignore[attr-defined]
            self._cache = DiskCache(self.cache_dir, ttl_hours=12)  # type: ignore[attr-defined]
        return self._cache  # type: ignore[attr-defined]

    def path(self, *parts: str) -> str:
        p = os.path.join(self.data_dir, *parts)
        os.makedirs(p if not os.path.splitext(p)[1] else os.path.dirname(p), exist_ok=True)
        return p


def _log_stage(ctx: PipelineContext, name: str) -> None:
    ctx.log("=" * 72)
    ctx.log(f"▶️ {name}")
    ctx.log("=" * 72)


# ======================================================================
# 各阶段
# ======================================================================
def stage_haixuan(ctx: PipelineContext, params: Dict[str, Any]) -> Dict[str, Any]:
    from . import screening
    _log_stage(ctx, "1. 海选公司（免费数据）")
    # 条件来源优先级：本次调用传入的结构化条件 > 传入的文本 > 用户配置里的结构化条件 > 配置文本 > 内置默认
    query = params.get("query") or ctx.config.get("haixuan_query") or cfg_mod.DEFAULT_WENCAI_QUERY
    conditions = params.get("conditions") or ctx.config.get("haixuan_conditions")
    pool = int(params.get("pool_size") or 300)
    ds_cfg = dict(ctx.config.get("data_source") or cfg_mod.DATA_SOURCE_DEFAULTS)
    ds_cfg.update(params.get("data_source") or {})
    res = screening.run_screening(
        query, ctx.data_dir, ctx.cache, pool_size=pool,
        workers=int(params.get("workers") or 6),
        log=ctx.log, progress=ctx.progress,
        extra_codes=params.get("extra_codes"),
        ds_cfg=ds_cfg,
        conditions=conditions,
    )
    return res


def stage_report(ctx: PipelineContext, params: Dict[str, Any]) -> Dict[str, Any]:
    from . import financial_search as fs
    _log_stage(ctx, "2. 年报数据搜索汇总（免费数据）")
    targets = params.get("targets") or []
    if not targets:
        # 自动从海选结果读取
        summary = os.path.join(ctx.data_dir, C.DIR_HAIXUAN, C.HAIXUAN_SUMMARY)
        if os.path.exists(summary):
            import pandas as pd
            # dtype=str：否则 pandas 会把纯数字的“代码”列读成 float（600887 → 600887.0）
            df = pd.read_excel(summary, dtype=str)
            targets = [{"code": str(r.get("代码", "")).zfill(6), "name": str(r.get("名称", "")),
                        "属性": C.TARGET_ATTR_PRIMARY} for _, r in df.iterrows()]
            ctx.log(f"   已从《{C.HAIXUAN_SUMMARY}》读取 {len(targets)} 家目标公司")
    if not targets:
        raise RuntimeError("没有目标公司：请先跑“海选公司”，或手工指定公司列表。")

    # 【代码 → 名称】补全：用户只填代码时（如“600887”），原逻辑会把代码当名称用，
    # 于是出现 “600887（600887）/ 600887_600887_深度排雷报告” 这类结果，
    # 下游的行业/同行/估值文件名也跟着错。这里统一用全市场清单补出真实简称。
    code2name: Dict[str, str] = {}
    try:
        from .universe import StockUniverse
        u = StockUniverse(ctx.cache)
        _udf = u.load()
        if _udf is not None and not _udf.empty and "code" in _udf.columns:
            code2name = {str(c).zfill(6): str(n).strip()
                         for c, n in zip(_udf["code"], _udf["name"])}
    except Exception as e:  # noqa: BLE001
        ctx.log(f"   ⚠️ 代码→名称 映射表获取失败（不影响提取，仅影响文件名显示）：{type(e).__name__}")

    fixed: List[Dict[str, Any]] = []
    filled = 0
    for t in targets:
        code = str(t.get("code") or "").strip()
        m = re.search(r"\d{6}", code)
        code = m.group(0) if m else code.zfill(6)
        name = str(t.get("name") or "").strip()
        # 名称为空、或名称本身就是那串代码 → 用清单里的简称替换
        if code in code2name and (not name or name == code or name.lstrip("0") == code.lstrip("0")
                                  or name in (code + ".0",) or not re.search(r"\D", name)):
            name = code2name[code]
            filled += 1
        fixed.append({**t, "code": code, "name": name or code})
    targets = fixed
    if filled:
        ctx.log(f"   ✅ 已为 {filled} 家公司补全名称（例如 {targets[0]['code']} → {targets[0]['name']}）")

    start_y = params.get("start_year")
    end_y = params.get("end_year")
    ds_cfg = dict(ctx.config.get("data_source") or cfg_mod.DATA_SOURCE_DEFAULTS)
    ds_cfg.update(params.get("data_source") or {})
    res = fs.run_annual_report_search(
        targets, ctx.data_dir, ctx.cache,
        years_back=int(params.get("years_back") or ds_cfg.get("years_back", 5)),
        start_year=int(start_y) if start_y else None,
        end_year=int(end_y) if end_y else None,
        log=ctx.log, progress=ctx.progress,
        include_peers=bool(params.get("include_peers", True)),
        ds_cfg=ds_cfg,
    )
    return res


def stage_fraud(ctx: PipelineContext, params: Dict[str, Any]) -> Dict[str, Any]:
    from . import fraud
    from . import employee_data as emp_mod
    _log_stage(ctx, "3. 造假排雷（18 项）")
    in_dir = os.path.join(ctx.data_dir, C.DIR_EXTRACT)
    if not os.path.isdir(in_dir):
        raise RuntimeError(f"未找到《报表提取完善》目录：{in_dir}，请先跑“年报数据搜索汇总”。")
    p = dict(ctx.config.get("fraud_params") or cfg_mod.default_fraud_params())
    p.update(params.get("fraud_params") or {})

    # 【员工人数补充】C1/C2 依赖员工人数，而免费数据源不提供。
    # 用户上传/随仓库放置的补充数据在这里注入，注入后 C1/C2 与桌面版口径一致。
    supplement, src = emp_mod.load_supplement(ctx.data_dir, _project_root())
    if supplement:
        ctx.log(f"   ℹ️ 已载入员工人数补充数据（{src}）")
    else:
        ctx.log("   ℹ️ 未提供员工人数补充数据：C1/C2 将标注“数据不足”"
                "（可在「造假排雷」页上传《员工人数补充.csv》恢复与桌面版一致的判定）")

    res = fraud.run_all(in_dir, ctx.data_dir, p, log=ctx.log, progress=ctx.progress,
                        employee_supplement=supplement)
    return _jsonable(res)


def _project_root() -> str:
    """项目根（含 core/ 的目录），用于查找随仓库上传的补充文件。"""
    here = os.path.dirname(os.path.abspath(__file__))      # .../core
    return os.path.dirname(here)


def stage_dim16(ctx: PipelineContext, params: Dict[str, Any]) -> Dict[str, Any]:
    from . import dim16
    _log_stage(ctx, "4. 16 维度同行对比")
    in_dir = os.path.join(ctx.data_dir, C.DIR_EXTRACT)
    if not os.path.isdir(in_dir):
        raise RuntimeError(f"未找到《报表提取完善》目录：{in_dir}，请先跑“年报数据搜索汇总”。")
    cfg = dict(ctx.config.get("config_16dim") or cfg_mod.default_16dim_config())
    cfg.update(params.get("config_16dim") or {})
    res = dim16.run_all(in_dir, ctx.data_dir, cfg, log=ctx.log, progress=ctx.progress)
    return _jsonable(res)


def stage_enterprise(ctx: PipelineContext, params: Dict[str, Any]) -> Dict[str, Any]:
    from . import enterprise_ai
    _log_stage(ctx, "5. 企业 AI 深度分析")
    ai = dict(ctx.config.get("ai") or cfg_mod.AI_DEFAULTS)
    ai.update(params.get("ai") or {})
    if not (ai.get("api_key") or "").strip():
        ctx.log("⚠️ 未配置 AI api_key，跳过企业AI深度分析（可在“设置”里填入 DeepSeek/OpenAI Key）。")
        return {"ok": False, "skipped": True, "reason": "未配置 api_key", "reports": []}
    template = params.get("template") or ai.get("template") or cfg_mod.DEFAULT_ENTERPRISE_PROMPT
    in_dir = os.path.join(ctx.data_dir, C.DIR_EXTRACT)
    if not os.path.isdir(in_dir):
        raise RuntimeError(f"未找到《报表提取完善》目录：{in_dir}，请先跑“年报数据搜索汇总”。")
    res = enterprise_ai.run_all(in_dir, ctx.data_dir, template, ai,
                               log=ctx.log, progress=ctx.progress,
                               company_filter=params.get("company_filter"))
    return _jsonable(res)


def stage_valuation(ctx: PipelineContext, params: Dict[str, Any]) -> Dict[str, Any]:
    from . import valuation
    _log_stage(ctx, "6. 好价估值")
    cfg = dict(ctx.config.get("valuation") or cfg_mod.VALUATION_DEFAULTS)
    cfg.update(params.get("valuation") or {})
    # 默认用年报汇总出来的名单作为估值标的（与桌面版“自动联动下载年报名单”一致）
    if not cfg.get("targets"):
        import pandas as pd
        list_path = os.path.join(ctx.data_dir, C.TARGET_LIST_FILENAME)
        if os.path.exists(list_path):
            # dtype=str：否则“代码”列被读成 float，str() 后变成 "600887.0"，
            # 后续报价接口与文件名都会带上 .0（日志里出现 600887.0 即此原因）
            df = pd.read_excel(list_path, dtype=str)
            tks = []
            for _, r in df.iterrows():
                raw = str(r.get("代码", "")).strip()
                m = re.search(r"\d{1,6}", raw)
                code = m.group(0).zfill(6) if m else ""
                name = str(r.get("名称", "")).strip()
                if code and code != "000000":
                    tks.append({"code": code, "name": name})
            if tks:
                cfg["targets"] = tks
                ctx.log(f"   已从《{C.TARGET_LIST_FILENAME}》读取 {len(tks)} 只估值标的")

    # 若无显式标的、也没有年报名单 → 才使用内置默认标的（平安银行/贵州茅台/腾讯控股…）。
    # 有明确名单时**不再混入默认标的**，否则用户明明只传了 600887，结果却出现 7 条记录。
    using_defaults = False
    if not cfg.get("targets"):
        using_defaults = True
        ctx.log("   ℹ️ 未指定估值标的，使用内置默认标的（可在「好价估值」页手工填写覆盖）")

    # 名称补全：用户只填代码（或填了简称但配不上代码）时，统一用全市场清单补齐简称，
    # 否则表格里会出现只有数字、或“无法识别标的：微软”这类情况。
    try:
        from .universe import StockUniverse
        _u = StockUniverse(ctx.cache)
        _udf = _u.load()
        if _udf is not None and not _udf.empty and "code" in _udf.columns:
            _c2n = {str(c).zfill(6): str(n).strip() for c, n in zip(_udf["code"], _udf["name"])}
            _n2c = {str(n).strip(): str(c).zfill(6) for c, n in zip(_udf["code"], _udf["name"])}
            _fixed = []
            for t in (cfg.get("targets") or []):
                code = str(t.get("code") or "").strip()
                name = str(t.get("name") or "").strip()
                m = re.search(r"\d{6}", code)
                code = m.group(0) if m else code
                if not code and name in _n2c:          # 只填了中文简称
                    code = _n2c[name]
                if code and (not name or name == code or not re.search(r"\D", name)):
                    name = _c2n.get(code, name)
                # 估值表格里显示“简称(代码)”更易读（原来只显示 600887）
                if code and name and name != code:
                    _fixed.append({**t, "code": code, "name": name,
                                   "display": f"{name}({code})"})
                else:
                    _fixed.append({**t, "code": code, "name": name})
            cfg["targets"] = _fixed
    except Exception as e:  # noqa: BLE001
        ctx.log(f"   ⚠️ 估值标的名称补全失败（不影响取值）：{type(e).__name__}")

    # 港股/美股默认标的：只在“使用默认标的”时才带上，避免污染用户指定名单
    if not using_defaults:
        for k in ("a_stocks", "hk_stocks", "us_stocks"):
            cfg.pop(k, None)

    cfg["_cache"] = ctx.cache
    out_dir = os.path.join(ctx.data_dir, C.DIR_VALUATION)
    res = valuation.run_valuation(cfg, out_dir, log=ctx.log, progress=ctx.progress)
    return _jsonable(res)


_STAGE_FUNCS = {
    "haixuan": stage_haixuan,
    "report": stage_report,
    "fraud": stage_fraud,
    "dim16": stage_dim16,
    "enterprise": stage_enterprise,
    "valuation": stage_valuation,
}


def _jsonable(obj: Any) -> Any:
    """把结果里可能出现的 NaN/Inf/DataFrame 转成可 JSON 序列化（FastAPI 严格模式需要）。"""
    import math
    import pandas as pd
    if isinstance(obj, dict):
        return {str(k): _jsonable(v) for k, v in obj.items()}
    if isinstance(obj, (list, tuple)):
        return [_jsonable(v) for v in obj]
    if isinstance(obj, pd.DataFrame):
        try:
            return _jsonable(obj.where(pd.notna(obj), None).to_dict("records"))
        except Exception:
            return []
    if isinstance(obj, float):
        return None if (math.isnan(obj) or math.isinf(obj)) else obj
    if obj is None or isinstance(obj, (str, int, bool)):
        return obj
    try:
        import numpy as np
        if isinstance(obj, (np.integer,)):
            return int(obj)
        if isinstance(obj, (np.floating,)):
            f = float(obj)
            return None if (math.isnan(f) or math.isinf(f)) else f
        if isinstance(obj, (np.bool_,)):
            return bool(obj)
    except Exception:
        pass
    return str(obj)


# ======================================================================
# 运行入口
# ======================================================================
def run_stages(ctx: PipelineContext, stage_keys: List[str],
               params: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
    """按给定顺序运行指定阶段，收集每阶段结果；单阶段失败不阻断后续（可自行决定）。"""
    params = params or {}
    results: Dict[str, Any] = {}
    for key in stage_keys:
        fn = _STAGE_FUNCS.get(key)
        if fn is None:
            ctx.log(f"⚠️ 未知阶段：{key}（已跳过）")
            continue
        try:
            results[key] = fn(ctx, params.get(key, {}) if isinstance(params.get(key), dict) else params)
        except Exception as e:  # noqa: BLE001
            ctx.log(f"❌ 阶段[{key}]执行失败: {type(e).__name__}: {e}")
            ctx.log(traceback.format_exc()[:1500])
            results[key] = {"ok": False, "error": f"{type(e).__name__}: {e}"}
            # 关键前置阶段失败则中止，避免后续空跑
            if key in ("report",) and not results[key].get("ok"):
                ctx.log("⛔ 年报数据汇总未成功，后续依赖它的阶段已中止。")
                break
    ok = all((r or {}).get("ok", False) for r in results.values() if isinstance(r, dict)) if results else False
    return {"ok": ok, "results": _jsonable(results)}


def make_context(data_dir: str, cache_dir: str, user_config: Optional[Dict[str, Any]] = None,
                 log: Optional[Callable[[str], None]] = None,
                 progress: Optional[Callable[[int, int, str], None]] = None) -> PipelineContext:
    cfg = user_config or cfg_mod.load_user_config(data_dir)
    # 「设置 → 数据源 → 行业分类数据源」的持久设置，同步到环境变量供 core.universe 读取。
    # （环境变量优先级最高，所以这里只在用户明确选过时写入。）
    try:
        _src = str((cfg.get("data_source") or {}).get("industry_source") or "").strip().lower()
        if _src in ("cninfo", "eastmoney"):
            os.environ["VIM_INDUSTRY_SOURCE"] = _src
        elif _src == "":
            os.environ.pop("VIM_INDUSTRY_SOURCE", None)
    except Exception:  # noqa: BLE001
        pass
    return PipelineContext(data_dir=data_dir, cache_dir=cache_dir, config=cfg,
                           log=log or make_logger(), progress=progress)
