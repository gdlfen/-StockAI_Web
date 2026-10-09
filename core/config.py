# -*- coding: utf-8 -*-
"""
配置层：完整保留原桌面版的两套参数（造假排雷 18 项 / 16 维度阈值），
并新增云端版数据源与 AI 配置。所有默认值与原程序逐字一致。
"""
from __future__ import annotations

import json
import os
from typing import Any, Dict

# ======================================================================
# 一、造假排雷（原 FraudDetectionApp.params）—— 22 项阈值，键名与桌面版逐字一致
# ======================================================================
FRAUD_PARAM_DEFS: Dict[str, Dict[str, Any]] = {
    "c1_emp_drop": {"name": "1. 员工人数减少大于(%)", "val": 10.0},
    "c1_rev_inc": {"name": "1. 且营业收入增加大于(%)", "val": 10.0},
    "c2_salary_diff": {"name": "2. 薪酬低于行业平均大于(%)", "val": 30.0},
    "c3_related_tx": {"name": "3. 重大关联交易次数(>=)", "val": 2},
    "c5_rev_growth": {"name": "5. 营收增长率小于(%)", "val": 15.0},
    "c5_margin": {"name": "5. 或毛利率小于(%)", "val": 30.0},
    "c5_inv_ratio": {"name": "5. 或存货/总资产大于(%)", "val": 15.0},
    "c7_cash_ratio": {"name": "7. (货币+金融资产)/总资产小于(%)", "val": 10.0},
    "c8_ar_ratio": {"name": "8. 应收账款/总资产大于(%)", "val": 15.0},
    "c9_fa_ratio": {"name": "9. 固定资产/总资产大于(%)", "val": 40.0},
    "c10_cip_ratio": {"name": "10. 在建工程/总资产大于(%)", "val": 5.0},
    "c11_prepay_ratio": {"name": "11. (预付+其他应收)/总资产大于(%)", "val": 10.0},
    "c12_fa_cip_ratio": {"name": "12. (固资+在建)/总资产大于(%)", "val": 45.0},
    "c13_ar_limit": {"name": "13. 应收/总资>5%或存货/总资大于(%)", "val": 15.0},
    "c13_fa_cip_inc": {"name": "13. 且(固资+在建)增幅大于(%)", "val": 20.0},
    "c14_int_rate": {"name": "14. 利息收入/货币资金小于(%)", "val": 2.0},
    "c15_cash_limit": {"name": "15. 货币/总资>10%且借款/总资大于(%)", "val": 10.0},
    "c16_impair_ratio": {"name": "16. 减值损失/各项资产总和小于(%)", "val": 0.1},
    "c17_margin_jump": {"name": "17. 毛利率增幅大于(%)", "val": 20.0},
    "c17_margin_diff": {"name": "17. 且偏离行业平均毛利大于(%)", "val": 15.0},
    "c18_ar_diff": {"name": "18. 应收占比偏离行业平均大于(%)", "val": 15.0},
    "c18_inv_diff": {"name": "18. 或存货占比偏离行业平均大于(%)", "val": 15.0},
    # 15/16 条的科目清单（桌面版为独立 tk 变量，这里并入同一字典，便于前端统一渲染）
    "c15_components_str": {
        "name": "15. 有息负债科目(逗号分隔)",
        "val": "短期借款,一年内到期的非流动负债,长期借款,应付债券,租赁负债,应付票据,交易性金融负债,长期应付款",
    },
    "c16_components_str": {
        "name": "16. 减值相关资产科目(逗号分隔)",
        "val": ("应收账款,应收票据,其他应收款,存货,固定资产,在建工程,债权投资,其他债权投资,长期债权投资,"
                "投资性房地产,无形资产,生产性生物资产,长期应收款,合同资产,可供出售金融资产,持有至到期投资,"
                "工程物资,油气资产,商誉"),
    },
}

# 造假排雷用的行业均值库（原程序内置：平均工资 + C17/C18 所需的三张均值表）
MACRO_AVG_WAGE: Dict[str, Dict[str, Any]] = {
    "默认_平均工资": {"2018": 82461, "2019": 90501, "2020": 97379, "2021": 106837, "2022": 114029,
                      "2023": 120698, "2024": 124110},
    "默认_毛利率": {"2018": 30.0, "2019": 30.0, "2020": 30.0, "2021": 30.0, "2022": 30.0,
                    "2023": 30.0, "2024": 30.0},
    "默认_应收占比": {"2018": 10.0, "2019": 10.0, "2020": 10.0, "2021": 10.0, "2022": 10.0,
                      "2023": 10.0, "2024": 10.0},
    "默认_存货占比": {"2018": 12.0, "2019": 12.0, "2020": 12.0, "2021": 12.0, "2022": 12.0,
                      "2023": 12.0, "2024": 12.0},
}

# ======================================================================
# 二、16 维度分析（原 FULL_CONFIG）—— 逐字保留
# ======================================================================
FULL_CONFIG: Dict[str, Dict[str, Any]] = {
    "f_dim2_money": {"val": "货币资金,交易性金融资产", "desc": "维2: 准货币资金(逗号分隔)", "tab": "自定义公式", "is_pct": False, "is_text": True},
    "f_dim2_debt": {"val": "短期借款,一年内到期的非流动负债,长期借款,应付债券,长期应付款,应付票据,交易性金融负债", "desc": "维2: 有息负债", "tab": "自定义公式", "is_pct": False, "is_text": True},
    "f_dim3_pay": {"val": "应付票据,应付账款,预收款项,合同负债", "desc": "维3: 应付预收项(逗号分隔)", "tab": "自定义公式", "is_pct": False, "is_text": True},
    "f_dim3_recv": {"val": "应收票据,应收账款,应收款项融资,预付款项,合同资产", "desc": "维3: 应收预付项(逗号分隔)", "tab": "自定义公式", "is_pct": False, "is_text": True},
    "f_dim6_inv": {"val": "以公允价值计量且其变动计入当期损益的金融资产,债权投资,其他债权投资,可供出售金融资产,持有至到期投资,长期股权投资,其他权益工具投资,其他非流动金融资产,投资性房地产", "desc": "维6: 投产类资产", "tab": "自定义公式", "is_pct": False, "is_text": True},
    "f_dim12_op_plus": {"val": "营业总收入,其他收益,投资收益,汇兑收益,净敞口套期收益,公允价值变动收益,信用减值损失,资产减值损失,资产处置收益", "desc": "维12: 利润加项", "tab": "自定义公式", "is_pct": False, "is_text": True},
    "f_dim12_op_minus": {"val": "营业总成本", "desc": "维12: 利润减项", "tab": "自定义公式", "is_pct": False, "is_text": True},
    "d1_asset_poor_ratio": {"val": 0.80, "desc": "维1: 实力差淘汰线", "tab": "维度 1-4", "is_pct": True},
    "d1_asset_score_base": {"val": 10.0, "desc": "维1: 资产规模基础分", "tab": "维度 1-4", "is_pct": False},
    "d2_debt_safe": {"val": 0.40, "desc": "维2: 负债率极安全线", "tab": "维度 1-4", "is_pct": True},
    "d2_debt_warn": {"val": 0.60, "desc": "维2: 负债率淘汰线", "tab": "维度 1-4", "is_pct": True},
    "d2_debt_danger": {"val": 0.70, "desc": "维2: 负债率危险线", "tab": "维度 1-4", "is_pct": True},
    "d3_diff_score_base": {"val": 10.0, "desc": "维3: 差额基础分", "tab": "维度 1-4", "is_pct": False},
    "d3_diff_score_step": {"val": 10.0, "desc": "维3: 差额计分乘数", "tab": "维度 1-4", "is_pct": False},
    "d4_ar_1": {"val": 0.01, "desc": "维4: 应收占资产-极畅销线", "tab": "维度 1-4", "is_pct": True},
    "d4_ar_2": {"val": 0.03, "desc": "维4: 应收占资产-优秀线", "tab": "维度 1-4", "is_pct": True},
    "d4_ar_3": {"val": 0.10, "desc": "维4: 应收占资产-一般线", "tab": "维度 1-4", "is_pct": True},
    "d4_ar_out": {"val": 0.15, "desc": "维4: 应收占资产淘汰线", "tab": "维度 1-4", "is_pct": True},
    "d4_ar_4": {"val": 0.20, "desc": "维4: 应收占资产-滞销线", "tab": "维度 1-4", "is_pct": True},
    "d5_fix_1": {"val": 0.20, "desc": "维5: 固定资产-轻资产线", "tab": "维度 5-8", "is_pct": True},
    "d5_fix_out": {"val": 0.40, "desc": "维5: 固定资产淘汰线", "tab": "维度 5-8", "is_pct": True},
    "d5_fix_2": {"val": 0.50, "desc": "维5: 固定资产-重资产线", "tab": "维度 5-8", "is_pct": True},
    "d6_inv_best": {"val": 0.00, "desc": "维6: 投资类资产-极优秀线", "tab": "维度 5-8", "is_pct": True},
    "d6_inv_out": {"val": 0.10, "desc": "维6: 投资类资产淘汰线", "tab": "维度 5-8", "is_pct": True},
    "d7_ar_risk": {"val": 0.05, "desc": "维7: 爆雷-应收高危线", "tab": "维度 5-8", "is_pct": True},
    "d7_inv_risk": {"val": 0.15, "desc": "维7: 爆雷-存货高危线", "tab": "维度 5-8", "is_pct": True},
    "d7_gw_risk": {"val": 0.10, "desc": "维7: 爆雷-商誉淘汰线", "tab": "维度 5-8", "is_pct": True},
    "d8_rev_poor_ratio": {"val": 0.20, "desc": "维8: 营收差淘汰线", "tab": "维度 5-8", "is_pct": True},
    "d8_rev_score_base": {"val": 10.0, "desc": "维8: 营收规模基础分", "tab": "维度 5-8", "is_pct": False},
    "d8_rev_score_step": {"val": 10.0, "desc": "维8: 营收计分乘数", "tab": "维度 5-8", "is_pct": False},
    "d8_growth_good": {"val": 0.10, "desc": "维8: 营收高成长优良线", "tab": "维度 5-8", "is_pct": True},
    "d9_gm_out": {"val": 0.40, "desc": "维9: 毛利率淘汰线", "tab": "维度 9-12", "is_pct": True},
    "d9_gm_vol_safe": {"val": 0.10, "desc": "维9: 毛利波动正常线", "tab": "维度 9-12", "is_pct": True},
    "d9_gm_vol_out": {"val": 0.20, "desc": "维9: 毛利波动淘汰线", "tab": "维度 9-12", "is_pct": True},
    "d10_exp_gm_safe": {"val": 0.40, "desc": "维10: 费用毛利比优秀线", "tab": "维度 9-12", "is_pct": True},
    "d10_exp_gm_out": {"val": 0.60, "desc": "维10: 费用毛利比淘汰线", "tab": "维度 9-12", "is_pct": True},
    "d11_sales_exp_1": {"val": 0.15, "desc": "维11: 销售费率优秀线", "tab": "维度 9-12", "is_pct": True},
    "d11_sales_exp_out": {"val": 0.30, "desc": "维11: 销售费率淘汰线", "tab": "维度 9-12", "is_pct": True},
    "d11_trend_out": {"val": 0.00, "desc": "维11: 销售费变动趋势分界", "tab": "维度 9-12", "is_pct": True},
    "d11_trend_score_base": {"val": 10.0, "desc": "维11: 趋势变动基础分", "tab": "维度 9-12", "is_pct": False},
    "d11_trend_score_step": {"val": 0.10, "desc": "维11: 趋势计分步长", "tab": "维度 9-12", "is_pct": True},
    "d12_core_out": {"val": 0.15, "desc": "维12: 主营利润率及格线", "tab": "维度 9-12", "is_pct": True},
    "d12_profit_out": {"val": 0.80, "desc": "维12: 利润质量及格线", "tab": "维度 9-12", "is_pct": True},
    "d13_non_op_out": {"val": 0.05, "desc": "维13: 营业外净额淘汰线", "tab": "维度 13-16", "is_pct": True},
    "d14_np_poor_ratio": {"val": 0.20, "desc": "维14: 净利差淘汰线", "tab": "维度 13-16", "is_pct": True},
    "d14_np_score_base": {"val": 10.0, "desc": "维14: 净利规模基础分", "tab": "维度 13-16", "is_pct": False},
    "d14_np_score_step": {"val": 10.0, "desc": "维14: 净利计分乘数", "tab": "维度 13-16", "is_pct": False},
    "d14_np_growth": {"val": 0.10, "desc": "维14: 净利持续增长优秀线", "tab": "维度 13-16", "is_pct": True},
    "d14_np_growth_out": {"val": 0.00, "desc": "维14: 净利增长淘汰线", "tab": "维度 13-16", "is_pct": True},
    "d15_capex_slow": {"val": 0.03, "desc": "维15: CAPEX过慢界限", "tab": "维度 13-16", "is_pct": True},
    "d15_capex_safe": {"val": 0.60, "desc": "维15: CAPEX扩张安全线", "tab": "维度 13-16", "is_pct": True},
    "d15_capex_out": {"val": 1.00, "desc": "维15: CAPEX激进界限", "tab": "维度 13-16", "is_pct": True},
    "d15_cash_trend_out": {"val": 0.00, "desc": "维15: 销售现金变动界限", "tab": "维度 13-16", "is_pct": True},
    "d15_cash_score_base": {"val": 10.0, "desc": "维15: 现金变动基础分", "tab": "维度 13-16", "is_pct": False},
    "d15_cash_score_step": {"val": 0.10, "desc": "维15: 现金计分步长", "tab": "维度 13-16", "is_pct": True},
    "d16_div_out_low": {"val": 0.30, "desc": "维16: 分红率下限淘汰线", "tab": "维度 13-16", "is_pct": True},
    "d16_div_out_high": {"val": 0.70, "desc": "维16: 分红率上限淘汰线", "tab": "维度 13-16", "is_pct": True},
}


def default_16dim_config() -> Dict[str, Any]:
    return {k: v["val"] for k, v in FULL_CONFIG.items()}


def default_fraud_params() -> Dict[str, Any]:
    return {k: v["val"] for k, v in FRAUD_PARAM_DEFS.items()}


# ======================================================================
# 三、海选默认条件（问财口径，与原程序模板一致）
# ======================================================================
DEFAULT_WENCAI_QUERY = (
    "连续5年加权roe>20，连续5年净利润现金含量>100，连续5年毛利率>40，上市时间>3年，剔除北交所，非金融股"
)
DEFAULT_AI_QUERY = (
    "请在A股中严格筛选同时符合以下条件的股票并作分析：2021、2022、2023、2024、2025年（五个完整会计年度）"
    "roe_weighted>40%、净利润现金含量>80%，gross_profit_margin>40%，上市时间>3年，剔除北交所中的公司和金融股。"
    "如果没有符合条件的则跳过，不能擅自乱抓其他不完全符合条件的公司。"
)
HAIXUAN_PRESETS: Dict[str, str] = {
    "默认A股多条件": "连续5年加权roe>25，连续5年净利润现金含量>80，连续5年毛利率>40，上市时间>3年，剔除北交所，非金融股",
    "优势大市值": "连续5年加权roe>15%，剔除北交所，非金融股\n量价齐升，且市值大于100亿\n近3年净利润增长率大于20%",
    "高股息低估值": "连续3年股息率>5%\n市盈率<15\n市值>200亿",
    "标准条件(原程序)": DEFAULT_WENCAI_QUERY,
}

# ======================================================================
# 四、好价估值默认设置
# ======================================================================
VALUATION_DEFAULTS: Dict[str, Any] = {
    "sz_pe": True, "hs_pe": True, "sp_pe": True,
    "cn_10y": True, "us_10y": True, "fed": True,
    "silent": True,
    "a_stocks": ["平安银行", "贵州茅台"],
    "hk_stocks": ["腾讯控股", "03968"],
    "us_stocks": ["微软", "AAPL"],
}

# ======================================================================
# 五、AI 配置（企业AI深度分析；云端版已取消AI文本分析）
# ======================================================================
AI_DEFAULTS: Dict[str, Any] = {
    "base_url": "https://api.deepseek.com/v1",
    "api_key": "",
    "model": "deepseek-chat",
    "temperature": 0.3,
    "web_search": True,
    "web_results": 3,
}

DEFAULT_ENTERPRISE_PROMPT = (
    "你是一位资深的企业分析师，请根据以下资料，围绕指定的维度和重点对目标企业【{company}】"
    "（参考同行【{competitors}】）作全面深入中肯的分析：\n\n"
    "【参考资料】\n[全网检索情报]:\n{web_info}\n\n[本地财报等内部数据]:\n{doc_info}\n\n"
    "【分析要求与维度】\n（一）行业与竞争环境分析。\n"
    "1、行业空间：包括市场规模、成熟阶段、现已占份额，尚有空间等。\n"
    "2、竞争格局：包括市场份额集中度（如CR5）、波特五力模型分析。\n"
    "3、行业政策与趋势：包括监管导向、技术变革（如数字化、低碳化等）对行业的影响。\n\n"
    "（二）企业全面对比分析。\n"
    "1、企业领导者个人的特质：包括领导者的志向、见识、恒心、自身学习等；\n"
    "2、企业文化：包括使命、愿景、核心价值观等；\n"
    "3、企业治理与管理层分析：包括股权结构、管理层背景及过往战略执行能力、信息披露与诚信等；\n"
    "4、企业的商业模式：用户模式、产品模式、推广模式、盈利模式等；\n"
    "5、主营业务情况：包括收入结构、核心产品或服务等；\n"
    "6、竞争优势：包括品牌、技术专利、网络效应、特许经营权等护城河情况；\n"
    "7、重要报表数据：剖析利润表、资产负债表、现金流量表等核心指标变动；\n"
    "8、关键财务比率：包括盈利能力(ROE/ROA)、偿债能力、运营效率、价值指标等；\n"
    "9、企业团队与内外沟通：组织架构、激励机制、人际关系治理等；\n"
    "10、企业现金流：开源与节流情况能力；\n"
    "11、企业系统与法规：企业规则制度、遵法守法情况等；\n"
    "12、风险与未来成长性：包括风险识别、成长驱动（新产能/新市场）、战略规划可行性等。\n\n"
    "（三）综合判断。\n"
    "对以上信息实行交叉验证，勾稽核对。总结提炼公司的核心优势与致命短板，结合成长性与风险评估当前估值水平，"
    "最终得出公司是否具备长期投资价值的结论。"
)

# ======================================================================
# 六、数据源开关（全部免费）
# ======================================================================
DATA_SOURCE_DEFAULTS: Dict[str, Any] = {
    "enable_akshare": True,      # 主数据源
    "enable_baostock": True,     # 备用（未安装则自动跳过）
    "enable_tushare": False,     # 需 token，默认关闭
    "tushare_token": "",
    "enable_ths": True,          # 同花顺口径（分红/主营）
    "enable_eastmoney": True,    # 东财数据中心
    "request_timeout": 20,
    "max_retries": 3,
    "sleep_between": 0.35,       # 反爬节流
    "years_back": 5,             # 默认取最近 5 年
    "max_peers": 3,              # 同行前三
}

# ======================================================================
# 七、运行期配置读写（供 Streamlit/FastAPI 复用）
# ======================================================================
CONFIG_FILENAME = "cloud_config.json"


def config_path(data_dir: str) -> str:
    return os.path.join(data_dir, CONFIG_FILENAME)


def load_user_config(data_dir: str) -> Dict[str, Any]:
    """读取用户配置；读不到就返回默认值（不抛异常，云端环境友好）。"""
    path = config_path(data_dir)
    cfg: Dict[str, Any] = {
        "fraud_params": default_fraud_params(),
        "config_16dim": default_16dim_config(),
        "valuation": dict(VALUATION_DEFAULTS),
        "ai": dict(AI_DEFAULTS),
        "data_source": dict(DATA_SOURCE_DEFAULTS),
        "haixuan_query": DEFAULT_WENCAI_QUERY,
        "haixuan_preset": "默认A股多条件",
    }
    try:
        if os.path.exists(path):
            with open(path, "r", encoding="utf-8") as f:
                saved = json.load(f)
            for k, v in saved.items():
                if isinstance(v, dict) and isinstance(cfg.get(k), dict):
                    cfg[k].update(v)
                else:
                    cfg[k] = v
    except Exception:
        pass
    return cfg


def save_user_config(data_dir: str, cfg: Dict[str, Any]) -> str:
    os.makedirs(data_dir, exist_ok=True)
    path = config_path(data_dir)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)
    return path
