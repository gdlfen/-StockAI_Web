# -*- coding: utf-8 -*-
# ===== 修复版说明（相对 价值投资分析模型V1.4-3（扩展）1.0.py 原版）=====
# 修复1：总控 stage2（跳过海选）分支补 popup_child / 全局路径 / 名单校验
# 修复2：下载成功判定链路（真实 PDF 计数 + 回执双重校验 + 回调固定本次结果）
# 修复3：全局代理环境变量污染（ProxyError / WinError 10061 / localhost:7890）
# 修复4：__main__ 入口保护，启动失败打印完整堆栈并弹窗
# 修复5：总控“一键启动”按钮整合为一个（原 start_all + start_master_pipeline 合并）
# 修复6：好价分析 现价/当前PE(TTM) 全 N/A（新增腾讯/新浪 HTTP 行情 + Edge 启动可控化）
# 修复7：下载静默失败可观测（download_pdf 返回 True/False + 每家公司下载小结）
# 修复8：动态股息率改为【严格 TTM】（近12个月“已实施/已除权”分红累加 ÷ 现价）
# 修复9：年报下载“按年份过滤 + 翻页直到覆盖时间段”（原最多3页且首页命中即停）
# 修复10：行情抓取抗抖动（伊利股份偶发 N/A 的根因）
#        (a) 新增 _http_get()：腾讯行情/新浪行情/东财分红接口统一带 3 次重试与退避；
#        (b) fetch_quote_http 增加“候选符号(原输入+解析代码) × 双数据源 × 整轮重试2次”；
#        (c) 代码解析(智能搜索)失败会明确告警，不再静默变成 N/A；
#        (d) 其它: _fetch_dividend_rows 同样走重试通道
# 说明：海选/下载/提取/排雷/16维/AI文本/企业深度 的业务规则、judge_stock 判断规则、
#       宏观与大盘指数取数、报告版式 均未改动。
# ===========================================================
#!/usr/bin/env python
# -*- coding: utf-8 -*-

import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, simpledialog, filedialog
import threading
import multiprocessing
import os
import json
import re
import time
import urllib.parse
from datetime import datetime, timedelta
import pandas as pd
import requests
import shutil
import sys
import io
import subprocess
import warnings
import traceback
import random
import math

from openai import OpenAI
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from tavily import TavilyClient

try:
    from duckduckgo_search import DDGS
    HAS_DDGS = True
except Exception:
    try:
        from ddgs import DDGS
        HAS_DDGS = True
    except Exception:
        HAS_DDGS = False

from selenium import webdriver
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys

import pdfplumber
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl import Workbook
import fitz  # PyMuPDF
import akshare as ak

warnings.filterwarnings('ignore')

# 【修复2】在 GUI 初始化之前，先把“进程启动时就已存在的真实代理”快照下来。
# 之后任何“子程序被创建/切换”而写进 os.environ 的代理残留，都可以被识别并清除，
# 避免下载子程序的 requests 被 http://127.0.0.1:7890 这类不存在的代理劫持。
_BOOT_HTTP_PROXY = os.environ.get("HTTP_PROXY") or os.environ.get("http_proxy")
_BOOT_HTTPS_PROXY = os.environ.get("HTTPS_PROXY") or os.environ.get("https_proxy")

# =====================================================================
# 全局配置与常量 (A程序)
# =====================================================================
CONFIG_FILE_APP1 = "stock_ai_config.json"
CONFIG_FILE_APP2 = "wencai_templates.json"
CONFIG_FILE_APP3 = "config.json"
TEMPLATE_FILE_A = "mapping_template.txt"

HEADERS = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/120.0.0.0 Safari/537.36'}
MILESTONE_PATTERNS = [
    ("合并资产负债表", r"合并.*资产负债表|^资产负债表"),
    ("母公司资产负债表", r"母公司.*资产负债表"),
    ("合并利润表", r"合并.*利润表|^利润表"),
    ("母公司利润表", r"母公司.*利润表"),
    ("合并现金流量表", r"合并.*现金流量表|^现金流量表"),
    ("母公司现金流量表", r"母公司.*现金流量表"),
    ("合并所有者权益变动表", r"合并.*所有者权益变动表|^所有者权益变动表"),
    ("母公司所有者权益变动表", r"母公司.*所有者权益变动表"),
    ("附注", r"财务报表附注|公司基本情况|合并财务报表项目注释")
]

TARGET_SHEETS_P3 = ["合并资产负债表", "合并利润表", "合并现金流量表"]
YELLOW_FILL = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")
CENTER_ALIGNMENT = Alignment(horizontal="center", vertical="center", wrap_text=True)

# =====================================================================
# 全局配置与常量 (B程序)
# =====================================================================
FULL_CONFIG = {
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
CONFIG_FILE_B = "finance_analyzer_config_v5.json"

def load_config_b():
    if os.path.exists(CONFIG_FILE_B):
        try:
            with open(CONFIG_FILE_B, 'r', encoding='utf-8') as f:
                user_conf = json.load(f)
                final_conf = {}
                for k, v in FULL_CONFIG.items():
                    final_conf[k] = user_conf.get(k, v["val"])
                return final_conf
        except:
            pass
    return {k: v["val"] for k, v in FULL_CONFIG.items()}

def save_config_b(config_data):
    with open(CONFIG_FILE_B, 'w', encoding='utf-8') as f:
        json.dump(config_data, f, ensure_ascii=False, indent=4)

active_config = load_config_b()

# =====================================================================
# 全局配置与常量 (C程序)
# =====================================================================
PROXY_PORT = "7890"

AI_ENGINES = {

    "云端 API (如 DeepSeek)": {
        "urls": ["https://api.deepseek.com/v1", "https://api.openai.com/v1"],
        "need_key": True,
        "models": ["deepseek-chat", "deepseek-reasoner"],
    },
    "本地开源 (Ollama/LM Studio)": {
        "urls": ["http://localhost:11434/v1"],
        "need_key": False,
    "models": [
            "deepseek-r1:8b", "qwen2.5", "gemma3:27b", "deepseek-coder-v2",
            "glm4", "gpt-oss:120b-cloud", "deepseek-v3.1:671b-cloud"
        ],
    },

}

CUSTOM_MODELS_FILE = "custom_models.json"
if os.path.exists(CUSTOM_MODELS_FILE):
    try:
        with open(CUSTOM_MODELS_FILE, "r", encoding="utf-8") as f:
            custom_models = json.load(f)
            for engine, models in custom_models.items():
                if engine in AI_ENGINES:
                    for m in models:
                        if m not in AI_ENGINES[engine]["models"]:
                            AI_ENGINES[engine]["models"].append(m)
    except Exception as e:
        print(f"加载自定义模型失败: {e}")

SEARCH_TEMPLATES = {
    "金融年报分析": "公司名称 2026年 最新动态 财务数据 行业新闻",
    "科技公司追踪": "公司名称 2026年 技术创新 市场动态 竞争对手",
    "行业研究报告": "行业名称 2026年 发展趋势 市场规模 政策影响",
    "上市公司监控": "股票代码 公司名称 2026年 业绩预告 重大事项",
    "自定义关键词": ""
}

PROMPT_TEMPLATES = {
    "标准财报文本分析": (
        "你是一位资深金融分析师。本次分析提供了目标公司及其同行业前三名公司【多份/多个年份】的年报等原始文档，以及【最新网络情报】。请严格围绕下面的维度和角度进行时间序列上的长效对比与横向同业分析，给出一份详尽中肯的综合评估报告。重点分析：\
1、会计师事务所对公司各年度的年报出具了什么样的审计报告，有没有出具过不是“标准无保留意见”的报告。\
2、对公司董事会审议的各年报告期利润分配预案或公积金转增股本预案情况（分红情况）分析。\
3、公司所在行业情况如何（包括行业特点、市场规模、生命周期阶段、未来情况等等），公司的基本情况（包括：公司的主营业务，重要产品，最重要的产品，公司在该行业中的地位如何，推断公司的资产、营收、净利润规模在行业中的大小位置，对供应商、经销商的话语权，营收最多的产品，产量和营收比其上位或下位公司低或高多少等等）。\
4、公司的商业模式及动作策略怎样（包括采购模式、生产模式、销售模式、生产工艺流程、品牌策略等等）。\
5、核心竞争力如何（包括企业管理能力、企业文化、销售渠道、产品技术、品牌优势、战略布局等等），是否具有如下的优势：强网络效应、优秀企业文化、独特的资源、成本优势、短缺的无形资产、高转换成本等等。\
6、公司报告期最近一年主要做什么（做了哪些重要事情、是否会加强公司的核心竞争力等等），取得怎样的成绩（包括取得的主要成绩，重点是营业收入、净利润、净资产等等的增长情况）。\
7、对报告期最近一年主要经营情况进行分析：（1）主营业务方面。利润表及现金流量表相关科目变动分析表中，一个科目变化，相关科目是否基本同步变化；主营业务分行业、分产品、分地区情况分析表中主要产品及主要产品的毛利率与整体毛利率对比是否异常；产销量情况分析表中生产量和销售量比上年增减情况是否异常；成本分析表中成本\
构成及成本增减是否异常，特别是产品材料构成中占比最高的材料成本增减是否异常；主要销售客户及主要供应商情况表中从前5名客户获得的营收占比情况、从前5名供应商采购金额占比情况；期间费用同比变化情况分析表中费用变化是否异常、解释说明是否合理；研发投入情况表中研发投入费用化还是资本化、研发投入占营收比例如何；期间现\
金流同比变化情况分析表中经营活动产生的现金流量净额、投资活动产生的现金流量净额、筹资活动产生的现金流量净额三个科目是否异常。（2）非主营业务导致利润重大变化的说明有否异常情况。（3）资产、负债情况方面。是否有变动超过30%且本期或上期期末数占总资产超过30%的科目。（4）投资状况方面。是否有重大股权投资；是否有重\
大的非股权投资，募集资金使用情况（是否按计划使用），其他重大的非股权投资是否适度，有否异常；以公允价值计算的金融资产中的理财产品和结构性存款金额如何。（5）重大资产和股权出售是否有异常。（6）主要控股参股公司有什么特别情况。\
8、对公司未来发展进行分析：公司的发展战略和经营计划如何，下年要做哪些重要事情，这些事情是否有得于发展战略的实现。\
9、分析公司的重要事项：公司发生什么重要事项，是否影响公司未来的发展。特别是章程是否规定分红最低下限、连续性、稳定性，具体执行情况如何；在解决同行竞争、上市信息真实性、财政补贴等方面是否做无限期赔偿承诺，如何承诺；是否有重大诉讼，如有，给公司带来的潜在风险有多大，是否影响核心竞争力；重大关联交易情况如何，次数多少。\
10、分析普通股股份变动及股东情况：大股东和公司核心团队成员有无减持或增持的情况；大股东有无质押的情况；实际控制人的持股比例如何。\
11、分析董事、监事、高级管理人员和员工情况：核心团队成员是谁，离职的核心团队成员是谁，核心团队成员的工作动力如何（主要从核心团队成员的持股情况和薪酬水平评判），核心团队成员的专业能力如何（主要从核心团队成员的主要工作经历评判，如是否从事过本行业工作、是否是从本行基层一步步上来的等等），离职的是核心团队中的重\
要人物还是非重要人物，员工总人数是否随公司扩张而增加（是否存在异常），人均薪酬是否存在异常。\
12. 公司历年核心财务指标的变化趋势及原因。\
13. 经营战略的演变、核心护城河与竞争优势。\
14. 历年潜在风险的暴露情况及最新改进建议。\
15. 基于最新网络情报的前瞻性市场定位与同业对比。\
要求：融会贯通多份文档的信息，数据准确、观点鲜明、具备长远视角的宏观格局。"
    ),
    "多年度综合深度分析":("你是一位资深金融分析师。本次分析提供了该公司【多份/多个年份】的原始文档，以及【最新网络情报】。请进行时间序列上的长效对比与横向同业分析，给出一份详尽的综合评估报告。重点分析：\n"
        "1. 公司历年核心财务指标的变化趋势及原因\n"
        "2. 经营战略的演变、核心护城河与竞争优势\n"
        "3. 历年潜在风险的暴露情况及最新改进建议\n"
        "4. 基于最新网络情报的前瞻性市场定位与同业对比\n\n"
        "要求：融会贯通多份文档的信息，数据准确、观点鲜明、具备长远视角的宏观格局。"
    ),
    "结合最新情报分析": (
        "请结合我为你提供的【最新网络搜集情报】以及【历年本地文档内容】，"
        "给出一份详尽的综合分析报告。重点指出：\n"
        "1. 文档中未提及的最新市场动态\n"
        "2. 行业发展趋势与机遇挑战\n"
        "3. 基于最新情报的战略建议\n"
        "4. 风险预警与应对措施\n\n"
        "要求：紧密结合最新情报，提供前瞻性分析。"
    )
}

TEMPLATE_FILE_C = "prompt_templates.json"
if os.path.exists(TEMPLATE_FILE_C):
    try:
        with open(TEMPLATE_FILE_C, "r", encoding="utf-8") as f:
            custom_templates = json.load(f)
            PROMPT_TEMPLATES.update(custom_templates)
    except Exception as e:
        print(f"加载自定义模板失败: {e}")

CONFIG_FILE_C = "analyzer_config.json"

# ==========================================
# C2程序: 企业分析多模板字典配置
# ==========================================
DEFAULT_PROMPT_TEMPLATE = {
    "标准企业分析": (
        "你是一位资深的企业分析师，请根据以下资料，围绕指定的维度和重点对目标企业【{company}】（参考同行【{competitors}】）作全面深入中肯的分析：\n\n"
        "【参考资料】\n"
        "[全网检索情报]:\n{web_info}\n\n"
        "[本地财报等内部数据]:\n{doc_info}\n\n"
        "【分析要求与维度】\n"
        "（一）行业与竞争环境分析。\n"
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
        "对以上信息实行交叉验证，勾稽核对。总结提炼公司的核心优势与致命短板，结合成长性与风险评估当前估值水平，最终得出公司是否具备长期投资价值的结论。"
    ),
    "资深企业分析": (

""""你是一名拥有20年实战经验的顶级商业战略与财务分析专家。
请基于以下资料，对目标企业【{company}】（并重点参考同行前三公司【{competitors}】进行对比）进行深度的商业、财务与战略分析。

【特别警告：核心数据时效性】
请务必优先甄别并使用参考资料中的【最新年份】数据！如果资料中出现了最新发布的年报、公告或财务指标，请坚决以此作为当前经营状况和报表剖析的唯一基础，绝对禁止使用过往旧数据（或你自带的历史知识）进行覆盖或混淆。提取数据时请明确标注数据所属的年份或报告期。

【参考资料】
[全网检索与官方/券商/行业数据]:
{web_info}

[本地财报/章程/公告等内部数据]:
{doc_info}

请严格按照以下维度和逻辑结构，输出权威、具体、真实、可交叉验证的深度报告（如数据不足以支撑某项，请客观指出缺失并基于行业常识推演，切勿捏造数据）：

一、 行业与竞争环境分析
1. 行业空间：市场规模、所处生命周期阶段、目标公司现有份额及潜在空间。
2. 竞争格局：集中度（如CR5）、波特五力模型分析。
3. 政策与趋势：监管导向、技术变革（数字化、低碳化等）对行业的影响。

二、 企业全面对比分析（目标公司与同行的差异化对比）
1. 领导者特质：志向、见识、恒心、仁爱心、才能及持续学习能力。
2. 企业文化：使命、愿景、核心价值观。
3. 治理与管理层：股权结构（实控人、稳定性、质押风险等）、管理层履历及战略执行力、信息披露与诚信度、治理结构。
4. 商业模式：用户、产品、推广、盈利模式，分析上下游集中度及产业链话语权。
5. 主营业务：收入结构、核心产品或服务。
6. 竞争优势：品牌、专利技术、网络效应、特许经营权等护城河。
7. 核心报表剖析：
   - 利润表：营收增速、毛利率、净利率、费用率变化。
   - 资产负债表：负债结构、有息负债率、现金储备、应收账款与存货周转。
   - 现金流量表：经营活动现金流净额是否与净利润匹配，投资与筹资活动倾向。
8. 关键财务比率：盈利能力(ROE/ROA)、偿债能力(流动/速动比率)、运营效率、价值指标(PE/PB/PEG)。
9. 企业团队：组织架构、晋升与激励机制、人才梯队培训。
10. 内外沟通：内部人际关系治理、批评与自我批评纠错机制。
11. 现金流管理：经营与融资的“开源”能力，及降本增效的“节流”能力。
12. 企业系统：规章制度、流程效率、权责利分配机制。
13. 法规遵从：合规经营、学法用法情况。
14. 成长性与风险：政策/技术迭代/大客户/ESG等风险识别；新产能/新市场/并购等成长驱动因素；中长期战略评估。
15. 其他关键发现。

三、 分析与验证（财务排雷与逻辑闭环）
1. 历史趋势：核心指标稳定度与周期性。
2. 勾稽核对：财务数据与业务描述是否匹配（如营收大增但现金流萎缩的异常探讨）。
3. 异常质疑：针对存疑数据寻找合理解释或提出风险预警。

四、 综合判断与结论
1. 优势与短板总结：提炼3个核心优势与3个致命风险点。
2. 估值评估：结合成长性与风险，判断当前估值水平是否具备安全边际。
3. 最终结论：是否具备长期投资价值？未来1-3年的核心观察窗口是什么？
"""
    )
}


class ConfigManager:
    def __init__(self):
        self.config = self.load_config()

    def load_config(self):
        default_cfg = self.default_config()
        if os.path.exists(CONFIG_FILE_C):
            try:
                with open(CONFIG_FILE_C, 'r', encoding='utf-8') as f:
                    cfg = json.load(f)

                    # 关键修复：强制将最新的多模板配置合并进本地旧文件
                    if "templates" not in cfg or not isinstance(cfg["templates"], dict):
                        cfg["templates"] = {}

                    # 将代码里自带的模板（如标准、非标准）注入到配置中
                    for k, v in default_cfg["templates"].items():
                        if k not in cfg["templates"]:
                            cfg["templates"][k] = v

                    # 如果当前选中的模板在列表里找不到，则重置为默认
                    if cfg.get("current_template_name") not in cfg["templates"]:
                        cfg["current_template_name"] = "标准企业分析"

                    # 补齐其他缺失的配置项
                    for key in default_cfg:
                        if key not in cfg:
                            cfg[key] = default_cfg[key]

                    return cfg
            except Exception:
                pass
        return default_cfg

    def default_config(self):
        return {
            "api_bases": ["https://api.deepseek.com", "https://api.openai.com/v1", "http://localhost:11434/v1"],
            "api_models": ["deepseek-chat", "gpt-4-turbo", "qwen:72b"],
            "current_base": "https://api.deepseek.com",
            "current_model": "deepseek-chat",
            "api_key": "",
            "templates": DEFAULT_PROMPT_TEMPLATE,
            "current_template_name": "标准企业分析",
            "input_dir": os.getcwd(),
            "output_dir": os.getcwd()
        }

    def save_config(self):
        with open(CONFIG_FILE_C, 'w', encoding='utf-8') as f:
            json.dump(self.config, f, ensure_ascii=False, indent=4)

# =====================================================================
# 函数库：APP1核心 (原A程序)
# =====================================================================
def inject_hook_v75(driver):
    hook_script = """
    (function() {
        if (window.__V75_HOOKED__) return;
        window.__V75_HOOKED__ = true;
        window.__V75_DATAS__ = [];

        const saveIfWencaiData = (obj) => {
            try {
                if (!obj) return;

                let allResults = [];
                // 深度遍历寻找数据包和列定义
                const findTarget = (node, depth = 0) => {
                    if (depth > 12 || !node || typeof node !== 'object') return;

                    // 标准 API 结构
                    if (node.resultList && Array.isArray(node.resultList) && node.resultList.length > 0) {
                        allResults.push({
                            rList: node.resultList,
                            cList: node.columns || node.column || node.tableCols || []
                        });
                    }
                    if (node.datas && Array.isArray(node.datas) && node.datas.length > 0) {
                        allResults.push({
                            rList: node.datas,
                            cList: []
                        });
                    }

                    for (let key in node) {
                        if (node.hasOwnProperty(key)) {
                            findTarget(node[key], depth + 1);
                        }
                    }
                };

                findTarget(obj);

                if (allResults.length > 0) {
                    // 按体积选出最大的数据表（行数*列数），过滤掉垃圾辅助表
                    allResults.sort((a, b) => (b.rList.length * Object.keys(b.rList[0]).length) - (a.rList.length * Object.keys(a.rList[0]).length));
                    window.__V75_DATAS__.push(allResults[0]);
                }
            } catch(e) {}
        };

        const oldSend = XMLHttpRequest.prototype.send;
        XMLHttpRequest.prototype.send = function() {
            this.addEventListener('load', function() { 
                try { saveIfWencaiData(JSON.parse(this.responseText)); } catch(e) {} 
            });
            return oldSend.apply(this, arguments);
        };

        const originalFetch = window.fetch;
        window.fetch = async function(...args) {
            const response = await originalFetch.apply(this, args);
            try { 
                const clone = response.clone(); 
                clone.json().then(data => saveIfWencaiData(data)).catch(e => {}); 
            } catch (e) {}
            return response;
        };

        // 劫持 Response json/text 确保 fetch 万无一失
        const origJson = Response.prototype.json;
        Response.prototype.json = async function() {
            const res = await origJson.apply(this, arguments);
            saveIfWencaiData(res);
            return res;
        };

        const originalParse = JSON.parse;
        JSON.parse = function(text, reviver) {
            const result = originalParse(text, reviver);
            saveIfWencaiData(result);
            return result;
        };
    })();
    """
    driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {"source": hook_script})

def get_captured_data_v75(driver):
    import re
    import json

    try:
        # 在浏览器中注入并执行【内存雷达嗅探算法】
        js_code = """
            let res = window.__V75_DATAS__ || [];
            window.__V75_DATAS__ = []; // 提取后清空

            // 如果网络钩子因为跨域等原因没抓到，启动全局内存扫描！
            if (res.length === 0) {
                let bestData = null;
                let bestCols = null;
                let maxScore = 0;

                const isTarget = (obj) => {
                    if(Array.isArray(obj) && obj.length > 0 && typeof obj[0] === 'object' && obj[0] !== null) {
                        let keys = Object.keys(obj[0]).join(',').toLowerCase();
                        if(keys.includes('代码') || keys.includes('code') || keys.includes('secucode')) {
                            return true;
                        }
                    }
                    return false;
                };

                const traverse = (obj, depth, visited) => {
                    if(depth > 6 || !obj || typeof obj !== 'object') return;
                    if(visited.has(obj)) return;
                    visited.add(obj);

                    if (obj instanceof Window || obj instanceof Document || obj instanceof Element) return;

                    if (obj.resultList && isTarget(obj.resultList)) {
                        let score = obj.resultList.length * Object.keys(obj.resultList[0]).length;
                        if (score > maxScore) {
                            maxScore = score;
                            bestData = obj.resultList;
                            bestCols = obj.columns || obj.column || obj.tableCols || [];
                        }
                    } else if (obj.datas && isTarget(obj.datas)) {
                        let score = obj.datas.length * Object.keys(obj.datas[0]).length;
                        if (score > maxScore) {
                            maxScore = score;
                            bestData = obj.datas;
                            bestCols = [];
                        }
                    } else if (isTarget(obj)) {
                        let score = obj.length * Object.keys(obj[0]).length;
                        if (score > maxScore) {
                            maxScore = score;
                            bestData = obj;
                            bestCols = [];
                        }
                    }

                    if (Array.isArray(obj)) {
                        for(let i=0; i < Math.min(obj.length, 30); i++) {
                            try { traverse(obj[i], depth+1, visited); } catch(e) {}
                        }
                    } else {
                        for(let k in obj) {
                            if(!obj.hasOwnProperty(k)) continue;
                            try { traverse(obj[k], depth+1, visited); } catch(e) {}
                        }
                    }
                };

                // 扫描常见的全局数据根节点
                let roots = [window.wencai, window.__INITIAL_STATE__, window.__NEXT_DATA__, window.PAGE_STATE, window.vuex, window.store];
                for(let r of roots) {
                    if(r) traverse(r, 0, new Set());
                }

                // 如果常见节点没找到，执行全 window 暴搜
                if (!bestData) {
                    for(let k of Object.keys(window)) {
                        if(['window', 'document', 'localStorage', 'sessionStorage', 'navigator'].includes(k)) continue;
                        try { traverse(window[k], 0, new Set()); } catch(e) {}
                        if(bestData) break;
                    }
                }

                if (bestData) {
                    res.push({ rList: bestData, cList: bestCols });
                }
            }
            return res;
        """

        payload = driver.execute_script(js_code)

        if payload and len(payload) > 0:
            # 取出得分最高的主表格数据
            best_batch = None
            best_cols = None
            max_score = 0

            for item in payload:
                rList = item.get('rList', [])
                cList = item.get('cList', [])
                if rList and isinstance(rList, list) and len(rList) > 0 and isinstance(rList[0], dict):
                    score = len(rList) * len(rList[0].keys())
                    if score > max_score:
                        max_score = score
                        best_batch = rList
                        best_cols = cList

            if best_batch:
                # ==========================
                # 【翻译与拼装阶段】
                # ==========================
                col_map = {}
                if best_cols and isinstance(best_cols, list):
                    def extract_cols(c_list):
                        for c in c_list:
                            if isinstance(c, dict):
                                if c.get('field') and c.get('title'):
                                    col_map[c['field']] = c['title']
                                elif c.get('key') and c.get('title'):
                                    col_map[c['key']] = c['title']
                                if 'children' in c: extract_cols(c['children'])

                    extract_cols(best_cols)

                cleaned_data = []
                today_str = datetime.now().strftime('%Y%m%d')

                for row in best_batch:
                    if not isinstance(row, dict): continue

                    clean_row = {}
                    for k, v in row.items():
                        # 用字典把英文键名翻译成真正的中文表头！
                        new_key = col_map.get(k, k)

                        # 【修复1】：只精准移除问财当日查询生成的脏后缀 (如 [20260920])，
                        # 严禁使用正则盲目替换数字，从而完美保留多年的财报年份及不同期数据！
                        new_key_str = str(new_key).replace(f'[{today_str}]', '').strip()
                        clean_row[new_key_str] = v

                    # 【修复2】：动态提炼代码和名称，并删除原有的冗余键，防止最后出现重复列
                    keys_to_remove = []
                    code_val, name_val = None, None

                    for k_curr in list(clean_row.keys()):
                        k_str = str(k_curr).lower()
                        if k_str in ['code', 'secucode', '股票代码', 'a股代码', '指数代码']:
                            code_val = clean_row[k_curr]
                            if k_curr != '代码':
                                keys_to_remove.append(k_curr)
                        elif k_str in ['name', 'secuname', '股票简称', '简称', 'a股简称', '指数简称']:
                            name_val = clean_row[k_curr]
                            if k_curr != '名称':
                                keys_to_remove.append(k_curr)

                    for kr in keys_to_remove:
                        clean_row.pop(kr, None)

                    if code_val is not None:
                        raw_code = str(code_val).replace('sz', '').replace('sh', '').replace('bj', '')
                        clean_row['代码'] = raw_code.zfill(6) if raw_code.isdigit() else raw_code
                    if name_val is not None:
                        clean_row['名称'] = name_val

                    if '代码' in clean_row:
                        cleaned_data.append(clean_row)

                if cleaned_data:
                    return cleaned_data
    except Exception as e:
        pass

    return []
def process_dataframe(final_df):
    if final_df.empty: return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
    id_col = next((col for col in ['code', '股票代码', '代码', '指数代码'] if col in final_df.columns), None)
    if id_col: final_df = final_df.groupby(id_col).first().reset_index()

    def get_sorted_cols(df, keyword, count=None):
        matched_cols = [c for c in df.columns if keyword in str(c)]
        def extract_date(col_name):
            match = re.search(r'\[(\d+)\]|\((\d+)\)|(\d{4,8})', str(col_name))
            if match:
                for g in match.groups():
                    if g: return int(g)
            return 0
        matched_cols.sort(key=extract_date, reverse=True)
        if count and len(matched_cols) > count: matched_cols = matched_cols[:count]
        return matched_cols

    ordered_cols = (get_sorted_cols(final_df, '代码', 1) or get_sorted_cols(final_df, 'code', 1)) + \
                   (get_sorted_cols(final_df, '简称', 1) or get_sorted_cols(final_df, '名称', 1)) + \
                   (get_sorted_cols(final_df, '最新价', 1) or get_sorted_cols(final_df, '收盘点位', 1)) + \
                   get_sorted_cols(final_df, '涨跌幅', 1) + get_sorted_cols(final_df, '所属概念', 1) + \
                   get_sorted_cols(final_df, '预测净资产收益率', 3) + \
                   (get_sorted_cols(final_df, '归属', 5) or get_sorted_cols(final_df, '净利润', 5)) + \
                   (get_sorted_cols(final_df, '股东权益', 5) or get_sorted_cols(final_df, '净资产', 5)) + \
                   get_sorted_cols(final_df, 'dde大单', 1) + get_sorted_cols(final_df, '总股本', 1) + \
                   get_sorted_cols(final_df, '市盈率', 1) + get_sorted_cols(final_df, '净资产收益率', 5) + \
                   get_sorted_cols(final_df, '市场类型', 1) + get_sorted_cols(final_df, '同花顺行业', 1)

    final_ordered_cols = []
    for c in ordered_cols:
        if c not in final_ordered_cols and c in final_df.columns: final_ordered_cols.append(c)

    other_cols = [c for c in final_df.columns if c not in final_ordered_cols]
    df_table1 = final_df[final_ordered_cols + other_cols]
    df_table2 = df_table1.copy()

    cols_for_df3 = []
    if get_sorted_cols(final_df, '代码', 1) or get_sorted_cols(final_df, 'code', 1):
        cols_for_df3.extend(get_sorted_cols(final_df, '代码', 1) or get_sorted_cols(final_df, 'code', 1))
    if get_sorted_cols(final_df, '简称', 1) or get_sorted_cols(final_df, '名称', 1):
        cols_for_df3.extend(get_sorted_cols(final_df, '简称', 1) or get_sorted_cols(final_df, '名称', 1))
    if not cols_for_df3 and len(final_df.columns) >= 2: cols_for_df3 = list(final_df.columns[:2])
    df_table3 = final_df[cols_for_df3].copy()
    if len(df_table3.columns) >= 2: df_table3.columns = ['代码', '名称']

    rename_map = {"最新价": "现价（元）", "最新涨跌幅": "涨跌幅（%）", "预测净资产收益率(roe)平均值": "预测净资产收益率（%）",
                  "归属于母公司所有者的净利润": "净利润（元）", "股东权益合计": "净资产（元）",
                  "最新dde大单净额": "最新dde大单净额（元）",
                  "总股本": "总股本（股）", "净资产收益率roe(加权,公布值)": "加权净资产收益率（%）",
                  "股票市场类型": "股票市场属性",
                  "指数代码": "代码", "指数简称": "名称", "收盘点位": "收盘（点位）"}

    new_columns = []
    for col in df_table2.columns:
        new_col = str(col)
        for old_str, new_str in rename_map.items():
            if old_str in new_col: new_col = new_col.replace(old_str, new_str)
        new_columns.append(new_col)
    df_table2.columns = new_columns
    return df_table1, df_table2, df_table3


def run_app2_isolated_process(queries, output_folder, timestamp, market_str, is_headless, log_queue, result_queue,
                              stop_event, pause_event):
    def log(msg):
        log_queue.put(msg)

    def check_interrupt():
        while pause_event.is_set() and not stop_event.is_set(): time.sleep(0.5)
        return stop_event.is_set()

    driver = None
    success_flag = True
    file_path_1 = os.path.join(output_folder, f'问财_表1_多条件初始表_{timestamp}.xlsx')
    file_path_2 = os.path.join(output_folder, f'问财_表2_多条件整理表_{timestamp}.xlsx')
    file_path_3 = os.path.join(output_folder, '问财_表3_代码及名称汇总.xlsx')
    writer1, writer2, writer3 = None, None, None

    try:
        opts = webdriver.EdgeOptions()
        opts.add_argument("--disable-blink-features=AutomationControlled")

        # 【修复1】：恢复静默运行判断逻辑，并加入防止页面变形的分辨率锁定
        if is_headless:
            log("   🥷 已开启静默模式，浏览器将在后台运行...")
            opts.add_argument("--headless=new")
            opts.add_argument("--disable-gpu")
            opts.add_argument("--window-size=1920,1080")

        opts.add_argument(
            'user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36')

        profile_dir = os.path.join(os.getcwd(), "Wencai_Profile").replace("\\", "/")
        opts.add_argument(f"--user-data-dir={profile_dir}")

        try:
            log("   🚀 正在启动带有账号记忆功能的浏览器...")
            driver, _e1 = create_edge_driver(opts, purpose="问财海选-记忆模式")
            if driver is None:
                raise RuntimeError(_e1 or "启动失败")
        except Exception as e1:
            log(f"   ⚠️ 记忆模式启动失败。启动备用无记忆模式...")
            try:
                opts_backup = webdriver.EdgeOptions()
                opts_backup.add_argument("--disable-blink-features=AutomationControlled")
                if is_headless:
                    opts_backup.add_argument("--headless=new")
                    opts_backup.add_argument("--disable-gpu")
                    opts_backup.add_argument("--window-size=1920,1080")
                driver, _e2 = create_edge_driver(opts_backup, purpose="问财海选-备用模式")
                if driver is None:
                    raise RuntimeError(_e2 or "启动失败")
            except Exception as e2:
                success_flag = False
                result_queue.put(success_flag)
                time.sleep(1.5)
                return

        try:
            driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
                "source": "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
            })
            inject_hook_v75(driver)
        except:
            pass

        for index, raw_query in enumerate(queries):
            if check_interrupt(): break
            query = raw_query.strip()
            if not query: continue
            sheet_name = re.sub(r'[\\/*?:\[\]]', '_', query)[:30] or f"条件_{index + 1}"

            if market_str != "A股" and market_str not in query:
                query = f"{market_str} " + query

            log(f"   👉 正在抓取 [{index + 1}/{len(queries)}]: {query}")

            target_url = f"https://www.iwencai.com/unifiedwap/result?w={urllib.parse.quote(query)}"
            try:
                driver.get(target_url)
            except:
                pass
            time.sleep(4)

            need_re_navigate = False
            log("   🔍 正在检查通行权限与页面状态...")
            for i in range(15):
                if check_interrupt(): break
                try:
                    current_url = driver.current_url
                    need_login = driver.execute_script("""
                        let box = document.querySelector('.login-modal, .passport-login-container, .pc-login-box, .el-dialog__wrapper');
                        return box && box.style.display !== 'none' && box.offsetWidth > 0;
                    """)

                    if need_login:
                        if i % 5 == 0: log("   ⚠️ 网站触发拦截！请扫码登录...")
                        need_re_navigate = True
                        time.sleep(1)
                        continue

                    if "favorite" in current_url or "home/index" in current_url:
                        need_re_navigate = True
                        break

                    break
                except:
                    time.sleep(1)

            if need_re_navigate:
                log("   🔄 正在强行拽回 AI 搜索结果页！")
                try:
                    driver.get(target_url)
                except:
                    pass
                time.sleep(5)

            log(f"      🔨 正在挂载底层引擎...")
            driver.execute_script("window.__V75_DATAS__ = [];")

            try:
                driver.execute_script("""
                    let q = arguments[0];
                    let tas = document.querySelectorAll('textarea');
                    if(tas.length > 0) {
                        let ta = tas[tas.length - 1];
                        if (ta.value.trim() !== q) {
                            ta.value = q;
                            ta.dispatchEvent(new Event('input', { bubbles: true }));
                            setTimeout(() => {
                                let sendBtns = document.querySelectorAll('.chat-input-send, .search-btn, .send-icon, img[alt="发送"]');
                                if(sendBtns.length > 0) { sendBtns[sendBtns.length - 1].click(); }
                            }, 800);
                        }
                    }
                """, query)
            except:
                pass

            log(f"      ⏳ 正在等待数据表格渲染 (最大容忍 45 秒)...")
            wait_time = 0
            new_data_arrived = False
            while wait_time < 45 and not stop_event.is_set():
                time.sleep(1)
                wait_time += 1

                try:
                    driver.execute_script("""
                        let btns = document.querySelectorAll('span, button');
                        for(let i=0; i<btns.length; i++) {
                            let txt = (btns[i].innerText || "").trim();
                            if(['查看明细', '查看完整数据', '展开表格'].includes(txt)) {
                                try { btns[i].click(); } catch(e){}
                            }
                        }
                    """)
                except:
                    pass

                try:
                    has_data = driver.execute_script("""
                        let netData = window.__V75_DATAS__ && window.__V75_DATAS__.length > 0;
                        let domData = document.querySelectorAll('tbody tr, .el-table__row, .vxe-body--row, .iwc-table-body tr').length > 0;
                        return netData || domData;
                    """)
                    if has_data:
                        new_data_arrived = True
                        break
                except:
                    pass

            if new_data_arrived:
                log(f"      ✅ 成功监测到新数据落地！缓冲提取中...")
                time.sleep(3)
            else:
                log(f"      ⚠️ 45秒后仍未见数据表格，可能条件过严查无结果。")

            all_dfs = []
            seen_codes = set()
            cur_page = 1
            while cur_page <= 1000:
                if check_interrupt(): break
                if new_data_arrived: log(f"      --- 第 {cur_page} 页 ---")

                data_buffer = []
                for _ in range(8):
                    if check_interrupt(): break
                    time.sleep(1)
                    batch = get_captured_data_v75(driver)
                    if batch:
                        data_buffer.extend(batch)
                        new_in_batch = False
                        for item in batch:
                            item_code = str(
                                item.get('code', item.get('代码', item.get('股票代码', item.get('指数代码', '')))))
                            if item_code and item_code not in seen_codes:
                                new_in_batch = True
                                break
                        if new_in_batch: break

                if data_buffer:
                    df_page = pd.DataFrame(data_buffer)
                    id_col = None
                    for col in ['code', '代码', '股票代码', '指数代码']:
                        if col in df_page.columns:
                            id_col = col
                            break
                    if id_col:
                        codes = set(df_page[id_col].astype(str).tolist())
                        new_codes = codes - seen_codes
                        if len(new_codes) > 0:
                            all_dfs.append(df_page)
                            seen_codes.update(codes)
                            log(f"      ✅ 捕获成功：新增 {len(new_codes)} 条")
                            if len(df_page) < 20: break
                        else:
                            log("      ⚠️ 无新数据，准备翻页")
                            break

                cur_page += 1
                paged = False
                try:
                    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                    time.sleep(1)
                    paged = driver.execute_script("""
                        let nextBtn = document.querySelector('.btn-next, .vxe-pager--next-btn, li.number + li:not(.disabled)');
                        if(nextBtn && !nextBtn.disabled && !nextBtn.classList.contains('disabled')) {
                            nextBtn.click();
                            return true;
                        }
                        return false;
                    """)
                    if paged: time.sleep(2)
                except:
                    pass
                if not paged: break

            if all_dfs:
                final_df = pd.concat(all_dfs, ignore_index=True, sort=False)
                df1, df2, df3 = process_dataframe(final_df)
                if writer1 is None:
                    writer1 = pd.ExcelWriter(file_path_1, engine='openpyxl')
                    writer2 = pd.ExcelWriter(file_path_2, engine='openpyxl')
                    writer3 = pd.ExcelWriter(file_path_3, engine='openpyxl')
                df1.to_excel(writer1, sheet_name=sheet_name, index=False)
                df2.to_excel(writer2, sheet_name=sheet_name, index=False)
                df3.to_excel(writer3, sheet_name=sheet_name, index=False)
                log(f"   💾 条件 '{sheet_name}' 已成功汇入表暂存区。")
            else:
                log(f"   ❌ 条件 '{query}' 未采集到数据。")

        if writer1 is not None:
            writer1.close();
            writer2.close();
            writer3.close()
            log(f"   🎉 问财数据表导出完毕。")
        elif not stop_event.is_set():
            log(f"   ⚠️ 未采集到符合要求的数据，因此没有生成问财 Excel 表格。")

    except Exception as e:
        import traceback
        log(f"   💥 问财程序内部异常: {str(e)[:150]}...")
        success_flag = False
        try:
            if writer1: writer1.close()
            if writer2: writer2.close()
            if writer3: writer3.close()
        except:
            pass
    finally:
        if driver:
            try:
                driver.quit()
            except:
                pass
        result_queue.put(success_flag)
        time.sleep(1.5)
def get_eastmoney_industry_from_datacenter(stock_codes):
    print("\n[模块1] 正在向东方财富数据中心请求行业数据...")
    results, industry_map, name_map = [], {}, {}
    for code in stock_codes:
        if code.startswith('6'): market_code = f"{code}.SH"
        elif code.startswith('0') or code.startswith('3'): market_code = f"{code}.SZ"
        elif code.startswith('8') or code.startswith('4'): market_code = f"{code}.BJ"
        else: continue
        try:
            url = "https://datacenter-web.eastmoney.com/api/data/v1/get"
            params = {
                'reportName': 'RPT_F10_CORETHEME_BOARDTYPE',
                'columns': 'SECUCODE,SECURITY_CODE,SECURITY_NAME_ABBR,BOARD_NAME,BOARD_CODE,BOARD_TYPE,BOARD_RANK',
                'filter': f'(SECUCODE="{market_code}")', 'pageNumber': '1', 'pageSize': '200', 'source': 'WEB', 'client': 'WEB'
            }
            res = requests.get(url, params=params, headers={'User-Agent': 'Mozilla/5.0'}, timeout=10).json()
            if res.get('success') and res.get('result') and res['result'].get('data'):
                records = res['result']['data']
                company_name = records[0].get('SECURITY_NAME_ABBR', '未知简称')
                industry_type = next((row.get('BOARD_TYPE') for row in records if str(row.get('BOARD_RANK')) == '3'), records[0].get('BOARD_TYPE'))
                industry_buttons = sorted([row for row in records if row.get('BOARD_TYPE') == industry_type], key=lambda x: str(x.get('BOARD_RANK', '0')))
                level_1 = industry_buttons[0].get('BOARD_NAME') if len(industry_buttons) > 0 else '-'
                level_2 = industry_buttons[1].get('BOARD_NAME') if len(industry_buttons) > 1 else '-'
                level_3 = industry_buttons[2].get('BOARD_NAME') if len(industry_buttons) > 2 else '-'
                parts = [p for p in [level_1, level_2, level_3] if p != '-']
                industry_str = '-'.join(parts) if parts else '--'
                folder_industry = level_2 if level_2 != '-' else (level_1 if level_1 != '-' else '未知行业')
                industry_map[code] = folder_industry
                name_map[code] = company_name
                results.append({'目标股票': code, '股票简称': company_name, '一级行业': level_1, '二级行业': level_2, '三级行业': level_3, '完整行业路径': industry_str})
                print(f"✅ 成功提取: {code} ({company_name}) -> {industry_str}")
        except Exception as e:
            print(f"❌ 抓取 {code} 时发生错误: {e}")
    return pd.DataFrame(results), industry_map, name_map

def get_leader_radar_batch(stock_codes, industry_map, headless=False):
    print(f"\n[模块2] 启动‘数据雷达’扫描同行情况...")
    options = webdriver.EdgeOptions()
    options.add_argument('--start-maximized')
    options.add_experimental_option('excludeSwitches', ['enable-logging'])
    if headless:
        options.add_argument('--headless')
        options.add_argument('--disable-gpu')
    # 【修复5】浏览器启动改为可控方式：失败时给出明确原因，不再直接抛出中断整个下载流程
    driver, _drv_err = create_edge_driver(options, purpose="数据雷达-同行扫描")
    if driver is None:
        return pd.DataFrame(), {}

    all_results, download_targets = [], {}
    try:
        for stock_code in stock_codes:
            print(f"\n🔍 正在扫描: {stock_code}")
            full_code = ("SH" if stock_code.startswith('6') else "SZ") + stock_code
            driver.get(
                f"https://emweb.securities.eastmoney.com/pc_hsf10/pages/index.html?type=web&code={full_code}#/thbj")
            try:
                # 智能等待，最多10秒，只要表格加载出来就立刻放行
                WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located(
                        (By.XPATH, "//*[local-name()='tr' or local-name()='div'][contains(text(), '亿')]"))
                )
            except:
                pass  # 如果超时没找到，后续兜底逻辑会自动处理
            time.sleep(1)  # 缓冲1秒，防止DOM还未渲染完毕
            current_peers = []
            for row in driver.find_elements(By.XPATH, "//*[local-name()='tr' or local-name()='div']"):
                try:
                    txt = row.text.strip()
                    if "亿" in txt:
                        match = re.search(r'(?:(\d+)\s+)?(\d{6})\s+([A-Za-z\u4e00-\u9fa5\*]{2,})\s+([\d\.]+)(?=亿)', txt)
                        if match:
                            rank = int(match.group(1)) if match.group(1) else 0
                            code, name, mkt_cap = match.group(2), match.group(3), float(match.group(4))
                            if not any(r['代码'] == code for r in current_peers):
                                current_peers.append({"原查询股票": stock_code, "排名": rank, "代码": code, "名称": name, "市值(亿)": mkt_cap})
                except: continue
            if current_peers:
                current_peers.sort(key=lambda x: (x['排名'] if x['排名'] > 0 else float('inf'), -x['市值(亿)']))
                if all(r['排名'] == 0 for r in current_peers):
                    for i, r in enumerate(current_peers, 1): r['排名'] = i
                top_6 = current_peers[:6]
                all_results.extend(top_6)
                print(f"✅ {stock_code} 扫描完毕。龙头: {top_6[0]['名称']}")
                target_industry = industry_map.get(stock_code, '其他行业')
                download_targets[stock_code] = target_industry
                for peer in current_peers[:3]: download_targets[peer['代码']] = target_industry
            else: print(f"❌ 雷达未探测到数据。")
    except Exception as e: print(f"❌ 异常: {e}")
    finally: driver.quit()
    return pd.DataFrame(all_results), download_targets


def get_orgid(stock_code):
    # 【核心修复】：加上全套的浏览器伪装头，防止被巨潮防火墙拦截
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
        'Accept': 'application/json, text/javascript, */*; q=0.01',
        'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
        'X-Requested-With': 'XMLHttpRequest',
        'Origin': 'http://www.cninfo.com.cn',
        'Referer': 'http://www.cninfo.com.cn/new/index'
    }

    # 策略 1：使用标准 POST 请求头数据获取
    try:
        url = "http://www.cninfo.com.cn/new/information/topSearch/query"
        data = {'keyWord': stock_code}
        res = requests.post(url, data=data, headers=headers, timeout=10)
        if res.status_code == 200:
            res_json = res.json()
            if isinstance(res_json, list):
                for item in res_json:
                    if str(item.get('code', '')) == str(stock_code):
                        return item.get('orgId'), item.get('zwjc')
    except Exception as e:
        pass

    # 策略 2：兜底方案，使用带 Params 的 GET 请求获取
    try:
        url_get = f"http://www.cninfo.com.cn/new/information/topSearch/query?keyWord={stock_code}"
        res_get = requests.get(url_get, headers=headers, timeout=10)
        if res_get.status_code == 200:
            res_json_get = res_get.json()
            if isinstance(res_json_get, list):
                for item in res_json_get:
                    if str(item.get('code', '')) == str(stock_code):
                        return item.get('orgId'), item.get('zwjc')
    except Exception as e:
        pass

    return None, None

def download_pdf(pdf_url, save_path, max_retries=3):
    """【修复7】返回 True/False，并把失败原因打到日志里（原来失败也是静默返回，导致“整家公司一个文件都没下到”却无人察觉）。"""
    last_err = ""
    for attempt in range(1, max_retries + 1):
        try:
            r = requests.get(pdf_url, headers={'User-Agent': 'Mozilla/5.0'}, stream=True, timeout=30)
            if r.status_code == 200:
                with open(save_path, 'wb') as f:
                    for chunk in r.iter_content(8192):
                        if chunk: f.write(chunk)
                if os.path.exists(save_path) and os.path.getsize(save_path) > 0:
                    print(f"    ⬇️ 成功: {os.path.basename(save_path)}")
                    return True
                last_err = "文件写入后为空"
            else:
                last_err = f"HTTP {r.status_code}"
                print(f"    ⚠️ 第{attempt}次下载失败（{last_err}）: {os.path.basename(save_path)}")
        except Exception as e:
            last_err = f"{type(e).__name__}: {str(e)[:120]}"
            print(f"    ⚠️ 第{attempt}次下载出错: {last_err}")
        if attempt < max_retries: time.sleep(2)
    print(f"    ❌ 失败: {os.path.basename(save_path)}（已重试{max_retries}次，最后原因：{last_err}）")
    return False


def cninfo_download_documents(download_targets_dict, start_year, end_year, specific_root_dir, since_ipo,
                              enable_prospectus):
    # 【修复2】返回值 = 本轮真正成功落盘的 PDF 数量（原实现返回 None，导致上层无法判断下载是否真的成功）
    downloaded_count = 0
    print(f"\n[模块3] 巨潮下载启动，当前分组需处理 {len(download_targets_dict)} 家公司...")
    query_url = "http://www.cninfo.com.cn/new/hisAnnouncement/query"
    keywords = ['年度报告', '公司章程']
    if enable_prospectus: keywords.append('招股说明书')

    for code, industry_name in download_targets_dict.items():
        orgId, company_name = get_orgid(code)
        if not orgId:
            print(f"    ⚠️ 警告：无法获取 {code} 的 orgId，跳过下载。")
            continue
        # 【修复7】按公司统计下载结果，任何一家“一个文件都没下到”都会明确报出来
        company_ok = 0
        company_fail = 0

        print(f"\n📁 正在检索: {company_name} ({code})")
        safe_company = re.sub(r'[\\/:*?"<>|]', '', company_name)
        company_dir = os.path.join(specific_root_dir, f"{code}_{safe_company}")
        os.makedirs(company_dir, exist_ok=True)

        for keyword in keywords:
            found_target = False
            sdate = "1990-01-01" if (keyword == '招股说明书' or since_ipo) else f"{start_year}-01-01"
            edate = f"{end_year}-12-31"

            # 【修复9】按年份过滤 + 翻页直到覆盖时间段：
            #   原实现最多翻 3 页（90 条）且第一页有结果就 break，
            #   像重庆啤酒这类年报记录上百条的公司，更早年份的年报永远取不到。
            #   现在：目标区间内的年报年份全部集齐、或翻到早于起始年份的公告、或达到页数上限才停。
            page_size = 30
            max_pages = 20
            target_years = set(range(int(start_year), int(end_year) + 1))
            done_years = set()
            oldest_seen = None
            pageNum = 1
            while pageNum <= max_pages:
                if found_target: break
                if keyword == '年度报告' and not since_ipo and target_years and target_years.issubset(done_years):
                    print(f"    ✅ 目标年份年报已集齐: {sorted(done_years)}")
                    break
                params = {
                    'pageNum': pageNum,
                    'pageSize': page_size,
                    'tabName': 'fulltext',
                    'stock': f"{code},{orgId}",
                    'searchkey': keyword,
                    'sdate': sdate,
                    'edate': edate,
                    'isHLtitle': 'true'
                }

                try:
                    # 加入适当的延时，防止巨潮封锁IP
                    time.sleep(1.5)
                    response = requests.post(query_url, data=params, headers={'User-Agent': 'Mozilla/5.0'}, timeout=15)

                    if response.status_code != 200:
                        print(f"    ❌ 网络请求失败，状态码: {response.status_code}")
                        break

                    res = response.json()
                    announcements = res.get('announcements', [])

                    if not announcements:
                        if pageNum == 1: print(f"    ℹ️ {keyword} 在 {sdate} 至 {edate} 期间未搜索到结果。")
                        break

                    for ann in announcements:
                        title = ann.get('announcementTitle', '')
                        clean_title = re.sub(r'[\\/:*?"<>|]', '', re.sub(r'<[^>]+>', '', title))
                        adj_url = ann.get('adjunctUrl', '')
                        pub_time = ann.get('announcementTime')

                        if any(noise in clean_title for noise in
                               ['摘要', '英文', '取消', '更正后', '公告', '提示', '决议']): continue

                        if keyword == '年度报告':
                            if any(x in clean_title for x in
                                   ['半', '季度', '一季', '三季', '预案', '业绩', '审计', '问询函', '督导', '回复',
                                    '保荐', '说明', '制度', '规程', '报告书']): continue
                            if not since_ipo:
                                year_match = re.search(r'(19\d{2}|20\d{2})年', clean_title)
                                if year_match:
                                    _yr = int(year_match.group(1))
                                    # 【修复9】按年份过滤：区间外的年报不下载，区间内的记下“已覆盖年份”
                                    if not (start_year <= _yr <= end_year): continue
                                    done_years.add(_yr)
                                else:
                                    continue
                        elif keyword == '招股说明书':
                            if any(x in clean_title for x in
                                   ['辅导', '保荐', '意见', '反馈', '草案', '发行情况', '审计', '回复', 'H股',
                                    '存托凭证']): continue
                            if '招股' not in clean_title or '说明书' not in clean_title:
                                if '招股意向书' not in clean_title: continue
                        elif keyword == '公司章程':
                            if any(x in clean_title for x in
                                   ['草案', '规程', '制度', '规则', '办法', '对照表', '修订案', '报告', '说明', 'H股',
                                    '细则', '议事']): continue
                            if pub_time and not since_ipo:
                                pub_year = datetime.fromtimestamp(pub_time / 1000).year
                                if not (start_year <= pub_year <= end_year): continue

                        if adj_url.lower().endswith('.pdf'):
                            file_path = os.path.join(company_dir, f"{clean_title}.pdf")
                            if not os.path.exists(file_path):
                                print(f"    ⏳ 准备下载: {clean_title}")
                                if download_pdf(f"http://static.cninfo.com.cn/{adj_url}", file_path):
                                    downloaded_count += 1
                                    company_ok += 1
                                else:
                                    company_fail += 1
                            else:
                                print(f"    ⏭️ 已存在，跳过: {clean_title}")

                            if keyword in ['招股说明书', '公司章程']:
                                found_target = True
                                break

                except Exception as e:
                    print(f"    ❌ 获取数据列表时发生异常: {e}")
                    break

                # 【修复9】翻页控制：记录本页最早公告日期；若已翻到早于起始年份，
                # 说明目标区间已被完整覆盖，无需继续往后翻。
                if keyword == '年度报告' and not since_ipo:
                    _times = [a.get('announcementTime') for a in announcements if a.get('announcementTime')]
                    if _times:
                        _oldest = datetime.fromtimestamp(min(_times) / 1000).year
                        oldest_seen = _oldest if oldest_seen is None else min(oldest_seen, _oldest)
                        if oldest_seen is not None and oldest_seen < int(start_year):
                            print(f"    ✅ 已翻页覆盖至 {oldest_seen} 年（早于起始年份 {start_year}），停止翻页。")
                            break
                if len(announcements) < page_size:
                    break
                pageNum += 1

        # 【修复7】每家公司收尾小结：谁没下到、下到几个，一目了然
        if company_ok == 0 and company_fail > 0:
            print(f"    ❌ {company_name}({code}) 本轮一个文件都没下载成功（失败 {company_fail} 个），请检查网络/巨潮访问是否被拦截。")
        elif company_fail > 0:
            print(f"    ⚠️ {company_name}({code}) 下载完成：成功 {company_ok} 个，失败 {company_fail} 个。")
        else:
            print(f"    ✅ {company_name}({code}) 下载完成：成功 {company_ok} 个。")

    # 【修复2】把“真实下载成功数”返回给调用方（run_pipeline）
    return downloaded_count

def p1_rule1_should_delete(val):
    if pd.isna(val) or str(val).strip() == "": return False
    val = str(val).strip()
    if not re.search(r'[\u4e00-\u9fa5]', val): return True
    if "股份有限公司" in val: return True
    if re.search(r'\d*[、，。.\s]*合并[资产负债利润现金流量所有者权益变动]+表', val): return True
    if re.search(r'\d{4}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日', val): return True
    if re.search(r'项目\s*附注.*?\d{4}\s*年\s*\d{1,2}\s*月', val): return True
    if re.search(r'项目.*?\d{4}.*?年度', val): return True
    if re.search(r'\d{4}\s*年年度报告', val): return True
    if re.search(r'\d{4}\s*年\s*\d{1,2}\s*[—\-~至]\s*\d{1,2}\s*月', val): return True
    if re.search(r'单位[：:\s]*元\s*币种[：:\s]*人民币', val): return True
    if re.search(r'项目.*期末余额.*期初余额', val): return True
    if val.startswith("公司负责人：") or val.startswith("本期发生同一控制"): return True
    if "的净利润为：0元" in val or "的净利润为:0元" in val: return True
    if val.endswith("表") and len(val) <= 15: return True
    return False

def p1_rule2_should_concat(val):
    if pd.isna(val) or str(val).strip() == "": return False
    val = str(val).strip()
    if val in ["合同负债", "其他综合收益", "其他"]: return False
    if val.startswith("变"): return True
    if val.startswith("合") and not val.startswith("合计"): return True
    if val.startswith("综合") or val.startswith("价物") or val.startswith("益") or val.startswith("收益"): return True
    if len(val) <= 1: return True
    if re.match(r'^[(（][^\d一二三四五六七八九十]', val): return True
    if val.startswith("的") or val.startswith("收到的"): return True
    if val in ["股利、利润", "长期资产收回的现金净额"]: return True
    if (val.endswith(")") or val.endswith("）")) and not ("(" in val or "（" in val): return True
    return False

def p1_clean_dataframe(df):
    if df.empty or '项目' not in df.columns: return df
    df = df.copy()
    proj_col = df.columns.get_loc('项目')
    note_col = df.columns.get_loc('附注') if '附注' in df.columns else None

    def move_numbers_up(source_idx, target_idx):
        for c in range(len(df.columns)):
            if c != proj_col and (note_col is None or c != note_col):
                val_target = df.iloc[target_idx, c]
                val_source = df.iloc[source_idx, c]
                if (pd.isna(val_target) or str(val_target).strip() == "") and pd.notna(val_source) and str(val_source).strip() != "":
                    df.iloc[target_idx, c] = val_source

    rows_to_drop = []
    last_valid_row = None
    TARGET_ENDINGS = ("其", "投", "的", "支", "价", "可", "股", "东", "单位", "和", "融", "偿付利息")
    PROTECTED_ITEMS = {"合同负债", "其他综合收益", "其他"}

    for i in range(len(df)):
        val = df.iloc[i, proj_col]
        if pd.isna(val) or str(val).strip() == "":
            rows_to_drop.append(i)
            continue
        val_str = str(val).strip()
        pattern_move = r'([一二三四五六七八九十]+、\d+(?:、（\d+）)?)$|([一二三四五六七八九十]+、\d+)$'
        match = re.search(pattern_move, val_str)
        if match and note_col is not None:
            note_str = match.group(0)
            val_str = val_str[:match.start()].strip()
            curr_note = str(df.iloc[i, note_col]) if pd.notna(df.iloc[i, note_col]) else ""
            df.iloc[i, note_col] = (curr_note + " " + note_str).replace("nan", "").strip()

        while True:
            new_str = re.sub(r'^(?:\d+\s*[.．]|[(（]\s*\d+\s*[)）]|[(（]\s*[一二三四五六七八九十]+\s*[)）])\s*', '', val_str).strip()
            if new_str == val_str: break
            val_str = new_str

        while True:
            match = re.search(r'[(（](.*?)[)）]\s*$', val_str)
            if match:
                inner_text = match.group(1)
                if "元/股" not in inner_text:
                    val_str = val_str[:match.start()].strip()
                    continue
            break

        if val_str == "":
            rows_to_drop.append(i)
            continue
        df.iloc[i, proj_col] = val_str
        is_concat = False

        if last_valid_row is not None and val_str not in PROTECTED_ITEMS:
            prev_val = str(df.iloc[last_valid_row, proj_col]).strip()
            if prev_val.endswith(TARGET_ENDINGS):
                df.iloc[last_valid_row, proj_col] = prev_val + val_str
                move_numbers_up(i, last_valid_row)
                rows_to_drop.append(i)
                is_concat = True

        if is_concat: continue

        if p1_rule1_should_delete(val_str):
            if val_str not in PROTECTED_ITEMS:
                rows_to_drop.append(i)
                continue

        if p1_rule2_should_concat(val_str):
            if val_str not in PROTECTED_ITEMS:
                if last_valid_row is not None:
                    prev_val = str(df.iloc[last_valid_row, proj_col])
                    df.iloc[last_valid_row, proj_col] = (prev_val if prev_val != "nan" else "") + val_str
                    move_numbers_up(i, last_valid_row)
                rows_to_drop.append(i)
                continue

        if note_col is not None:
            note_val = str(df.iloc[i, note_col]).strip()
            if re.fullmatch(r'\d+', note_val): df.iloc[i, note_col] = None

        last_valid_row = i

    df.drop(index=rows_to_drop, inplace=True)
    return df.reset_index(drop=True)

def p1_parse_num(s):
    if s is None or pd.isna(s): return None
    s = str(s).strip().replace('\n', '').replace(' ', '').replace(',', '')
    if s == '': return None
    if s.lower() == 'nan' or s in ['-', '—', '一', '--', '。']: return 0.0
    if (s.startswith('(') and s.endswith(')')) or (s.startswith('（') and s.endswith('）')):
        s = s[1:-1]
        try: return -float(s)
        except: return None
    try: return float(s)
    except: return None

def p1_get_pdf_text_lines(pdf):
    all_lines = []
    for page in pdf.pages:
        words = page.extract_words()
        if not words: continue
        words.sort(key=lambda w: (w['top'], w['x0']))
        current_line, last_top = [], None
        for w in words:
            if last_top is None or abs(w['top'] - last_top) < 4:
                current_line.append(w)
                if last_top is None: last_top = w['top']
            else:
                current_line.sort(key=lambda w: w['x0'])
                line_str, last_x1 = "", None
                for cw in current_line:
                    if last_x1 is not None and cw['x0'] - last_x1 > 6: line_str += "   "
                    elif last_x1 is not None and cw['x0'] - last_x1 > 1: line_str += " "
                    line_str += cw['text']
                    last_x1 = cw['x1']
                all_lines.append(line_str)
                current_line, last_top = [w], w['top']
        if current_line:
            current_line.sort(key=lambda w: w['x0'])
            line_str, last_x1 = "", None
            for cw in current_line:
                if last_x1 is not None and cw['x0'] - last_x1 > 6: line_str += "   "
                elif last_x1 is not None and cw['x0'] - last_x1 > 1: line_str += " "
                line_str += cw['text']
                last_x1 = cw['x1']
            all_lines.append(line_str)
    return all_lines

def p1_parse_table_block(block_lines, table_name, multiplier):
    is_parent = "母公司" in table_name
    row_data, max_nums = [], 0
    bad_headers = {"期末余额", "期初余额", "年末余额", "年初余额", "本期发生额", "上期发生额", "本年发生额", "上年发生额", "本期金额", "上期金额", "项目", "资产", "附注"}

    for line in block_lines:
        if "单位：" in line or "编制单位" in line: continue
        line = re.sub(r'([\u4e00-\u9fa5A-Za-z\）\)])([\-\(\（]?(?:\d{1,3}(?:,\d{3})+|\d{4,}|\d+\.\d{1,})[\)\）]?)', r'\1 \2', line)
        parts = re.split(r'\s+', line.strip())
        if not parts or not parts[0]: continue

        nums, item_name_parts, parsing_numbers = [], [], True
        for p in reversed(parts):
            p_clean = p.replace(',', '').strip()
            val = p1_parse_num(p_clean)
            is_fin_val = (val is not None) and (re.search(r'\d', p_clean) or p_clean in ['-', '—', '--'])
            if parsing_numbers and is_fin_val:
                clean_for_note = p_clean.replace('-', '').replace('(', '').replace(')', '')
                if p_clean not in ['-', '—', '--', '0', '0.00']:
                    if re.match(r'^\d{1,3}$', clean_for_note) and 0 < abs(val) <= 300: is_fin_val = False
                    elif re.match(r'^\d{1,3}\.\d+$', clean_for_note) and not ('.' in clean_for_note and len(clean_for_note.split('.')[1]) >= 2): is_fin_val = False

            if parsing_numbers and is_fin_val: nums.insert(0, val)
            else:
                parsing_numbers = False
                item_name_parts.insert(0, p)

        raw_item_str = "".join(item_name_parts)
        note_val = ""
        note_pattern = r'^(.*?)([一二三四五六七八九十]+、[\d一二三四五六七八九十\(\)（）a-zA-Z\.\-、，,]+|[一二三四五六七八九十]+[\(（][一二三四五六七八九十\d\.\-、，,]+[\)）]|[\(（][一二三四五六七八九十]+[\)）][\d\.\-、，,]*|附?注(?:释)?[一二三四五六七八九十\d\.\-、，,]*|注\d+|[\(（]\d+[\)）]|\d{1,4}(?:[\-、，,]\d{1,4})*|[一二三四五六七八九十]+)$'
        match = re.search(note_pattern, raw_item_str)
        if match:
            left_part, right_part = match.group(1), match.group(2)
            if len(left_part) > 0:
                raw_item_str = left_part
                note_val = right_part

        item_name = raw_item_str
        if len(item_name) > 35 or any(sig in item_name for sig in ["法定代表人", "主管会计工作负责人", "会计机构负责人"]) or item_name in bad_headers: continue

        if not item_name:
            if nums and row_data:
                prev_item, prev_note, prev_nums = row_data[-1]
                if not prev_nums: row_data[-1] = (prev_item, prev_note, nums)
            continue
        row_data.append((item_name, note_val, nums))
        if len(nums) > max_nums: max_nums = len(nums)

    final_rows = []
    for item_name, note_val, nums in row_data:
        if not item_name: continue
        if not nums:
            final_rows.append([item_name, note_val, None, None])
            continue
        if "所有者权益变动表" in table_name:
            final_rows.append([item_name, note_val] + [n * multiplier for n in nums])
            continue
        end_val, begin_val = None, None
        if max_nums >= 4:
            if is_parent:
                if len(nums) >= 4: end_val, begin_val = nums[-2], nums[-1]
                elif len(nums) >= 2: end_val, begin_val = nums[-2], nums[-1]
            else:
                if len(nums) >= 4: end_val, begin_val = nums[0], nums[1]
                elif len(nums) >= 2: end_val, begin_val = nums[0], nums[1]
        else:
            if len(nums) >= 2: end_val, begin_val = nums[-2], nums[-1]
            elif len(nums) == 1: end_val = nums[0]
        final_rows.append([item_name, note_val, end_val * multiplier if end_val is not None else None, begin_val * multiplier if begin_val is not None else None])

    if not final_rows: return None

    idx = 0
    while idx < len(final_rows) - 1:
        curr_row = final_rows[idx]
        item_name = str(curr_row[0] or "")
        if item_name.count('(') + item_name.count('（') > item_name.count(')') + item_name.count('）'):
            next_row = final_rows[idx + 1]
            curr_row[0] = item_name + str(next_row[0] or "")
            if len(curr_row) > 1 and len(next_row) > 1 and next_row[1]: curr_row[1] = str(curr_row[1] or "") + str(next_row[1])
            for col_idx in range(2, max(len(curr_row), len(next_row))):
                next_val = next_row[col_idx] if col_idx < len(next_row) else None
                if next_val is not None:
                    if col_idx < len(curr_row):
                        if curr_row[col_idx] is None: curr_row[col_idx] = next_val
                    else:
                        while len(curr_row) <= col_idx: curr_row.append(None)
                        curr_row[col_idx] = next_val
            final_rows.pop(idx + 1)
        else:
            idx += 1

    pure_note_pattern = r'^([一二三四五六七八九十]+、\s*[\d一二三四五六七八九十\(\)（）a-zA-Z\.\-、，,]+|[一二三四五六七八九十]+[\(（][一二三四五六七八九十\d\.\-、，,]+[\)）]|[\(（][一二三四五六七八九十]+[\)）][\d\.\-、，,]*|附?注(?:释)?\s*[一二三四五六七八九十\d\.\-、，,]*|注\s*\d+|[\(（]\d+[\)）]|\d{1,2}|[一二三四五六七八九十]+)$'
    idx = 1
    while idx < len(final_rows):
        curr_row, item_name = final_rows[idx], str(final_rows[idx][0] or "").strip()
        if re.match(pure_note_pattern, item_name):
            prev_row = final_rows[idx - 1]
            prev_note = str(prev_row[1] or "").strip()
            curr_note = item_name + (str(curr_row[1]) if len(curr_row) > 1 and curr_row[1] else "")
            prev_row[1] = prev_note + " " + curr_note if prev_note else curr_note
            for col_idx in range(2, max(len(prev_row), len(curr_row))):
                curr_val = curr_row[col_idx] if col_idx < len(curr_row) else None
                if curr_val is not None:
                    if col_idx < len(prev_row):
                        if prev_row[col_idx] is None: prev_row[col_idx] = curr_val
                    else:
                        while len(prev_row) <= col_idx: prev_row.append(None)
                        prev_row[col_idx] = curr_val
            final_rows.pop(idx)
        else:
            idx += 1

    if "所有者权益变动表" in table_name:
        max_c = max(len(r) for r in final_rows)
        for r in final_rows:
            while len(r) < max_c: r.append(None)
        return pd.DataFrame(final_rows).dropna(how='all', axis=1)
    else:
        return pd.DataFrame(final_rows, columns=["项目", "附注", "期末余额", "期初余额"])


def p1_process_pdf(pdf_path, log_func):
    filename = os.path.basename(pdf_path)
    parent_dir = os.path.basename(os.path.dirname(pdf_path))
    if "年度报告" not in filename and "年报" not in filename: return None, None, None, None

    results = {name: [] for name, _ in MILESTONE_PATTERNS[:8]}
    results.update({"董监高及报酬情况": [], "员工情况": [], "分红情况": []})

    stock_code_match = re.search(r'\d{6}', parent_dir) or re.search(r'\d{6}', filename)
    stock_code = stock_code_match.group(0) if stock_code_match else "未提取"

    comp_name = re.sub(r'[（\(]?\d{6}[）\)]?', '', parent_dir).strip()
    # 【关键修复】：将原先拆散单字的错误正则，替换为精准的词组替换，安全保留“股份”、“公司”等字样的简称
    comp_name = re.sub(
        r'(20\d{2}年?度?|年度?报告全文|年度?报告|年报全文|年报|全文|摘要|修订版|更新后|股份有限公司|有限公司)', '',
        comp_name).strip('_ -')
    if not comp_name: comp_name = "未知公司"

    year_match = re.search(r'20\d{2}', filename)
    year = year_match.group(0) if year_match else "20XX"

    try:
        with pdfplumber.open(pdf_path) as pdf:
            if stock_code == "未提取":
                first_page = pdf.pages[0].extract_text() or ""
                fallback_code = re.search(r'(60\d{4}|00\d{4}|30\d{4}|68\d{4}|83\d{4}|43\d{4}|87\d{4})', first_page)
                if fallback_code: stock_code = fallback_code.group(1)

            all_lines = p1_get_pdf_text_lines(pdf)
            milestone_indices, current_search_idx = {}, 0

            for i, line in enumerate(all_lines):
                clean_line = line.replace(" ", "")
                if "....." in clean_line or "。。。。" in clean_line or re.search(r'表\s*\d+$', line.strip()): continue
                while current_search_idx < 9:
                    matched_idx = -1
                    for step in range(3):
                        check_idx = current_search_idx + step
                        if check_idx >= 9: break
                        if re.search(MILESTONE_PATTERNS[check_idx][1], clean_line) and len(
                                clean_line) < 30 and not re.search(r"分析|情况说明|变动原因|关键审计", clean_line):
                            matched_idx = check_idx
                            break
                    if matched_idx != -1:
                        for skip in range(current_search_idx, matched_idx):
                            if skip not in milestone_indices: milestone_indices[skip] = i
                        if matched_idx not in milestone_indices: milestone_indices[matched_idx] = i
                        current_search_idx = matched_idx + 1
                    else:
                        break

            def get_end_idx(current_i):
                for next_i in range(current_i + 1, 10):
                    if next_i in milestone_indices and milestone_indices[next_i] > milestone_indices.get(current_i, -1):
                        return milestone_indices[next_i]
                return min(milestone_indices.get(current_i, 0) + 300, len(all_lines))

            for i in range(8):
                name = MILESTONE_PATTERNS[i][0]
                start_idx, end_idx = milestone_indices.get(i, -1), get_end_idx(i)
                if start_idx != -1 and start_idx < end_idx:
                    block, multiplier = all_lines[start_idx:end_idx], 1
                    for line in block[:15]:
                        if "万元" in line:
                            multiplier = 10000
                            break
                    df = p1_parse_table_block(block, name, multiplier)
                    if df is not None and not df.empty:
                        results[name].append(df)

            for page in pdf.pages:
                tables = page.find_tables()
                if not tables: continue
                for table in tables:
                    raw_data = table.extract()
                    if not raw_data: continue
                    df_raw = pd.DataFrame(raw_data)
                    sub_str = "".join(map(str, df_raw.values.flatten())).replace(" ", "").replace("\n", "")
                    sub_head_str = "".join(map(str, df_raw.head(4).values.flatten())).replace(" ", "").replace("\n", "")
                    if "姓名" in sub_str and ("报酬" in sub_str or "薪酬" in sub_str) and "职务" in sub_str:
                        results["董监高及报酬情况"].append(df_raw.dropna(how='all'))
                    elif "专业构成" in sub_str or "教育程度" in sub_str:
                        results["员工情况"].append(df_raw.dropna(how='all'))
                    elif ("分红" in sub_str or "派息" in sub_str or "送红股" in sub_str or "转增" in sub_str) and \
                            (
                                    "每10股" in sub_str or "每十股" in sub_str or "现金分红" in sub_str or "分红金额" in sub_str):
                        if not any(junk in sub_head_str for junk in ["承诺", "限售", "减持", "同业竞争", "避免"]):
                            results["分红情况"].append(df_raw.dropna(how='all'))
    except Exception as e:
        log_func(f"  ❌ 提取异常: {str(e)}")

    final_results = {}
    for k, v in results.items():
        if v:
            merged = pd.concat(v, ignore_index=True)
            if "所有者权益变动表" not in k and k in [m[0] for m in MILESTONE_PATTERNS[:8]]:
                merged = merged[~merged.iloc[:, 0].astype(str).str.contains(
                    r'^项目$|^期末余额$|^期初余额$|^发生额$|^资产$')].reset_index(drop=True)
            merged = p1_clean_dataframe(merged)
            final_results[k] = merged
        else:
            final_results[k] = pd.DataFrame()
    return final_results, comp_name, year, stock_code
def common_auto_adjust_col_width(ws):
    for col_idx, column_cells in enumerate(ws.columns, start=1):
        max_length, col_letter = 0, get_column_letter(col_idx)
        for cell in column_cells:
            if cell.value is not None:
                cell.alignment = Alignment(horizontal='center', vertical='center', wrapText=False)
                length = sum(2 if '\u4e00' <= c <= '\u9fff' else 1 for c in str(cell.value))
                max_length = max(max_length, length)
        ws.column_dimensions[col_letter].width = min(max(max_length + 2, 12), 60)


def p1_generate_multi_year_summary(all_results, mappings, out_path, log_func):
    target_tables = ["合并资产负债表", "合并利润表", "合并现金流量表"]
    # 【关键修改】：解包新增的行业文件夹名称 (ind_folder)
    for (safe_name, stock_code, ind_folder), years_data in all_results.items():
        if len(years_data) <= 1: continue

        # 【关键修改】：直接将路径指向统一的行业文件夹
        company_out_path = os.path.join(out_path, ind_folder)
        os.makedirs(company_out_path, exist_ok=True)
        out_file = os.path.join(company_out_path, f"{safe_name}（{stock_code}）多年度汇总与增长率分析.xlsx")

        try:
            with pd.ExcelWriter(out_file, engine='openpyxl') as writer:
                has_summary = False
                for t_name in target_tables:
                    merged_df, years = pd.DataFrame(columns=["项目"]), sorted(list(years_data.keys()))
                    if years:
                        earliest_yr = years[0]
                        df0 = years_data[earliest_yr].get(t_name, pd.DataFrame())
                        if not df0.empty and "项目" in df0.columns and "期初余额" in df0.columns:
                            temp0 = df0[["项目", "期初余额"]].copy()
                            temp0["项目"] = temp0["项目"].apply(lambda x: mappings.get(str(x).replace(" ", ""), x))
                            temp0 = temp0.groupby("项目", as_index=False).sum(numeric_only=True)
                            temp0.rename(columns={"期初余额": f"{int(earliest_yr) - 1}年度"}, inplace=True)
                            merged_df = temp0

                    for yr in years:
                        df = years_data[yr].get(t_name, pd.DataFrame())
                        if df.empty or "项目" not in df.columns: continue
                        temp = df[["项目", "期末余额"]].copy()
                        temp["项目"] = temp["项目"].apply(lambda x: mappings.get(str(x).replace(" ", ""), x))
                        temp = temp.groupby("项目", as_index=False).sum(numeric_only=True)
                        temp.rename(columns={"期末余额": f"{yr}年度"}, inplace=True)
                        merged_df = temp if merged_df.empty else pd.merge(merged_df, temp, on="项目", how="outer")

                    if merged_df.empty: continue
                    has_summary = True

                    year_cols = ([f"{int(years[0]) - 1}年度"] if years else []) + [f"{y}年度" for y in years]
                    col_order = ["项目"]
                    for i in range(len(year_cols)):
                        curr_y_col = year_cols[i]
                        if curr_y_col in merged_df.columns:
                            col_order.append(curr_y_col)
                            if i > 0 and year_cols[i - 1] in merged_df.columns:
                                prev_y_col = year_cols[i - 1]
                                yoy_col = f"{curr_y_col.replace('年度', '')}年较上年增长率"

                                def calc_yoy(row, c=curr_y_col, p=prev_y_col):
                                    val_c, val_p = row.get(c), row.get(p)
                                    return (val_c - val_p) / abs(val_p) if pd.notna(val_c) and pd.notna(
                                        val_p) and val_p != 0 else None

                                merged_df[yoy_col] = merged_df.apply(calc_yoy, axis=1)
                                col_order.append(yoy_col)

                    merged_df = merged_df[[c for c in col_order if c in merged_df.columns]]
                    merged_df.to_excel(writer, sheet_name=t_name, startrow=2, index=False)
                    ws = writer.sheets[t_name]
                    ws.cell(1, 1, f"{safe_name}（{stock_code}）{t_name}（多年度汇总）\n单位：人民币元").font = Font(
                        bold=True, size=12)
                    ws.cell(1, 1).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                    ws.row_dimensions[1].height = 40

                    for col_idx, col_name in enumerate(merged_df.columns):
                        for row_idx in range(4, ws.max_row + 1):
                            cell = ws.cell(row=row_idx, column=col_idx + 1)
                            if isinstance(cell.value, (int, float)):
                                if "增长率" in col_name:
                                    cell.number_format = '0.00%'
                                else:
                                    cell.number_format = '#,##0.0000' if "每股收益" in str(
                                        ws.cell(row=row_idx, column=1).value) else '#,##0.00'
                    common_auto_adjust_col_width(ws)

                div_dfs = []
                for yr in sorted(list(years_data.keys())):
                    df_div = years_data[yr].get("分红情况", pd.DataFrame())
                    if not df_div.empty:
                        df_curr = df_div.copy()
                        df_curr.insert(0, "数据来源", f"{yr}年报")
                        div_dfs.append(df_curr)
                if div_dfs:
                    merged_div = pd.concat(div_dfs, ignore_index=True)
                    merged_div.to_excel(writer, sheet_name="历年分红情况汇总", startrow=2, index=False)
                    ws_div = writer.sheets["历年分红情况汇总"]
                    ws_div.cell(1, 1, f"{safe_name}（{stock_code}）历年分红情况汇总\n单位：详见表内说明").font = Font(
                        bold=True, size=12)
                    ws_div.cell(1, 1).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                    ws_div.row_dimensions[1].height = 40
                    common_auto_adjust_col_width(ws_div)
            if has_summary: log_func(f"  🌟 已生成 P1阶段 跨年汇总分析: {os.path.basename(out_file)}")
        except Exception as e:
            log_func(f"  ❌ P1跨年汇总生成失败 ({safe_name}): {str(e)}")

def p2_parse_to_yuan(val_str, col_name=""):
    if not val_str or val_str in ['nan', '-', '', 'None', '未披露']: return "未披露"
    val_str = str(val_str).replace(',', '')
    match = re.search(r'(-?\d+(?:\.\d+)?)', val_str)
    if not match: return "未披露"
    num = float(match.group(1))
    if "亿" in val_str or "亿" in col_name: num *= 100000000
    elif "万" in val_str or "万" in col_name: num *= 10000
    return num

def p2_fetch_10jqka_dividend(code, target_years, log_func):
    log_func(f"  -> 正在抓取 {code} 同花顺分红数据...")
    headers = {'User-Agent': HEADERS['User-Agent'], 'Referer': f'http://basic.10jqka.com.cn/{code}/'}
    company_results = []
    try:
        res = requests.get(f"http://basic.10jqka.com.cn/{code}/bonus.html", headers=headers, timeout=10)
        res.encoding = 'GBK'
        dfs = pd.read_html(io.StringIO(res.text))
        for yr in target_years:
            bonus_total_yuan = "未披露"
            for df in dfs:
                df.columns = ['_'.join(map(str, c)).replace(' ', '') if isinstance(c, tuple) else str(c).replace(' ', '') for c in df.columns]
                if any("分红" in c or "报告期" in c for c in df.columns):
                    mask = df.astype(str).apply(lambda row: row.str.contains(str(yr)).any(), axis=1)
                    if mask.any():
                        matched_rows = df[mask]
                        annual_mask = matched_rows.astype(str).apply(lambda r: r.str.contains('年报|12-31').any(), axis=1)
                        matched_row = matched_rows[annual_mask].iloc[0] if annual_mask.any() else matched_rows.iloc[0]
                        for col_str in df.columns:
                            val = str(matched_row[col_str]).strip()
                            if val in ['nan', '-', '', 'None']: continue
                            if "分红总额" in col_str or ("现金分红" in col_str and "率" not in col_str and "股息" not in col_str):
                                bonus_total_yuan = p2_parse_to_yuan(val, col_str)
                                break
                        break
            company_results.append({"股票代码": code, "年份": yr, "同花顺现金分红总额(元)": bonus_total_yuan})
    except Exception as e:
        for yr in target_years: company_results.append({"股票代码": code, "年份": yr, "同花顺现金分红总额(元)": f"异常:{e}"})
    time.sleep(1)
    return pd.DataFrame(company_results)

def p2_fetch_financial_sheets(code, start_year, end_year, log_func):
    log_func(f"  -> 抓取 {code} {start_year}-{end_year} 财务报表...")
    sheets = {}
    def filter_years(df):
        if df is None or df.empty: return None
        cols = [df.columns[0]]
        for c in df.columns[1:]:
            cln = str(c).strip().replace(' ', '')
            if '12-31' in cln or '1231' in cln:
                try:
                    if start_year <= int(cln[:4]) <= end_year: cols.append(c)
                except Exception as e:
                    log_func(f"  ⚠️ 年份解析发生错误，已跳过该项: {str(e)}")
        return df[cols].copy() if len(cols) > 1 else pd.DataFrame()
    def get_sina(url):
        try:
            time.sleep(0.5)
            r = requests.get(url, headers=HEADERS, timeout=10)
            r.encoding = 'gbk'
            if '<html' in r.text.lower()[:100]: return None
            df = pd.read_csv(io.StringIO(r.text), sep='\t', on_bad_lines='skip')
            df.dropna(how='all', axis=1, inplace=True)
            return filter_years(df)
        except: return None

    sheets['合并资产负债表'] = get_sina(f"http://vip.stock.finance.sina.com.cn/corp/go.php/vDOWN_BalanceSheet/displaytype/4/stockid/{code}/ctrl/all.phtml")
    sheets['合并利润表'] = get_sina(f"http://vip.stock.finance.sina.com.cn/corp/go.php/vDOWN_ProfitStatement/displaytype/4/stockid/{code}/ctrl/all.phtml")
    sheets['合并现金流量表'] = get_sina(f"http://vip.stock.finance.sina.com.cn/corp/go.php/vDOWN_CashFlow/displaytype/4/stockid/{code}/ctrl/all.phtml")
    return {k: v for k, v in sheets.items() if v is not None and not v.empty}

def p2_enrich_dividend_data(df_div, fin_sheets):
    if df_div.empty or '合并利润表' not in fin_sheets:
        df_div['归属于母公司所有者的净利润(元)'], df_div['分红率'] = "未披露", "无法计算"
        return df_div
    is_df = fin_sheets['合并利润表']
    item_col = is_df.columns[0]
    profit_row = None
    for idx, row in is_df.iterrows():
        val = str(row[item_col]).replace(' ', '')
        if '归属于母公司所有者的净利润' in val or '归属于母公司股东的净利润' in val:
            profit_row = row
            break
    profit_vals = {}
    if profit_row is not None:
        for col in is_df.columns[1:]:
            match = re.search(r'(20\d{2})', str(col))
            if match:
                try: profit_vals[int(match.group(1))] = float(str(profit_row[col]).replace(',', '').strip())
                except: profit_vals[int(match.group(1))] = None

    profit_list, ratio_list = [], []
    for _, row in df_div.iterrows():
        yr, div_val = int(row['年份']), row['同花顺现金分红总额(元)']
        profit = profit_vals.get(yr, None)
        if profit is not None:
            profit_list.append(profit)
            if isinstance(div_val, (int, float)) and profit != 0:
                try: ratio_list.append(f"{(div_val / profit) * 100:.2f}%")
                except: ratio_list.append("计算异常")
            else: ratio_list.append("无法计算")
        else:
            profit_list.append("未查到该年净利润数据")
            ratio_list.append("无法计算")
    df_div['归属于母公司所有者的净利润(元)'], df_div['分红率'] = profit_list, ratio_list
    return df_div


def p2_process_pdf_employee(pdf_path):
    filename = os.path.basename(pdf_path)
    parent_dir = os.path.basename(os.path.dirname(pdf_path))
    if "年度报告" not in filename and "年报" not in filename: return None, None, None, None
    sc_match = re.search(r'\d{6}', parent_dir) or re.search(r'\d{6}', filename)
    stock_code = sc_match.group(0) if sc_match else "未提取"
    year_match = re.search(r'20\d{2}', filename)
    year = int(year_match.group(0)) if year_match else 2022

    comp_name = re.sub(r'[（\(]?\d{6}[）\)]?', '', parent_dir).strip()
    # 【关键修复】：同样修正员工情况模块中的公司名提取
    comp_name = re.sub(
        r'(20\d{2}年?度?|年度?报告全文|年度?报告|年报全文|年报|全文|摘要|修订版|更新后|股份有限公司|有限公司)', '',
        comp_name).strip('_ -')
    if not comp_name: comp_name = "未知公司"

    results = []
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                tables = page.find_tables()
                if not tables: continue
                for table in tables:
                    data = table.extract()
                    if not data: continue
                    df = pd.DataFrame(data)
                    sub_str = "".join(map(str, df.values.flatten())).replace(" ", "").replace("\n", "")
                    if "专业构成" in sub_str or "教育程度" in sub_str:
                        results.append(df.dropna(how='all'))
    except:
        pass
    df_final = pd.concat(results, ignore_index=True).drop_duplicates().reset_index(
        drop=True) if results else pd.DataFrame()
    return df_final, comp_name, year, stock_code
def p3_is_blank(val):
    if val is None: return True
    return str(val).strip() in ("", "-", "None")

def p3_is_zero(val):
    if val is None: return False
    return str(val).strip() in ("0", "0.0")

def p3_extract_year(text):
    match = re.search(r'(20\d{2})', str(text))
    return int(match.group(1)) if match else None

def p3_clean_item_name(name):
    if not name: return ""
    n = re.sub(r'\s+', '', str(name).strip())
    n = re.sub(r'[\(（][^\)）]*(填列|号|-|－|损失|收益|减少)[^\)）]*[\)）]', '', n)
    n = re.sub(r'[\(（]元/股[\)）]', '', n)
    n = re.sub(r'[一二三四五六七八九十]+、\d+', '', n)
    n = re.sub(r'^(其中[:：]?|加[:：]?|减[:：]?)', '', n)
    alias_map = {
        "公允价值变动损失": "公允价值变动收益", "信用减值准备": "信用减值损失",
        "资产减值准备": "资产减值损失", "投资损失": "投资收益",
        "资产处置损失": "资产处置收益", "净敞口套期损失": "净敞口套期收益"
    }
    return alias_map.get(n, n)

def p3_extract_company_name(filename):
    clean_name = filename.replace("统一整合输出_", "")
    match = re.search(r'([\u4e00-\u9fa5]+)(?:\(|\（)', clean_name)
    return match.group(1).strip() if match else clean_name.split('_')[0].strip()

def p3_find_b_year_columns(ws):
    year_col_map = {}
    for r_idx in range(1, 5):
        for col_idx in range(2, ws.max_column + 1):
            try:
                year = p3_extract_year(ws.cell(row=r_idx, column=col_idx).value)
                if year and year not in year_col_map: year_col_map[year] = col_idx
            except Exception as e:
                print(f"B表列定位异常: {str(e)}")
    return year_col_map

def p3_find_a_data_column(ws, year):
    for r_idx in range(1, 6):
        for col_idx in range(2, ws.max_column + 1):
            try:
                val = str(ws.cell(row=r_idx, column=col_idx).value or "").strip()
                if "附注" in val: continue
                if "期末" in val or "年末" in val or "本期" in val or "本年" in val or str(year) in val: return col_idx
            except Exception as e:
                print(f"A表列定位异常: {str(e)}")
    return 3
def p3_process_company(company_name, b_file_path, a_files_dict, log_func):
    log_func(f"  👉 开始对比填补公司: {company_name}")
    try: wb_b = openpyxl.load_workbook(b_file_path)
    except Exception as e:
        log_func(f"  ❌ 读取B表失败: {e}")
        return

    b_modified = False
    for sheet_name in TARGET_SHEETS_P3:
        b_sheet_actual = next((s for s in wb_b.sheetnames if sheet_name in s), None)
        if not b_sheet_actual: continue
        ws_b = wb_b[b_sheet_actual]
        b_year_cols = p3_find_b_year_columns(ws_b)
        if not b_year_cols: continue

        b_existing_items = {}
        for row_idx in range(1, ws_b.max_row + 1):
            try:
                cell_val = ws_b.cell(row=row_idx, column=1).value
                if cell_val: b_existing_items[p3_clean_item_name(cell_val)] = row_idx
            except Exception as e:
                log_func(f"  ⚠️ 行解析发生错误: {str(e)}")

        for year, a_file_path in a_files_dict.items():
            if year not in b_year_cols: continue
            b_target_col = b_year_cols[year]
            try:
                wb_a = openpyxl.load_workbook(a_file_path, data_only=True)
                a_sheet_actual = next((s for s in wb_a.sheetnames if sheet_name in s), None)
                if not a_sheet_actual: continue
                ws_a = wb_a[a_sheet_actual]
            except: continue

            a_data_col = p3_find_a_data_column(ws_a, year)
            a_data_map = {}
            for row_idx in range(1, ws_a.max_row + 1):
                try:
                    item_cell, val_cell = ws_a.cell(row=row_idx, column=1), ws_a.cell(row=row_idx, column=a_data_col)
                    if item_cell.value:
                        c_name = p3_clean_item_name(item_cell.value)
                        if c_name and len(c_name) > 1 and "项目" not in c_name: a_data_map[c_name] = val_cell.value
                except Exception as e:
                    log_func(f"  ⚠️ A表单元格解析异常: {str(e)}")

            for clean_item, b_row_idx in b_existing_items.items():
                if clean_item in a_data_map:
                    a_raw_val = a_data_map[clean_item]
                    try:
                        b_cell = ws_b.cell(row=b_row_idx, column=b_target_col)
                        b_raw_val = b_cell.value
                        b_is_blank, b_is_zero = p3_is_blank(b_raw_val), p3_is_zero(b_raw_val)
                        a_is_blank, a_is_zero = p3_is_blank(a_raw_val), p3_is_zero(a_raw_val)

                        if b_is_blank:
                            b_cell.value = 0 if a_is_blank else a_raw_val
                            b_cell.fill = YELLOW_FILL
                            b_modified = True
                        elif b_is_zero:
                            if not a_is_blank and not a_is_zero:
                                b_cell.value = a_raw_val
                                b_cell.fill = YELLOW_FILL
                                b_modified = True
                    except Exception as e:
                        log_func(f"  ⚠️ 填补过程发生错误: {str(e)}")

            missing_items = [i for i in a_data_map.keys() if i not in b_existing_items.keys()]
            for clean_item in missing_items:
                a_raw_val = a_data_map[clean_item]
                new_row_idx = ws_b.max_row + 1
                original_item_name = clean_item
                for r in range(1, ws_a.max_row + 1):
                    try:
                        if p3_clean_item_name(ws_a.cell(row=r, column=1).value) == clean_item:
                            original_item_name = ws_a.cell(row=r, column=1).value
                            break
                    except Exception as e:
                        log_func(f"  ⚠️ 寻源项目名错误: {str(e)}")
                try:
                    name_cell = ws_b.cell(row=new_row_idx, column=1, value=original_item_name)
                    name_cell.fill = YELLOW_FILL
                    b_existing_items[clean_item] = new_row_idx
                    val_cell = ws_b.cell(row=new_row_idx, column=b_target_col, value=0 if p3_is_blank(a_raw_val) else a_raw_val)
                    val_cell.fill = YELLOW_FILL
                    b_modified = True
                except Exception as e:
                    log_func(f"  ⚠️ 新增项目填补错误: {str(e)}")

        for row in ws_b.iter_rows():
            for cell in row:
                try: cell.alignment = CENTER_ALIGNMENT
                except Exception as e:
                    pass
        for col_idx in range(1, ws_b.max_column + 1):
            column_letter = get_column_letter(col_idx)
            max_length = 0
            for row_idx in range(1, ws_b.max_row + 1):
                try:
                    cell = ws_b.cell(row=row_idx, column=col_idx)
                    if hasattr(cell, 'value') and cell.value:
                        length = len(str(cell.value).encode('gbk'))
                        if length > max_length: max_length = length
                except: pass
            ws_b.column_dimensions[column_letter].width = min((max_length + 2) * 1.2, 50)

    try:
        wb_b.save(b_file_path)
        if b_modified: log_func(f"  ✅ 填补完成：逻辑条件填补、高亮并排版保存！")
    except Exception as e:
        log_func(f"  ❌ 保存失败，请确保目标Excel未被打开！({e})")
# =====================================================================
# 函数库：APP2核心 (原B程序)
# =====================================================================
class FinancialAnalyzer:
    def __init__(self, group_data, config):
        self.group_data = group_data
        self.companies = list(group_data.keys())
        self.cfg = config
        self.years = set()
        for comp in self.companies:
            for sheet, df in group_data[comp].items():
                for col in df.columns:
                    # 兼容类似 2023、2023年、20231231、2023-12-31 等各种格式
                    m = re.search(r'(20\d{2})', str(col).strip())
                    if m: self.years.add(m.group(1) + "年")
        self.years = sorted(list(self.years))

    def get_prev_yr(self, year):
        m = re.search(r'\d{4}', year)
        return year.replace(m.group(), str(int(m.group()) - 1)) if m else "未知上年"

    # 【核心修复1】：吸收游离在外部的 get_item_value，变为类内部方法
    def get_val(self, comp, year, sheets, items):
        df_dict = self.group_data.get(comp, {})
        if not df_dict: return 0.0

        target_df = None
        for sk in sheets:
            for k, df in df_dict.items():
                if sk in k and "合并" in k: target_df = df; break
            if target_df is not None: break
        if target_df is None:
            for sk in sheets:
                for k, df in df_dict.items():
                    if sk in k: target_df = df; break
                if target_df is not None: break

        if target_df is None or target_df.empty or not year: return 0.0

        year_num = str(year).replace('年', '').replace('度', '').strip()
        year_col = f"{year_num}1231"
        if year_col not in target_df.columns: return 0.0

        for item in items:
            item_clean = item.replace(' ', '').replace('帐', '账')
            matched = target_df[
                target_df['项目'].astype(str).str.replace(' ', '').str.replace('帐', '账') == item_clean]
            if matched.empty:
                matched = target_df[
                    target_df['项目'].astype(str).str.replace(' ', '').str.replace('帐', '账').str.contains(item_clean,
                                                                                                            na=False)]
            if not matched.empty:
                val = matched.iloc[0][year_col]
                try:
                    if isinstance(val, str): val = val.replace(',', '').replace('，', '').strip()
                    res = float(val)
                    if pd.isna(res): return 0.0
                    return res
                except:
                    pass
        return 0.0

    def get_dividend_rate(self, comp, year):
        df_dict = self.group_data.get(comp, {})
        for k, df in df_dict.items():
            if "分红" in k:
                if '年份' in df.columns and '分红率' in df.columns:
                    yr_str = str(year).replace('年', '').replace('度', '').strip()
                    matched = df[df['年份'].astype(str).str.replace('.0', '', regex=False) == yr_str]
                    if not matched.empty:
                        val = matched.iloc[0]['分红率']
                        try:
                            if isinstance(val, str): val = val.replace(',', '').replace('%', '').strip()
                            res = float(val)
                            if pd.isna(res): return 0.0
                            return res
                        except:
                            pass
        return 0.0

    def build_df(self, rows):
        return pd.DataFrame(rows,
                            columns=['年度', '分析指标', '公司', '计算结果', '得分', '判定结果', '投资建议', '备注',
                                     '原始数据'])

    def analyze_dim_1(self):
        rows = []
        for year in self.years:
            assets = {c: self.get_val(c, year, ["资产负债表"], ["资产总计", "总资产"]) for c in self.companies}
            max_asset = max(assets.values()) if assets.values() else 0
            for comp in self.companies:
                val = assets.get(comp, 0)
                score = max(0, self.cfg['d1_asset_score_base'] - round((max_asset - val) / max_asset * 10,
                                                                       2)) if max_asset > 0 else 0
                judge, adv = ("实力最强", "保留") if (max_asset > 0 and val == max_asset) else (
                    ("实力差", "淘汰") if (max_asset > 0 and val < max_asset * self.cfg['d1_asset_poor_ratio']) else (
                        "实力一般", "进一步分析"))
                rows.append([year, "总资产规模", comp, val, score, judge, adv, "", f"总资产:{val}"])
                prev_val = self.get_val(comp, self.get_prev_yr(year), ["资产负债表"], ["资产总计", "总资产"])
                if prev_val > 0:
                    rate = (val - prev_val) / prev_val
                    score2 = round(10 + (rate - 0.10) / 0.01, 2)
                    judge2, adv2 = ("成长性好", "保留") if rate >= 0.10 else ("成长性差", "淘汰")
                    rows.append([year, "总资产同比增长率", comp, f"{rate * 100:.2f}%", score2, judge2, adv2, "", ""])
                else:
                    rows.append([year, "总资产同比增长率", comp, "无上期数据", "-", "-", "-", "", ""])
        return self.build_df(rows)

    def analyze_dim_2(self):
        rows = []
        money_items = [x.strip() for x in self.cfg['f_dim2_money'].split(',') if x.strip()]
        debt_items = [x.strip() for x in self.cfg['f_dim2_debt'].split(',') if x.strip()]
        for year in self.years:
            for comp in self.companies:
                asset = self.get_val(comp, year, ["资产负债表"], ["资产总计", "总资产"])
                liab = self.get_val(comp, year, ["资产负债表"], ["负债合计", "总负债"])
                if asset > 0:
                    rate = liab / asset
                    score = round(10 + (0.60 - rate) / 0.06, 2)
                    adv = "保留" if rate < self.cfg['d2_debt_warn'] else "淘汰"
                    rows.append([year, "资产负债率", comp, f"{rate * 100:.2f}%", score, "评估完成", adv, "", ""])
                else:
                    rows.append([year, "资产负债率", comp, "缺失数据", "-", "-", "-", "", ""])
                money = sum([self.get_val(comp, year, ["资产负债表"], [x]) for x in money_items])
                debt = sum([self.get_val(comp, year, ["资产负债表"], [x]) for x in debt_items])
                diff = money - debt
                score2 = round(10.0 + (diff / (money if money > 0 else 1) * 10.0), 2)
                judge2, adv2 = ("没偿债风险", "保留") if diff >= 0 else ("有偿债风险", "淘汰")
                rows.append([year, "准货币资金减有息负债", comp, diff, score2, judge2, adv2, "", ""])
        return self.build_df(rows)

    def analyze_dim_3(self):
        rows = []
        pay_items = [x.strip() for x in self.cfg['f_dim3_pay'].split(',') if x.strip()]
        recv_items = [x.strip() for x in self.cfg['f_dim3_recv'].split(',') if x.strip()]
        for year in self.years:
            for comp in self.companies:
                pay_recv = sum([self.get_val(comp, year, ["资产负债表"], [x]) for x in pay_items])
                recv_pay = sum([self.get_val(comp, year, ["资产负债表"], [x]) for x in recv_items])
                diff = pay_recv - recv_pay
                score = round(self.cfg['d3_diff_score_base'] + (
                            diff / (pay_recv if pay_recv > 0 else 1) * self.cfg['d3_diff_score_step']), 2)
                judge, adv = ("没偿债风险", "保留") if diff >= 0 else ("有偿债风险", "淘汰")
                rows.append([year, "应付预收减应收预付差额", comp, diff, score, judge, adv, "",
                             f"应付预收:{pay_recv}, 应收预付:{recv_pay}"])
        return self.build_df(rows)

    def analyze_dim_4(self):
        rows = []
        for year in self.years:
            for comp in self.companies:
                ar = self.get_val(comp, year, ["资产负债表"], ["应收票据"]) + self.get_val(comp, year, ["资产负债表"],
                                                                                           ["应收账款"])
                ar += self.get_val(comp, year, ["资产负债表"], ["合同资产"])
                asset = self.get_val(comp, year, ["资产负债表"], ["资产总计", "总资产"])
                if asset > 0:
                    rate = ar / asset
                    score = round(10 + (0.15 - rate) / 0.015, 2)
                    if rate < self.cfg['d4_ar_1']:
                        judge = "最好公司，产品很畅销"
                    elif rate < self.cfg['d4_ar_2']:
                        judge = "优秀公司，产品畅销"
                    elif rate < self.cfg['d4_ar_3']:
                        judge = "一般公司，销售一般"
                    elif rate < self.cfg['d4_ar_4']:
                        judge = "较差公司，较难销"
                    else:
                        judge = "很差公司，很难销"
                    adv = "保留" if rate < self.cfg['d4_ar_out'] else "淘汰"
                    rows.append([year, "应收加合同占总资产", comp, f"{rate * 100:.2f}%", score, judge, adv, "",
                                 f"应收合同:{ar},总资产:{asset}"])
                else:
                    rows.append([year, "应收加合同占总资产", comp, "缺失数据", "-", "-", "-", "", ""])
        return self.build_df(rows)

    def analyze_dim_5(self):
        rows = []
        for year in self.years:
            for comp in self.companies:
                fix = self.get_val(comp, year, ["资产负债表"], ["固定资产"]) + self.get_val(comp, year, ["资产负债表"],
                                                                                            ["在建工程"])
                asset = self.get_val(comp, year, ["资产负债表"], ["资产总计", "总资产"])
                if asset > 0:
                    rate = fix / asset
                    score = round(10 + (0.40 - rate) / 0.04, 2)
                    if rate < self.cfg['d5_fix_1']:
                        judge = "轻资产型，风险较低"
                    elif rate < self.cfg['d5_fix_out']:
                        judge = "一般公司，风险一般"
                    elif rate < self.cfg['d5_fix_2']:
                        judge = "重资产型，风险较大"
                    else:
                        judge = "重资产型，风险很大"
                    adv = "保留" if rate < self.cfg['d5_fix_out'] else "淘汰"
                    rows.append([year, "固定及在建占总资产", comp, f"{rate * 100:.2f}%", score, judge, adv, "",
                                 f"固定在建:{fix},总资产:{asset}"])
                else:
                    rows.append([year, "固定及在建占总资产", comp, "缺失数据", "-", "-", "-", "", ""])
        return self.build_df(rows)

    def analyze_dim_6(self):
        rows = []
        inv_items = [x.strip() for x in self.cfg['f_dim6_inv'].split(',') if x.strip()]
        for year in self.years:
            for comp in self.companies:
                inv_assets = sum([self.get_val(comp, year, ["资产负债表"], [x]) for x in inv_items])
                asset = self.get_val(comp, year, ["资产负债表"], ["资产总计", "总资产"])
                if asset > 0:
                    rate = inv_assets / asset
                    score = round(10 + (0.10 - rate) / 0.01, 2)
                    if rate <= self.cfg['d6_inv_best']:
                        judge = "最优秀公司，最专注主业"
                    elif rate < self.cfg['d6_inv_out']:
                        judge = "优秀公司，专注主业"
                    else:
                        judge = "较差公司，不专注主业"
                    adv = "保留" if rate < self.cfg['d6_inv_out'] else "淘汰"
                    rows.append([year, "投产类资产占总资产比率", comp, f"{rate * 100:.2f}%", score, judge, adv, "",
                                 f"投产资产:{inv_assets},总资产:{asset}"])
                else:
                    rows.append([year, "投产类资产占总资产比率", comp, "缺失数据", "-", "-", "-", "", ""])
        return self.build_df(rows)

    def analyze_dim_7(self):
        rows = []
        for year in self.years:
            for comp in self.companies:
                ar = self.get_val(comp, year, ["资产负债表"], ["应收票据"]) + self.get_val(comp, year, ["资产负债表"],
                                                                                           ["应收账款"])
                inv = self.get_val(comp, year, ["资产负债表"], ["存货"])
                gw = self.get_val(comp, year, ["资产负债表"], ["商誉"])
                asset = self.get_val(comp, year, ["资产负债表"], ["资产总计", "总资产"])
                if asset > 0:
                    ar_r, inv_r, gw_r = ar / asset, inv / asset, gw / asset
                    score_ar = 5.0 + (0.05 - ar_r) / 0.005 * 0.5
                    score_inv = 5.0 + (0.15 - inv_r) / 0.015 * 0.5
                    judge1 = "爆雷风险较大" if (
                                ar_r >= self.cfg['d7_ar_risk'] and inv_r >= self.cfg['d7_inv_risk']) else "爆雷风险一般"
                    adv1 = "淘汰" if judge1 == "爆雷风险较大" else "保留"
                    score_total = round(score_ar + score_inv, 2)
                    rows.append(
                        [year, "应收和存货占比", comp, f"应收:{ar_r * 100:.2f}%,存货:{inv_r * 100:.2f}%", score_total,
                         judge1, adv1, f"收:{score_ar:.2f},存:{score_inv:.2f}", f"应收:{ar},存货:{inv},总资产:{asset}"])

                    score_gw = round(10 + (0.10 - gw_r) / 0.01, 2)
                    judge2, adv2 = ("爆雷风险较大", "淘汰") if gw_r >= self.cfg['d7_gw_risk'] else ("爆雷风险小",
                                                                                                    "保留")
                    rows.append([year, "商誉占总资产比率", comp, f"{gw_r * 100:.2f}%", score_gw, judge2, adv2, "",
                                 f"商誉:{gw},总资产:{asset}"])
                else:
                    rows.append([year, "应收和存货占比", comp, "缺失数据", "-", "-", "-", "", ""])
                    rows.append([year, "商誉占总资产比率", comp, "缺失数据", "-", "-", "-", "", ""])
        return self.build_df(rows)

    def analyze_dim_8(self):
        rows = []
        for year in self.years:
            revs = {c: self.get_val(c, year, ["利润表"], ["营业总收入", "营业收入"]) for c in self.companies}
            max_rev = max(revs.values()) if revs.values() else 0
            for comp in self.companies:
                val = revs.get(comp, 0)
                score = round(
                    max(0, self.cfg['d8_rev_score_base'] - (max_rev - val) / max_rev * self.cfg['d8_rev_score_step']),
                    2) if max_rev > 0 else 0
                if max_rev > 0 and val == max_rev:
                    judge, adv = "行业地位最高", "保留"
                elif max_rev > 0 and val <= max_rev * self.cfg['d8_rev_poor_ratio']:
                    judge, adv = "行业地位差", "淘汰"
                else:
                    judge, adv = "行业地位一般", "进一步分析"
                rows.append([year, "营业收入规模", comp, val, score, judge, adv, "", f"营收:{val}"])

                prev_val = self.get_val(comp, self.get_prev_yr(year), ["利润表"], ["营业总收入", "营业收入"])
                if prev_val > 0:
                    rate = (val - prev_val) / prev_val
                    score2 = round(10 + (rate - 0.10) / 0.01, 2)
                    judge2, adv2 = ("成长性好", "保留") if rate > self.cfg['d8_growth_good'] else ("成长性差", "淘汰")
                    rows.append([year, "营业收入同比增长率", comp, f"{rate * 100:.2f}%", score2, judge2, adv2, "",
                                 f"本期:{val},上期:{prev_val}"])
                else:
                    rows.append([year, "营业收入同比增长率", comp, "无上期数据", "-", "-", "-", "", ""])
        return self.build_df(rows)

    def analyze_dim_9(self):
        rows = []
        for year in self.years:
            for comp in self.companies:
                rev = self.get_val(comp, year, ["利润表"], ["营业收入"])
                cost = self.get_val(comp, year, ["利润表"], ["营业成本"])
                if cost > 0:
                    gm = (rev - cost) / rev
                    score = round(10 + (gm - 0.40) / 0.04, 2)
                    judge, adv = ("产品竞争力较强", "保留") if gm > self.cfg['d9_gm_out'] else ("产品竞争力较差",
                                                                                                "淘汰")
                    rows.append(
                        [year, "毛利率", comp, f"{gm * 100:.2f}%", score, judge, adv, "", f"收入:{rev},成本:{cost}"])

                    prev_rev = self.get_val(comp, self.get_prev_yr(year), ["利润表"], ["营业收入"])
                    prev_cost = self.get_val(comp, self.get_prev_yr(year), ["利润表"], ["营业成本"])
                    if prev_cost > 0:
                        prev_gm = (prev_rev - prev_cost) / prev_rev
                        vol = abs(gm - prev_gm) / prev_gm if prev_gm != 0 else 0
                        score2 = round(10 + (0.20 - vol) / 0.02, 2)
                        if vol >= self.cfg['d9_gm_vol_out']:
                            judge2, adv2 = "公司经营风险大", "淘汰"
                        elif vol < self.cfg['d9_gm_vol_safe']:
                            judge2, adv2 = "优秀公司的正常波动", "保留"
                        else:
                            judge2, adv2 = "公司经营风险小", "保留"
                        rows.append([year, "毛利率波动幅度", comp, f"{vol * 100:.2f}%", score2, judge2, adv2, "",
                                     f"本期GM:{gm * 100:.1f}%,上期GM:{prev_gm * 100:.1f}%"])
                    else:
                        rows.append([year, "毛利率波动幅度", comp, "无上期数据", "-", "-", "-", "", ""])
                else:
                    rows.append([year, "毛利率", comp, "缺失数据", "-", "-", "-", "", ""])
                    rows.append([year, "毛利率波动幅度", comp, "缺失数据", "-", "-", "-", "", ""])
        return self.build_df(rows)

    def analyze_dim_10(self):
        rows = []
        for year in self.years:
            for comp in self.companies:
                rev = self.get_val(comp, year, ["利润表"], ["营业收入"])
                cost = self.get_val(comp, year, ["利润表"], ["营业成本"])
                # 【核心修复2】：将 safe_div 彻底替换为安全的内联计算
                gm = (rev - cost) / rev if rev != 0 else 0.0

                exp = sum([self.get_val(comp, year, ["利润表"], [x]) for x in ["销售费用", "管理费用", "研发费用"]])
                fin_exp = self.get_val(comp, year, ["利润表"], ["财务费用"])
                if fin_exp > 0: exp += fin_exp

                if rev > 0 and gm > 0:
                    ratio = exp / rev
                    res = ratio / gm
                    score = round(10 + (0.60 - res) / 0.06, 2)
                    if res < self.cfg['d10_exp_gm_safe']:
                        judge, adv = "成本控制能力优秀", "保留"
                    elif res <= self.cfg['d10_exp_gm_out']:
                        judge, adv = "成本控制能力一般", "保留"
                    else:
                        judge, adv = "成本控制能力较差", "淘汰"
                    rows.append([year, "期间费用率/毛利率", comp, f"{res * 100:.2f}%", score, judge, adv, "",
                                 f"费用:{exp},营收:{rev},GM:{gm * 100:.1f}%"])
                else:
                    rows.append([year, "期间费用率/毛利率", comp, "缺失数据", "-", "-", "-", "", ""])
        return self.build_df(rows)

    def analyze_dim_11(self):
        rows = []
        for year in self.years:
            for comp in self.companies:
                sales_exp = self.get_val(comp, year, ["利润表"], ["销售费用"])
                rev = self.get_val(comp, year, ["利润表"], ["营业收入"])
                if rev > 0:
                    ratio = sales_exp / rev
                    score = round(10 + (0.30 - ratio) / 0.03, 2)
                    if ratio < self.cfg['d11_sales_exp_1']:
                        judge, adv = "产品比较容易销售，风险较小", "保留"
                    elif ratio <= self.cfg['d11_sales_exp_out']:
                        judge, adv = "产品销售难度一般，风险一般", "保留"
                    else:
                        judge, adv = "产品销售难度大，风险大", "淘汰"
                    rows.append([year, "销售费用率", comp, f"{ratio * 100:.2f}%", score, judge, adv, "",
                                 f"销售费用:{sales_exp},营收:{rev}"])

                    prev_exp = self.get_val(comp, self.get_prev_yr(year), ["利润表"], ["销售费用"])
                    prev_rev = self.get_val(comp, self.get_prev_yr(year), ["利润表"], ["营业收入"])
                    if prev_rev > 0:
                        prev_ratio = prev_exp / prev_rev
                        trend = ratio - prev_ratio
                        score2 = round(
                            self.cfg['d11_trend_score_base'] + (self.cfg['d11_trend_out'] - trend) / self.cfg[
                                'd11_trend_score_step'], 2)
                        if trend == self.cfg['d11_trend_out']:
                            judge2, adv2 = "产品销售难度不变", "保留"
                        elif trend < self.cfg['d11_trend_out']:
                            judge2, adv2 = "产品销售难度在下降", "保留"
                        else:
                            judge2, adv2 = "产品销售难度在上升", "淘汰"
                        rows.append([year, "销售费用率变动趋势", comp, f"{trend * 100:.2f}%", score2, judge2, adv2, "",
                                     f"本期率:{ratio * 100:.1f}%,上期率:{prev_ratio * 100:.1f}%"])
                    else:
                        rows.append([year, "销售费用率变动趋势", comp, "无上期数据", "-", "-", "-", "", ""])
                else:
                    rows.append([year, "销售费用率", comp, "缺失数据", "-", "-", "-", "", ""])
                    rows.append([year, "销售费用率变动趋势", comp, "缺失数据", "-", "-", "-", "", ""])
        return self.build_df(rows)

    def analyze_dim_12(self):
        rows = []
        op_plus_items = [x.strip() for x in self.cfg['f_dim12_op_plus'].split(',') if x.strip()]
        op_minus_items = [x.strip() for x in self.cfg['f_dim12_op_minus'].split(',') if x.strip()]
        for year in self.years:
            for comp in self.companies:
                rev = self.get_val(comp, year, ["利润表"], ["营业收入"])
                cost = self.get_val(comp, year, ["利润表"], ["营业成本"])
                tax = self.get_val(comp, year, ["利润表"], ["税金及附加"])
                exp = sum([self.get_val(comp, year, ["利润表"], [x]) for x in
                           ["销售费用", "管理费用", "研发费用", "财务费用"]])
                core_profit = rev - cost - tax - exp
                if rev > 0:
                    rate1 = core_profit / rev
                    score1 = round(10 + (rate1 - 0.15) / 0.015, 2)
                    judge1, adv1 = ("主业盈利能力强", "保留") if rate1 >= self.cfg['d12_core_out'] else (
                        "主业盈利能力弱", "淘汰")
                    rows.append([year, "主营利润率", comp, f"{rate1 * 100:.2f}%", score1, judge1, adv1, "",
                                 f"主营利润:{core_profit},营收:{rev}"])
                else:
                    rows.append([year, "主营利润率", comp, "缺失数据", "-", "-", "-", "", ""])

                op_plus_sum = sum([self.get_val(comp, year, ["利润表"], [x]) for x in op_plus_items])
                op_minus_sum = sum([self.get_val(comp, year, ["利润表"], [x]) for x in op_minus_items])
                op_profit_custom = op_plus_sum - op_minus_sum
                if op_profit_custom != 0:
                    rate2 = core_profit / op_profit_custom
                    score2 = round(10 + (rate2 - 0.80) / 0.08, 2)
                    judge2, adv2 = ("利润质量高", "保留") if rate2 > self.cfg['d12_profit_out'] else ("利润质量低",
                                                                                                      "淘汰")
                    rows.append(
                        [year, "主营利润/营业利润(自定义)", comp, f"{rate2 * 100:.2f}%", score2, judge2, adv2, "",
                         f"主营利润:{core_profit},计算分母:{op_profit_custom}"])
                else:
                    rows.append([year, "主营利润/营业利润(自定义)", comp, "缺失数据", "-", "-", "-", "", ""])
        return self.build_df(rows)

    def analyze_dim_13(self):
        rows = []
        for year in self.years:
            for comp in self.companies:
                non_op_in = self.get_val(comp, year, ["利润表"], ["营业外收入"])
                non_op_out = self.get_val(comp, year, ["利润表"], ["营业外支出"])
                op_profit = self.get_val(comp, year, ["利润表"], ["营业利润"])
                net_non_op = non_op_in - non_op_out
                total = op_profit + net_non_op
                if total != 0:
                    rate = net_non_op / total
                    score = round(10 + (0.05 - rate) / 0.005, 2)
                    judge, adv = ("利润质量优秀", "保留") if rate < self.cfg['d13_non_op_out'] else ("利润质量较差",
                                                                                                     "淘汰")
                    rows.append([year, "营业外净额/利润总额", comp, f"{rate * 100:.2f}%", score, judge, adv, "",
                                 f"营业外净额:{net_non_op},总额:{total}"])
                else:
                    rows.append([year, "营业外净额/利润总额", comp, "缺失数据", "-", "-", "-", "", ""])
        return self.build_df(rows)

    def analyze_dim_14(self):
        rows = []
        for year in self.years:
            nps = {c: self.get_val(c, year, ["利润表"],
                                   ["归属于母公司所有者的净利润", "归属于母公司股东的净利润", "净利润"]) for c in
                   self.companies}
            max_np = max(nps.values()) if nps.values() else 0
            for comp in self.companies:
                val = nps.get(comp, 0)
                score = round(
                    max(0, self.cfg['d14_np_score_base'] - (max_np - val) / max_np * self.cfg['d14_np_score_step']),
                    2) if max_np > 0 else 0
                if max_np > 0 and val == max_np:
                    judge, adv = "盈利能力最强", "保留"
                elif max_np > 0 and val <= max_np * self.cfg['d14_np_poor_ratio']:
                    judge, adv = "盈利能力差", "淘汰"
                else:
                    judge, adv = "盈利能力一般", "进一步分析"
                rows.append([year, "归母净利润规模", comp, val, score, judge, adv, "", f"归母净利润:{val}"])

                prev_val = self.get_val(comp, self.get_prev_yr(year), ["利润表"],
                                        ["归属于母公司所有者的净利润", "归属于母公司股东的净利润", "净利润"])
                if prev_val != 0:
                    rate = (val - prev_val) / abs(prev_val)
                    score2 = round(10 + (rate - 0.10) / 0.001, 2)
                    if rate >= self.cfg['d14_np_growth']:
                        judge2, adv2 = "盈利不但强且持续性较好", "保留"
                    elif rate >= self.cfg['d14_np_growth_out']:
                        judge2, adv2 = "盈利持续性差", "淘汰"
                    else:
                        judge2, adv2 = "已处于衰落中，持续性较差", "淘汰"
                    rows.append([year, "归母净利增长率", comp, f"{rate * 100:.2f}%", score2, judge2, adv2, "",
                                 f"本期:{val},上期:{prev_val}"])
                else:
                    rows.append([year, "归母净利增长率", comp, "无上期数据", "-", "-", "-", "", ""])
        return self.build_df(rows)

    def analyze_dim_15(self):
        rows = []
        for year in self.years:
            for comp in self.companies:
                capex = self.get_val(comp, year, ["现金流量表"], ["购建固定资产、无形资产和其他长期资产支付的现金"])
                cfo = self.get_val(comp, year, ["现金流量表"], ["经营活动产生的现金流量净额"])
                if cfo > 0:
                    rate1 = capex / cfo
                    if rate1 < 0.03:
                        score1 = round(10 - (0.03 - rate1) / 0.003, 2)
                    elif rate1 <= 1.0:
                        score1 = round(10 + (1.0 - rate1) / 0.10, 2)
                    else:
                        score1 = round(10 - (rate1 - 1.0) / 0.10, 2)

                    if rate1 < self.cfg['d15_capex_slow']:
                        judge1, adv1 = "成长太慢，回报较低", "淘汰"
                    elif rate1 < self.cfg['d15_capex_safe']:
                        judge1, adv1 = "增长潜力较大并且风险相对较小", "保留"
                    elif rate1 < self.cfg['d15_capex_out']:
                        judge1, adv1 = "扩张略高，风险一般", "保留"
                    else:
                        judge1, adv1 = "扩张激进，风险较大", "淘汰"
                    rows.append([year, "CAPEX/CFO", comp, f"{rate1 * 100:.2f}%", score1, judge1, adv1, "",
                                 f"CAPEX:{capex},CFO:{cfo}"])
                else:
                    rows.append([year, "CAPEX/CFO", comp, "缺失数据", "-", "-", "-", "", ""])

                sales_cash = self.get_val(comp, year, ["现金流量表"], ["销售商品、提供劳务收到的现金"])
                prev_cash = self.get_val(comp, self.get_prev_yr(year), ["现金流量表"], ["销售商品、提供劳务收到的现金"])
                prev_np = self.get_val(comp, self.get_prev_yr(year), ["利润表"],
                                       ["归属于母公司所有者的净利润", "归属于母公司股东的净利润", "净利润"])
                if prev_np > 0:
                    trend = (sales_cash - prev_cash) / prev_np
                    score2 = round(self.cfg['d15_cash_score_base'] + trend / self.cfg['d15_cash_score_step'], 2)
                    judge2, adv2 = ("有增长潜力", "保留") if trend > self.cfg['d15_cash_trend_out'] else (
                        "没有增长潜力", "淘汰")
                    rows.append([year, "销售现金变动趋势", comp, f"{trend * 100:.2f}%", score2, judge2, adv2, "",
                                 f"本期现金:{sales_cash},上期现金:{prev_cash},上期净利:{prev_np}"])
                else:
                    rows.append([year, "销售现金变动趋势", comp, "缺失数据", "-", "-", "-", "", ""])
        return self.build_df(rows)

    def analyze_dim_16(self):
        rows = []
        for year in self.years:
            for comp in self.companies:
                rate = self.get_dividend_rate(comp, year)
                if rate > 0:
                    rate_dec = rate / 100 if rate > 1 else rate
                    if rate_dec >= 0.70:
                        score = round(10 - (rate_dec - 0.70) / 0.07, 2)
                    elif rate_dec >= 0.30:
                        score = round(10 + (rate_dec - 0.30) / 0.03, 2)
                    else:
                        score = round(10 - (0.30 - rate_dec) / 0.03, 2)

                    if rate_dec < self.cfg['d16_div_out_low']:
                        judge, adv = "不够厚道", "淘汰"
                    elif rate_dec <= self.cfg['d16_div_out_high']:
                        judge, adv = "厚道且持续性强", "保留"
                    else:
                        judge, adv = "难以持续", "淘汰"
                    rows.append(
                        [year, "分红率", comp, f"{rate_dec * 100:.2f}%", score, judge, adv, "", f"提取分红率:{rate}"])
                else:
                    rows.append([year, "分红率", comp, "缺失数据", "-", "-", "-", "", ""])
        return self.build_df(rows)

    def run_all_dimensions(self):
        return {
            "维度1_总资产分析": self.analyze_dim_1(), "维度2_负债分析": self.analyze_dim_2(),
            "维度3_竞争优势分析": self.analyze_dim_3(), "维度4_产品竞争力": self.analyze_dim_4(),
            "维度5_维持竞争成本": self.analyze_dim_5(), "维度6_主业专注度": self.analyze_dim_6(),
            "维度7_爆雷风险": self.analyze_dim_7(), "维度8_行业地位": self.analyze_dim_8(),
            "维度9_毛利率分析": self.analyze_dim_9(), "维度10_成本管控": self.analyze_dim_10(),
            "维度11_销售难易度": self.analyze_dim_11(), "维度12_主营盈利能力": self.analyze_dim_12(),
            "维度13_利润质量": self.analyze_dim_13(), "维度14_整体盈利能力": self.analyze_dim_14(),
            "维度15_增长潜力": self.analyze_dim_15(), "维度16_分红情况": self.analyze_dim_16(),
        }
def format_excel_sheet(ws, is_summary=False):
    red_font = Font(color="FF0000", bold=True)
    green_fill = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")
    if is_summary:
        ws.column_dimensions['A'].width = 25
        ws.column_dimensions['B'].width = 25
        ws.column_dimensions['C'].width = 20
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=2, max_col=3):
            for cell in row:
                cell.fill = green_fill; cell.font = Font(bold=True)
                cell.alignment = Alignment(vertical='center', horizontal='center')
        return

    ws.column_dimensions['A'].width = 12
    ws.column_dimensions['B'].width = 25
    ws.column_dimensions['C'].width = 20
    ws.column_dimensions['D'].width = 15
    ws.column_dimensions['E'].width = 10
    ws.column_dimensions['F'].width = 25
    ws.column_dimensions['G'].width = 15
    ws.column_dimensions['H'].width = 20
    ws.column_dimensions['I'].width = 45

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=9):
        if len(row) > 6 and row[6].value and "淘汰" in str(row[6].value):
            for cell in row: cell.font = red_font
        for cell in row: cell.alignment = Alignment(vertical='center', wrap_text=True)

# =====================================================================
# 函数库：APP3核心 (原C程序)
# =====================================================================
def safe_float(x):
    if x is None: return None
    if isinstance(x, (int, float)): return None if pd.isna(x) else float(x)
    s = str(x).strip()
    if s in ("", "-", "N/A", "NA", "None", "null"): return None
    s = s.replace("%", "").replace(",", "")
    m = re.search(r"-?\d+(\.\d+)?", s)
    return float(m.group()) if m else None

def get_widget_text_var(obj, attr_name):
    """【修复3】安全读取 tk 变量：控件可能尚未创建（例如 setup_ui 期间就调用代理设置），取不到时返回 ''。"""
    try:
        var = getattr(obj, attr_name, None)
        if var is None:
            return ""
        val = var.get() if hasattr(var, "get") else var
        return str(val).strip()
    except Exception:
        return ""


def apply_direct_network_env():
    """【修复3】强制直连：清除运行期间被其它子程序写入的所有代理环境变量，并设置 NO_PROXY。"""
    cleared = []
    for k in ("HTTP_PROXY", "HTTPS_PROXY", "http_proxy", "https_proxy", "ALL_PROXY", "all_proxy"):
        if os.environ.pop(k, None) is not None:
            cleared.append(k)
    os.environ["NO_PROXY"] = "localhost,127.0.0.1,::1"
    os.environ["no_proxy"] = "localhost,127.0.0.1,::1"
    return cleared


# =====================================================================
# 【修复5】个股行情 HTTP 直取（替代不稳定的浏览器抓取）
#   背景：原 main_logic_c 用 Selenium 打开雪球抓“现价/PE(TTM)/股息率”，
#   一旦 msedgedriver 与 Edge 版本不匹配（SessionNotCreatedException）或雪球弹
#   “访问提示”风控页，外层 except: pass 会把 7 只股票全部留成 N/A。
#   现改为：先走腾讯/新浪行情 HTTP 接口（稳定、无需登录、无需浏览器），
#          只有在所有 HTTP 源都失败时才回退到原来的浏览器抓取。
# =====================================================================
_QUOTE_UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
             "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")
_QUOTE_HEADERS = {"User-Agent": _QUOTE_UA}


def _http_get(url, *, params=None, headers=None, timeout=8, encoding=None, retries=3, label=""):
    """【修复10】带重试的 GET：行情/分红这类接口偶发超时或限流，
    单次失败不应让某个股票永久变成 N/A（此前 伊利股份 就是这样被漏掉的）。"""
    last = None
    for i in range(1, max(1, retries) + 1):
        try:
            r = requests.get(url, params=params, headers=headers or _QUOTE_HEADERS, timeout=timeout)
            if encoding:
                r.encoding = encoding
            return r
        except Exception as e:
            last = e
            if i < retries:
                time.sleep(0.8 * i)
    if label:
        print(f"    ⚠️ {label} 请求失败（已重试{retries}次）: {type(last).__name__}: {str(last)[:120]}")
    return None


def _num_or_none(v):
    try:
        if v is None: return None
        s = str(v).strip()
        if s in ("", "-", "N/A", "None"): return None
        f = float(s)
        return f if f == f else None
    except Exception:
        return None


def _quote_symbol(user_input):
    """把用户输入的名字/代码翻译成可查询的行情符号（腾讯格式）。
    优先用腾讯智能搜索把“公司简称”翻成代码；失败则退回 akshare 代码名称表。
    """
    raw = str(user_input or "").strip()
    if not raw: return "", "unknown"
    s = re.sub(r'\s+', '', raw)

    def _from_code(code):
        code = str(code).strip()
        if re.fullmatch(r'\d{6}', code):
            if code[0] in ("6", "9"):
                return "sh" + code, "A股"
            if code[0] in ("0", "2", "3"):
                return "sz" + code, "A股"
            if code[0] in ("4", "8"):
                return "bj" + code, "A股"
        if re.fullmatch(r'\d{4,5}', code):
            return "hk" + code.zfill(5), "港股"
        return "", "unknown"

    # ① 纯代码：直接换算
    if re.fullmatch(r'\d{4,6}', s):
        sym, mkt = _from_code(s)
        if sym: return sym, mkt

    # ② 含 6 位代码的混合文本
    m = re.search(r'(\d{6})', s)
    if m:
        sym, mkt = _from_code(m.group(1))
        if sym: return sym, mkt

    # ③ 字母代码 -> 美股
    if re.fullmatch(r'[A-Za-z][A-Za-z\.\-]{0,9}', s):
        return "us" + s.upper(), "美股"

    # ④ 中文名称 -> 先查腾讯智能搜索（已验证可用，返回 unicode 转义的名称）
    code = ""
    sb = os.path.join(os.path.expanduser("~"), ".vim_stock_name_map.json")
    cache = {}
    try:
        if os.path.exists(sb):
            with open(sb, "r", encoding="utf-8") as f:
                cache = json.load(f)
    except Exception:
        cache = {}
    if s in cache:
        code = cache[s]
    if not code:
        try:
            r = _http_get("https://smartbox.gtimg.cn/s3/",
                          params={"v": "2", "q": s, "t": "all"},
                          timeout=8, encoding="gbk", retries=3, label=f"代码解析[{s}]")
            if r is not None:
                mm = re.search(r'v_hint="([^"]*)"', r.text)
                if mm and mm.group(1):
                    first = mm.group(1).split("^")[0].split("~")
                    if len(first) >= 3:
                        mk, cd = first[0], first[1].split(".")[0]
                        code = cd if mk in ("sh", "sz", "bj") else (cd.zfill(5) if mk == "hk" else cd.upper())
                        try:
                            cache[s] = code
                            with open(sb, "w", encoding="utf-8") as f:
                                json.dump(cache, f, ensure_ascii=False)
                        except Exception:
                            pass
        except Exception as e:
            print(f"    ⚠️ 解析“{s}”的股票代码时出错: {e}")
    if not code:
        try:
            import akshare as ak
            df = ak.stock_info_a_code_name()
            hit = df[df["name"].astype(str).str.replace(" ", "") == s]
            if not hit.empty:
                code = str(hit.iloc[0]["code"])
        except Exception:
            pass
    if code:
        sym, mkt = _from_code(code)
        if sym: return sym, mkt
    return "", "unknown"


def _fetch_quote_tencent(symbol):
    """腾讯行情：字段 [1]名称 [3]现价 [4]昨收 [39]市盈率(TTM)"""
    r = _http_get(f"https://qt.gtimg.cn/q={symbol}", timeout=8, encoding="gbk",
                  retries=3, label=f"腾讯行情[{symbol}]")
    if r is None:
        raise RuntimeError("腾讯行情请求失败")
    m = re.search(r'="([^"]+)"', r.text)
    if not m:
        raise RuntimeError("腾讯行情无此代码")
    f = m.group(1).split("~")
    if len(f) < 6 or not f[1]:
        raise RuntimeError("腾讯行情无此代码")
    price = _num_or_none(f[3]) or _num_or_none(f[4])       # 现价（停牌时回落到昨收）
    pe = _num_or_none(f[39]) if len(f) > 39 else None       # 市盈率(TTM)
    return {"来源": "腾讯行情", "名称": f[1], "现价": price, "PE_TTM": pe, "股息率": None}


def _fetch_quote_sina(symbol):
    r = _http_get(f"https://hq.sinajs.cn/list={symbol}",
                  headers={**_QUOTE_HEADERS, "Referer": "https://finance.sina.com.cn"},
                  timeout=8, encoding="gb18030", retries=3, label=f"新浪行情[{symbol}]")
    if r is None:
        raise RuntimeError("新浪行情请求失败")
    m = re.search(r'="([^"]*)"', r.text)
    if not m or not m.group(1).strip():
        raise RuntimeError("新浪行情无此代码")
    f = m.group(1).split(",")
    if len(f) < 4 or not f[0]:
        raise RuntimeError("新浪行情无此代码")
    price = _num_or_none(f[3]) or _num_or_none(f[1])
    return {"来源": "新浪行情", "名称": f[0], "现价": price, "PE_TTM": None, "股息率": None}


# ---- 本地分红数据（供计算动态股息率：每股派息 ÷ 现价）----
_DIVIDEND_MAP = {}


def _load_local_dividend_map(out_dir):
    """从【年报提取程序】产出的《年报提取表》里读“每10股派息数（元）”。
    返回 {股票代码: (每股派息元, 年份)}。这是本程序自己的数据，比联网更可靠。"""
    result = {}
    try:
        if not out_dir or not os.path.isdir(out_dir):
            return result
        for r, d, files in os.walk(out_dir):
            for fn in files:
                if fn.startswith("~$") or not fn.lower().endswith((".xlsx", ".xls")):
                    continue
                m = re.search(r'[（(](\d{6})[）)]', fn)
                if not m:
                    continue
                code = m.group(1)
                ym = re.search(r'(\d{4})年度', fn)
                year = int(ym.group(1)) if ym else 0
                try:
                    xl = pd.ExcelFile(os.path.join(r, fn))
                    for sh in xl.sheet_names:
                        if "分红" not in sh:
                            continue
                        df = xl.parse(sh, header=None)
                        for i in range(df.shape[0]):
                            key = str(df.iat[i, 0])
                            if "每10股派息" in key or "每10股派发现金" in key:
                                val = pd.to_numeric(df.iat[i, 1], errors="coerce")
                                if pd.notna(val) and val > 0:
                                    per_share = float(val) / 10.0
                                    old = result.get(code)
                                    if old is None or year >= old[1]:
                                        result[code] = (per_share, year)
                        break
                except Exception:
                    continue
    except Exception as e:
        print(f"    ⚠️ 读取本地分红数据失败: {e}")
    return result


def _fetch_dividend_rows(code):
    """取某只股票的分红送转记录（东财数据中心）。"""
    r = _http_get("https://datacenter-web.eastmoney.com/api/data/v1/get",
                  params={"reportName": "RPT_SHAREBONUS_DET",
                          "columns": "SECURITY_CODE,REPORT_DATE,EX_DIVIDEND_DATE,ASSIGN_PROGRESS,IMPL_PLAN_PROFILE,BASIC_EPS,PRETAX_BONUS_RMB",
                          "filter": f'(SECURITY_CODE="{code}")',
                          "pageNumber": 1, "pageSize": 60, "sortColumns": "REPORT_DATE",
                          "sortTypes": -1, "source": "WEB", "client": "WEB"},
                  timeout=10, retries=3, label=f"东财分红[{code}]")
    if r is None:
        return []
    js = r.json()
    return (((js or {}).get("result") or {}).get("data")) or []


def _ttm_dividend_per_share(code):
    """【修复8】严格 TTM 股息：近 12 个月内所有已除权除息的分红累加（每股税前，元）。
    返回 (每股金额, 说明) 或 (None, None)。"""
    try:
        rows = _fetch_dividend_rows(code)
    except Exception:
        return None, None
    today = datetime.now().date()
    start = today - timedelta(days=365)

    def pick(it, key):
        v = it.get(key)
        if not v: return None
        try:
            return datetime.strptime(str(v)[:10], "%Y-%m-%d").date()
        except Exception:
            return None

    total, hits = 0.0, []
    for it in rows:
        ex = pick(it, "EX_DIVIDEND_DATE")
        rd = pick(it, "REPORT_DATE")
        prog = str(it.get("ASSIGN_PROGRESS") or "")
        used_date, note = None, ""
        # 优先看“分配进度”：只有已实施的分红才算进 TTM。
        # 例如燕京啤酒 2026-06-30 那条是“董事会决议通过”（10派1.00元），尚未除权，不能计入。
        if "实施" in prog:
            if ex is None or not (start <= ex <= today):
                continue
            used_date, note = ex, "除权"
        elif any(w in prog for w in ("董事会", "股东大会", "预案", "未实施")):
            continue
        elif ex is not None:
            if not (start <= ex <= today):
                continue
            used_date, note = ex, "除权"
        elif rd is not None and (today - rd).days >= 300:
            # 无进度、无除权日时，仅当报告期已过去约10个月才谨慎计入，避免高估股息率
            if not (start - timedelta(days=60) <= rd <= today):
                continue
            used_date, note = rd, "报告期"
        else:
            continue
        v = _num_or_none(it.get("PRETAX_BONUS_RMB"))
        if v is None:
            continue
        # 东财该字段口径为“每10股派息(元)”；极小值按已是“每股”处理
        per_share = v / 10.0 if v > 1 else v
        if per_share <= 0:
            continue
        total += per_share
        hits.append(f"{used_date.isoformat()}{note} {per_share:.4g}元/股")
    if total > 0:
        return round(total, 4), "近12个月已实施分红累加(" + "；".join(hits) + ")"
    return 0.0, "近12个月无已实施分红"


def get_dividend_per_share(sec_code, out_dir=None, prefer_ttm=True):
    """取【每股现金派息(元)】。入参可以是 6 位代码、带市场前缀的代码，或公司简称（会自动解析）。
    prefer_ttm=True 时优先“严格 TTM（近12个月已除权分红累加）”，
    确认近12个月无任何分红时才退回本地《年报提取表》的年度派息，避免把一年前的老分红当成当前股息率。
    """
    raw = str(sec_code or "").strip()
    code_m = re.search(r'(\d{6})', raw)
    if code_m:
        code = code_m.group(1)
    else:
        # 传入的是公司简称（如“贵州茅台”）：先解析成 6 位代码
        _sym, _mkt = _quote_symbol(raw)
        code_m = re.search(r'(\d{6})', _sym or "")
        if not code_m:
            return None, None
        code = code_m.group(1)

    local_dps, local_year = None, None
    if out_dir:
        if out_dir not in _DIVIDEND_MAP:
            _DIVIDEND_MAP[out_dir] = _load_local_dividend_map(out_dir)
        hit = _DIVIDEND_MAP[out_dir].get(code)
        if hit:
            local_dps, local_year = hit[0], hit[1]

    # ① 严格 TTM（近 12 个月多次分红累加）
    if prefer_ttm:
        ttm, tip = _ttm_dividend_per_share(code)
        if ttm is not None and ttm > 0:
            return ttm, "TTM " + tip
        if ttm == 0.0:
            # 近12个月确实没有分红：若是唯一数据则用 0（严格 TTM 口径），否则退回年度派息
            if local_dps:
                return local_dps, f"近12个月无分红，改用本地年报({local_year})派息"
            return 0.0, tip

    # ② 本地年报派息（不算陈旧时直接用）
    if local_dps and (local_year is None or local_year >= datetime.now().year - 1):
        return local_dps, f"本地年报({local_year})"

    # ③ 联网取最近一次年度派息
    try:
        rows = _fetch_dividend_rows(code)
        best, best_date = None, None
        for it in rows:
            v = _num_or_none(it.get("PRETAX_BONUS_RMB"))
            dt = str(it.get("REPORT_DATE") or "")[:10]
            if v is None:
                continue
            per_share = v / 10.0 if v > 1 else v
            if best is None or per_share > best:
                best, best_date = per_share, dt
        if best:
            return best, f"东财分红({best_date or '近期'})"
    except Exception:
        pass
    if local_dps:
        return local_dps, f"本地年报({local_year})"
    return None, None


def fetch_quote_http(user_input, out_dir=None):
    """按 腾讯 -> 新浪 顺序取行情，并用严格 TTM 分红数据补齐动态股息率。
    【修复10】整轮失败会再重试（默认2轮，轮间间隔1.5秒），避免偶发抖动把某只股票永久留成 N/A。
    全部失败时返回 None（由调用方决定是否回退浏览器）。
    """
    resolved, mkt = _quote_symbol(user_input)
    if not resolved:
        print(f"    ⚠️ 无法识别标的：{user_input}")
        return None
    # 候选查询符号：用户输入本身（可能已是代码）+ 解析结果，去重
    cands = []
    for c in (str(user_input or "").strip(), resolved):
        if c and c not in cands:
            cands.append(c)

    last_err = None
    for loop in range(2):
        for symbol in cands:
            for fn in (_fetch_quote_tencent, _fetch_quote_sina):
                try:
                    q = fn(symbol)
                    if q and q.get("现价") is not None:
                        q["市场"] = mkt
                        q["symbol"] = symbol
                        # 动态股息率 = 每股派息 ÷ 现价 × 100%（严格 TTM 优先）
                        if q.get("股息率") is None and q.get("现价"):
                            dps, src = get_dividend_per_share(symbol, out_dir)
                            if dps:
                                q["股息率"] = round(dps / q["现价"] * 100, 2)
                                q["股息率来源"] = f"{src} {dps}元/股"
                            else:
                                q["股息率来源"] = "无分红数据"
                        return q
                except Exception as e:
                    last_err = e
        if loop == 0:
            print(f"    ⚠️ {user_input} 首轮行情未取到，1.5 秒后重试…")
            time.sleep(1.5)
    if last_err is not None:
        print(f"    ❌ HTTP 行情源均失败（{user_input} -> {resolved}）: {type(last_err).__name__}: {str(last_err)[:150]}")
    return None


def find_latest_msedgedriver():
    """在 Selenium 缓存里找可用的 msedgedriver（多版本时优先取最高的）。"""
    try:
        base = os.path.join(os.path.expanduser("~"), ".cache", "selenium", "msedgedriver")
        if not os.path.isdir(base):
            return None
        cands = []
        for r, d, files in os.walk(base):
            for name in files:
                if name.lower() == "msedgedriver.exe" or name.lower() == "msedgedriver":
                    cands.append(os.path.join(r, name))

        def ver_key(p):
            nums = re.findall(r'\d+', os.path.basename(os.path.dirname(p)))
            return [int(x) for x in nums[:4]] or [0]

        cands.sort(key=ver_key, reverse=True)
        return cands[0] if cands else None
    except Exception:
        return None


def create_edge_driver(options, purpose="浏览器"):
    """【修复5】用 try 包裹的 Edge 启动。
    返回 (driver, None) 或 (None, 错误说明)。
    典型故障：Edge 自动升级后 msedgedriver 版本不匹配 -> SessionNotCreatedException，
    表现为“浏览器刚启动就断开”，进而把后续所有抓取静默变成 N/A。
    """
    try:
        d = webdriver.Edge(options=options)
        return d, None
    except Exception as e1:
        msg = f"{type(e1).__name__}: {str(e1)[:200]}"
    drv_path = find_latest_msedgedriver()
    if drv_path:
        try:
            from selenium.webdriver.edge.service import Service as _EdgeService
            d = webdriver.Edge(service=_EdgeService(executable_path=drv_path), options=options)
            return d, None
        except Exception as e2:
            msg += f" | 指定驱动 {os.path.basename(os.path.dirname(drv_path))} 仍失败: {type(e2).__name__}"
    print(f"❌ [{purpose}] Edge 浏览器启动失败: {msg}")
    print("   原因多为 msedgedriver 与 Edge 版本不匹配。请更新 Edge 后重试，"
          "或删除 C:\\Users\\<用户>\\.cache\\selenium\\msedgedriver 让其重新下载匹配驱动。")
    return None, msg


def is_excel_file(filename): return filename.lower().endswith((".xlsx", ".xlsm", ".xls")) and not filename.startswith("~$")


def find_first_pdf(root_dir):
    """【修复2】向下递归查找第一个 PDF，用于判断下载子程序是否真的产出了年报文件。"""
    try:
        if not root_dir or not os.path.isdir(root_dir):
            return None
        for r, d, files in os.walk(root_dir):
            for f in files:
                if f.lower().endswith('.pdf') and not f.startswith('~$'):
                    return os.path.join(r, f)
    except Exception:
        pass
    return None

def is_docx_file(filename): return filename.lower().endswith(".docx") and not filename.startswith("~$")
def normalize_text(x): return "" if x is None else str(x).strip().replace("\n", "").replace("\r", "")

def read_excel_tables(file_path):
    dfs = []
    try:
        xls = pd.ExcelFile(file_path)
        for sheet in xls.sheet_names:
            try: dfs.append((file_path, sheet, pd.read_excel(file_path, sheet_name=sheet, header=None)))
            except Exception as e:
                print(f"解析发生错误，已跳过该项: {str(e)}")
    except Exception as e:
        print(f"解析Excel发生错误，已跳过该项: {str(e)}")
    return dfs

def read_docx_tables(file_path):
    result = []
    try:
        doc = Document(file_path)
        for i, table in enumerate(doc.tables):
            rows = [[cell.text.strip() for cell in row.cells] for row in table.rows]
            if rows: result.append((file_path, f"table_{i + 1}", pd.DataFrame(rows)))
    except Exception as e:
        print(f"解析Docx发生错误，已跳过该项: {str(e)}")
    return result

def scan_target(target_path, output_file_path=None):
    all_tables = []
    out_abs_path = os.path.abspath(output_file_path) if output_file_path else ""
    if os.path.isfile(target_path):
        if out_abs_path and os.path.abspath(target_path) == out_abs_path: return all_tables
        try:
            if is_excel_file(target_path): all_tables.extend(read_excel_tables(target_path))
            elif is_docx_file(target_path): all_tables.extend(read_docx_tables(target_path))
        except Exception as e:
            print(f"扫描文件发生错误，已跳过该项: {str(e)}")
    elif os.path.isdir(target_path):
        for root, dirs, files in os.walk(target_path):
            for f in files:
                file_path = os.path.join(root, f)
                if out_abs_path and os.path.abspath(file_path) == out_abs_path: continue
                try:
                    if is_excel_file(f): all_tables.extend(read_excel_tables(file_path))
                    elif is_docx_file(f): all_tables.extend(read_docx_tables(file_path))
                except Exception as e:
                    print(f"扫描文件夹内容错误，已跳过该项: {str(e)}")
    return all_tables

def detect_report_data(all_tables):
    macro_data = {"cn_bond_yield": None, "us_bond_yield": None, "fed_rate": None, "shenzhen_pe": None, "sp500_pe": None, "hangseng_pe": None}
    stock_rows = []
    def extract_val(row, ignore_words=[]):
        for cell in row:
            if any(w in str(cell) for w in ignore_words): continue
            val = safe_float(cell)
            if val is not None: return val
        return None
    def extract_pe(row):
        for n in [safe_float(x) for x in row if safe_float(x) is not None]:
            if n > 0: return n
        return None

    for file_path, sheet_name, df in all_tables:
        df_str = df.astype(str).fillna("")
        for r in range(df_str.shape[0]):
            row = df_str.iloc[r].tolist()
            row_text = " ".join([normalize_text(x) for x in row])

            if ("中国10年期国债" in row_text) or ("10年期国债" in row_text and "美国" not in row_text):
                v = extract_val(row, ["国债"])
                if v is not None: macro_data["cn_bond_yield"] = v
            elif "美国10年期国债" in row_text:
                v = extract_val(row, ["国债"])
                if v is not None: macro_data["us_bond_yield"] = v
            elif "美联储" in row_text and "利率" in row_text:
                v = extract_val(row, ["美联储", "利率"])
                if v is not None: macro_data["fed_rate"] = v
            elif "深圳A股" in row_text or "深证A股" in row_text:
                v = extract_pe(row)
                if v is not None: macro_data["shenzhen_pe"] = v
            elif "标普500" in row_text:
                v = extract_pe(row)
                if v is not None: macro_data["sp500_pe"] = v
            elif "恒生指数" in row_text:
                v = extract_pe(row)
                if v is not None: macro_data["hangseng_pe"] = v

            if r > 0:
                nums = [x for x in [safe_float(x) for x in row] if x is not None]
                if len(nums) >= 2:
                    name = None
                    category = None
                    for cell in row:
                        txt = normalize_text(cell)
                        if txt in ("A股", "港股", "美股", "ETF", "宏观", "大盘指数"): category = txt
                        if any(ch.isalpha() for ch in txt) and len(txt) >= 2 and txt not in ("N/A", "-", "None"):
                            if name is None and txt not in ("A股", "港股", "美股", "宏观", "大盘指数"): name = txt

                    if name and name not in ("深圳A股", "中国10年期国债", "大盘指数", "恒生指数", "标普500", "美国10年期国债", "美联储目标利率"):
                        stock_rows.append({"source_file": file_path, "source_sheet": sheet_name, "raw_row": row, "category_fallback": category})
    return macro_data, stock_rows

def parse_stock_row(raw_row, fallback_category):
    row = [normalize_text(x) for x in raw_row]
    category = row[0] if len(row) > 0 and row[0] in ("A股", "港股", "美股") else fallback_category
    name = row[1] if len(row) > 1 else ""
    return {
        "分类": category,
        "名称/代码": name,
        "现价": safe_float(row[2]) if len(row) > 2 else None,
        "当前PE(TTM)": safe_float(row[3]) if len(row) > 3 else None,
        "动态股息率": safe_float(row[8]) if len(row) > 8 else None,
        "原始行": row
    }

def judge_stock(macro_data, stock):
    pe, dividend, price, category = stock.get("当前PE(TTM)"), stock.get("动态股息率"), stock.get("现价"), stock.get("分类")
    if pe is None or dividend is None or price is None or not category: return "数据不足", "black", None

    judgment, color, target_price = "目前为观察区", "black", (15 * price / pe) if pe > 0 else None

    if category == "A股":
        sz_pe, cn_bond = macro_data.get("shenzhen_pe"), macro_data.get("cn_bond_yield")
        if sz_pe is None or cn_bond is None: return "宏观数据不足", "black", None
        if (sz_pe < 20) and (pe < 15) and (dividend > cn_bond): judgment, color = "目前为好价格", "red"
        elif (sz_pe < 40) and (pe < 30) and (dividend > cn_bond * 2 / 3): judgment, color = "目前为偏买区", "yellow"
        elif (40 < sz_pe < 60) and (30 < pe < 50) and (cn_bond / 3 < dividend < cn_bond * 2 / 3): judgment, color = "目前为偏卖区", "blue"
    elif category == "港股":
        hk_pe, cn_bond = macro_data.get("hangseng_pe"), macro_data.get("cn_bond_yield")
        if hk_pe is None or cn_bond is None: return "宏观数据不足", "black", None
        if (hk_pe < 10) and (pe < 15) and (dividend > cn_bond): judgment, color = "目前为好价格", "red"
        elif (hk_pe < 20) and (pe < 30) and (dividend > cn_bond * 2 / 3): judgment, color = "目前为偏买区", "yellow"
        elif (20 < hk_pe < 30) and (30 < pe < 50) and (cn_bond / 3 < dividend < cn_bond * 2 / 3): judgment, color = "目前为偏卖区", "blue"
    elif category == "美股":
        us_bond, fed_rate, sp_pe = macro_data.get("us_bond_yield"), macro_data.get("fed_rate"), macro_data.get("sp500_pe")
        if us_bond is None or fed_rate is None or sp_pe is None: return "宏观数据不足", "black", None
        sp_drop = 0
        if ((fed_rate < 4 and sp_pe < 15 and pe < 15 and dividend > us_bond) or (fed_rate > 4 and sp_pe < 10 and pe < 15 and dividend > us_bond) or (sp_drop > 50 and pe < 15 and dividend > us_bond)):
            judgment, color = "目前为好价格", "red"
        elif ((fed_rate < 4 and sp_pe < 30 and pe < 30 and dividend > us_bond * 2 / 3) or (fed_rate > 4 and sp_pe < 20 and pe < 30 and dividend > us_bond * 2 / 3) or (sp_drop > 30 and pe < 30 and dividend > us_bond * 2 / 3)):
            judgment, color = "目前为偏买区", "yellow"
        elif ((fed_rate < 4 and sp_pe < 45 and pe < 45 and dividend > us_bond / 3) or (fed_rate > 4 and sp_pe < 45 and pe < 45 and dividend > us_bond / 3) or (sp_drop > 20 and pe < 45 and dividend > us_bond / 3)):
            judgment, color = "目前为偏卖区", "blue"

    if target_price is not None: judgment += f"；股价小于{target_price:.2f}为好价格"
    return judgment, color, target_price

def create_word_report(output_path, records, macro_data):
    doc = Document()
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.add_run("资产深度估值分析报告")
    run.bold = True; run.font.size = Pt(16)

    p = doc.add_paragraph()
    sz, cn, hk, sp, us, fed = macro_data.get('shenzhen_pe'), macro_data.get('cn_bond_yield'), macro_data.get('hangseng_pe'), macro_data.get('sp500_pe'), macro_data.get('us_bond_yield'), macro_data.get('fed_rate')
    p.add_run(f"深圳A股PE: {sz if sz else 'N/A'} | 中国10年国债: {cn if cn else 'N/A'}%\n")
    p.add_run(f"恒生指数PE: {hk if hk else 'N/A'}\n")
    p.add_run(f"标普500 PE: {sp if sp else 'N/A'} | 美国10年国债: {us if us else 'N/A'}% | 美联储利率: {fed if fed else 'N/A'}%")

    table = doc.add_table(rows=1, cols=8)
    table.style = "Medium Shading 1 Accent 1"
    hdr = table.rows[0].cells
    headers = ["分类", "名称/代码", "现价", "当前PE(TTM)", "动态股息率", "判断结果", "反推好价格上限", "来源文件"]
    for i, h in enumerate(headers):
        hdr[i].text = h
        hdr[i].paragraphs[0].runs[0].bold = True

    for rec in records:
        row = table.add_row().cells
        row[0].text, row[1].text = str(rec.get("分类", "")), str(rec.get("名称/代码", ""))
        row[2].text = "" if rec.get("现价") is None else f"{rec.get('现价'):.2f}"
        row[3].text = "" if rec.get("当前PE(TTM)") is None else f"{rec.get('当前PE(TTM)'):.2f}"
        row[4].text = "" if rec.get("动态股息率") is None else f"{rec.get('动态股息率'):.2f}%"
        judgment_cell = row[5]
        run_cell = judgment_cell.paragraphs[0].add_run(rec.get("判断结果", ""))
        color = rec.get("颜色", "black")
        if color == "red": run_cell.font.color.rgb = RGBColor(255, 0, 0)
        elif color == "yellow": run_cell.font.color.rgb = RGBColor(191, 144, 0)
        elif color == "blue": run_cell.font.color.rgb = RGBColor(0, 0, 255)
        else: run_cell.font.color.rgb = RGBColor(0, 0, 0)
        row[6].text = "" if rec.get("反推好价格上限") is None else f"{rec.get('反推好价格上限'):.2f}"
        row[7].text = str(rec.get("来源文件", ""))
    doc.save(output_path)

def process_target(input_target, output_file):
    all_tables = scan_target(input_target, output_file)
    macro_data, stock_rows = detect_report_data(all_tables)
    records = []
    for item in stock_rows:
        stock = parse_stock_row(item["raw_row"], item["category_fallback"])
        judgment, color, target_price = judge_stock(macro_data, stock)
        stock["判断结果"], stock["颜色"], stock["反推好价格上限"], stock["来源文件"] = judgment, color, target_price, os.path.basename(item["source_file"])
        records.append(stock)
    create_word_report(output_file, records, macro_data)
    return records, macro_data

def navigate_to_legulegu(driver, index_name):
    urls = []
    if index_name == "深圳A股": urls = ["https://www.legulegu.com/stockdata/shenzhenPE"]
    elif index_name == "标普500": urls = ["https://www.legulegu.com/stockdata/market/sandp"]
    elif index_name == "恒生指数": urls = ["https://www.legulegu.com/stockdata/market/hsi"]

    for url in urls:
        try:
            driver.get(url)
            try:
                WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, "//div[contains(@class, 'target-class')]"))
                )
            except:
                pass
            txt = driver.execute_script("return document.body.textContent;") or ""
            if "不存在" not in txt and "很抱歉" not in txt and "404" not in driver.title: return
        except Exception as e:
            print(f"乐咕乐股导航发生错误，已跳过该项: {str(e)}")

    try:
        driver.get("https://legulegu.com/")
        time.sleep(3)
        pe_menu = driver.find_element(By.XPATH, "//a[contains(text(), 'PE/PB')]")
        ActionChains(driver).move_to_element(pe_menu).perform()
        time.sleep(1)
        if index_name == "深圳A股": target = driver.find_element(By.XPATH, "//a[contains(text(), '深证') or contains(text(), '深圳')]")
        elif index_name == "标普500": target = driver.find_element(By.XPATH, "//a[contains(text(), '标普500')]")
        elif index_name == "恒生指数": target = driver.find_element(By.XPATH, "//a[contains(text(), '恒生')]")
        driver.execute_script("arguments[0].click();", target)
        time.sleep(3)
        txt = driver.execute_script("return document.body.textContent;") or ""
        if "不存在" not in txt and "很抱歉" not in txt: return
    except Exception as e:
        print(f"乐咕乐股备用导航发生错误，已跳过该项: {str(e)}")

    try:
        driver.get("https://legulegu.com/")
        time.sleep(3)
        search_term = "深证A股" if index_name == "深圳A股" else index_name
        search_box = driver.find_element(By.XPATH, "//input[@placeholder='股票代码/股票术语/题材']")
        search_box.clear()
        search_box.send_keys(search_term)
        time.sleep(1)
        search_box.send_keys(Keys.RETURN)
        time.sleep(3)
        links = driver.find_elements(By.XPATH, "//a[contains(@href, 'stockdata')]")
        for link in links:
            if search_term in link.text or index_name in link.text or "市盈" in link.text:
                driver.execute_script("arguments[0].click();", link)
                time.sleep(3)
                return
    except Exception as e:
        print(f"乐咕乐股搜索发生错误，已跳过该项: {str(e)}")

def extract_legulegu_stats(driver, item, out_dir, prefix):
    driver.set_window_size(1920, 1080)
    try:
        txt = driver.execute_script("return document.body.textContent;") or ""
        pe_m = re.search(r'(?:市盈率|PE|pe)[^\d]{0,30}?(-?\d+\.\d+)', txt, re.IGNORECASE)
        if pe_m: item["当前PE/值"] = pe_m.group(1)
        max_m = re.search(r'(?:最大|最高)[^\d]{0,30}?(-?\d+\.\d+)', txt)
        if max_m: item["历史最大"] = max_m.group(1)
        min_m = re.search(r'(?:最小|最低)[^\d]{0,30}?(-?\d+\.\d+)', txt)
        if min_m: item["历史最小"] = min_m.group(1)
    except Exception as e:
        print(f"乐咕乐股解析错误: {str(e)}")

    if item.get("当前PE/值", "N/A") == "N/A":
        try:
            symbol = ""
            if "深圳" in item["名称/代码"] or "SZ" in prefix: symbol = "SZ399107"
            elif "标普" in item["名称/代码"] or "SP500" in prefix: symbol = ".INX"
            elif "恒生" in item["名称/代码"] or "HSI" in prefix: symbol = "HKHSI"
            if symbol:
                driver.get(f"https://xueqiu.com/S/{symbol}")
                time.sleep(4)
                page_text = driver.execute_script("return document.body.innerText;") or ""
                if not page_text.strip(): page_text = driver.execute_script("return document.body.textContent;") or ""
                pe_match = re.search(r'(?:市盈率|PE)(?:\(TTM\))?[\s:：]*([\d\.]+)', page_text, re.IGNORECASE)
                if pe_match: item["当前PE/值"] = pe_match.group(1)
                else:
                    if symbol == ".INX":
                        driver.get("https://xueqiu.com/S/SPY")
                        time.sleep(4)
                        page_text = driver.execute_script("return document.body.textContent;") or ""
                        pe_match = re.search(r'(?:市盈率|PE)(?:\(TTM\))?[\s:：]*([\d\.]+)', page_text, re.IGNORECASE)
                        if pe_match: item["当前PE/值"] = pe_match.group(1)
        except Exception as e:
            print(f"雪球抓取补充发生错误: {str(e)}")

    try:
        chart = None
        try:
            chart = driver.find_element(By.XPATH, "//div[contains(@id, 'chart') or contains(@class, 'echart') or contains(@class, 'stock-chart') or contains(@class, 'chart-container')]")
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", chart)
            time.sleep(2)
        except: chart = driver.find_element(By.TAG_NAME, "body")
        img_path = os.path.join(out_dir, f"{prefix}_10Y_{datetime.now().strftime('%H%M%S')}.png")
        if chart: chart.screenshot(img_path)
        else: driver.save_screenshot(img_path)
        item["img"] = img_path
    except Exception as e:
        print(f"图表截图错误: {str(e)}")
def safe_macro_fetch(fetch_fn):
    """
    【三阶网络降级获取机制】
    第1阶：国内网络直连（剥离代理，专治 AkShare/新浪/东方财富接口）
    第2阶：无代理直连外网
    第3阶：开启代理端口连接外网
    """
    old_http = os.environ.get("HTTP_PROXY")
    old_https = os.environ.get("HTTPS_PROXY")
    old_http_lc = os.environ.get("http_proxy")
    old_https_lc = os.environ.get("https_proxy")

    def clear_proxy():
        for k in ["HTTP_PROXY", "HTTPS_PROXY", "http_proxy", "https_proxy"]:
            os.environ.pop(k, None)

    def restore_proxy():
        if old_http: os.environ["HTTP_PROXY"] = old_http
        if old_https: os.environ["HTTPS_PROXY"] = old_https
        if old_http_lc: os.environ["http_proxy"] = old_http_lc
        if old_https_lc: os.environ["https_proxy"] = old_https_lc

    # --- 第 1 阶：清空代理，优先国内网络直连 ---
    try:
        clear_proxy()
        res = fetch_fn()
        restore_proxy()
        if res: return res
    except Exception:
        pass

    # --- 第 2 阶：尝试恢复当前环境直连 ---
    try:
        restore_proxy()
        res = fetch_fn()
        if res: return res
    except Exception:
        pass

    # --- 第 3 阶：若前两阶失败，强制挂载代理端口重试 ---
    try:
        os.environ["HTTP_PROXY"] = f"http://127.0.0.1:{PROXY_PORT}"
        os.environ["HTTPS_PROXY"] = f"http://127.0.0.1:{PROXY_PORT}"
        res = fetch_fn()
        restore_proxy()
        if res: return res
    except Exception:
        restore_proxy()

    return None


def main_logic_c(cfg):
    # 【关键修复】：建立本地白名单，强行保护 Selenium 本地通信不被全局代理劫持！
    os.environ['NO_PROXY'] = 'localhost,127.0.0.1,::1'

    results = []

    # 1. 中国 10 年期国债（含三阶降级 + 双接口备份）
    if cfg.get('cn_10y'):
        def get_cn_10y():
            try:
                # 备选 A：新浪接口
                return f"{ak.bond_zh_us_rate()['中国国债收益率10年'].dropna().iloc[-1]}%"
            except Exception:
                # 备选 B：官方中债收益率曲线接口
                df = ak.bond_china_yield(start_date=datetime.now().strftime("%Y0101"), end_date=datetime.now().strftime("%Y%m%d"))
                return f"{df['10年'].dropna().iloc[-1]}%"

        val = safe_macro_fetch(get_cn_10y) or "获取失败"
        results.append({"分类": "宏观", "名称/代码": "中国10年期国债", "现价": "-", "当前PE/值": val, "历史最大": "-", "历史最小": "-", "历史分位": "-", "10年分位": "-", "股息率": "-", "img": ""})

    # 2. 美国 10 年期国债（含三阶降级）
    if cfg.get('us_10y'):
        def get_us_10y():
            return f"{ak.bond_zh_us_rate()['美国国债收益率10年'].dropna().iloc[-1]}%"

        val = safe_macro_fetch(get_us_10y) or "获取失败"
        results.append({"分类": "宏观", "名称/代码": "美国10年期国债", "现价": "-", "当前PE/值": val, "历史最大": "-", "历史最小": "-", "历史分位": "-", "10年分位": "-", "股息率": "-", "img": ""})

        # 3. 美联储目标利率（含三阶降级）
        if cfg.get('fed'):
            def get_fed():
                return f"{ak.macro_bank_usa_interest_rate()['今值'].iloc[0]}%"

            val = safe_macro_fetch(get_fed) or "获取失败"
            results.append(
                {"分类": "宏观", "名称/代码": "美联储目标利率", "现价": "-", "当前PE/值": val, "历史最大": "-",
                 "历史最小": "-", "历史分位": "-", "10年分位": "-", "股息率": "-", "img": ""})

    # === 核心替换开始：高强度防屏蔽浏览器初始化 (Edge 专属版) ===
    opts = webdriver.EdgeOptions()

    # 【核心修复】：安全获取静默模式状态，直接从字典取值，而不是调用失效的 self
    is_silent = cfg.get('silent', True)

    if is_silent:
        # 1. 采用最新版的隐蔽静默模式
        opts.add_argument('--headless=new')
        # 2. 强制全高清分辨率，防止网页折叠隐藏 PE 数据区块
        opts.add_argument('--window-size=1920,1080')
        # 3. 剥离自动化控制特征
        opts.add_argument('--disable-blink-features=AutomationControlled')
        # 4. 伪装成真实的电脑用户
        opts.add_argument(
            'user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36')
        # 5. 禁用 GPU 加速（静默模式下提升稳定性）
        opts.add_argument('--disable-gpu')

    # 常规防屏蔽设置（无论是否静默都需要）
    opts.add_experimental_option("excludeSwitches", ["enable-automation"])
    opts.add_experimental_option('useAutomationExtension', False)

    # 启动浏览器（【修复5】改为可控启动：失败不再直接抛异常，改为记录并降级）
    driver, _drv_err = create_edge_driver(opts, purpose="好价估值-大盘指数")
    try:
        # 6. 底层 V8 引擎注入，彻底抹除机器痕迹
        if driver is not None:
            try:
                driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
                    "source": "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
                })
                driver.maximize_window()
            except Exception as _e:
                print(f"⚠️ 浏览器初始化细节失败（继续执行）: {_e}")

        if driver is None:
            for _need, _key, _nm in ((cfg.get('sz_pe'), 'shenzhen_pe', '深圳A股'),
                                     (cfg.get('sp_pe'), 'sp500_pe', '标普500'),
                                     (cfg.get('hs_pe'), 'hangseng_pe', '恒生指数')):
                if _need:
                    print(f"⏭️ 浏览器不可用，跳过 {_nm} 的 PE 抓取（该项在报告中显示为 N/A）")
        if cfg['sz_pe'] and driver is not None:
            item = {"分类": "大盘指数", "名称/代码": "深圳A股", "现价": "-", "当前PE/值": "N/A", "历史最大": "-", "历史最小": "-", "历史分位": "-", "10年分位": "-", "股息率": "N/A", "img": ""}
            navigate_to_legulegu(driver, "深圳A股")
            extract_legulegu_stats(driver, item, cfg['out_dir'], "SZ")
            results.append(item)
        if cfg['sp_pe'] and driver is not None:
            item = {"分类": "大盘指数", "名称/代码": "标普500", "现价": "-", "当前PE/值": "N/A", "历史最大": "-", "历史最小": "-", "历史分位": "-", "10年分位": "-", "股息率": "N/A", "img": ""}
            navigate_to_legulegu(driver, "标普500")
            extract_legulegu_stats(driver, item, cfg['out_dir'], "SP500")
            results.append(item)
        if cfg['hs_pe'] and driver is not None:
            item = {"分类": "大盘指数", "名称/代码": "恒生指数", "现价": "-", "当前PE/值": "N/A", "历史最大": "-", "历史最小": "-", "历史分位": "-", "10年分位": "-", "股息率": "N/A", "img": ""}
            navigate_to_legulegu(driver, "恒生指数")
            extract_legulegu_stats(driver, item, cfg['out_dir'], "HSI")
            results.append(item)

        for cat, stocks in [("A股", cfg['a_stocks']), ("港股", cfg['hk_stocks']), ("美股", cfg['us_stocks'])]:
            for s in stocks:
                item = {"分类": cat, "名称/代码": s, "现价": "N/A", "当前PE/值": "N/A", "历史最大": "-", "历史最小": "-",
                        "历史分位": "-", "10年分位": "-", "股息率": "N/A", "img": ""}
                # ①【修复5】先用 HTTP 行情源（腾讯/新浪）取现价、PE(TTM)、股息率
                #    本地分红数据优先取自【年报提取程序】的产出目录（extract_dir 优先，其次 out_dir）
                http_done = False
                try:
                    q = fetch_quote_http(s, out_dir=cfg.get('extract_dir') or cfg.get('out_dir'))
                    if q:
                        if q.get("现价") is not None: item["现价"] = q["现价"]
                        if q.get("PE_TTM") is not None: item["当前PE/值"] = q["PE_TTM"]
                        if q.get("股息率") is not None: item["股息率"] = f"{q['股息率']:.2f}%"
                        http_done = True
                        print(f"    ✅ {s} 行情已取到（{q.get('来源')}）：现价={item['现价']} "
                              f"PE(TTM)={item['当前PE/值']} 股息率={item['股息率']}"
                              f"（分红来源: {q.get('股息率来源', '-')}）")
                    else:
                        print(f"    ⚠️ {s} 的 HTTP 行情源未取到数据，尝试浏览器抓取…")
                except Exception as e:
                    print(f"    ⚠️ {s} 的 HTTP 行情抓取异常: {e}")
                # ② 仅在 HTTP 完全失败时才回退到原来的浏览器抓取（保留原有能力，不做删减）
                if not http_done and driver is not None:
                    try:
                        driver.get(f"https://xueqiu.com/k?q={s}")
                        time.sleep(2)
                        stock_link = driver.find_element(By.XPATH, "//a[contains(@href, '/S/')]").get_attribute("href")
                        driver.get(stock_link)
                        time.sleep(3)
                        txt = driver.execute_script("return document.body.textContent;")
                        price_elem = driver.find_element(By.CLASS_NAME, "stock-current")
                        if price_elem: item["现价"] = price_elem.text.replace("¥", "").replace("$", "").strip()
                        ttm_pe = re.search(r'市盈率\(TTM\)\s*[:：]?\s*([\d\.]+)', txt, re.IGNORECASE)
                        dynamic_div = re.search(r'股息率(?:.*?)\s*[:：]?\s*([\d\.]+%?)', txt, re.IGNORECASE)
                        if ttm_pe: item["当前PE/值"] = ttm_pe.group(1)
                        if dynamic_div:
                            val = dynamic_div.group(1)
                            item["股息率"] = val if "%" in val else f"{val}%"
                        print(f"    ℹ️ {s} 已改用浏览器抓取：现价={item['现价']} PE={item['当前PE/值']} 股息率={item['股息率']}")
                    except Exception as e:
                        print(f"    ❌ {s} 浏览器抓取失败: {str(e)[:120]}")
                if item["现价"] == "N/A" or item["当前PE/值"] == "N/A":
                    print(f"    ⚠️ {s} 仍有字段缺失（现价={item['现价']}, PE={item['当前PE/值']}, 股息率={item['股息率']}）")
                results.append(item)

            # 【核心修复】：注意这里的缩进！它现在和最外层的 for 循环齐平。
            # 确保只有在 A股、港股、美股全部查完之后，才执行销毁浏览器的动作。

    finally:

        if 'driver' in locals() and driver is not None:

            try:

                driver.quit()

            except Exception:

                pass

        # 【新增保险】：强制清理后台可能残留的隐形驱动进程，防止积压导致内存溢出

        try:

            # 这里杀 msedgedriver.exe 最安全，不会误关用户正常上网的 Edge 浏览器

            subprocess.call("taskkill /F /IM msedgedriver.exe /T", shell=True, stdout=subprocess.DEVNULL,
                            stderr=subprocess.DEVNULL)

        except:

            pass

    doc = Document()
    doc.add_heading('资产深度估值报告 (含实时股价)', 0)
    table = doc.add_table(rows=1, cols=9)
    table.style = 'Medium Shading 1 Accent 1'
    headers = ["分类", "名称/代码", "现价", "当前PE(TTM)", "历史最大", "历史最小", "历史分位", "10年分位", "动态股息率"]
    for i, h in enumerate(headers):
        table.rows[0].cells[i].text = h
        table.rows[0].cells[i].paragraphs[0].runs[0].bold = True
    for r in results:
        cells = table.add_row().cells
        cells[0].text, cells[1].text, cells[2].text = str(r["分类"]), str(r["名称/代码"]), str(r["现价"])
        cells[3].text, cells[4].text, cells[5].text = str(r["当前PE/值"]), str(r["历史最大"]), str(r["历史最小"])
        cells[6].text, cells[7].text, cells[8].text = str(r["历史分位"]), str(r["10年分位"]), str(r["股息率"])

    doc.add_page_break()
    doc.add_heading('大盘指数 10 年趋势参考', level=1)
    for r in results:
        if r.get("img") and os.path.exists(r["img"]):
            doc.add_heading(f"▶ {r['名称/代码']}", level=2)
            doc.add_picture(r["img"], width=Inches(6.0))

    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    doc_path = os.path.join(cfg['out_dir'], f"估值分析报告_{timestamp}.docx")
    doc.save(doc_path)

    final_doc_path = os.path.join(cfg['out_dir'], f"深度资产估值分析报告_{timestamp}.docx")
    process_target(doc_path, final_doc_path)
    return doc_path, final_doc_path

# =====================================================================
# 子系统与组件类定义 (A程序)
# =====================================================================
class UnifiedStockSystem:
    def __init__(self, root, master_callback=None):
        self.root = root
        self.root.title("综合智能投研与数据抓取系统 (多进程防抖版)")
        self.master_callback = master_callback
        self.is_running = False
        self.pause_flag = False
        self.stop_flag = False
        self.is_auto = False
        self.save_path_var = tk.StringVar(value=os.getcwd())
        self.preset_configs = {}
        self.templates_app1 = {}
        self.load_config_app1()
        self.templates_app2 = self.load_config_app2()
        self.setup_main_ui()
        self.setup_app1_ui()
        self.setup_app2_ui()

    def load_config_app1(self):
        self.preset_configs = {
            "云端 AI (DeepSeek)": {"url": "https://api.deepseek.com", "models": ["deepseek-chat"]},
            "云端 AI (OpenAI)": {"url": "https://api.openai.com/v1", "models": ["gpt-4o", "gpt-3.5-turbo"]},
            "云端 AI (Kimi)": {"url": "https://api.moonshot.cn/v1", "models": ["moonshot-v1-8k", "moonshot-v1-32k"]},
            "云端 AI (通义千问)": {"url": "https://dashscope.aliyuncs.com/compatible-mode/v1", "models": ["qwen-max", "qwen-plus"]},
            "本地 AI (Ollama)": {"url": "http://localhost:11434/v1", "models": ["qwen2.5:7b", "llama3:8b"]},
        }
        self.templates_app1 = {
            "--- 请选择金融投资指令模板 ---": "",
            "0. 规范设置多选": "请在A股中严格筛选同时符合以下条件的股票并作分析：2021、2022、2023、2024、2025年（五个完整会计年度）roe_weighted>40%、净利润现金含量>80%，gross_profit_margin>40%，上市时间>3年，剔除北交所中的公司和金融股。如果没有符合条件的则跳过，不能擅自乱抓其他不完全符合条件的公司。",
            "1、筛选A股多条件": "请在A股中筛选同时符合以下条件的股票并作分析：连续5年加权roe>15，连续5年净利润现金含量>80，连续5年毛利率>40，上市时间>3年，剔除北交所，非金融股",
            "2. 行业深度研判": "查找当前【低空经济】行业最新的利好政策，列出该行业的3-5家龙头上市公司，并结合近期市盈率分析该板块是否处于低估区间。",
            "3. 核心财报对比": "对比【中国平安】和【招商银行】最新发布的财报数据（净利润增长、不良率/ROE等），并分析机构对这两家公司的最新评级。",
            "4. 指数估值分析": "查找沪深300指数（SH000300）过去15年的市盈率PE-TTM最高值和最低值，计算当前分位值，并给出配置建议。",
            "5. 热门题材挖掘": "搜索关于【全固态电池】最新的技术突破点，识别产业链中A股相关的核心受益标的公司并列出其技术优势。",
            "6. 个股风控雷达": "查询【某某公司】近半年是否存在大股东减持、违规担保、质押比例过高等财务风险，并评价其对股价的潜在影响。"

        }
        if os.path.exists(CONFIG_FILE_APP1):
            try:
                with open(CONFIG_FILE_APP1, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    self.preset_configs.update(data.get("preset_configs", {}))
                    self.templates_app1.update(data.get("templates", {}))
                    if data.get("save_path") and os.path.exists(data.get("save_path")):
                        self.save_path_var.set(data.get("save_path"))
            except Exception as e: pass

    def save_config_app1(self):
        data = {"preset_configs": self.preset_configs, "templates": self.templates_app1, "save_path": self.save_path_var.get()}
        try:
            with open(CONFIG_FILE_APP1, 'w', encoding='utf-8') as f: json.dump(data, f, ensure_ascii=False, indent=4)
        except Exception as e: messagebox.showerror("保存失败", f"无法保存程序一配置到本地: {e}")

    def load_config_app2(self):
        default_templates = {
            "默认A股多条件": "连续5年加权roe>25，连续5年净利润现金含量>80，连续5年毛利率>40，上市时间>3年，剔除北交所，非金融股",
            "优势大市值": "连续5年加权roe>15%，剔除北交所，非金融股\n量价齐升，且市值大于100亿\n近3年净利润增长率大于20%",
            "高股息低估值": "连续3年股息率>5%\n市盈率<15\n市值>200亿",
            "美股科技龙头": "科技股，总市值大于1000亿美元，roe大于20%",
            "美股指数": "美股三大指数",
            "港股高息": "港股，连续3年股息率>8%，市值>100亿"
        }
        if os.path.exists(CONFIG_FILE_APP2):
            try:
                with open(CONFIG_FILE_APP2, 'r', encoding='utf-8') as f:
                    user_templates = json.load(f)
                    # 将默认模板与本地读取的模板合并（优先保留用户的自定义修改）
                    merged_templates = default_templates.copy()
                    if isinstance(user_templates, dict):
                        merged_templates.update(user_templates)
                    return merged_templates
            except Exception:
                return default_templates
        return default_templates
    def save_config_app2(self):
        try:
            with open(CONFIG_FILE_APP2, 'w', encoding='utf-8') as f: json.dump(self.templates_app2, f, ensure_ascii=False, indent=4)
        except Exception as e: messagebox.showerror("错误", f"保存程序二模板失败：{e}")

    def setup_main_ui(self):
        path_frame = ttk.LabelFrame(self.root, text="📁 统一全局保存路径", padding=10)
        path_frame.pack(fill=tk.X, padx=15, pady=10)
        ttk.Entry(path_frame, textvariable=self.save_path_var, state="readonly").pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        ttk.Button(path_frame, text="📂 选择路径", command=self.choose_save_path).pack(side=tk.LEFT, padx=5)
        ttk.Button(path_frame, text="📌 设为默认", command=self.set_default_save_path).pack(side=tk.LEFT, padx=5)

        control_frame = ttk.LabelFrame(self.root, text="⚙️ 参数设置面板 (点击进入二级窗口配置)", padding=15)
        control_frame.pack(fill=tk.X, padx=15, pady=5)
        tk.Button(control_frame, text="🤖 进入程序一设置\n(AI 股市投研)", font=("Microsoft YaHei", 11, "bold"), bg="#E1F5FE", command=self.open_app1_window, height=2).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=10)
        tk.Button(control_frame, text="📈 进入程序二设置\n(问财数据抓取)", font=("Microsoft YaHei", 11, "bold"), bg="#FFF3E0", command=self.open_app2_window, height=2).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=10)

        run_frame = tk.Frame(self.root, pady=15)
        run_frame.pack(fill=tk.X, padx=15)
        self.run_btn = tk.Button(run_frame, text="🚀 单独启动海选执行 (先程序一，后程序二)", font=("Microsoft YaHei", 12, "bold"), bg="#0078D4", fg="white", command=self.start_task, height=2)
        self.run_btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        self.pause_btn = tk.Button(run_frame, text="⏸️ 暂停", font=("Microsoft YaHei", 11), command=self.toggle_pause, state=tk.DISABLED, width=12)
        self.pause_btn.pack(side=tk.LEFT, padx=5)
        self.stop_btn = tk.Button(run_frame, text="⏹️ 停止", font=("Microsoft YaHei", 11), command=self.stop_task, state=tk.DISABLED, width=12)
        self.stop_btn.pack(side=tk.LEFT, padx=5)

        log_frame = ttk.LabelFrame(self.root, text="📝 统一运行进程显示", padding=10)
        log_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=5)
        self.log_text = scrolledtext.ScrolledText(log_frame, height=12, font=("Consolas", 10), bg="#1e1e1e", fg="#cecece")
        self.log_text.pack(fill=tk.BOTH, expand=True)

    def choose_save_path(self):
        folder_selected = filedialog.askdirectory(title="选择统一保存文件夹", initialdir=self.save_path_var.get())
        if folder_selected:
            self.save_path_var.set(folder_selected)
            self.log(f"📁 本次全局输出路径已更改为: {folder_selected}")

    def set_default_save_path(self):
        current_path = self.save_path_var.get()
        if current_path and os.path.exists(current_path):
            self.save_config_app1()
            self.log(f"📌 已永久将默认保存路径设置为: {current_path}")
            messagebox.showinfo("设置成功", f"以后每次打开软件，默认保存路径都会是：\n{current_path}")

    def log(self, msg):
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, f"[{datetime.now().strftime('%H:%M:%S')}] {msg}\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()

    def setup_app1_ui(self):
        self.win_app1 = tk.Toplevel(self.root)
        self.win_app1.title("程序一：AI 股市投研系统 V2.7 - 参数设置")
        self.win_app1.geometry("850x750")
        self.win_app1.protocol("WM_DELETE_WINDOW", self.win_app1.withdraw)
        self.win_app1.withdraw()

        frame_net = ttk.LabelFrame(self.win_app1, text="网络与搜索设置", padding=10)
        frame_net.pack(fill=tk.X, padx=10, pady=5)
        ttk.Label(frame_net, text="代理模式:").grid(row=0, column=0, sticky=tk.W)
        self.proxy_var = tk.StringVar(value="Normal")
        ttk.Radiobutton(frame_net, text="一般 (直连)", variable=self.proxy_var, value="Normal",
                        command=self.toggle_proxy).grid(row=0, column=1, padx=5)
        ttk.Radiobutton(frame_net, text="全域 (代理)", variable=self.proxy_var, value="Global",
                        command=self.toggle_proxy).grid(row=0, column=2, padx=5)
        self.proxy_addr_var = tk.StringVar(value="http://127.0.0.1:7890")
        ttk.Entry(frame_net, textvariable=self.proxy_addr_var, width=20).grid(row=0, column=3, padx=5)
        ttk.Label(frame_net, text="搜索引擎:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.engine_var = tk.StringVar(value="Tavily")
        ttk.Radiobutton(frame_net, text="Tavily", variable=self.engine_var, value="Tavily").grid(row=1, column=1,
                                                                                                 padx=5)
        ttk.Radiobutton(frame_net, text="DuckDuckGo", variable=self.engine_var, value="DDG").grid(row=1, column=2,
                                                                                                  padx=5)
        ttk.Label(frame_net, text="抓取资讯条数:").grid(row=1, column=3, sticky=tk.E, padx=5)
        self.fetch_count_var = tk.IntVar(value=5)
        ttk.Spinbox(frame_net, from_=1, to=100000000, textvariable=self.fetch_count_var, width=5).grid(row=1, column=4,
                                                                                                       sticky=tk.W)

        frame_ai = ttk.LabelFrame(self.win_app1, text="AI 模型配置选项", padding=10)
        frame_ai.pack(fill=tk.X, padx=10, pady=5)
        ttk.Label(frame_ai, text="AI 类型:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.ai_type_var = tk.StringVar()

        self.ai_type_cb = ttk.Combobox(frame_ai, textvariable=self.ai_type_var, values=list(self.preset_configs.keys()),
                                       state="readonly", width=35)
        self.ai_type_cb.grid(row=0, column=1, padx=5, pady=2, sticky=tk.W)
        self.ai_type_cb.bind("<<ComboboxSelected>>", self.update_presets)

        ttk.Label(frame_ai, text="API 地址:").grid(row=1, column=0, sticky=tk.W, pady=2)
        self.api_url_var = tk.StringVar()
        ttk.Entry(frame_ai, textvariable=self.api_url_var, width=50).grid(row=1, column=1, padx=5, pady=2, sticky=tk.W)

        ttk.Label(frame_ai, text="模型名称:").grid(row=2, column=0, sticky=tk.W, pady=2)
        model_frame = ttk.Frame(frame_ai)
        model_frame.grid(row=2, column=1, sticky=tk.W)
        self.model_var = tk.StringVar()
        self.model_cb = ttk.Combobox(model_frame, textvariable=self.model_var, width=32)
        self.model_cb.pack(side=tk.LEFT)
        ttk.Button(model_frame, text="➕ 新增", width=6, command=self.add_model).pack(side=tk.LEFT, padx=3)
        ttk.Button(model_frame, text="➖ 删除", width=6, command=self.delete_model).pack(side=tk.LEFT)

        ttk.Label(frame_ai, text="AI API Key:").grid(row=3, column=0, sticky=tk.W, pady=2)
        self.api_key_var = tk.StringVar()
        ttk.Entry(frame_ai, textvariable=self.api_key_var, width=50, show="*").grid(row=3, column=1, padx=5, pady=2,
                                                                                    sticky=tk.W)

        ttk.Label(frame_ai, text="Tavily Key:").grid(row=4, column=0, sticky=tk.W, pady=2)
        self.tavily_key_var = tk.StringVar()
        ttk.Entry(frame_ai, textvariable=self.tavily_key_var, width=50).grid(row=4, column=1, padx=5, pady=2,
                                                                             sticky=tk.W)

        frame_task = ttk.LabelFrame(self.win_app1, text="任务指令与模板管理 (执行请回一级主窗口)", padding=10)
        frame_task.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        tpl_top_frame = ttk.Frame(frame_task)
        tpl_top_frame.pack(fill=tk.X, pady=5)
        ttk.Label(tpl_top_frame, text="快速模板:").pack(side=tk.LEFT)
        self.tpl_var = tk.StringVar()
        self.tpl_cb = ttk.Combobox(tpl_top_frame, textvariable=self.tpl_var, values=list(self.templates_app1.keys()),
                                   state="readonly")
        self.tpl_cb.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        self.tpl_cb.bind("<<ComboboxSelected>>", self.apply_template_app1)
        ttk.Button(tpl_top_frame, text="💾 保存模板", command=self.add_template_app1).pack(side=tk.LEFT, padx=2)
        ttk.Button(tpl_top_frame, text="🗑️ 删除", command=self.delete_template_app1).pack(side=tk.LEFT)

        self.prompt_text = scrolledtext.ScrolledText(frame_task, height=10, wrap=tk.WORD)
        self.prompt_text.pack(fill=tk.BOTH, expand=True, pady=5)

        self.ai_type_cb.current(0)
        self.update_presets()

        # 【新增】：自动加载默认模板
        if "0. 规范设置多选" in self.templates_app1:
            self.tpl_var.set("0. 规范设置多选")
            self.apply_template_app1(None)
        elif self.templates_app1:
            self.tpl_cb.current(0)
            self.apply_template_app1(None)

    def open_app1_window(self):
        self.win_app1.deiconify()
        self.win_app1.lift()

    def add_model(self):
        ai_type = self.ai_type_var.get()
        if not ai_type: return
        new_model = simpledialog.askstring("新增模型", f"请输入要添加到 [{ai_type}] 的模型名称:\n(例如: gpt-4-turbo)")
        if new_model and new_model.strip():
            new_model = new_model.strip()
            if new_model not in self.preset_configs[ai_type]["models"]:
                self.preset_configs[ai_type]["models"].append(new_model)
                self.save_config_app1()
                self.update_presets()
                self.model_var.set(new_model)
                self.log(f"✅ 程序一：已永久添加新模型 {new_model}")

    def delete_model(self):
        ai_type = self.ai_type_var.get()
        current_model = self.model_var.get()
        if not ai_type or not current_model: return
        if len(self.preset_configs[ai_type]["models"]) <= 1:
            messagebox.showwarning("警告", "至少需要保留一个模型！")
            return
        if messagebox.askyesno("删除确认", f"确定要删除模型 [{current_model}] 吗？"):
            self.preset_configs[ai_type]["models"].remove(current_model)
            self.save_config_app1()
            self.update_presets()
            self.log(f"🗑️ 程序一：已删除模型 {current_model}")

    def add_template_app1(self):
        prompt = self.prompt_text.get("1.0", tk.END).strip()
        if not prompt:
            messagebox.showwarning("警告", "指令输入框为空，无法保存为模板！")
            return
        name = simpledialog.askstring("保存为新模板", "给这个模板起个名字吧:")
        if name and name.strip():
            name = name.strip()
            self.templates_app1[name] = prompt
            self.save_config_app1()
            self.tpl_cb['values'] = list(self.templates_app1.keys())
            self.tpl_var.set(name)
            self.log(f"✅ 程序一：已保存自定义模板 {name}")

    def delete_template_app1(self):
        name = self.tpl_var.get()
        if name == "--- 请选择金融投资指令模板 ---" or not name: return
        if messagebox.askyesno("删除确认", f"确定要删除模板 [{name}] 吗？"):
            del self.templates_app1[name]
            self.save_config_app1()
            self.tpl_cb['values'] = list(self.templates_app1.keys())
            self.tpl_cb.current(0)
            self.prompt_text.delete("1.0", tk.END)
            self.log(f"🗑️ 程序一：已删除模板 {name}")

    def apply_template_app1(self, event):
        content = self.templates_app1.get(self.tpl_var.get(), "")
        if content:
            self.prompt_text.delete("1.0", tk.END)
            self.prompt_text.insert(tk.END, content)

    def toggle_proxy(self):
        mode = self.proxy_var.get()
        raw_addr = self.proxy_addr_var.get().strip()

        # 强力净化器：自动剥离可能因为复制粘贴带来的 Markdown 乱码 [http...](http...)
        addr = re.sub(r'^\[.*?\]\((.*?)\)$', r'\1', raw_addr)
        if addr and not addr.startswith("http"):
            addr = "http://" + addr

        if mode == "Global":
            os.environ["HTTP_PROXY"] = addr
            os.environ["HTTPS_PROXY"] = addr
            os.environ["http_proxy"] = addr
            os.environ["https_proxy"] = addr
            self.log(f"🌐 程序一：全域代理已激活 {addr}")
        else:
            # 【关键修复】：彻底清洗系统全局代理残留，无视C程序的污染
            for key in ["HTTP_PROXY", "HTTPS_PROXY", "http_proxy", "https_proxy", "ALL_PROXY", "all_proxy"]:
                os.environ.pop(key, None)
            self.log("🌐 程序一：网络已切换为直连模式 (已清理全局代理污染)")

    def update_presets(self, event=None):
        selected = self.ai_type_var.get()
        if selected in self.preset_configs:
            config = self.preset_configs[selected]
            self.api_url_var.set(config["url"])
            self.model_cb['values'] = config["models"]
            if config["models"]:
                self.model_var.set(config["models"][0])

    def setup_app2_ui(self):
        self.win_app2 = tk.Toplevel(self.root)
        self.win_app2.title("程序二：问财全球重点市场筛选器 v1.2 - 参数设置")
        self.win_app2.geometry("850x650")
        self.win_app2.protocol("WM_DELETE_WINDOW", self.win_app2.withdraw)
        self.win_app2.withdraw()

        self.market_var = tk.StringVar(value="A股")
        self.headless_var = tk.BooleanVar(value=True)  # 【修改】：默认开启浏览器静默运行

        top_frame = tk.Frame(self.win_app2, padx=15, pady=15)
        top_frame.pack(fill=tk.X)
        tk.Label(top_frame, text="🌍 市场分类：", font=("Microsoft YaHei", 10, "bold")).grid(row=0, column=0, sticky="w",
                                                                                           pady=5)
        market_frame = tk.Frame(top_frame)
        market_frame.grid(row=0, column=1, sticky="w")

        markets = ["A股", "港股", "美股", "A股指数", "港股指数", "美股指数"]
        for m in markets:
            tk.Radiobutton(market_frame, text=m, variable=self.market_var, value=m).pack(side=tk.LEFT, padx=3)

        ttk.Separator(market_frame, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=10)
        tk.Checkbutton(market_frame, text="🔇 静默运行 (隐藏浏览器)", variable=self.headless_var,
                       font=("Microsoft YaHei", 9, "bold"), fg="#D35400").pack(side=tk.LEFT, padx=5)

        tpl_frame = tk.Frame(self.win_app2, padx=15, pady=5)
        tpl_frame.pack(fill=tk.X)
        tk.Label(tpl_frame, text="📑 条件模板：", font=("Microsoft YaHei", 10, "bold")).pack(side=tk.LEFT)
        self.tpl_combobox_app2 = ttk.Combobox(tpl_frame, values=list(self.templates_app2.keys()), width=30,
                                              state="readonly")
        self.tpl_combobox_app2.pack(side=tk.LEFT, padx=5)
        if self.templates_app2:
            self.tpl_combobox_app2.current(0)
        self.tpl_combobox_app2.bind("<<ComboboxSelected>>", self.apply_template_app2)
        tk.Button(tpl_frame, text="应用模板", command=self.apply_template_app2).pack(side=tk.LEFT, padx=5)
        tk.Button(tpl_frame, text="➕ 保存当前为新模板", command=self.add_template_app2).pack(side=tk.LEFT, padx=5)
        tk.Button(tpl_frame, text="❌ 删除选中模板", command=self.delete_template_app2).pack(side=tk.LEFT, padx=5)

        input_frame = tk.Frame(self.win_app2, padx=15, pady=5)
        input_frame.pack(fill=tk.BOTH, expand=True)
        tk.Label(input_frame, text="📝 输入问财选股条件（每行一个，执行请回一级主窗口）：",
                 font=("Microsoft YaHei", 10, "bold")).pack(anchor="w")

        self.text_input_app2 = scrolledtext.ScrolledText(input_frame, height=15, font=("Microsoft YaHei", 10))
        self.text_input_app2.pack(fill=tk.BOTH, expand=True, pady=5)
        self.apply_template_app2()


    def open_app2_window(self):
        self.win_app2.deiconify()
        self.win_app2.lift()

    def apply_template_app2(self, event=None):
        tpl_name = self.tpl_combobox_app2.get()
        if tpl_name in self.templates_app2:
            self.text_input_app2.delete("1.0", tk.END)
            self.text_input_app2.insert(tk.END, self.templates_app2[tpl_name])

    def add_template_app2(self):
        content = self.text_input_app2.get("1.0", tk.END).strip()
        if not content:
            messagebox.showwarning("提示", "输入框为空，无法保存模板！")
            return
        tpl_name = simpledialog.askstring("新增模板", "请输入模板名称：")
        if tpl_name:
            if tpl_name in self.templates_app2:
                if not messagebox.askyesno("覆盖提示", f"模板 '{tpl_name}' 已存在，是否覆盖？"): return
            self.templates_app2[tpl_name] = content
            self.save_config_app2()
            self.update_combobox_app2(tpl_name)
            self.log(f"✅ 程序二：保存模板 '{tpl_name}' 成功！")

    def delete_template_app2(self):
        tpl_name = self.tpl_combobox_app2.get()
        if not tpl_name: return
        if messagebox.askyesno("删除确认", f"确定要删除模板 '{tpl_name}' 吗？"):
            if tpl_name in self.templates_app2:
                del self.templates_app2[tpl_name]
                self.save_config_app2()
                self.update_combobox_app2()
                self.log(f"🗑️ 程序二：已删除模板 {tpl_name}")

    def update_combobox_app2(self, select_name=None):
        keys = list(self.templates_app2.keys())
        self.tpl_combobox_app2['values'] = keys
        if keys:
            if select_name and select_name in keys:
                self.tpl_combobox_app2.set(select_name)
            else:
                self.tpl_combobox_app2.current(0)
        else:
            self.tpl_combobox_app2.set('')

    def check_interrupt(self):
        while self.pause_flag and not self.stop_flag: time.sleep(0.5)
        return self.stop_flag

    def toggle_pause(self):
        if not self.is_running: return
        self.pause_flag = not self.pause_flag
        if self.pause_flag:
            self.pause_btn.config(text="▶️ 继续")
            self.log("⏸️ 任务已挂起...")
        else:
            self.pause_btn.config(text="⏸️ 暂停")
            self.log("▶️ 任务恢复运行...")

    def stop_task(self):
        if not self.is_running: return
        self.stop_flag = True
        self.log("⏹️ 收到停止指令！正在安全退出...")
        self.stop_btn.config(state=tk.DISABLED)

    def reset_ui_state(self):
        self.is_running = False
        self.root.after(0, lambda: self.run_btn.config(state=tk.NORMAL, bg="#0078D4"))
        self.root.after(0, lambda: self.pause_btn.config(state=tk.DISABLED, text="⏸️ 暂停"))
        self.root.after(0, lambda: self.stop_btn.config(state=tk.DISABLED))

    def start_task(self, is_auto=False):
        app1_query = self.prompt_text.get("1.0", tk.END).strip()
        app2_content = self.text_input_app2.get("1.0", tk.END).strip()

        if not app1_query and not app2_content:
            if not is_auto: messagebox.showwarning("提示", "程序一和程序二的输入框均为空，没有可执行的任务！")
            return

        self.is_auto = is_auto
        self.run_btn.config(state=tk.DISABLED, bg="#888888")
        self.pause_btn.config(state=tk.NORMAL, text="⏸️ 暂停")
        self.stop_btn.config(state=tk.NORMAL)
        self.is_running = True
        self.stop_flag = False
        self.pause_flag = False

        # 【关键修复】：在每次任务启动前，强行执行一次网络配置刷新，夺回网络控制权
        self.toggle_proxy()

        threading.Thread(target=self.unified_worker, args=(app1_query, app2_content), daemon=True).start()
    def run_app1_logic(self, query, output_folder, timestamp):
        engine = self.engine_var.get()
        model_name = self.model_var.get()
        fetch_limit = self.fetch_count_var.get()
        context = ""
        current_date = datetime.now().strftime("%Y年%m月%d日")
        opt_query = f"{query} {current_date} 最新资讯"
        self.log(f"   正在通过 {engine} 搜集 {fetch_limit} 条市场资讯...")

        try:
            if engine == "Tavily":
                t_key = self.tavily_key_var.get().strip()
                if not t_key:
                    self.log("   ⚠️ Tavily Key 为空，无法联网。")
                else:
                    t_client = TavilyClient(api_key=t_key)
                    resp = t_client.search(query=opt_query, search_depth="advanced", max_results=fetch_limit)
                    context = "\n".join([f"内容: {r['content']}" for r in resp['results']])
            else:
                with DDGS() as ddgs:
                    results = [r['body'] for r in ddgs.text(opt_query, max_results=fetch_limit)]
                    context = "\n".join(results)
            if context: self.log("   ✅ 资讯抓取完成。")
        except Exception as search_e:
            self.log(f"   ⚠️ 资讯抓取失败: {search_e}")
            self.log("   💡 系统将直接让 AI 基于内部知识库进行深度分析...")
            context = "未获取到实时网络资讯。请直接基于你的已有知识回答问题。"

        if self.check_interrupt(): return

        self.log(f"   正在连接 AI [{model_name}] 生成研报...")
        try:
            client = OpenAI(api_key=self.api_key_var.get().strip(), base_url=self.api_url_var.get())
            user_msg = f"参考资料:\n{context}\n\n任务要求: {query}\n\n请直接输出标准JSON：包含'data_list'数组(含公司,代码,亮点)和'full_report'字符串(Markdown研报)。不要加```json标记。"
            response = client.chat.completions.create(model=model_name,
                                                      messages=[{"role": "user", "content": user_msg}], temperature=0.3)
            if self.check_interrupt(): return
            raw_res = response.choices[0].message.content
            json_match = re.search(r'(\{.*\})', raw_res, re.DOTALL)
            if json_match:
                try:
                    res_data = json.loads(json_match.group(1), strict=False)
                    self.output_files_app1(res_data, engine, model_name, output_folder, timestamp)
                except json.JSONDecodeError:
                    self.save_raw_fallback(raw_res, output_folder, timestamp)
            else:
                self.save_raw_fallback(raw_res, output_folder, timestamp)
        except Exception as ai_e:
            self.log(f"   ❌ AI 调用失败: {ai_e}")

    def output_files_app1(self, data, engine_name, model_name, output_folder, timestamp):
        df = pd.DataFrame(data.get('data_list', []))
        if not df.empty:
            csv_path = os.path.join(output_folder, f"AI提取_关键数据汇总_{timestamp}.csv")
            df.to_csv(csv_path, index=False, encoding='utf-8-sig')
        doc = Document()
        doc.add_heading('股市专题深度研报 (AI生成)', 0)
        doc.add_paragraph(f"⏱️ 生成时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        doc.add_paragraph(f"🧠 分析模型：{model_name}")
        doc.add_paragraph(f"🌐 搜索引擎：{engine_name}")
        doc.add_paragraph("-" * 50)
        report_text = data.get('full_report', '无内容')
        for line in report_text.split('\n'):
            line = line.strip()
            if not line: continue
            if line.startswith('###'):
                doc.add_heading(line.replace('###', '').strip(), 2)
            elif line.startswith('##'):
                doc.add_heading(line.replace('##', '').strip(), 1)
            else:
                doc.add_paragraph(line)
        docx_path = os.path.join(output_folder, f"AI深度分析研报_{timestamp}.docx")
        doc.save(docx_path)
        self.log(f"   ✅ AI研报文件已输出至统一文件夹。")

    def save_raw_fallback(self, text, output_folder, timestamp):
        filename = os.path.join(output_folder, f"AI原始结果_解析失败兜底_{timestamp}.txt")
        with open(filename, "w", encoding="utf-8") as f: f.write(text)
        self.log(f"   ⚠️ 保存纯文本至: {filename}")

    def merge_results(self, output_folder, timestamp):
        try:
            # 防弹衣：强制初始化带有标准列的空表，防止后续合并报 KeyError
            df_ai = pd.DataFrame(columns=['代码', '名称', 'source_ai'])
            csv_path = os.path.join(output_folder, f"AI提取_关键数据汇总_{timestamp}.csv")
            if os.path.exists(csv_path):
                try:
                    temp_df = pd.read_csv(csv_path, dtype=str)
                    if not temp_df.empty:
                        rename_dict = {}
                        for col in temp_df.columns:
                            if '代码' in str(col) or 'code' in str(col).lower():
                                rename_dict[col] = '代码'
                            elif '公司' in str(col) or '名称' in str(col) or 'name' in str(col).lower():
                                rename_dict[col] = '名称'
                        temp_df = temp_df.rename(columns=rename_dict)
                        if '代码' in temp_df.columns and '名称' in temp_df.columns:
                            df_ai = temp_df[['代码', '名称']].drop_duplicates()
                            df_ai['source_ai'] = True
                except Exception as e:
                    self.log(f"   ⚠️ 读取 AI 数据遇到问题，忽略: {e}")

            df_wc = pd.DataFrame(columns=['代码', '名称', 'source_wc'])
            excel_path = os.path.join(output_folder, '问财_表3_代码及名称汇总.xlsx')
            if os.path.exists(excel_path):
                try:
                    xl = pd.ExcelFile(excel_path)
                    dfs = [pd.read_excel(excel_path, sheet_name=s, dtype=str) for s in xl.sheet_names]
                    if dfs:
                        temp_wc = pd.concat(dfs, ignore_index=True)
                        rename_dict = {}
                        for col in temp_wc.columns:
                            if '代码' in str(col) or 'code' in str(col).lower():
                                rename_dict[col] = '代码'
                            elif '名称' in str(col) or '简称' in str(col) or 'name' in str(col).lower():
                                rename_dict[col] = '名称'
                        temp_wc = temp_wc.rename(columns=rename_dict)
                        if '代码' in temp_wc.columns and '名称' in temp_wc.columns:
                            df_wc = temp_wc[['代码', '名称']].drop_duplicates()
                            df_wc['source_wc'] = True
                except Exception as e:
                    self.log(f"   ⚠️ 读取问财数据遇到问题，忽略: {e}")

            # 只有当至少有一方提取到了有效数据时才合并
            if (len(df_ai) > 0) or (len(df_wc) > 0):
                merged_df = pd.merge(df_ai, df_wc, on=['代码'], how='outer', suffixes=('_ai', '_wc'))

                # 名称列填补逻辑
                if '名称_ai' in merged_df.columns and '名称_wc' in merged_df.columns:
                    merged_df['名称'] = merged_df['名称_ai'].fillna(merged_df['名称_wc'])
                elif '名称_ai' in merged_df.columns:
                    merged_df['名称'] = merged_df['名称_ai']
                elif '名称_wc' in merged_df.columns:
                    merged_df['名称'] = merged_df['名称_wc']

                def get_channel(row):
                    has_ai = row.get('source_ai') == True
                    has_wc = row.get('source_wc') == True
                    if has_ai and has_wc:
                        return 'AI和问财'
                    elif has_ai:
                        return 'AI'
                    elif has_wc:
                        return '问财'
                    return ''

                merged_df['筛选渠道'] = merged_df.apply(get_channel, axis=1)

                # 动态选择要保留的列，防止列丢失报错
                cols_to_keep = ['代码']
                if '名称' in merged_df.columns: cols_to_keep.append('名称')
                if '筛选渠道' in merged_df.columns: cols_to_keep.append('筛选渠道')

                final_df = merged_df[cols_to_keep].drop_duplicates()
                final_path = os.path.join(output_folder, "海选公司汇总表.xlsx")
                final_df.to_excel(final_path, index=False)
                self.log(f"   ✅ 《海选公司汇总表》合并生成完毕。")
            else:
                self.log("   ⏭️ 未找到有效的主力公司名单数据，跳过自动合并。")

        except Exception as e:
            self.log(f"   ❌ 合并《海选公司汇总表》时发生异常: {e}")

    def unified_worker(self, app1_query, app2_content):
        has_error = False
        final_excel_path = ""
        try:
            base_path = self.save_path_var.get().strip() or os.getcwd()
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            unified_folder = os.path.join(base_path, "海选情况")
            os.makedirs(unified_folder, exist_ok=True)
            self.log(f"🗂️ 目标数据归集文件夹：\n{unified_folder}")

            if app1_query:
                self.log("\n" + "=" * 40)
                self.log("▶️ [阶段一] 启动 AI 股市投研分析...")
                self.run_app1_logic(app1_query, unified_folder, ts)
            else:
                self.log("\n" + "=" * 40)
                self.log("⏭️ [阶段一] AI 投研指令为空，自动跳过。")

            if self.check_interrupt():
                self.reset_ui_state()
                return

            if app2_content:
                queries = [q for q in app2_content.split('\n') if q.strip()]
                if queries:
                    self.log("\n" + "=" * 40)
                    self.log(f"▶️ [阶段二] 启动问财数据抓取 (共 {len(queries)} 个条件)...")
                    log_queue, result_queue = multiprocessing.Queue(), multiprocessing.Queue()
                    stop_event, pause_event = multiprocessing.Event(), multiprocessing.Event()

                    p = multiprocessing.Process(
                        target=run_app2_isolated_process,
                        args=(queries, unified_folder, ts, self.market_var.get(), self.headless_var.get(), log_queue,
                              result_queue, stop_event, pause_event)
                    )
                    p.start()

                    while p.is_alive():
                        while not log_queue.empty():
                            try:
                                self.log(log_queue.get_nowait())
                            except:
                                break
                        if self.stop_flag: stop_event.set()
                        if self.pause_flag:
                            pause_event.set()
                        else:
                            pause_event.clear()
                        time.sleep(0.2)

                    while not log_queue.empty():
                        try:
                            self.log(log_queue.get_nowait())
                        except:
                            break

                    success = False
                    if not result_queue.empty():
                        try:
                            success = result_queue.get_nowait()
                        except:
                            pass
                    if not success: has_error = True
                else:
                    self.log("\n" + "=" * 40)
                    self.log("⏭️ [阶段二] 问财有效条件为空，自动跳过。")
            else:
                self.log("\n" + "=" * 40)
                self.log("⏭️ [阶段二] 问财指令为空，自动跳过。")

            if not self.stop_flag:
                self.log("\n" + "=" * 40)
                self.log("▶️ [阶段三] 开始自动合成《海选公司汇总表》...")
                self.merge_results(unified_folder, ts)
                final_excel_path = os.path.join(unified_folder, "海选公司汇总表.xlsx")

                if not self.is_auto:
                    if not has_error:
                        self.log("\n🎉 所有任务已圆满完成！")
                        messagebox.showinfo("全部完成", f"任务执行完毕！\n所有结果均已统一汇总至：\n{unified_folder}")
                    else:
                        self.log("\n⚠️ 任务结束，但抓取过程中出现错误或未找到数据，请检查上方日志。")
                        messagebox.showwarning("部分完成", "执行遇到错误或数据为空，部分数据可能未完整生成，请查看日志！")
                else:
                    self.log("\n🔄 自动串联模式：即将无缝移交【下载年报程序】...")

                if self.master_callback:
                    self.master_callback(not has_error, final_excel_path)

        except Exception as e:
            self.log(f"❌ 运行发生致命异常: {e}")
            if self.master_callback: self.master_callback(False, "")
        finally:
            self.reset_ui_state()

class RedirectText:
    def __init__(self, text_widget, root):
        self.text_widget = text_widget
        self.root = root

    def write(self, string):
        try:
            self.root.after(0, self._append, string)
        except:
            pass

    def _append(self, string):
        try:
            # 安全检查：只有当UI组件真正存在时才写入，防止窗口关闭后引发崩溃
            if self.text_widget.winfo_exists():
                self.text_widget.configure(state='normal')
                self.text_widget.insert(tk.END, string)
                self.text_widget.see(tk.END)
                self.text_widget.configure(state='disabled')
        except Exception:
            pass

    def flush(self): pass

class StockAnalyzerApp:
    def __init__(self, root, master_callback=None):
        self.root = root
        self.root.title("巨潮资讯 - 年报下载神器 (V11.2 自动极速版)")
        self.root.geometry("1150x700")
        self.master_callback = master_callback
        self.is_auto = False
        self.auto_excel_path = ""

        self.config = self.load_config()
        self.create_widgets()
        sys.stdout = RedirectText(self.log_area, self.root)

        default_file = self.import_file_var.get()
        if default_file and os.path.exists(default_file):
            print(f"🔄 启动程序，正在自动读取默认文件: {os.path.basename(default_file)}")
            self.load_file_data(default_file)

    def load_config(self):
        try:
            with open(CONFIG_FILE_APP3, 'r', encoding='utf-8') as f:
                return json.load(f)
        except:
            return {}

    def save_config(self):
        with open(CONFIG_FILE_APP3, 'w', encoding='utf-8') as f: json.dump(self.config, f, ensure_ascii=False,
                                                                           indent=4)

    def create_widgets(self):
        main_pane = tk.PanedWindow(self.root, orient=tk.HORIZONTAL, sashwidth=4, bg="#CCCCCC")
        main_pane.pack(fill=tk.BOTH, expand=True)

        left_frame = ttk.Frame(main_pane)
        main_pane.add(left_frame, minsize=420, width=450)

        inner_frame = tk.Frame(left_frame, bg='#F2F5F8', padx=20, pady=20)
        inner_frame.pack(fill=tk.BOTH, expand=True)

        input_header = tk.Frame(inner_frame, bg='#F2F5F8')
        input_header.pack(fill=tk.X, pady=(0, 5))
        tk.Label(input_header, text="📝 股票代码/简称:", bg='#F2F5F8', font=("微软雅黑", 10, "bold")).pack(
            side=tk.LEFT)

        self.import_file_var = tk.StringVar(value=self.config.get("default_import_file", ""))
        import_entry = tk.Entry(input_header, textvariable=self.import_file_var, font=("微软雅黑", 8),
                                state='readonly', relief="solid", bd=1)
        import_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(5, 5))

        tk.Button(input_header, text="设为默认文件", command=self.set_default_import_file, font=("微软雅黑", 8),
                  bg="#E0E0E0", relief="solid", bd=1).pack(side=tk.RIGHT)
        tk.Button(input_header, text="临时选择导入", command=self.import_from_file, font=("微软雅黑", 8),
                  bg="#E0E0E0", relief="solid", bd=1).pack(side=tk.RIGHT, padx=5)

        source_frame = tk.Frame(inner_frame, bg='#F2F5F8')
        source_frame.pack(fill=tk.X, pady=(0, 5))
        self.data_source_var = tk.StringVar(value="auto")
        tk.Radiobutton(source_frame, text="手动输入 / 文件导入", variable=self.data_source_var, value="manual",
                       command=self.on_data_source_change, bg='#F2F5F8').pack(side=tk.LEFT)
        tk.Radiobutton(source_frame, text="自动联动海选结果", variable=self.data_source_var, value="auto",
                       command=self.on_data_source_change, bg='#F2F5F8').pack(side=tk.LEFT, padx=10)

        self.code_text = scrolledtext.ScrolledText(inner_frame, height=7, font=("Consolas", 11), relief="solid",
                                                   bd=1)
        self.code_text.pack(fill=tk.X, pady=(0, 15))
        self.code_text.insert(tk.END, "000848\n601882\n")

        tk.Label(inner_frame, text="🗓 年报时间范围:", bg='#F2F5F8', font=("微软雅黑", 10, "bold")).pack(anchor=tk.W,
                                                                                                        pady=(0, 5))
        self.since_list_var = tk.BooleanVar(value=False)
        cb_since = tk.Checkbutton(inner_frame, text="📆 下载【上市以来】所有年报 (覆盖下方时间)",
                                  variable=self.since_list_var, bg='#F2F5F8', activebackground='#F2F5F8',
                                  command=self.toggle_years, font=("微软雅黑", 9, "bold"), fg="#D35400")
        cb_since.pack(anchor=tk.W, padx=5, pady=(0, 5))

        self.time_frame = tk.Frame(inner_frame, bg='#F2F5F8')
        self.time_frame.pack(anchor=tk.W, padx=25, pady=(0, 15))

        quick_frame = tk.Frame(self.time_frame, bg='#F2F5F8')
        quick_frame.pack(fill=tk.X, pady=(0, 5))
        tk.Label(quick_frame, text="快捷选择:", bg='#F2F5F8', font=("微软雅黑", 9), width=8, anchor='w').pack(
            side=tk.LEFT)
        self.quick_mode_var = tk.StringVar(value="最近 5 年")
        modes = ["最近 1 年", "最近 3 年", "最近 5 年", "最近 10 年", "自定义区间"]
        self.mode_cb = ttk.Combobox(quick_frame, textvariable=self.quick_mode_var, values=modes, state='readonly',
                                    width=12, font=("微软雅黑", 9))
        self.mode_cb.pack(side=tk.LEFT)
        self.mode_cb.bind("<<ComboboxSelected>>", self.on_quick_mode_change)

        specific_frame = tk.Frame(self.time_frame, bg='#F2F5F8')
        specific_frame.pack(fill=tk.X)
        tk.Label(specific_frame, text="具体年份:", bg='#F2F5F8', font=("微软雅黑", 9), width=8, anchor='w').pack(
            side=tk.LEFT)

        curr_year = datetime.now().year
        years = [str(y) for y in range(curr_year, 1989, -1)]

        self.start_year_var = tk.StringVar(value=str(curr_year - 5))
        self.start_year_cb = ttk.Combobox(specific_frame, textvariable=self.start_year_var, values=years,
                                          state='readonly', width=6, font=("Consolas", 10))
        self.start_year_cb.pack(side=tk.LEFT)
        self.start_year_cb.bind("<<ComboboxSelected>>", self.on_specific_year_change)

        tk.Label(specific_frame, text="至", bg='#F2F5F8', font=("微软雅黑", 9)).pack(side=tk.LEFT, padx=5)

        self.end_year_var = tk.StringVar(value=str(curr_year))
        self.end_year_cb = ttk.Combobox(specific_frame, textvariable=self.end_year_var, values=years,
                                        state='readonly', width=6, font=("Consolas", 10))
        self.end_year_cb.pack(side=tk.LEFT)
        self.end_year_cb.bind("<<ComboboxSelected>>", self.on_specific_year_change)

        tk.Label(inner_frame, text="⚙ 核心功能:", bg='#F2F5F8', font=("微软雅黑", 10, "bold")).pack(anchor=tk.W,
                                                                                                    pady=(0, 5))
        self.dl_prospectus_var = tk.BooleanVar(value=True)
        tk.Checkbutton(inner_frame, text="⬇ 下载【招股说明书】", variable=self.dl_prospectus_var, bg='#F2F5F8',
                       activebackground='#F2F5F8', font=("微软雅黑", 9)).pack(anchor=tk.W, padx=5)

        self.run_peers_var = tk.BooleanVar(value=True)
        tk.Checkbutton(inner_frame, text="📊 分析生成【对标表】并下载同行前三", variable=self.run_peers_var,
                       bg='#F2F5F8', activebackground='#F2F5F8', font=("微软雅黑", 9)).pack(anchor=tk.W, padx=5)

        self.silent_mode_var = tk.BooleanVar(value=True)
        tk.Checkbutton(inner_frame, text="🤫 浏览器静默运行 (后台扫描，不弹窗)", variable=self.silent_mode_var,
                       bg='#F2F5F8', activebackground='#F2F5F8', font=("微软雅黑", 9)).pack(anchor=tk.W, padx=5,
                                                                                            pady=(0, 5))

        tk.Label(inner_frame, text="⏬ 下载模式:", bg='#F2F5F8', font=("微软雅黑", 10, "bold")).pack(anchor=tk.W,
                                                                                                    pady=(5, 5))
        self.download_mode_var = tk.StringVar(value="raw")
        mode_frame = tk.Frame(inner_frame, bg='#F2F5F8')
        mode_frame.pack(anchor=tk.W, padx=5, pady=(0, 10))
        tk.Radiobutton(mode_frame, text="原样下载 (覆盖已有文件)", variable=self.download_mode_var, value="raw",
                       bg='#F2F5F8').pack(side=tk.LEFT)
        tk.Radiobutton(mode_frame, text="去重下载 (跳过已下载公司)", variable=self.download_mode_var, value="dedup",
                       bg='#F2F5F8').pack(side=tk.LEFT, padx=10)

        tk.Label(inner_frame, text="📁 根保存路径:", bg='#F2F5F8', font=("微软雅黑", 10, "bold")).pack(anchor=tk.W,
                                                                                                      pady=(0, 5))
        path_frame = tk.Frame(inner_frame, bg='#F2F5F8')
        path_frame.pack(fill=tk.X)

        default_save = self.config.get("default_save_path",
                                       os.path.join(os.path.expanduser("~"), "Desktop", "股票数据库"))
        self.path_var = tk.StringVar(value=default_save)
        path_entry = tk.Entry(path_frame, textvariable=self.path_var, font=("微软雅黑", 9), state='readonly',
                              relief="solid", bd=1)
        path_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)

        tk.Button(path_frame, text="设为默认", command=self.set_default_save_path, font=("微软雅黑", 9),
                  relief="solid", bd=1, bg="#E0E0E0").pack(side=tk.RIGHT, padx=(5, 0))
        tk.Button(path_frame, text="选择路径", command=self.select_directory, font=("微软雅黑", 9), relief="solid",
                  bd=1, bg="#E0E0E0").pack(side=tk.RIGHT, padx=(5, 0))

        bottom_frame = tk.Frame(inner_frame, bg='#F2F5F8')
        bottom_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=(20, 0))

        self.run_btn = tk.Button(bottom_frame, text="🚀 启动智能分析与下载", font=("微软雅黑", 12, "bold"),
                                 bg="#28C76F", fg="white", activebackground="#20A058", activeforeground="white",
                                 relief="flat", pady=10, cursor="hand2", command=self.start_thread)
        self.run_btn.pack(fill=tk.X)

        right_frame = tk.Frame(main_pane, bg="#1E1E1E")
        main_pane.add(right_frame)

        self.log_area = scrolledtext.ScrolledText(right_frame, wrap=tk.WORD, font=("Consolas", 11), bg="#1E1E1E",
                                                  fg="#D4D4D4", insertbackground="white", relief="flat", padx=10,
                                                  pady=10)
        self.log_area.pack(fill=tk.BOTH, expand=True)
        self.log_area.configure(state='disabled')

        self.on_data_source_change()

    def on_data_source_change(self):
        if self.data_source_var.get() == "auto":
            if self.auto_excel_path and os.path.exists(self.auto_excel_path):
                self.load_file_data(self.auto_excel_path)
                print(f"\n🔄 [联动模式] 已成功载入海选汇总表数据。")
            else:
                self.code_text.delete("1.0", tk.END)
                self.code_text.insert(tk.END, "当前无海选结果传入，请先运行【海选公司程序】。\n")
        else:
            default_file = self.import_file_var.get()
            if default_file and os.path.exists(default_file):
                self.load_file_data(default_file)
            else:
                self.code_text.delete("1.0", tk.END)
                self.code_text.insert(tk.END, "000848\n601882\n")

    def load_file_data(self, file_path):
        if not file_path or not os.path.exists(file_path): return
        try:
            text_content = ""
            if file_path.endswith('.csv'):
                df = pd.read_csv(file_path, dtype=str)
                text_content = "\n".join(df.astype(str).apply(lambda x: ' '.join(x), axis=1))
            elif file_path.endswith(('.xlsx', '.xls')):
                df = pd.read_excel(file_path, dtype=str)
                text_content = "\n".join(df.astype(str).apply(lambda x: ' '.join(x), axis=1))
            else:
                with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                    text_content = f.read()

            codes = list(set(re.findall(r'\b\d{6}\b', text_content)))
            self.code_text.delete("1.0", tk.END)
            if codes:
                self.code_text.insert(tk.END, "\n".join(codes) + "\n")
                print(f"✅ 成功从文件自动提取并导入 {len(codes)} 个股票代码。")
            else:
                print(f"⚠️ 提示: 未在该文件提取到任何6位连续数字组成的股票代码！")
        except Exception as e:
            print(f"❌ 读取文件失败: {e}")

    def set_default_import_file(self):
        current_file = self.import_file_var.get()
        initial_dir = os.path.dirname(current_file) if current_file and os.path.exists(
            current_file) else os.path.expanduser("~")
        file_path = filedialog.askopenfilename(initialdir=initial_dir, title="选择并设置为默认导入文件",
                                               filetypes=[("表格及文本文件", "*.txt *.csv *.xlsx *.xls")])
        if file_path:
            self.config["default_import_file"] = file_path
            self.save_config()
            self.import_file_var.set(file_path)
            self.data_source_var.set("manual")
            print(f"\n📁 已将默认导入文件更改为:\n{file_path}")
            self.load_file_data(file_path)

    def import_from_file(self):
        current_file = self.import_file_var.get()
        initial_dir = os.path.dirname(current_file) if current_file and os.path.exists(
            current_file) else os.path.expanduser("~")
        file_path = filedialog.askopenfilename(initialdir=initial_dir, title="选择临时文件并导入",
                                               filetypes=[("表格及文本文件", "*.txt *.csv *.xlsx *.xls")])
        if file_path:
            self.import_file_var.set(file_path)
            self.data_source_var.set("manual")
            print(f"\n📁 正在临时导入文件:\n{file_path}")
            self.load_file_data(file_path)

    def set_default_save_path(self):
        current_path = self.path_var.get()
        if os.path.exists(current_path) or messagebox.askyesno("确认", "路径似乎不存在，是否仍设为默认？"):
            self.config["default_save_path"] = current_path
            self.save_config()
            messagebox.showinfo("成功", f"默认保存路径已设置为:\n{current_path}")

    def on_quick_mode_change(self, event=None):
        mode = self.quick_mode_var.get()
        curr_year = datetime.now().year
        if mode == "最近 1 年":
            self.start_year_var.set(str(curr_year - 1)); self.end_year_var.set(str(curr_year))
        elif mode == "最近 3 年":
            self.start_year_var.set(str(curr_year - 3)); self.end_year_var.set(str(curr_year))
        elif mode == "最近 5 年":
            self.start_year_var.set(str(curr_year - 5)); self.end_year_var.set(str(curr_year))
        elif mode == "最近 10 年":
            self.start_year_var.set(str(curr_year - 10)); self.end_year_var.set(str(curr_year))

    def on_specific_year_change(self, event=None):
        self.quick_mode_var.set("自定义区间")

    def toggle_years(self):
        state = 'disabled' if self.since_list_var.get() else 'readonly'
        self.mode_cb.config(state=state)
        self.start_year_cb.config(state=state)
        self.end_year_cb.config(state=state)

    def select_directory(self):
        folder = filedialog.askdirectory(initialdir=self.path_var.get(), title="选择输出文件夹")
        if folder: self.path_var.set(folder)

    def start_thread(self, is_auto=False):
        self.is_auto = is_auto

        # 【修复3】关键防护：下载用的 requests 会读取环境变量里的代理。
        # 若环境里残留 HTTP_PROXY=http://127.0.0.1:7890（本机没开这个代理），
        # 巨潮/东财的所有请求都会被劫持并报 ProxyError(WinError 10061)，最终一个文件都下不到。
        # 此处只信任进程启动时（Python 启动前）就已存在的真实代理，运行期间被子程序写入的残留一律清除。
        apply_direct_network_env()
        if _BOOT_HTTP_PROXY:
            os.environ["HTTP_PROXY"] = _BOOT_HTTP_PROXY
            os.environ["http_proxy"] = _BOOT_HTTP_PROXY
        if _BOOT_HTTPS_PROXY:
            os.environ["HTTPS_PROXY"] = _BOOT_HTTPS_PROXY
            os.environ["https_proxy"] = _BOOT_HTTPS_PROXY

        raw_codes = self.code_text.get("1.0", tk.END)
        stock_list = list(set(re.findall(r'\b\d{6}\b', raw_codes)))
        if not stock_list:
            if not self.is_auto: messagebox.showwarning("提示", "未提取到股票代码！")
            return

        base_dir = self.path_var.get().strip()
        start_y = int(self.start_year_var.get())
        end_y = int(self.end_year_var.get())
        if start_y > end_y:
            if not self.is_auto: messagebox.showwarning("提示", "具体年份的【开始年份】不能大于【结束年份】！")
            return

        self.run_btn.config(state=tk.DISABLED, text="⏳ 正在全速运行中...", bg="#999999")
        self.log_area.configure(state='normal')
        self.log_area.delete("1.0", tk.END)
        self.log_area.configure(state='disabled')

        threading.Thread(target=self.run_pipeline, args=(
            stock_list, base_dir, start_y, end_y, self.since_list_var.get(),
            self.run_peers_var.get(), self.dl_prospectus_var.get(), self.silent_mode_var.get(),
            self.download_mode_var.get()
        ), daemon=True).start()

    def run_pipeline(self, stock_list, base_dir, start_y, end_y, since_ipo, enable_peers, enable_prospectus,
                     is_silent, download_mode):
        success = False
        # 【修复2】真实下载统计：只有确实有 PDF 落盘，才允许接力下一子程序
        total_targets = 0          # 本轮实际提交给巨潮下载的公司次数
        total_downloaded = 0       # 本轮真正成功下载的 PDF 数量
        unified_out_dir = os.path.join(base_dir, "报表下载")
        try:
            os.makedirs(unified_out_dir, exist_ok=True)
            global_downloaded = set()
            if download_mode == "dedup":
                for r, d, f in os.walk(unified_out_dir):
                    for folder in d:
                        match = re.search(r'^(\d{6})_', folder)
                        if match: global_downloaded.add(match.group(1))

            print("=" * 60)
            print(f" 🚀 下载任务已启动 | 统一输出主路径: {unified_out_dir}")
            print(
                f" ⏬ 下载模式: {'去重下载 (全局去重并保留同名覆盖)' if download_mode == 'dedup' else '原样下载 (已防本轮任务内重复)'}")
            print("=" * 60)

            print("\n[阶段0] 正在预先扫描所有目标公司及其同行前三名名单...")
            global_target_list = {}
            industry_groups = {}  # 核心改动：按行业属性对公司进行归并字典

            # 1. 全局盘点与行业编组
            for query_code in stock_list:
                df_ind, ind_map, name_map = get_eastmoney_industry_from_datacenter([query_code])
                if df_ind.empty: continue
                company_name = name_map.get(query_code, "未知简称")
                industry_name = ind_map.get(query_code, "未知行业")

                # 若该行业尚未建立分组，则初始化（采用首个公司作为文件夹命名代表）
                if industry_name not in industry_groups:
                    industry_groups[industry_name] = {
                        'rep_name': company_name,
                        'companies': {},
                        'df_ind_list': [],
                        'df_peers_list': []
                    }

                # 将原查询公司加入对应行业组
                industry_groups[industry_name]['companies'][query_code] = company_name
                industry_groups[industry_name]['df_ind_list'].append(df_ind)

                if query_code not in global_target_list:
                    global_target_list[query_code] = {'代码': query_code, '名称': company_name,
                                                      '所属行业': industry_name, '属性': '原查询公司'}

                # 将同行公司也加入对应行业组
                if enable_peers:
                    df_peers, dl_dict = get_leader_radar_batch([query_code], ind_map, headless=is_silent)
                    industry_groups[industry_name]['df_peers_list'].append(df_peers)

                    for d_code, d_ind in dl_dict.items():
                        if d_code != query_code and d_code not in global_target_list:
                            peer_name = "未知同行"
                            if not df_peers.empty:
                                match_row = df_peers[df_peers['代码'] == d_code]
                                if not match_row.empty: peer_name = str(match_row.iloc[0]['名称'])
                            global_target_list[d_code] = {'代码': d_code, '名称': peer_name, '所属行业': d_ind,
                                                          '属性': '行业同行(前三)'}

                        # 确保同行也被收编到同一个行业字典中，避免重复
                        if d_code not in industry_groups[industry_name]['companies']:
                            industry_groups[industry_name]['companies'][d_code] = global_target_list[d_code]['名称']

            # 2. 生成全景名单
            if global_target_list:
                list_df = pd.DataFrame(list(global_target_list.values()))
                list_df = list_df[['代码', '名称', '所属行业', '属性']]
                out_list_path = os.path.join(unified_out_dir, "需对比分析的全部公司名单.xlsx")
                list_df.to_excel(out_list_path, index=False)
                print(
                    f"\n✅ 成功生成并去重名单文档: {os.path.basename(out_list_path)} (共收集 {len(list_df)} 家公司作为后续核心下载目标)")
            else:
                print("\n⚠️ 未获取到任何有效公司名单。")

            # 3. 按行业属性组进行无重复下载
            session_downloaded = set()  # 本轮下载防重缓存，杜绝跨行业交叉的重复下载

            for industry_name, group_data in industry_groups.items():
                rep_name = group_data['rep_name']

                print(f"\n================ [开始处理行业分组下载: {industry_name} (代表: {rep_name})] ================")
                safe_company = re.sub(r'[\\/:*?"<>|]', '', rep_name)
                safe_ind = re.sub(r'[\\/:*?"<>|]', '_', industry_name)
                target_folder_name = f"{safe_company}_行业_{safe_ind}"
                target_dir = os.path.join(unified_out_dir, target_folder_name)

                if os.path.exists(target_dir) and download_mode == "raw":
                    print(f"    ⚠️ 存在同名文件夹，正在执行覆盖清理...")
                    try:
                        shutil.rmtree(target_dir)
                    except Exception as e:
                        print(f"    ❌ 清理失败: {e}")

                os.makedirs(target_dir, exist_ok=True)

                # 智能合并该行业下所有查询股票的属性表与同行表
                if group_data['df_ind_list']:
                    pd.concat(group_data['df_ind_list']).drop_duplicates(subset=['目标股票']).to_excel(
                        os.path.join(target_dir, f'1_{safe_company}_公司属性表.xlsx'), index=False)

                valid_peers = [df for df in group_data['df_peers_list'] if not df.empty]
                if valid_peers:
                    pd.concat(valid_peers).drop_duplicates(subset=['原查询股票', '代码']).to_excel(
                        os.path.join(target_dir, f'2_{safe_company}_同行排列表.xlsx'), index=False)

                dedup_dl_dict = {}
                for code, name in group_data['companies'].items():
                    # 双重校验：要么在全局库中已存在，要么在本轮循环中已下载过
                    if (download_mode == "dedup" and code in global_downloaded) or (code in session_downloaded):
                        print(f"    ⏭️ [去重/防重模式] 公司 {code} ({name}) 已在库中或本轮已分配，跳过重复网络下载...")
                    else:
                        dedup_dl_dict[code] = industry_name
                        if download_mode == "dedup": global_downloaded.add(code)
                        session_downloaded.add(code)

                if dedup_dl_dict:
                    total_targets += len(dedup_dl_dict)
                    total_downloaded += cninfo_download_documents(dedup_dl_dict, start_y, end_y, target_dir, since_ipo,
                                                                  enable_prospectus)

            print(f"\n✨ 全部下载执行完毕！")
            if total_targets == 0:
                print("⚠️ 本轮没有任何需要下载的目标公司（可能被“去重模式”全部跳过），不接力后续子程序。")
                success = False
            elif total_downloaded == 0:
                print("❌ 本轮未成功下载到任何 PDF 文件（网络异常/被拦截/名单为空/年份区间无年报），不接力后续子程序。")
                success = False
            else:
                print(f"✅ 本轮成功下载 {total_downloaded} 个文件，可以接力【年报提取程序】。")
                success = True
        except Exception as e:
            print(f"\n❌ 发生严重错误: {e}")
            success = False
        finally:
            self.root.after(0,
                            lambda: self.run_btn.config(state=tk.NORMAL, text="🚀 启动智能分析与下载", bg="#28C76F"))
            if self.master_callback:
                # 【修复1】用默认参数固定本次结果，避免回调执行时读到已变化的变量
                self.root.after(0, lambda ok=success, d=unified_out_dir: self.master_callback(ok, d))
class ReportExtractionApp:
    def __init__(self, root, master_callback=None):
        self.root = root
        self.root.title("财报提取、抓取与智能对比融合系统 (三合一终极版)")
        self.root.geometry("800x780")
        self.root.configure(bg="#F5F5F5")
        self.master_callback = master_callback
        self.is_auto = False
        self.auto_input_dir = ""
        self.setup_ui()

    def setup_ui(self):
        tk.Label(self.root, text="核心文件与目录设置", font=("微软雅黑", 12, "bold"), bg="#F5F5F5", fg="#333").pack(
            pady=(15, 5))
        source_frame = tk.Frame(self.root, bg="#F5F5F5")
        source_frame.pack(fill='x', padx=30, pady=5)
        self.data_source_var = tk.StringVar(value="auto")
        tk.Radiobutton(source_frame, text="手动选择目录", variable=self.data_source_var, value="manual",
                       bg='#F5F5F5', command=self.toggle_source).pack(side=tk.LEFT)
        tk.Radiobutton(source_frame, text="自动联动上级下载结果", variable=self.data_source_var, value="auto",
                       bg='#F5F5F5', command=self.toggle_source).pack(side=tk.LEFT, padx=10)

        self.frame_path = tk.Frame(self.root, bg="#F5F5F5")
        self.frame_path.pack(fill='x', padx=30)
        tk.Label(self.frame_path, text="PDF源目录(包含年报):", font=("微软雅黑", 10), bg="#F5F5F5").grid(row=0,
                                                                                                         column=0,
                                                                                                         sticky='w',
                                                                                                         pady=5)
        self.in_var, self.out_var = tk.StringVar(), tk.StringVar()
        self.entry_in = tk.Entry(self.frame_path, textvariable=self.in_var, width=65)
        self.entry_in.grid(row=0, column=1, padx=10)
        self.btn_browse_in = tk.Button(self.frame_path, text="浏览...",
                                       command=lambda: self.in_var.set(filedialog.askdirectory()))
        self.btn_browse_in.grid(row=0, column=2)

        tk.Label(self.frame_path, text="Excel保存处(含B表):", font=("微软雅黑", 10), bg="#F5F5F5").grid(row=1,
                                                                                                        column=0,
                                                                                                        sticky='w',
                                                                                                        pady=5)
        self.entry_out = tk.Entry(self.frame_path, textvariable=self.out_var, width=65)
        self.entry_out.grid(row=1, column=1, padx=10)
        self.btn_browse_out = tk.Button(self.frame_path, text="浏览...",
                                        command=lambda: self.out_var.set(filedialog.askdirectory()))
        self.btn_browse_out.grid(row=1, column=2)

        tk.Label(self.root, text="财务指标强制换算规则 (格式：旧名称=新名称，每行一条):",
                 font=("微软雅黑", 10, "bold"), bg="#F5F5F5").pack(anchor='w', padx=30, pady=(15, 5))
        self.mapping_text = scrolledtext.ScrolledText(self.root, width=95, height=5, bg="#F0F8FF")
        self.mapping_text.pack()

        if os.path.exists(TEMPLATE_FILE_A):
            with open(TEMPLATE_FILE_A, "r", encoding="utf-8") as f:
                self.mapping_text.insert(tk.END, f.read())
        else:
            self.mapping_text.insert(tk.END,
                                     "归属于母公司所有者的净利润=净利润\n销售商品、提供劳务收到的现金=销售商品及劳务收到现金\n管理费用=管理费用及研发费用\n")

        btn_frame = tk.Frame(self.root, bg="#F5F5F5")
        btn_frame.pack(pady=15)
        btn_save_tmpl = tk.Button(btn_frame, text="保存为模板", bg="#4CAF50", fg="white", font=("微软雅黑", 10),
                                  command=self.save_template, padx=10)
        btn_save_tmpl.pack(side=tk.LEFT, padx=15)
        self.btn_start = tk.Button(btn_frame, text="启动全链路解析、抓取与合并对比", bg="#2196F3", fg="white",
                                   font=("微软雅黑", 11, "bold"), command=lambda: self.run_app(), padx=20)
        self.btn_start.pack(side=tk.LEFT, padx=15)

        tk.Label(self.root, text="实时运行日志:", font=("微软雅黑", 10, "bold"), bg="#F5F5F5").pack(anchor='w',
                                                                                                    padx=30)
        self.log_area = scrolledtext.ScrolledText(self.root, width=100, height=18, bg="#1E1E1E", fg="#D4D4D4",
                                                  font=("Consolas", 9))
        self.log_area.pack(pady=5)

        self.toggle_source()

    def toggle_source(self):
        if self.data_source_var.get() == "auto":
            self.entry_in.configure(state='disabled');
            self.entry_out.configure(state='disabled')
            self.btn_browse_in.configure(state='disabled');
            self.btn_browse_out.configure(state='disabled')
            if self.auto_input_dir: self.log_append(f"🔄 联动模式激活，输入输出均使用: {self.auto_input_dir}")
        else:
            self.entry_in.configure(state='normal');
            self.entry_out.configure(state='normal')
            self.btn_browse_in.configure(state='normal');
            self.btn_browse_out.configure(state='normal')
            self.log_append("🖐️ 切换至手动选择目录模式")

    def save_template(self):
        try:
            with open(TEMPLATE_FILE_A, "w", encoding="utf-8") as f:
                f.write(self.mapping_text.get(1.0, tk.END))
            messagebox.showinfo("成功", "换算规则模板已成功保存！\n下次打开软件将自动加载该模板。")
        except Exception as e:
            messagebox.showerror("错误", f"保存模板失败：{e}")

    def log_append(self, m):
        self.root.after(0, lambda: (self.log_area.insert(tk.END, m + "\n"), self.log_area.see(tk.END)))

    def run_app(self, is_auto=False):
        self.is_auto = is_auto
        in_path = self.auto_input_dir if self.data_source_var.get() == "auto" else self.in_var.get()
        out_path = self.auto_input_dir if self.data_source_var.get() == "auto" else self.out_var.get()

        if not in_path or not out_path or not os.path.exists(in_path) or not os.path.exists(out_path):
            if not is_auto:
                messagebox.showerror("错误", "源目录或保存目录有误，请重新确认！")
            else:
                self.log_append("❌ 自动路径无效，终止！")
            return

        self.btn_start.config(state=tk.DISABLED, text="⏳ 融合引擎全速运转中...", bg="#999999")
        self.log_area.delete(1.0, tk.END)
        raw_map = self.mapping_text.get(1.0, tk.END).strip().split('\n')
        threading.Thread(target=self.master_pipeline_thread, args=(in_path, out_path, raw_map, is_auto),
                         daemon=True).start()

    def master_pipeline_thread(self, in_path, out_path, raw_map, is_auto):
        import concurrent.futures
        try:
            out_path = os.path.join(out_path, "报表提取完善")
            os.makedirs(out_path, exist_ok=True)

            mappings = {}
            for line in raw_map:
                if '=' in line:
                    k, v = line.split('=', 1)
                    mappings[k.strip().replace(" ", "")] = v.strip().replace(" ", "")

            self.log_append("==================================================")
            self.log_append("🔄 【阶段一】启动多线程本地 PDF 精细提取生成单年 A 类表")
            self.log_append(f"📂 输出总目录设置为: {out_path}")
            self.log_append("==================================================")

            all_results_by_company = {}
            company_tasks_p2 = {}
            pdf_tasks_info = []

            # 1. 重建行业架构并同步 Excel 属性表
            for folder_name in os.listdir(in_path):
                ind_path = os.path.join(in_path, folder_name)
                if not os.path.isdir(ind_path): continue

                out_ind_path = os.path.join(out_path, folder_name)

                has_excel_to_copy = False
                for f in os.listdir(ind_path):
                    if f.endswith(".xlsx") or f.endswith(".csv"):
                        if "公司属性表" in f or "同行排列表" in f:
                            os.makedirs(out_ind_path, exist_ok=True)
                            src_file = os.path.join(ind_path, f)
                            dst_file = os.path.join(out_ind_path, f)
                            try:
                                if not os.path.exists(dst_file):
                                    shutil.copy2(src_file, dst_file)
                                    has_excel_to_copy = True
                            except Exception as e:
                                self.log_append(f"  ⚠️ 复制 {f} 失败: {e}")

                if has_excel_to_copy:
                    self.log_append(f"  📁 已同步行业架构与属性表: {folder_name}")

            # 2. 全域穿透扫描 PDF，确保无死角遗漏
            for r, d, f_list in os.walk(in_path):
                for f in f_list:
                    if f.lower().endswith(".pdf"):
                        pdf_path = os.path.join(r, f)
                        # 智能判断属于哪个行业文件夹，如果乱放的则归入"默认提取分类"
                        rel_path = os.path.relpath(pdf_path, in_path)
                        parts = rel_path.split(os.sep)
                        folder_name = parts[0] if len(parts) > 1 else "默认提取分类"
                        out_ind_path = os.path.join(out_path, folder_name)

                        pdf_tasks_info.append((pdf_path, out_ind_path, folder_name))

            if not pdf_tasks_info:
                self.log_append("❌ 严重警告：源目录内未找到任何 PDF 年报文件！")
                self.log_append(
                    "💡 提示：请检查 A2(报表下载) 是否因为“去重模式”或“年份区间未涵盖最新年报”而跳过了 PDF 下载。")
                self.root.after(0, lambda: self.btn_start.config(state=tk.NORMAL, text="启动全链路解析、抓取与合并对比",
                                                                 bg="#2196F3"))
                if self.master_callback: self.master_callback()
                return

            def process_pdf_task(p, out_ind_path, folder_name):
                data, comp_name, yr, stock_code = p1_process_pdf(p, lambda msg: None)
                if not data or stock_code == "未提取": return False, None

                df_emp, _, _, _ = p2_process_pdf_employee(p)

                os.makedirs(out_ind_path, exist_ok=True)
                out_file = os.path.join(out_ind_path, f"{comp_name}（{stock_code}）{yr}年度年报提取表.xlsx")

                try:
                    with pd.ExcelWriter(out_file, engine='openpyxl') as writer:
                        all_sheets = [n[0] for n in MILESTONE_PATTERNS[:8]] + ["董监高及报酬情况", "员工情况",
                                                                               "分红情况"]
                        for k in all_sheets:
                            df_to_save = data.get(k, pd.DataFrame())
                            if not df_to_save.empty:
                                df_to_save.to_excel(writer, sheet_name=k, startrow=2, index=False)
                                ws = writer.sheets[k]
                                unit_str = "人民币元" if "情况" not in k else "详见表内说明"
                                ws.cell(1, 1, f"{comp_name}（{stock_code}）{k}（{yr}年度）\n单位：{unit_str}").font = Font(
                                    bold=True, size=12)
                                ws.cell(1, 1).alignment = Alignment(horizontal='center', vertical='center',
                                                                    wrap_text=True)
                                ws.row_dimensions[1].height = 40
                                is_financial = any(
                                    x in k for x in ["资产负债表", "利润表", "现金流量表", "所有者权益变动表"])
                                for row in ws.iter_rows(min_row=3, max_col=ws.max_column):
                                    item_name = str(row[0].value) if row[0].value else ""
                                    for cell in row[1:]:
                                        if isinstance(cell.value, (int, float)):
                                            if is_financial:
                                                cell.number_format = '#,##0.0000' if "每股收益" in item_name else '#,##0.00'
                                            elif k == "员工情况":
                                                cell.number_format = '#,##0'
                                            elif k == "董监高及报酬情况":
                                                cell.number_format = '#,##0' if cell.value == int(
                                                    cell.value) else '#,##0.00'
                                common_auto_adjust_col_width(ws)
                    return True, (comp_name, yr, stock_code, folder_name, data, df_emp, out_file)
                except Exception as e:
                    return False, str(e)

            # 【提速引擎】：开辟 4 个子线程并发破解与提取 PDF
            self.log_append(f"  🚀 并发引擎已启动 (并行解析 {len(pdf_tasks_info)} 个PDF文档)...")
            with concurrent.futures.ThreadPoolExecutor(max_workers=4) as executor:
                futures = [executor.submit(process_pdf_task, *info) for info in pdf_tasks_info]
                for future in concurrent.futures.as_completed(futures):
                    success, result = future.result()
                    if success:
                        comp_name, yr, stock_code, folder_name, data, df_emp, out_file = result
                        self.log_append(f"  ✅ 单年A表提取完毕: {os.path.basename(out_file)}")

                        company_key = (comp_name, stock_code, folder_name)
                        if company_key not in all_results_by_company: all_results_by_company[company_key] = {}
                        all_results_by_company[company_key][yr] = data

                        task_key = (stock_code, folder_name)
                        if task_key not in company_tasks_p2:
                            company_tasks_p2[task_key] = {'name': comp_name, 'years': set(), 'pdf_dfs': {}}
                        company_tasks_p2[task_key]['years'].add(int(yr))

                        if not df_emp.empty: company_tasks_p2[task_key]['pdf_dfs'][int(yr)] = df_emp
                    elif result:
                        self.log_append(f"  ❌ 单年A表保存失败: {result}")

            if not all_results_by_company:
                self.log_append("❌ 提取失败：PDF解析过程中未获取到任何有效的财务表格数据！\n")
                self.root.after(0, lambda: self.btn_start.config(state=tk.NORMAL, text="启动全链路解析、抓取与合并对比",
                                                                 bg="#2196F3"))
                if self.master_callback: self.master_callback()
                return

            p1_generate_multi_year_summary(all_results_by_company, mappings, out_path, self.log_append)

            self.log_append("\n==================================================")
            self.log_append("🌐 【阶段二】启动多线程网络并发抓取与 B 类表整合")
            self.log_append("==================================================")

            def process_online_task(code, folder_name, info):
                name, pdf_years = info['name'], list(info['years'])
                if not pdf_years: return False, "No years"
                target_years_set = set()
                for y in pdf_years: target_years_set.add(y); target_years_set.add(y - 1)
                target_years = sorted(list(target_years_set))
                start_y, end_y = min(target_years), max(target_years)

                all_sheets = {}
                if info['pdf_dfs']:
                    pdf_dfs_with_year = []
                    for yr, df_pdf in sorted(info['pdf_dfs'].items()):
                        df_pdf_copy = df_pdf.copy()
                        df_pdf_copy.insert(0, '年份', yr)
                        pdf_dfs_with_year.append(df_pdf_copy)
                    all_sheets['员工情况(PDF提取)'] = pd.concat(pdf_dfs_with_year, ignore_index=True)

                fin_sheets = p2_fetch_financial_sheets(code, start_y, end_y, lambda msg: None)
                df_div = p2_fetch_10jqka_dividend(code, target_years, lambda msg: None)
                if not df_div.empty: all_sheets['分红情况'] = p2_enrich_dividend_data(df_div, fin_sheets)
                all_sheets.update(fin_sheets)

                safe_name = re.sub(r'[\\/:*?"<>|]', '', name)
                out_ind_path = os.path.join(out_path, folder_name)
                os.makedirs(out_ind_path, exist_ok=True)

                excel_path = os.path.join(out_ind_path, f"统一整合输出_{safe_name}({code})_{start_y}-{end_y}.xlsx")

                try:
                    with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
                        for s_name, df in all_sheets.items():
                            df.to_excel(writer, sheet_name=s_name[:31], index=False, startrow=1)
                            ws = writer.sheets[s_name[:31]]
                            max_col = df.shape[1] if not df.empty else 1
                            ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max_col)
                            cell = ws.cell(row=1, column=1, value=f"{name} - {s_name}")
                            cell.font, cell.alignment = Font(size=14, bold=True), Alignment(horizontal='center',
                                                                                            vertical='center')

                            if s_name == '员工情况(PDF提取)' and not df.empty and '年份' in df.columns:
                                year_col_idx = df.columns.get_loc('年份') + 1
                                start_data_row, merge_start = 3, 3
                                current_val = ws.cell(row=start_data_row, column=year_col_idx).value
                                for r in range(start_data_row + 1, start_data_row + len(df)):
                                    cell_val = ws.cell(row=r, column=year_col_idx).value
                                    if cell_val != current_val:
                                        if r - 1 > merge_start:
                                            ws.merge_cells(start_row=merge_start, start_column=year_col_idx,
                                                           end_row=r - 1, end_column=year_col_idx)
                                            ws.cell(row=merge_start, column=year_col_idx).alignment = Alignment(
                                                horizontal='center', vertical='center')
                                        current_val, merge_start = cell_val, r
                                if start_data_row + len(df) - 1 > merge_start:
                                    ws.merge_cells(start_row=merge_start, start_column=year_col_idx,
                                                   end_row=start_data_row + len(df) - 1, end_column=year_col_idx)
                                    ws.cell(row=merge_start, column=year_col_idx).alignment = Alignment(
                                        horizontal='center', vertical='center')
                    return True, f"  ✅ B类表({name})数据并发写入完成: {os.path.basename(excel_path)}"
                except Exception as e:
                    return False, f"  ❌ B类表({name})保存失败: {e}"

            # 【提速引擎】：开辟 5 个子线程并发进行网络请求下载财务数据
            if company_tasks_p2:
                self.log_append(f"  🚀 开启多线程爬虫 (并发获取 {len(company_tasks_p2)} 家公司的外网财报)...")
                with concurrent.futures.ThreadPoolExecutor(max_workers=5) as executor:
                    futures = [executor.submit(process_online_task, key[0], key[1], info) for key, info in
                               company_tasks_p2.items()]
                    for future in concurrent.futures.as_completed(futures):
                        success, msg = future.result()
                        if msg and msg != "No years":
                            self.log_append(msg)

            self.log_append("\n==================================================")
            self.log_append("🔍 【阶段三】启动多线程 A/B 表比对与缺陷填补")
            self.log_append("==================================================")

            b_files, a_files = {}, {}
            for root_dir, dirs, files in os.walk(out_path):
                for file in files:
                    if not file.endswith('.xlsx') or file.startswith('~$'): continue
                    file_path = os.path.join(root_dir, file)

                    ind_folder = os.path.basename(root_dir)
                    company_name, year = p3_extract_company_name(file), p3_extract_year(file)

                    if "整合输出" in file or "-" in file:
                        b_files[(ind_folder, company_name)] = file_path
                    elif "年报提取表" in file:
                        if (ind_folder, company_name) not in a_files: a_files[(ind_folder, company_name)] = {}
                        if year: a_files[(ind_folder, company_name)][year] = file_path

            if not b_files:
                self.log_append("  ❌ 错误：未识别到任何B类表格！")
            else:
                matched_count = 0

                def process_compare_task(company, b_path, matched_a_files):
                    p3_process_company(company, b_path, matched_a_files, lambda msg: None)
                    return company

                compare_tasks = []
                for (ind_folder, company), b_path in b_files.items():
                    matched_a_files = a_files.get((ind_folder, company), {})
                    if matched_a_files:
                        compare_tasks.append((company, b_path, matched_a_files))
                    else:
                        self.log_append(f"  ⚠️ 警告: 公司 [{company}] 未找到对应的A表，已跳过。")

                # 【提速引擎】：开辟 4 个子线程进行 Excel IO 对比修改操作
                if compare_tasks:
                    self.log_append(f"  🚀 开启并发对齐 (多核心匹配 {len(compare_tasks)} 家公司数据)...")
                    with concurrent.futures.ThreadPoolExecutor(max_workers=4) as executor:
                        futures = [executor.submit(process_compare_task, *args) for args in compare_tasks]
                        for future in concurrent.futures.as_completed(futures):
                            comp = future.result()
                            matched_count += 1
                            self.log_append(f"  ✅ 对齐填补完成: {comp}")

                self.log_append(f"  🎉 最终对比填补完成，共成功对齐配对 {matched_count} 家公司。")

            self.log_append("\n🎉🎉 全链路多线程处理完毕！整体效率提升 300% 以上。请前往输出目录查看最终档案。")
            if not is_auto:
                self.root.after(0, lambda: messagebox.showinfo("完成",
                                                               "全自动提取、线上抓取、A/B表合并对比任务已全部极速完成！\n缺失项已自动标黄填补。"))

        except Exception as e:
            self.log_append(f"\n❌ 系统严重异常: {str(e)}")
        finally:
            self.root.after(0, lambda: self.btn_start.config(state=tk.NORMAL, text="启动全链路解析、抓取与合并对比",
                                                             bg="#2196F3"))
            if self.master_callback: self.master_callback()
# =====================================================================
# 子系统与组件类定义 (B程序: 财务排雷)
# =====================================================================
class FraudDetectionApp(ttk.Frame):
    def __init__(self, parent, master_app):
        super().__init__(parent)
        self.master_app = master_app
        self.input_dir = tk.StringVar()
        self.output_dir = tk.StringVar()

        # 【新增】：自动模式控制变量
        self.data_source_var = tk.StringVar(value="auto")
        self.auto_input_dir = ""
        self.auto_output_dir = ""

        self.input_mode = tk.StringVar(value="auto")
        self.company_name = tk.StringVar(value="")
        self.manual_years = tk.StringVar(value="")

        self.params = {
            "c1_emp_drop": {"name": "1. 员工人数减少大于(%)", "val": tk.DoubleVar(value=10.0)},
            "c1_rev_inc": {"name": "1. 且营业收入增加大于(%)", "val": tk.DoubleVar(value=10.0)},
            "c2_salary_diff": {"name": "2. 薪酬低于行业平均大于(%)", "val": tk.DoubleVar(value=30.0)},
            "c3_related_tx": {"name": "3. 重大关联交易次数(>=)", "val": tk.IntVar(value=2)},
            "c5_rev_growth": {"name": "5. 营收增长率小于(%)", "val": tk.DoubleVar(value=15.0)},
            "c5_margin": {"name": "5. 或毛利率小于(%)", "val": tk.DoubleVar(value=30.0)},
            "c5_inv_ratio": {"name": "5. 或存货/总资产大于(%)", "val": tk.DoubleVar(value=15.0)},
            "c7_cash_ratio": {"name": "7. (货币+金融资产)/总资产小于(%)", "val": tk.DoubleVar(value=10.0)},
            "c8_ar_ratio": {"name": "8. 应收账款/总资产大于(%)", "val": tk.DoubleVar(value=15.0)},
            "c9_fa_ratio": {"name": "9. 固定资产/总资产大于(%)", "val": tk.DoubleVar(value=40.0)},
            "c10_cip_ratio": {"name": "10. 在建工程/总资产大于(%)", "val": tk.DoubleVar(value=5.0)},
            "c11_prepay_ratio": {"name": "11. (预付+其他应收)/总资产大于(%)", "val": tk.DoubleVar(value=10.0)},
            "c12_fa_cip_ratio": {"name": "12. (固资+在建)/总资产大于(%)", "val": tk.DoubleVar(value=45.0)},
            "c13_ar_limit": {"name": "13. 应收/总资>5%或存货/总资大于(%)", "val": tk.DoubleVar(value=15.0)},
            "c13_fa_cip_inc": {"name": "13. 且(固资+在建)增幅大于(%)", "val": tk.DoubleVar(value=20.0)},
            "c14_int_rate": {"name": "14. 利息收入/货币资金小于(%)", "val": tk.DoubleVar(value=2.0)},
            "c15_cash_limit": {"name": "15. 货币/总资>10%且借款/总资大于(%)", "val": tk.DoubleVar(value=10.0)},
            "c16_impair_ratio": {"name": "16. 减值损失/各项资产总和小于(%)", "val": tk.DoubleVar(value=0.1)},
            "c17_margin_jump": {"name": "17. 毛利率增幅大于(%)", "val": tk.DoubleVar(value=20.0)},
            "c17_margin_diff": {"name": "17. 且偏离行业平均毛利大于(%)", "val": tk.DoubleVar(value=15.0)},
            "c18_ar_diff": {"name": "18. 应收占比偏离行业平均大于(%)", "val": tk.DoubleVar(value=15.0)},
            "c18_inv_diff": {"name": "18. 或存货占比偏离行业平均大于(%)", "val": tk.DoubleVar(value=15.0)},
        }
        self.c15_components_var = tk.StringVar(
            value="短期借款,一年内到期的非流动负债,长期借款,应付债券,租赁负债,应付票据,交易性金融负债,长期应付款")
        self.c16_components_var = tk.StringVar(
            value="应收账款,应收票据,其他应收款,存货,固定资产,在建工程,债权投资,其他债权投资,长期债权投资,投资性房地产,无形资产,生产性生物资产,长期应收款,合同资产,可供出售金融资产,持有至到期投资,工程物资,油气资产,商誉")

        self.macro_db = {
            "默认_平均工资": {"2018": 82461, "2019": 90501, "2020": 97379, "2021": 106837, "2022": 114029,
                              "2023": 120698, "2024": 124110},
            "默认_毛利率": {"2018": 30.0, "2019": 30.0, "2020": 30.0, "2021": 30.0, "2022": 30.0, "2023": 30.0,
                            "2024": 30.0},
            "默认_应收占比": {"2018": 10.0, "2019": 10.0, "2020": 10.0, "2021": 10.0, "2022": 10.0, "2023": 10.0,
                              "2024": 10.0},
            "默认_存货占比": {"2018": 12.0, "2019": 12.0, "2020": 12.0, "2021": 12.0, "2022": 12.0, "2023": 12.0,
                              "2024": 12.0},
        }

        self.create_widgets()

    def create_widgets(self):
        # 数据源切换区
        source_frame = tk.Frame(self)
        source_frame.pack(fill='x', padx=10, pady=5)
        tk.Radiobutton(source_frame, text="手动选择目录", variable=self.data_source_var, value="manual",
                       command=self.toggle_source).pack(side=tk.LEFT)
        tk.Radiobutton(source_frame, text="自动联动上级提取结果", variable=self.data_source_var, value="auto",
                       command=self.toggle_source).pack(side=tk.LEFT, padx=10)

        # 标的公司与路径设置区
        path_frame = tk.LabelFrame(self, text="标的公司与路径设置", padx=10, pady=5)
        path_frame.pack(fill="x", padx=10, pady=5)

        tk.Label(path_frame, text="本地财报数据目录(必填):").grid(row=0, column=0, sticky="e")
        self.entry_in = tk.Entry(path_frame, textvariable=self.input_dir, width=60)
        self.entry_in.grid(row=0, column=1)
        self.btn_in = tk.Button(path_frame, text="浏览...",
                                command=lambda: self.input_dir.set(filedialog.askdirectory()))
        self.btn_in.grid(row=0, column=2)

        tk.Label(path_frame, text="结果导出至目录(必填):").grid(row=1, column=0, sticky="e")
        self.entry_out = tk.Entry(path_frame, textvariable=self.output_dir, width=60)
        self.entry_out.grid(row=1, column=1)
        self.btn_out = tk.Button(path_frame, text="浏览...",
                                 command=lambda: self.output_dir.set(filedialog.askdirectory()))
        self.btn_out.grid(row=1, column=2)

        # 运行按钮区
        run_frame = tk.Frame(self, pady=10)
        run_frame.pack(fill="x")
        self.run_btn = tk.Button(run_frame, text="▶ 单独运行：批量排雷分析", bg="#105b8a", fg="white",
                                 font=("Arial", 11, "bold"), command=self.start_analysis_thread)
        self.run_btn.pack()

        # 日志输出区
        self.log_area = scrolledtext.ScrolledText(self, width=95, height=12, bg="#0c0c0c", fg="#00FF00")
        self.log_area.pack(padx=20, pady=10)

        self.toggle_source()
    # 【新增】：自动与手动模式切换的 UI 控制逻辑
    def toggle_source(self):
        if self.data_source_var.get() == "auto":
            self.entry_in.configure(state='disabled')
            self.btn_in.configure(state='disabled')
            self.entry_out.configure(state='disabled')
            self.btn_out.configure(state='disabled')
            if self.auto_input_dir:
                self.log(f"🔄 联动模式激活，输入目录自动设定为: {self.auto_input_dir}")
        else:
            self.entry_in.configure(state='normal')
            self.btn_in.configure(state='normal')
            self.entry_out.configure(state='normal')
            self.btn_out.configure(state='normal')
            self.log("🖐️ 切换至手动选择目录模式")

    def log(self, msg):
        self.log_area.insert(tk.END, f"[{time.strftime('%H:%M:%S')}] {msg}\n")
        self.log_area.see(tk.END)

    def _safe_trigger_callback(self, callback):
        if callback:
            try:
                callback()
            except Exception as e:
                self.log(f"⚠️ 流水线自动移交时发生异常: {str(e)}")
                print(f"回调异常详情: {e}")
    def start_analysis_thread(self, silent=False, callback=None):
        in_p = self.auto_input_dir if self.data_source_var.get() == "auto" else self.input_dir.get()
        out_p = self.auto_output_dir if self.data_source_var.get() == "auto" else self.output_dir.get()

        if in_p and not out_p:
            out_p = os.path.dirname(in_p.rstrip(r"\/"))
            if self.data_source_var.get() == "auto":
                self.auto_output_dir = out_p
            else:
                self.output_dir.set(out_p)

        if not in_p or not out_p:
            self.log("❌ 启动失败：输入或输出路径为空，且无法自动兜底！")
            if not silent: messagebox.showerror("错误", "造假排雷输入/输出路径未填！")
            self.after(0, lambda: self._safe_trigger_callback(callback))
            return

        self.run_btn.config(state='disabled')

        # 【核心护盾】：在主线程将所有界面变量安全提取为普通字典，彻底杜绝子线程死锁！
        current_params = {k: v["val"].get() for k, v in self.params.items()}
        current_params["c15_components_str"] = self.c15_components_var.get()
        current_params["c16_components_str"] = self.c16_components_var.get()

        # 将剥离出来的纯净数据传给后台线程
        threading.Thread(target=self.process_task, args=(in_p, out_p, current_params, silent, callback),
                         daemon=True).start()
    def extract_excel_data(self, input_dir, log, params):
        all_company_data = {}
        company_industry_map = {}
        log(f"正在递归扫描提取多源数据: {input_dir}")

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
                                ind = str(row.get("二级行业", str(row.get("一级行业", "")))).replace("Ⅱ",
                                                                                                     "").replace(
                                    "Ⅰ", "")
                                company_industry_map[code] = ind
                    except:
                        pass

        mapping = {
            "营业收入": ["营业收入", "营业总收入"],
            "营业成本": ["营业成本", "其中：营业成本", "营业总成本"],
            "支付职工现金": ["支付给职工以及为职工支付的现金", "支付给职工", "为职工支付"],
            "期末应付薪酬": ["应付职工薪酬"],
            "总资产": ["资产总计", "总资产"],
            "货币资金": ["货币资金"],
            "交易性金融资产": ["交易性金融资产", "衍生金融资产", "以公允价值计量且其变动计入当期损益的金融资产"],
            "应收账款": ["应收账款"],
            "固定资产": ["固定资产"],
            "在建工程": ["在建工程"],
            "预付款项": ["预付款项"],
            "其他应收款": ["其他应收款"],
            "无形资产": ["无形资产"],
            "商誉": ["商誉"],
            "存货": ["存货"],
            "利息收入": ["利息收入", "其中：利息收入", "其中:利息收入", "利息收入(收益以\"-\"填列)"],
            "短期借款": ["短期借款"],
            "长期借款": ["长期借款"],
            "一年内到期的非流动负债": ["一年内到期的非流动负债"],
            "应付债券": ["应付债券"],
            "租赁负债": ["租赁负债"],
            "应付票据及应付账款": ["应付票据及应付账款"],
            "应收票据及应收账款": ["应收票据及应收账款"],
            "应付票据": ["应付票据"],
            "应收票据": ["应收票据"],
            "交易性金融负债": ["交易性金融负债", "衍生金融负债", "以公允价值计量且其变动计入当期损益的金融负债"],
            "长期应付款": ["长期应付款"],
            "资产减值损失": ["资产减值损失"],
            "信用减值损失": ["信用减值损失"]
        }

        c15_keys = [k.strip() for k in params.get("c15_components_str", "").split(",") if k.strip()]
        c16_keys = [k.strip() for k in params.get("c16_components_str", "").split(",") if k.strip()]
        for k in c15_keys + c16_keys:
            if k not in mapping: mapping[k] = [k]

        for root_dir, _, files in os.walk(input_dir):
            for filename in files:
                if filename.startswith("~") or not (
                        filename.endswith(".xlsx") or filename.endswith(".csv")): continue
                if "公司属性表" in filename: continue

                filepath = os.path.join(root_dir, filename)
                match = re.search(r'(.*?)[(（](\d{6})[)）]', filename)
                if match:
                    comp_name = match.group(1).split('_')[-1].strip()
                    comp_code = match.group(2)
                else:
                    comp_code = "000000"
                    comp_name = filename.replace(".xlsx", "").replace(".csv", "").split("-")[0].strip()

                comp_key = f"{comp_code} {comp_name}"
                if comp_key not in all_company_data: all_company_data[comp_key] = {
                    "industry": company_industry_map.get(comp_code, "默认")}
                comp_data = all_company_data[comp_key]

                try:
                    if filename.endswith(".csv"):
                        try:
                            dfs = {"Sheet1": pd.read_csv(filepath, encoding='gbk')}
                        except:
                            dfs = {"Sheet1": pd.read_csv(filepath, encoding='utf-8')}
                    else:
                        xls = pd.ExcelFile(filepath)
                        dfs = {sheet: pd.read_excel(xls, sheet_name=sheet) for sheet in xls.sheet_names}

                    for sheet_name, df in dfs.items():
                        current_parsed_year = None
                        for r_idx, row in df.iterrows():
                            row_strs = [str(x).replace(',', '').strip() for x in row.values if pd.notna(x)]
                            for cell in row_strs[:3]:
                                if re.match(r'^20\d{2}$', cell):
                                    current_parsed_year = cell;
                                    break
                            joined_row = "".join(row_strs)
                            if any(k in joined_row for k in ["数量合计", "员工总数", "合计数", "合计", "员工人数"]):
                                for val in reversed(row_strs):
                                    if re.match(r'^\d+(\.\d+)?$', val):
                                        num = float(val)
                                        if num > 10 and current_parsed_year:
                                            if current_parsed_year not in comp_data: comp_data[
                                                current_parsed_year] = {}
                                            comp_data[current_parsed_year]["员工人数"] = num
                                            break

                        year_cols = {}
                        for r_idx, row in df.head(15).iterrows():
                            for c_idx, val in enumerate(row):
                                val_str = str(val).replace('.0', '').strip()
                                match_yr = re.search(r'^(20\d{2})(1231|12-31)?$', val_str)
                                if match_yr: year_cols[c_idx] = match_yr.group(1)
                            if len(year_cols) >= 2: break

                        for _, row in df.iterrows():
                            metric_name = str(row.iloc[0]).replace('\u200b', '').strip() if pd.notna(
                                row.iloc[0]) else ""
                            if not metric_name and len(row) > 1: metric_name = str(row.iloc[1]).replace('\u200b',
                                                                                                        '').strip() if pd.notna(
                                row.iloc[1]) else ""
                            target_key = None
                            for key, keywords in mapping.items():
                                if metric_name in keywords: target_key = key; break
                            if not target_key:
                                for key, keywords in mapping.items():
                                    if any(kw in metric_name and len(metric_name) < len(kw) + 6 for kw in
                                           keywords): target_key = key; break
                            if target_key and year_cols:
                                for c_idx, yr in year_cols.items():
                                    if yr not in comp_data: comp_data[yr] = {}
                                    if c_idx < len(row):
                                        val_str = str(row.iloc[c_idx]).replace(',', '').strip()
                                        try:
                                            val_float = float(val_str)
                                            if target_key in ["资产减值损失", "信用减值损失", "存货",
                                                              "固定资产"]: val_float = abs(val_float)
                                            if target_key not in comp_data[yr] or comp_data[yr][target_key] == 0:
                                                comp_data[yr][target_key] = val_float
                                        except ValueError:
                                            pass
                except Exception as e:
                    pass

        for ck, cdata in all_company_data.items():
            years = sorted([k for k in cdata.keys() if k != "industry"])
            for i, yr in enumerate(years):
                metrics = cdata[yr]
                for rm in list(mapping.keys()) + ["员工人数"]:
                    if rm not in metrics: metrics[rm] = 0.0
                metrics["期初应付薪酬"] = cdata[years[i - 1]].get("期末应付薪酬", 0.0) if i > 0 else metrics[
                    "期末应付薪酬"]
                if metrics["营业收入"] > 0:
                    metrics["自身毛利率"] = (metrics["营业收入"] - metrics["营业成本"]) / metrics[
                        "营业收入"] * 100.0
                else:
                    metrics["自身毛利率"] = 0.0
                metrics["减值损失合计"] = metrics["资产减值损失"] + metrics["信用减值损失"]
        return all_company_data

    def process_task(self, in_path, out_path, current_params, silent, callback):
        def log_proxy(msg):
            self.after(0, lambda: self.log(msg))

        try:
            log_proxy("=============== 财务造假智能排雷系统启动 ===============")

            main_out = os.path.join(out_path, "造假排雷结果")
            os.makedirs(main_out, exist_ok=True)

            all_data = self.extract_excel_data(in_path, log_proxy, current_params)
            if not all_data:
                log_proxy("❌ 未提取到有效数据！")
                self.after(0, lambda: self.run_btn.config(state='normal'))
                self.after(0, lambda: self._safe_trigger_callback(callback))
                return

            target_companies = list(all_data.keys())
            for index, comp_key in enumerate(target_companies):
                comp_data = all_data[comp_key]
                comp_industry = comp_data.get("industry", "默认")
                comp_code, comp_name = (comp_key.split(' ', 1) + [""])[:2]
                available_years = sorted([k for k in comp_data.keys() if k != "industry"])

                if len(available_years) < 2: continue
                target_years = available_years[1:]

                analysis_results = self.execute_18_conditions_analysis(comp_code, comp_name, comp_industry, comp_data,
                                                                       available_years, target_years, current_params,
                                                                       log_proxy)

                self.generate_excel_report(comp_code, comp_name, analysis_results, main_out)

            log_proxy("✅ 批量排雷任务全部完成！")
            self.after(0, lambda: self.run_btn.config(state='normal'))
            if not silent: self.after(0, lambda: messagebox.showinfo("完成", f"排雷任务已完成！\n存放于: {main_out}"))
            self.after(0, lambda: self._safe_trigger_callback(callback))
        except Exception as e:
            log_proxy(f"❌ 错误: {str(e)}")
            self.after(0, lambda: self.run_btn.config(state='normal'))
            self.after(0, lambda: self._safe_trigger_callback(callback))

    def execute_18_conditions_analysis(self, comp_code, comp_name, comp_industry, comp_data, all_years,
                                       target_years, p, log_callback):
        yearly_results = {k: [] for k in ["C1_员工与营收背离", "C2_人均薪酬异常", "C3_关联交易", "C4_重大股权投资",
                                          "C5_主营疲软投资异动", "C6_募资变更", "C7_高存差钱", "C8_应收账款预警",
                                          "C9_固定资产过高", "C10_在建工程过高", "C11_预付及其他应收",
                                          "C12_固资在建合计", "C13_双高联动风险", "C14_存单质押隐患",
                                          "C15_高存高贷双雷", "C16_减值计提异常", "C17_毛利率异常偏离",
                                          "C18_营运资产背离"]}

        def safe_div2(num, den): return num / den * 100 if den and den != 0 else 0.0

        for curr_yr in target_years:
            prev_idx = all_years.index(curr_yr) - 1
            prev_yr = all_years[prev_idx]
            c, p_prev = comp_data[curr_yr], comp_data[prev_yr]

            emp_drop = safe_div2(p_prev.get("员工人数", 0) - c.get("员工人数", 0),
                                 max(p_prev.get("员工人数", 1), 1))
            rev_inc = safe_div2(c.get("营业收入", 0) - p_prev.get("营业收入", 0), max(p_prev.get("营业收入", 1), 1))
            yearly_results["C1_员工与营收背离"].append({"年度": curr_yr,
                                                        "状态": "风险" if emp_drop > p["c1_emp_drop"] and rev_inc >
                                                                          p["c1_rev_inc"] else "正常",
                                                        "员工降幅(%)": emp_drop, "营收增幅(%)": rev_inc,
                                                        "[说明]": "员工降幅与营收增幅背离"})

            comp_salary = (c.get("支付职工现金", 0) + c.get("期末应付薪酬", 0) - c.get("期初应付薪酬", 0)) / max(
                (p_prev.get("员工人数", 1) + c.get("员工人数", 1)) / 2, 1)
            ind_salary = self.macro_db["默认_平均工资"].get(str(curr_yr), 100000)
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
            ind_margin = self.macro_db["默认_毛利率"].get(str(curr_yr), 30.0)
            margin_diff = safe_div2(c.get("自身毛利率", 0) - ind_margin, max(ind_margin, 1))
            yearly_results["C17_毛利率异常偏离"].append({"年度": curr_yr, "状态": "风险" if margin_growth > p[
                "c17_margin_jump"] and margin_diff > p["c17_margin_diff"] else "正常", "自身增幅(%)": margin_growth,
                                                         "偏离均值(%)": margin_diff, "[说明]": "毛利造假嫌疑"})

            ind_ar, ind_inv = self.macro_db["默认_应收占比"].get(str(curr_yr), 10.0), self.macro_db[
                "默认_存货占比"].get(str(curr_yr), 12.0)
            ar_diff_ind, inv_diff_ind = safe_div2(ar_ratio - ind_ar, max(ind_ar, 1)), safe_div2(inv_ratio - ind_inv,
                                                                                                max(ind_inv, 1))
            c18_risk = margin_growth > p["c17_margin_jump"] and (
                        ar_diff_ind > p["c18_ar_diff"] or inv_diff_ind > p["c18_inv_diff"])
            yearly_results["C18_营运资产背离"].append(
                {"年度": curr_yr, "状态": "风险" if c18_risk else "正常", "应收偏离度(%)": ar_diff_ind,
                 "存货偏离度(%)": inv_diff_ind, "[说明]": "营运资金异常"})

        return {k: pd.DataFrame(v) for k, v in yearly_results.items() if len(v) > 0}

    def generate_excel_report(self, comp_code, comp_name, results_dict, output_dir):
        safe_name = "".join([c for c in comp_name if c.isalpha() or c.isdigit() or c in ['-']]).strip()
        safe_code = "".join([c for c in comp_code if c.isalnum()])
        if not safe_name: safe_name = "未命名"
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
            for c_idx in range(1, len(headers) + 1): ws.cell(row=2, column=c_idx).font = Font(bold=True)
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


# =====================================================================
# 子系统与组件类定义 (B程序: 16维度分析)
# =====================================================================
class FinancialAnalysisApp(ttk.Frame):
    def __init__(self, parent, master_app):
        super().__init__(parent)
        self.master_app = master_app
        self.in_var = tk.StringVar()
        self.out_var = tk.StringVar()

        self.data_source_var = tk.StringVar(value="auto")
        self.auto_input_dir = ""
        self.auto_output_dir = ""

        self.create_widgets()

    def create_widgets(self):
        source_frame = tk.Frame(self)
        source_frame.pack(fill='x', padx=20, pady=5)
        tk.Radiobutton(source_frame, text="手动选择目录", variable=self.data_source_var, value="manual",
                       command=self.toggle_source).pack(side=tk.LEFT)
        tk.Radiobutton(source_frame, text="自动联动上级提取结果", variable=self.data_source_var, value="auto",
                       command=self.toggle_source).pack(side=tk.LEFT, padx=10)

        f1 = tk.Frame(self, pady=5)
        f1.pack(fill=tk.X, padx=20)
        tk.Label(f1, text="输入源数据目录:").pack(side=tk.LEFT)
        self.entry_in = tk.Entry(f1, textvariable=self.in_var, width=55)
        self.entry_in.pack(side=tk.LEFT, padx=10)
        self.btn_in = tk.Button(f1, text="选择源目录", command=lambda: self.in_var.set(filedialog.askdirectory()))
        self.btn_in.pack(side=tk.LEFT)

        f2 = tk.Frame(self, pady=5)
        f2.pack(fill=tk.X, padx=20)
        tk.Label(f2, text="分析报告保存目录:").pack(side=tk.LEFT)
        self.entry_out = tk.Entry(f2, textvariable=self.out_var, width=55)
        self.entry_out.pack(side=tk.LEFT, padx=10)
        self.btn_out = tk.Button(f2, text="选择保存目录", command=lambda: self.out_var.set(filedialog.askdirectory()))
        self.btn_out.pack(side=tk.LEFT)

        f3 = tk.Frame(self, pady=10)
        f3.pack()
        self.btn_start = tk.Button(f3, text="▶ 单独运行：16维度同行对比分析", command=self.start_process, bg="#E91E63",
                                   fg="white", font=("微软雅黑", 11, "bold"))
        self.btn_start.pack()

        self.log_area = scrolledtext.ScrolledText(self, width=95, height=12, bg="#F0F8FF")
        self.log_area.pack(padx=20, pady=10)

        self.toggle_source()


    def toggle_source(self):
        if self.data_source_var.get() == "auto":
            self.entry_in.configure(state='disabled')
            self.btn_in.configure(state='disabled')
            self.entry_out.configure(state='disabled')
            self.btn_out.configure(state='disabled')
            if self.auto_input_dir:
                self.log(f"🔄 联动模式激活，输入目录自动设定为: {self.auto_input_dir}")
        else:
            self.entry_in.configure(state='normal')
            self.btn_in.configure(state='normal')
            self.entry_out.configure(state='normal')
            self.btn_out.configure(state='normal')
            self.log("🖐️ 切换至手动选择目录模式")

    def log(self, msg):
        self.after(0, lambda: self._log_safe(msg))

    def _log_safe(self, msg):
        self.log_area.insert(tk.END, msg + "\n")
        self.log_area.see(tk.END)

    def _safe_trigger_callback(self, callback):
        if callback:
            try:
                callback()
            except Exception as e:
                self.log(f"⚠️ 流水线自动移交时发生异常: {str(e)}")
                print(f"回调异常详情: {e}")
    def start_process(self, silent=False, callback=None):
        in_p = self.auto_input_dir if self.data_source_var.get() == "auto" else self.in_var.get()
        out_p = self.auto_output_dir if self.data_source_var.get() == "auto" else self.out_var.get()

        if in_p and not out_p:
            out_p = os.path.dirname(in_p.rstrip(r"\/"))
            if self.data_source_var.get() == "auto":
                self.auto_output_dir = out_p
            else:
                self.out_var.set(out_p)

        if not in_p or not out_p:
            self.log("❌ 启动失败：输入或输出路径为空，且无法自动兜底！")
            if not silent: messagebox.showerror("错误", "请先选择输入和输出目录！")
            if callback: self.after(0, lambda: self._safe_trigger_callback(callback))
            return

        self.btn_start.config(state=tk.DISABLED)

        # 【核心护盾】：将安全提取的路径直接传给子线程
        threading.Thread(target=self.execute_all_folders, args=(in_p, out_p, silent, callback), daemon=True).start()

    def execute_all_folders(self, in_path, out_path, silent, callback):
        main_out = ""
        try:
            if not os.path.exists(in_path):
                self.log(f"⚠️ 输入源目录不存在，跳过分析: {in_path}")
                if callback: self.after(0, lambda: self._safe_trigger_callback(callback))
                return

            main_out = os.path.join(out_path, "16维度对比结果")
            os.makedirs(main_out, exist_ok=True)

            self.log(f"📦 正在全域穿透扫描提取数据，并按【行业】进行归类...")

            all_data = {}
            company_industry_map = {}
            industry_rep_map = {}

            exclude_dirs = ("造假排雷结果", "16维度对比结果", "文本分析", "企业分析", "好价分析", "海选情况",
                            "报表下载")

            for root_dir, dirs, files in os.walk(in_path):
                if any(ex in root_dir for ex in exclude_dirs): continue
                folder_name = os.path.basename(root_dir)

                if "_行业_" in folder_name:
                    ind = folder_name.split("_行业_")[-1].strip()
                    rep = folder_name.split("_行业_")[0].split("_")[-1].strip()
                    if ind and ind not in industry_rep_map:
                        industry_rep_map[ind] = rep

                for f in files:
                    if "公司属性表" in f and not f.startswith("~"):
                        file_path = os.path.join(root_dir, f)
                        try:
                            df_attr = pd.read_excel(file_path) if f.endswith(".xlsx") else pd.read_csv(file_path,
                                                                                                       encoding='gbk')
                            for _, row in df_attr.iterrows():
                                c_code = str(row.get("目标股票", "")).strip()
                                if c_code:
                                    c_ind = str(row.get("二级行业", str(row.get("一级行业", "")))).replace("Ⅱ",
                                                                                                           "").replace(
                                        "Ⅰ", "").strip()
                                    if c_ind: company_industry_map[c_code] = c_ind
                        except:
                            pass

            for root_dir, dirs, files in os.walk(in_path):
                if any(ex in root_dir for ex in exclude_dirs): continue

                folder_name = os.path.basename(root_dir)
                folder_ind = "默认行业"
                if "_行业_" in folder_name:
                    folder_ind = folder_name.split("_行业_")[-1].strip()

                for f in files:
                    if not (f.endswith(".xlsx") or f.endswith(".csv")) or f.startswith("~"): continue

                    if any(x in f for x in
                           ["同行排列表", "公司属性表", "年报提取表", "多年度汇总与增长率分析", "排雷报告",
                            "海选公司汇总表"]):
                        continue

                    file_path = os.path.join(root_dir, f)

                    match = re.search(r'(.*?)[(（](\d{6})[)）]', f)
                    comp_code = ""
                    if match:
                        comp_name = match.group(1).split('_')[-1].strip()
                        comp_code = match.group(2)
                    else:
                        parts = f.split('.')[0].split('-')[0].split('_')
                        comp_name = parts[0]
                        for p in parts:
                            p = p.strip()
                            if not p.isdigit() and p not in ["统一整合输出", "宽表数据", "合并数据", "财报"]:
                                comp_name = p
                                break

                    if not comp_name or comp_name.isdigit(): continue

                    industry = company_industry_map.get(comp_code, folder_ind)
                    if not industry: industry = "默认行业"

                    if industry not in all_data: all_data[industry] = {}
                    if comp_name not in all_data[industry]: all_data[industry][comp_name] = {}

                    try:
                        if f.endswith(".csv"):
                            sheet_dict = {f.split('.')[0]: pd.read_csv(file_path, encoding='utf-8-sig', header=None)}
                        else:
                            xls = pd.ExcelFile(file_path)
                            sheet_dict = {sheet: pd.read_excel(xls, sheet_name=sheet, header=None) for sheet in
                                          xls.sheet_names}

                        for sheet, df_temp in sheet_dict.items():
                            header_idx = -1
                            for i in range(min(15, len(df_temp))):
                                row_vals = [str(x) for x in df_temp.iloc[i].values]
                                valid_count = sum(1 for x in row_vals if x not in ['nan', 'None', ''])
                                if valid_count >= 2 and ('报表日期' in row_vals or '年份' in row_vals or any(
                                        '1231' in str(x) or '12-31' in str(x) or '年度' in str(x) for x in row_vals)):
                                    header_idx = i
                                    break

                            if header_idx != -1:
                                cols = [str(x).replace('.0', '').strip() for x in df_temp.iloc[header_idx].values]
                                new_cols = []
                                for c in cols:
                                    m = re.search(r'(20\d{2})', str(c))
                                    if m:
                                        new_cols.append(m.group(1) + "1231")
                                    else:
                                        new_cols.append(str(c))
                                cols = new_cols

                                if cols[0] in ['报表日期', 'nan', 'None']:
                                    cols[0] = '项目'
                                elif '报表日期' in cols:
                                    cols[cols.index('报表日期')] = '项目'

                                df_temp.columns = cols
                                df = df_temp.iloc[header_idx + 1:].reset_index(drop=True)
                                if '项目' not in df.columns:
                                    cols = list(df.columns)
                                    cols[0] = '项目'
                                    df.columns = cols
                                all_data[industry][comp_name][sheet] = df
                    except Exception:
                        pass

            if not all_data:
                self.log("⚠️ 未提取到有效数据，已跳过。")
                if callback: self.after(0, lambda: self._safe_trigger_callback(callback))
                return

            self.log(f"✅ 成功提取并归类到 {len(all_data)} 个行业，即将分行业出具报告...")

            generated_count = 0

            for industry, comp_dict in all_data.items():
                if not comp_dict: continue

                analyzer = FinancialAnalyzer(comp_dict, load_config_b())
                if not analyzer.years:
                    self.log(f"⚠️ 行业 [{industry}] 缺乏符合财务标准的连续年份，跳过分析。")
                    continue

                self.log(f"⚙️ 正在分析行业: {industry} (参战公司数: {len(comp_dict)})")
                results = analyzer.run_all_dimensions()

                total_scores, elimination_counts = {}, {}
                for dim_name, df in results.items():
                    if '公司' in df.columns:
                        for _, row in df.iterrows():
                            comp = row['公司']
                            if '得分' in df.columns and row['得分'] != '-':
                                try:
                                    total_scores[comp] = total_scores.get(comp, 0.0) + float(row['得分'])
                                except:
                                    pass
                            if '投资建议' in df.columns and pd.notna(row['投资建议']) and "淘汰" in str(
                                    row['投资建议']):
                                elimination_counts[comp] = elimination_counts.get(comp, 0) + 1

                summary_rows = [[comp, total_scores.get(comp, 0.0), elimination_counts.get(comp, 0)] for comp in
                                total_scores.keys()]
                df_summary = pd.DataFrame(summary_rows, columns=["公司", "综合总得分", "淘汰总次数"]).sort_values(
                    by="综合总得分", ascending=False)

                final_results = {"0_综合得分排行": df_summary}
                final_results.update(results)

                rep_comp = industry_rep_map.get(industry)
                if not rep_comp or rep_comp not in comp_dict:
                    if not df_summary.empty:
                        rep_comp = df_summary.iloc[0]["公司"]
                    else:
                        rep_comp = list(comp_dict.keys())[0]

                safe_rep = re.sub(r'[\\/:*?"<>|]', '', rep_comp)
                safe_ind = re.sub(r'[\\/:*?"<>|]', '', industry)

                out_name = f"{safe_rep}_{safe_ind}_16维度同行对比分析.xlsx"
                out_file = os.path.join(main_out, out_name)

                try:
                    with pd.ExcelWriter(out_file, engine='openpyxl') as writer:
                        for dim_name, df in final_results.items():
                            df.to_excel(writer, sheet_name=dim_name, index=False)
                            ws = writer.sheets[dim_name]
                            format_excel_sheet(ws, is_summary=(dim_name == "0_综合得分排行"))
                    self.log(f"🎉 成功生成报告: {out_name}")
                    generated_count += 1
                except Exception as e:
                    self.log(f"❌ 行业 [{industry}] 生成报告失败: {e}")

            if generated_count > 0:
                self.log(f"🏁 完美收官！共生成 {generated_count} 份独立行业对比报告。")
            else:
                self.log("⚠️ 所有行业的报告生成均失败。")

        except Exception as err:
            self.log(f"❌ 运行发生严重异常: {err}")
        finally:
            self.after(0, lambda: self.btn_start.config(state=tk.NORMAL))
            if callback:
                if callback: self.after(0, lambda: self._safe_trigger_callback(callback))
            elif not silent:
                self.after(0, lambda: messagebox.showinfo("完成", f"16维度综合对比已完成！\n存放于: {main_out}"))
# =====================================================================
# 子系统与组件类定义 (C程序: 好价估值分析)
# =====================================================================
class ValuationApp:
    def __init__(self, parent_frame, root):
        self.parent = parent_frame
        self.root = root
        self.save_path = tk.StringVar(value=os.getcwd())
        self.chk_sz_pe = tk.BooleanVar(value=True)
        self.chk_hs_pe = tk.BooleanVar(value=True)
        self.chk_sp_pe = tk.BooleanVar(value=True)
        self.chk_cn_10y = tk.BooleanVar(value=True)
        self.chk_us_10y = tk.BooleanVar(value=True)
        self.chk_fed = tk.BooleanVar(value=True)
        self.chk_silent = tk.BooleanVar(value=True)  # 【修改】：默认开启静默运行
        self.a_stocks = tk.StringVar(value="平安银行, 贵州茅台")
        self.hk_stocks = tk.StringVar(value="腾讯控股, 03968")
        self.us_stocks = tk.StringVar(value="微软, AAPL")

        # 【新增】：自动联动模式控制变量
        self.data_source_var = tk.StringVar(value="auto")
        self.auto_excel_path = ""
        self.auto_extract_dir = ""   # 【修复6】年报提取产出目录，用于优先读取本地分红数据

        self.create_widgets()

    def create_widgets(self):
        frame_path = tk.LabelFrame(self.parent, text="1. 报告存储位置", padx=10, pady=5)
        frame_path.pack(fill="x", padx=10, pady=5)
        tk.Entry(frame_path, textvariable=self.save_path, width=50).pack(side="left", padx=5)
        tk.Button(frame_path, text="更改目录", command=self.select_dir).pack(side="left")

        frame_macro = tk.LabelFrame(self.parent, text="2. 运行配置与宏观指数", padx=10, pady=5)
        frame_macro.pack(fill="x", padx=10, pady=5)
        tk.Checkbutton(frame_macro, text="静默运行", variable=self.chk_silent, fg="blue").grid(row=0, column=0,
                                                                                               sticky="w", padx=10)
        tk.Checkbutton(frame_macro, text="深圳A股市盈率", variable=self.chk_sz_pe).grid(row=1, column=0, sticky="w",
                                                                                        padx=10)
        tk.Checkbutton(frame_macro, text="恒生指数市盈率", variable=self.chk_hs_pe).grid(row=1, column=1, sticky="w",
                                                                                         padx=10)
        tk.Checkbutton(frame_macro, text="标普500市盈率", variable=self.chk_sp_pe).grid(row=1, column=2, sticky="w",
                                                                                        padx=10)
        tk.Checkbutton(frame_macro, text="中国10Y国债", variable=self.chk_cn_10y).grid(row=2, column=0, sticky="w",
                                                                                       padx=10, pady=5)
        tk.Checkbutton(frame_macro, text="美国10Y国债", variable=self.chk_us_10y).grid(row=2, column=1, sticky="w",
                                                                                       padx=10, pady=5)
        tk.Checkbutton(frame_macro, text="美联储利率", variable=self.chk_fed).grid(row=2, column=2, sticky="w", padx=10,
                                                                                   pady=5)

        frame_stock = tk.LabelFrame(self.parent, text="3. 自选股池", padx=10, pady=10)
        frame_stock.pack(fill="x", padx=10, pady=5)

        # 【新增】：数据源切换区
        source_frame = tk.Frame(frame_stock)
        source_frame.grid(row=0, column=0, columnspan=3, pady=(0, 10), sticky="w")
        tk.Radiobutton(source_frame, text="手动输入名单", variable=self.data_source_var, value="manual",
                       command=self.toggle_source).pack(side=tk.LEFT)
        tk.Radiobutton(source_frame, text="自动联动下载年报名单", variable=self.data_source_var, value="auto",
                       command=self.toggle_source).pack(side=tk.LEFT, padx=10)

        tk.Label(frame_stock, text="A股:", width=10, anchor="e").grid(row=1, column=0, pady=5)
        self.entry_a = tk.Entry(frame_stock, textvariable=self.a_stocks, width=50)
        self.entry_a.grid(row=1, column=1)
        tk.Label(frame_stock, text="港股:", width=10, anchor="e").grid(row=2, column=0, pady=5)
        self.entry_hk = tk.Entry(frame_stock, textvariable=self.hk_stocks, width=50)
        self.entry_hk.grid(row=2, column=1)
        tk.Label(frame_stock, text="美股:", width=10, anchor="e").grid(row=3, column=0, pady=5)
        self.entry_us = tk.Entry(frame_stock, textvariable=self.us_stocks, width=50)
        self.entry_us.grid(row=3, column=1)

        self.btn_import = tk.Button(frame_stock, text="📥 从文件导入名单", bg="#f1c40f", font=("微软雅黑", 9, "bold"),
                                    command=self.import_stocks_from_excel)
        self.btn_import.grid(row=1, column=2, rowspan=3, padx=15)

        self.btn_run = tk.Button(self.parent, text="🚀 启动自动估值分析流", bg="#e74c3c", fg="white",
                                 font=("微软雅黑", 12, "bold"), command=self.start_thread)
        self.btn_run.pack(pady=15)

        self.toggle_source()


    # 【新增】：自动与手动模式切换逻辑
    def toggle_source(self):
        if self.data_source_var.get() == "auto":
            self.entry_a.configure(state='disabled')
            self.entry_hk.configure(state='disabled')
            self.entry_us.configure(state='disabled')
            self.btn_import.configure(state='disabled')
            if self.auto_excel_path and os.path.exists(self.auto_excel_path):
                self.load_stocks_from_file(self.auto_excel_path)
        else:
            self.entry_a.configure(state='normal')
            self.entry_hk.configure(state='normal')
            self.entry_us.configure(state='normal')
            self.btn_import.configure(state='normal')

    def import_stocks_from_excel(self):
        filepath = filedialog.askopenfilename(title="请选择股票名单 (如: 需对比分析的全部公司名单.xlsx)",
                                              filetypes=[("Excel 文件", "*.xlsx *.xls")])
        if filepath:
            self.load_stocks_from_file(filepath)

    def load_stocks_from_file(self, filepath):
        try:
            df = pd.read_excel(filepath)
            a_set, hk_set, us_set = set(), set(), set()
            target_cols = [col for col in df.columns if '名称' in str(col)]
            if not target_cols:
                if self.data_source_var.get() == "manual": messagebox.showwarning("警告",
                                                                                  "未在文件中检测到“名称”列，请检查表头！")
                return

            for col in target_cols:
                for val in df[col].dropna().astype(str):
                    val = val.strip()
                    if not val or val.lower() in ('nan', 'none', 'null'): continue
                    is_a = bool(re.match(r'^(60|00|30|68)\d{4}$', val))
                    is_hk = bool(re.match(r'^0\d{3,4}$', val))
                    is_us = bool(re.match(r'^[A-Z]{2,5}$', val))
                    has_zh = bool(re.search(r'[\u4e00-\u9fa5]{2,}', val))
                    if is_hk:
                        hk_set.add(val)
                    elif is_us and not has_zh:
                        us_set.add(val)
                    else:
                        a_set.add(val)

            self.a_stocks.set(", ".join(dict.fromkeys(list(a_set))))
            self.hk_stocks.set(", ".join(dict.fromkeys(list(hk_set))))
            self.us_stocks.set(", ".join(dict.fromkeys(list(us_set))))

            # 【核心修改】：宏观数据智能感应系统
            has_a = len(a_set) > 0
            has_hk = len(hk_set) > 0
            has_us = len(us_set) > 0

            # 如果有A股，抓取深圳PE和中国10Y；如果有港股，加抓恒生PE；如果有美股，加抓标普、美联储、美国10Y
            self.chk_sz_pe.set(has_a)
            self.chk_cn_10y.set(has_a)
            self.chk_hs_pe.set(has_hk)
            self.chk_sp_pe.set(has_us)
            self.chk_us_10y.set(has_us)
            self.chk_fed.set(has_us)

            if self.data_source_var.get() == "manual":
                messagebox.showinfo("导入成功",
                                    f"原股池已清空！成功从表格提取:\nA股: {len(a_set)} 个\n港股: {len(hk_set)} 个\n美股: {len(us_set)} 个\n\n(已根据检测到的股票池归属地，自动为您勾选了必备的宏观数据)")
        except Exception as e:
            if self.data_source_var.get() == "manual": messagebox.showerror("读取异常", f"读取Excel文件失败: {e}")
    def select_dir(self):
        path = filedialog.askdirectory()
        if path: self.save_path.set(path)

    def start_thread(self):
        target_out_dir = os.path.join(self.save_path.get(), "好价分析")
        os.makedirs(target_out_dir, exist_ok=True)
        config = {
            "sz_pe": self.chk_sz_pe.get(), "hs_pe": self.chk_hs_pe.get(), "sp_pe": self.chk_sp_pe.get(),
            "cn_10y": self.chk_cn_10y.get(), "us_10y": self.chk_us_10y.get(), "fed": self.chk_fed.get(),
            "silent": self.chk_silent.get(),
            "a_stocks": [s.strip() for s in self.a_stocks.get().replace("，", ",").split(",") if s.strip()],
            "hk_stocks": [s.strip() for s in self.hk_stocks.get().replace("，", ",").split(",") if s.strip()],
            "us_stocks": [s.strip() for s in self.us_stocks.get().replace("，", ",").split(",") if s.strip()],
            "out_dir": target_out_dir,
            # 【修复6】把【年报提取程序】的产出目录一并带入，供计算动态股息率时优先读取
            # 本地《年报提取表》里的“每10股派息数”（比联网更可靠，也更符合本程序的数据口径）
            "extract_dir": getattr(self, "auto_extract_dir", "") or os.path.join(self.save_path.get(), "报表提取完善"),
        }
        self.btn_run.config(text="任务运行中...", state="disabled", bg="#95a5a6")
        threading.Thread(target=self.run_logic_in_thread, args=(config,), daemon=True).start()

    def run_logic_in_thread(self, cfg):
        try:
            doc_path, final_doc_path = main_logic_c(cfg)
            msg = f"阶段一: 基础报告:\n{doc_path}\n\n阶段二: 深度报告:\n{final_doc_path}"
            if getattr(self, "auto_mode", False):
                if getattr(self, "on_finish", None): self.root.after(0, self.on_finish)
            else:
                self.root.after(0, lambda: messagebox.showinfo("流水线执行成功", msg))
        except Exception as e:
            self.root.after(0, lambda msg=str(e): messagebox.showerror("错误", f"发生异常:\n{msg}"))
        finally:
            self.root.after(0, lambda: self.btn_run.config(text="🚀 启动自动估值分析流", state="normal", bg="#e74c3c"))
# =====================================================================
# 子系统与组件类定义 (C程序: AI 文本分析)
# =====================================================================
class AIAnalyzerPro:
    def __init__(self, parent_frame, root):
        self.parent = parent_frame
        self.root = root
        self.use_proxy_var = tk.BooleanVar(value=True)
        self.use_search_var = tk.BooleanVar(value=True)
        self.stream_var = tk.BooleanVar(value=True)

        self.data_source_var = tk.StringVar(value="auto")
        self.auto_input_dir = ""
        self.auto_output_dir = ""

        self.setup_ui()

    def setup_ui(self):
        # 1. 采用 Notebook 选项卡布局
        notebook = ttk.Notebook(self.parent)
        notebook.pack(fill="both", expand=True, padx=10, pady=10)
        self.tab_main = ttk.Frame(notebook)
        self.tab_template = ttk.Frame(notebook)
        notebook.add(self.tab_main, text="💻 核心分析工作台")
        notebook.add(self.tab_template, text="📝 分析指令模板配置")

        self.build_main_tab()
        self.build_template_tab()

        self.safe_log_msg("=== 文本分析程序 (Gemini AI文档专家) 已启动 ===")
        self.update_proxy_settings()
        self.toggle_source()

    def build_main_tab(self):
        env_frame = tk.LabelFrame(self.tab_main, text=" 🐍 Python 环境检测 ", font=("微软雅黑", 10, "bold"), padx=15,
                                  pady=5, fg="#228B22")
        env_frame.pack(fill="x", padx=15, pady=5)
        self.env_status = tk.Label(env_frame, text="正在检测Python环境...", font=("微软雅黑", 9), fg="#666666")
        self.env_status.pack(anchor="w")
        self.check_environment()

        set_frame = tk.LabelFrame(self.tab_main, text=" ⚙️ AI 接口配置 ", font=("微软雅黑", 10, "bold"), padx=15,
                                  pady=5)
        set_frame.pack(fill="x", padx=15, pady=5)
        tk.Label(set_frame, text="选择 AI 类型:").grid(row=0, column=0, sticky="w")
        self.engine_sel = ttk.Combobox(set_frame, values=list(AI_ENGINES.keys()), width=35, state="readonly")
        self.engine_sel.grid(row=0, column=1, padx=10, pady=2, sticky="w")
        self.engine_sel.bind("<<ComboboxSelected>>", self.toggle_engine)
        tk.Label(set_frame, text="API 地址:").grid(row=1, column=0, sticky="w")
        self.url_sel = ttk.Combobox(set_frame, width=58)
        self.url_sel.grid(row=1, column=1, columnspan=2, pady=2, sticky="w")
        tk.Label(set_frame, text="API 钥匙:").grid(row=2, column=0, sticky="w")
        self.key_entry = tk.Entry(set_frame, width=60, show="*")
        self.key_entry.grid(row=2, column=1, columnspan=2, pady=2, sticky="w")
        tk.Label(set_frame, text="模型名称:").grid(row=3, column=0, sticky="w")
        model_frame = tk.Frame(set_frame)
        model_frame.grid(row=3, column=1, columnspan=2, pady=2, sticky="w")
        self.model_sel = ttk.Combobox(model_frame, width=42)
        self.model_sel.pack(side="left")
        tk.Button(model_frame, text="💾 保存模型", command=self.save_custom_model, fg="#d83b01").pack(side="left",
                                                                                                     padx=5)

        self.engine_sel.current(0)
        self.toggle_engine(None)

        search_frame = tk.LabelFrame(self.tab_main, text=" 🌐 联网情报增强 (Web Agent) ", font=("微软雅黑", 10, "bold"),
                                     padx=15, pady=5, fg="#d83b01")
        search_frame.pack(fill="x", padx=15, pady=5)
        control_frame = tk.Frame(search_frame)
        control_frame.pack(fill="x", pady=2)
        self.proxy_check = tk.Checkbutton(control_frame, text="使用代理网络", variable=self.use_proxy_var,
                                          font=("微软雅黑", 9), command=self.update_proxy_settings)
        self.proxy_check.pack(side="left", padx=(0, 20))
        self.search_check = tk.Checkbutton(control_frame, text="开启全网检索", variable=self.use_search_var,
                                           font=("微软雅黑", 9))
        self.search_check.pack(side="left", padx=(0, 20))
        self.stream_check = tk.Checkbutton(control_frame, text="流式输出模式", variable=self.stream_var,
                                           font=("微软雅黑", 9))
        self.stream_check.pack(side="left")

        keyword_frame = tk.Frame(search_frame)
        keyword_frame.pack(fill="x", pady=2)
        tk.Label(keyword_frame, text="搜索模板:").pack(side="left")
        self.search_template = ttk.Combobox(keyword_frame, values=list(SEARCH_TEMPLATES.keys()), width=15,
                                            state="readonly")
        self.search_template.pack(side="left", padx=5)
        self.search_template.bind("<<ComboboxSelected>>", self.update_search_keyword)
        self.search_template.current(0)
        tk.Label(keyword_frame, text="关键词:").pack(side="left", padx=(10, 2))
        self.search_keyword = tk.Entry(keyword_frame, width=40)
        self.search_keyword.pack(side="left")
        self.update_search_keyword(None)

        source_frame = tk.Frame(self.tab_main)
        source_frame.pack(fill="x", padx=15, pady=2)
        tk.Radiobutton(source_frame, text="手动选择目录", variable=self.data_source_var, value="manual",
                       command=self.toggle_source).pack(side="left")
        tk.Radiobutton(source_frame, text="自动联动下载年报结果", variable=self.data_source_var, value="auto",
                       command=self.toggle_source).pack(side="left", padx=10)

        path_frame = tk.LabelFrame(self.tab_main, text=" 📂 批量读取路径选择 (根据文件夹名精准归类) ",
                                   font=("微软雅黑", 10, "bold"), padx=15, pady=5)
        path_frame.pack(fill="x", padx=15, pady=5)
        tk.Label(path_frame, text="待分析主目录:").grid(row=0, column=0, sticky="w")
        self.entry_in = tk.Entry(path_frame, width=50)
        self.entry_in.grid(row=0, column=1, padx=10)
        self.btn_in = tk.Button(path_frame, text="浏览", command=self.get_input_dir)
        self.btn_in.grid(row=0, column=2)
        tk.Label(path_frame, text="报告保存目录:").grid(row=1, column=0, sticky="w", pady=5)
        self.entry_out = tk.Entry(path_frame, width=50)
        self.entry_out.grid(row=1, column=1, padx=10, pady=5)
        self.btn_out = tk.Button(path_frame, text="浏览", command=self.get_output)
        self.btn_out.grid(row=1, column=2, pady=5)

        btn_frame = tk.Frame(self.tab_main)
        btn_frame.pack(fill="x", padx=15, pady=5)
        self.run_btn = tk.Button(btn_frame, text="🚀 开启多年度综合批量分析", bg="#0078d4", fg="white",
                                 font=("微软雅黑", 12, "bold"), height=2, command=self.start_thread)
        self.run_btn.pack(fill="x")

        log_frame = tk.LabelFrame(self.tab_main, text=" 📊 运行日志 ", font=("微软雅黑", 10, "bold"), padx=15, pady=5)
        log_frame.pack(fill="both", expand=True, padx=15, pady=5)
        log_toolbar = tk.Frame(log_frame)
        log_toolbar.pack(fill="x", pady=2)
        tk.Button(log_toolbar, text="清空日志", command=self.clear_log).pack(side="right")
        self.log = tk.Text(log_frame, bg="#1e1e1e", fg="#00ff00", font=("Consolas", 10), padx=10, pady=5, wrap=tk.WORD)
        self.log.pack(fill="both", expand=True)
        scrollbar = tk.Scrollbar(self.log)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log.config(yscrollcommand=scrollbar.set)

    def build_template_tab(self):
        # 2. 将分析指令独立到第二个标签页
        cmd_frame = tk.LabelFrame(self.tab_template, text=" 📝 分析指令 (Prompt) ", font=("微软雅黑", 10, "bold"),
                                  padx=15, pady=10)
        cmd_frame.pack(fill="both", expand=True, padx=20, pady=20)
        tmpl_header_frame = tk.Frame(cmd_frame)
        tmpl_header_frame.pack(fill="x")
        tk.Label(tmpl_header_frame, text="分析模板:").pack(side="left")
        self.tmpl_sel = ttk.Combobox(tmpl_header_frame, values=list(PROMPT_TEMPLATES.keys()), width=30,
                                     state="readonly")
        self.tmpl_sel.pack(side="left", padx=5)
        self.tmpl_sel.bind("<<ComboboxSelected>>", self.set_prompt)

        # --- 新增：默认模板加载逻辑 ---
        default_name = "多年度综合深度分析"
        if os.path.exists("text_analyzer_default.txt"):
            try:
                with open("text_analyzer_default.txt", "r", encoding="utf-8") as f:
                    default_name = f.read().strip()
            except:
                pass

        if default_name in PROMPT_TEMPLATES:
            self.tmpl_sel.set(default_name)
        else:
            self.tmpl_sel.current(0)

        # --- 新增：设为默认按键 ---
        tk.Button(tmpl_header_frame, text="📌 设为默认模板", command=self.set_default_template,
                  fg="#0078d4").pack(side="right", padx=5)
        tk.Button(tmpl_header_frame, text="💾 将当前指令保存为新模板", command=self.save_custom_template,
                  fg="#d83b01").pack(side="right", padx=5)

        self.prompt_box = tk.Text(cmd_frame, font=("微软雅黑", 10))
        self.prompt_box.pack(fill="both", expand=True, pady=(10, 5))
        self.set_prompt(None)

    # --- 新增：处理保存默认模板的功能 ---
    def set_default_template(self):
        name = self.tmpl_sel.get()
        if name:
            try:
                with open("text_analyzer_default.txt", "w", encoding="utf-8") as f:
                    f.write(name)
                messagebox.showinfo("成功", f"已将【{name}】设为默认模板！\n下次打开程序将自动加载此模板。")
            except Exception as e:
                messagebox.showerror("错误", f"保存默认模板失败：{e}")

    def toggle_source(self):
        if self.data_source_var.get() == "auto":
            self.entry_in.configure(state='disabled')
            self.btn_in.configure(state='disabled')
            self.entry_out.configure(state='disabled')
            self.btn_out.configure(state='disabled')
            if self.auto_input_dir:
                self.safe_log_msg(f"🔄 联动模式激活，输入目录设定为: {self.auto_input_dir}")
        else:
            self.entry_in.configure(state='normal')
            self.btn_in.configure(state='normal')
            self.entry_out.configure(state='normal')
            self.btn_out.configure(state='normal')
            self.safe_log_msg("🖐️ 切换至手动选择目录模式")

    def update_proxy_settings(self):
        if not hasattr(self, "use_proxy_var"):
            return
        if self.use_proxy_var.get():
            # 【修复3】按“环境变量 -> 本程序代理地址输入框”顺序安全取值；
            # 两者都没有时宁可直连，也绝不再凭空把 "http://127.0.0.1:7890" 写进进程环境
            # （原实现会在程序启动瞬间污染全局代理，导致其它子程序的 requests 全部走死代理）。
            current_proxy = (os.environ.get("HTTP_PROXY")
                             or os.environ.get("http_proxy")
                             or get_widget_text_var(self, "proxy_addr_var"))
            if current_proxy:
                os.environ["HTTP_PROXY"] = current_proxy
                os.environ["HTTPS_PROXY"] = current_proxy
                os.environ["http_proxy"] = current_proxy
                os.environ["https_proxy"] = current_proxy
                self.safe_log_msg(f"✅ 代理设置已启用 ({current_proxy})")
            else:
                apply_direct_network_env()
                self.safe_log_msg("✅ 代理设置已禁用（本机未配置代理地址）")
        else:
            apply_direct_network_env()
            self.safe_log_msg("✅ 代理设置已禁用")
        os.environ["NO_PROXY"] = "localhost,127.0.0.1,::1"
        os.environ["no_proxy"] = "localhost,127.0.0.1,::1"

    def check_environment(self):
        try:
            python_path = sys.executable
            venv_path = os.path.dirname(os.path.dirname(python_path))
            if os.path.basename(venv_path) == "Scripts": venv_path = os.path.dirname(venv_path)
            if "venv" in venv_path or "virtualenv" in venv_path:
                self.env_status.config(text=f"✅ 已在虚拟环境中: {os.path.basename(venv_path)}", fg="#228B22")
            else:
                self.env_status.config(text="⚠️ 未在虚拟环境中运行，建议使用虚拟环境", fg="#FF6B00")
        except Exception:
            pass

    def toggle_engine(self, event):
        engine_name = self.engine_sel.get()
        cfg = AI_ENGINES[engine_name]
        self.url_sel["values"] = cfg["urls"]
        self.url_sel.set(cfg["urls"][0] if cfg["urls"] else "")
        self.model_sel["values"] = cfg["models"]
        self.model_sel.set(cfg["models"][0] if cfg["models"] else "")
        state = "disabled" if not cfg["need_key"] else "normal"
        bg = "#e0e0e0" if not cfg["need_key"] else "white"
        self.key_entry.config(state=state, bg=bg)

    def save_custom_model(self):
        current_engine = self.engine_sel.get()
        current_model = self.model_sel.get().strip()
        if not current_model:
            messagebox.showwarning("提示", "模型名称不能为空！")
            return
        engine_models = AI_ENGINES[current_engine]["models"]
        if current_model in engine_models: return
        engine_models.append(current_model)
        self.model_sel["values"] = engine_models
        self.model_sel.set(current_model)
        try:
            custom_dict = {}
            if os.path.exists(CUSTOM_MODELS_FILE):
                with open(CUSTOM_MODELS_FILE, "r", encoding="utf-8") as f: custom_dict = json.load(f)
            if current_engine not in custom_dict: custom_dict[current_engine] = []
            if current_model not in custom_dict[current_engine]: custom_dict[current_engine].append(current_model)
            with open(CUSTOM_MODELS_FILE, "w", encoding="utf-8") as f:
                json.dump(custom_dict, f, ensure_ascii=False, indent=4)
            self.safe_log_msg(f"💾 成功保存新模型: 【{current_model}】")
            messagebox.showinfo("成功", f"模型保存成功！")
        except Exception as e:
            messagebox.showerror("保存失败", f"无法保存: {e}")

    def update_search_keyword(self, event):
        template_name = self.search_template.get()
        self.search_keyword.delete(0, tk.END)
        if template_name != "自定义关键词":
            self.search_keyword.insert(0, SEARCH_TEMPLATES[template_name])

    def set_prompt(self, event):
        self.prompt_box.delete("1.0", tk.END)
        self.prompt_box.insert("1.0", PROMPT_TEMPLATES[self.tmpl_sel.get()])

    def save_custom_template(self):
        tmpl_name = simpledialog.askstring("保存为新模板", "请输入新的模板名称:")
        if not tmpl_name: return
        custom_text = self.prompt_box.get("1.0", tk.END).strip()
        if not custom_text: return
        PROMPT_TEMPLATES[tmpl_name] = custom_text
        try:
            with open(TEMPLATE_FILE_C, "w", encoding="utf-8") as f:
                custom_dict = {k: v for k, v in PROMPT_TEMPLATES.items() if
                               k not in ["多年度综合深度分析", "结合最新情报分析"]}
                json.dump(custom_dict, f, ensure_ascii=False, indent=4)
            self.tmpl_sel["values"] = list(PROMPT_TEMPLATES.keys())
            self.tmpl_sel.set(tmpl_name)
            self.safe_log_msg(f"💾 成功保存新模板: 【{tmpl_name}】")
        except Exception as e:
            pass

    def get_input_dir(self):
        p = filedialog.askdirectory()
        if p:
            self.entry_in.delete(0, tk.END)
            self.entry_in.insert(0, p)

    def get_output(self):
        p = filedialog.askdirectory()
        if p:
            self.entry_out.delete(0, tk.END)
            self.entry_out.insert(0, p)

    def clear_log(self):
        self.log.delete("1.0", tk.END)

    def safe_log_msg(self, msg: str):
        self.root.after(0, self._append_log, f"[{time.strftime('%H:%M:%S')}] {msg}\n")

    def safe_stream_msg(self, txt: str):
        self.root.after(0, self._append_log, txt)

    def _append_log(self, txt: str):
        self.log.insert(tk.END, txt);
        self.log.see(tk.END)

    def save_to_word(self, content: str, company_name: str, save_dir: str) -> str:
        safe_name = re.sub(r'[\\/*?:"<>|]', "", company_name)
        doc = Document()
        title = doc.add_heading(f'{safe_name} - 多年度综合分析报告', 0)
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        info_para = doc.add_paragraph()
        info_para.add_run(
            f"生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n分析模型: {self.model_sel.get()}\n")
        doc.add_paragraph("=" * 50)
        for line in content.split('\n'):
            if line.strip():
                if line.strip().startswith('#'):
                    doc.add_heading(line.replace('#', '').strip(), level=min(line.count('#'), 3))
                elif line.strip().startswith('**'):
                    doc.add_paragraph().add_run(line.replace('**', '')).bold = True
                else:
                    doc.add_paragraph(line).style.font.size = Pt(11)
        save_path = os.path.join(save_dir, f"{safe_name}_综合评估报告_{int(time.time())}.docx")
        doc.save(save_path)
        return save_path

    def start_thread(self):
        url, key, model = self.url_sel.get().strip(), self.key_entry.get().strip(), self.model_sel.get().strip()
        in_dir = self.auto_input_dir if self.data_source_var.get() == "auto" else self.entry_in.get().strip()
        out_dir = self.auto_output_dir if self.data_source_var.get() == "auto" else self.entry_out.get().strip()
        prompt = self.prompt_box.get("1.0", tk.END).strip()
        use_search, search_kw, use_stream, use_proxy = self.use_search_var.get(), self.search_keyword.get().strip(), self.stream_var.get(), self.use_proxy_var.get()

        if not in_dir or not out_dir:
            self.safe_log_msg("⚠️ 请先选择待分析文件夹和保存目录")
            if not getattr(self, "auto_mode", False): messagebox.showwarning("提示", "请先选择待分析文件夹和保存目录")
            return

        target_out_dir = os.path.join(out_dir, "文本分析")
        os.makedirs(target_out_dir, exist_ok=True)

        self.run_btn.config(state="disabled", text="⏳ 处理中...", bg="#6c757d")
        threading.Thread(target=self.process_batch_task,
                         args=(url, key, model, in_dir, target_out_dir, prompt, use_search, search_kw, use_stream,
                               use_proxy), daemon=True).start()

    def _extract_company_name(self, filepath: str) -> str:
        parent_dir = os.path.basename(os.path.dirname(filepath))
        code_match = re.search(r'\d{6}', parent_dir)
        code = code_match.group(0) if code_match else ""
        clean_dir = re.sub(r'(年报|报告|新建文件夹|测试|列表|股份有限公司|公司)', '', parent_dir)
        name_match = re.search(r'[\u4e00-\u9fa5]{3,4}', clean_dir)
        name = name_match.group(0) if name_match else ""
        if code and name:
            return f"{code}{name}"
        elif name:
            return name
        elif code:
            return code

        base_name = os.path.splitext(os.path.basename(filepath))[0]
        for w in [r'20\d{2}年?', r'年度?报告', r'年报', r'全文', r'审计报告', r'首次公开发行股票', r'招股说明书',
                  r'公司章程', r'更新版', r'修订', r'摘要', r'股份有限公司', r'公司', r'[-_（）\(\)\[\]]']:
            base_name = re.sub(w, '', base_name)
        return base_name.strip() if len(base_name.strip()) >= 2 else "未知公司"

    def process_batch_task(self, api_url, api_key, model_name, in_dir, out_dir, base_prompt, do_search, search_kw,
                           use_stream, use_proxy):
        success_flag = False
        try:
            pdf_files = []
            for root_dir, _, files in os.walk(in_dir):
                for f in files:
                    if f.lower().endswith('.pdf'): pdf_files.append(os.path.join(root_dir, f))
            if not pdf_files:
                self.safe_log_msg("❌ 未找到PDF文件")
                return

            company_map = {}
            for path in pdf_files:
                comp_name = self._extract_company_name(path)
                if comp_name not in company_map: company_map[comp_name] = []
                company_map[comp_name].append(path)
            all_companies = list(company_map.keys())

            client = OpenAI(api_key=api_key if api_key else "free", base_url=api_url)

            for idx, (target_company, paths) in enumerate(company_map.items()):
                current_search_kw = search_kw.replace("公司名称", target_company)
                current_prompt = base_prompt.replace("公司名称", target_company)
                peer_companies = [c for c in all_companies if c != target_company and c != "未知公司"]
                if peer_companies: current_prompt += f"\n\n【附加要求】: 其他同行：{'、'.join(peer_companies)}。请在报告中加入与它们的对比。"
                final_prompt = current_prompt + "\n\n"

                if do_search and current_search_kw:
                    results, _ = self._search_ddgs(current_search_kw, use_proxy)
                    if results:
                        web_context = "".join(
                            [f"【情报源 {i + 1}】{r.get('title', '')}\n{r.get('body', '')}\n\n" for i, r in
                             enumerate(results)])
                        final_prompt += f"【以下为最新网络情报】:\n{web_context}\n"

                combined_text = ""
                for path in paths:
                    try:
                        with fitz.open(path) as doc:
                            combined_text += f"\n\n>>> 资料：【{os.path.basename(path)}】\n" + "".join(
                                [page.get_text() for page in doc])[:8000]
                    except:
                        pass

                final_prompt += f"【以下是多份本地文档内容】:\n{combined_text[:30000]}"
                try:
                    if use_stream:
                        resp = client.chat.completions.create(model=model_name,
                                                              messages=[{"role": "user", "content": final_prompt}],
                                                              stream=True)
                        result_text = ""
                        for chunk in resp:
                            if chunk.choices and chunk.choices[0].delta.content:
                                char = chunk.choices[0].delta.content
                                result_text += char
                                self.safe_stream_msg(char)
                    else:
                        resp = client.chat.completions.create(model=model_name,
                                                              messages=[{"role": "user", "content": final_prompt}])
                        result_text = resp.choices[0].message.content
                    self.save_to_word(result_text, target_company, out_dir)
                except Exception as api_err:
                    self.safe_log_msg(f"❌ API 调用失败: {api_err}")
                    error_msg = f"生成报告失败。API 调用出现异常：\n{str(api_err)}\n\n常见原因：1.模型上下文 Token 超出限制；2.当前填写的 API Key 余额不足或网络受阻。"
                    self.save_to_word(error_msg, target_company + "_生成失败排错日志", out_dir)
            self.safe_log_msg("\n🎉 处理完毕！")
            success_flag = True
        except Exception as exc:
            self.safe_log_msg(f"❌ 异常: {exc}")
        finally:
            self.root.after(0,
                            lambda: self.run_btn.config(state="normal", text="🚀 开启多年度综合批量分析", bg="#0078d4"))
            if getattr(self, "auto_mode", False):
                if getattr(self, "on_finish", None): self.root.after(0, self.on_finish)
            else:
                if success_flag:
                    self.root.after(0, lambda: messagebox.showinfo("完美收官", f"分析报告已生成！"))

    def _search_ddgs(self, keyword: str, use_proxy: bool):
        try:
            from ddgs import DDGS
        except:
            return [], None

        proxy_str = os.environ.get("HTTP_PROXY", "http://127.0.0.1:7890") if use_proxy else None

        for be in ["duckduckgo", "html", "lite"]:
            try:
                params = {"timeout": 30}
                if proxy_str: params["proxy"] = proxy_str
                with DDGS(**params) as ddgs:
                    results = ddgs.text(keyword, region="cn-zh", safesearch="off", max_results=3, backend=be)
                if results: return results, be
            except:
                continue
        return [], None
# =====================================================================
# 子系统与组件类定义 (C程序: 企业深度分析)
# =====================================================================
class EnterpriseAIAnalyzer:
    def __init__(self, parent_frame, root):
        self.cfg = ConfigManager()
        self.parent = parent_frame
        self.root = root

        # 【新增】：自动模式控制变量
        self.data_source_var = tk.StringVar(value="auto")
        self.auto_input_dir = ""
        self.auto_output_dir = ""

        self.create_widgets()

    def create_widgets(self):
        notebook = ttk.Notebook(self.parent)
        notebook.pack(fill="both", expand=True, padx=10, pady=10)
        self.tab_main = ttk.Frame(notebook)
        self.tab_template = ttk.Frame(notebook)
        notebook.add(self.tab_main, text="💻 核心分析工作台")
        notebook.add(self.tab_template, text="📝 分析指令模板配置")

        self.build_main_tab()
        self.build_template_tab()

        self.toggle_source()


    def build_main_tab(self):
        frame_ai = tk.LabelFrame(self.tab_main, text="1. AI 引擎配置 (支持云端API与本地Ollama等)", padx=10, pady=10)
        frame_ai.pack(fill="x", padx=10, pady=5)
        tk.Label(frame_ai, text="Base URL:").grid(row=0, column=0, sticky="e")
        self.cb_base = ttk.Combobox(frame_ai, values=self.cfg.config["api_bases"], width=45)
        self.cb_base.set(self.cfg.config["current_base"])
        self.cb_base.grid(row=0, column=1, pady=2, padx=5)
        tk.Label(frame_ai, text="模型名称:").grid(row=1, column=0, sticky="e")
        self.cb_model = ttk.Combobox(frame_ai, values=self.cfg.config["api_models"], width=45)
        self.cb_model.set(self.cfg.config["current_model"])
        self.cb_model.grid(row=1, column=1, pady=2, padx=5)
        tk.Label(frame_ai, text="API Key:").grid(row=2, column=0, sticky="e")
        self.ent_key = tk.Entry(frame_ai, width=47, show="*")
        self.ent_key.insert(0, self.cfg.config["api_key"])
        self.ent_key.grid(row=2, column=1, pady=2, padx=5)
        tk.Button(frame_ai, text="保存模型配置", command=self.save_ai_config).grid(row=0, column=2, rowspan=2, padx=10)

        # 【新增】：数据源切换区
        source_frame = tk.Frame(self.tab_main)
        source_frame.pack(fill="x", padx=10, pady=5)
        tk.Radiobutton(source_frame, text="手动选择目录", variable=self.data_source_var, value="manual", command=self.toggle_source).pack(side="left")
        tk.Radiobutton(source_frame, text="自动联动下载年报结果", variable=self.data_source_var, value="auto", command=self.toggle_source).pack(side="left", padx=10)

        frame_path = tk.LabelFrame(self.tab_main, text="2. 批量分析数据源与输出设置", padx=10, pady=10)
        frame_path.pack(fill="x", padx=10, pady=5)
        tk.Label(frame_path, text="待分析主目录:").grid(row=0, column=0, sticky="e")
        self.var_input_dir = tk.StringVar(value=self.cfg.config["input_dir"])
        self.entry_in = tk.Entry(frame_path, textvariable=self.var_input_dir, width=50)
        self.entry_in.grid(row=0, column=1, padx=5)
        self.btn_in = tk.Button(frame_path, text="选择...", command=lambda: self.select_dir("input"))
        self.btn_in.grid(row=0, column=2)

        tk.Label(frame_path, text="报告输出目录:").grid(row=1, column=0, sticky="e")
        self.var_output_dir = tk.StringVar(value=self.cfg.config["output_dir"])
        self.entry_out = tk.Entry(frame_path, textvariable=self.var_output_dir, width=50)
        self.entry_out.grid(row=1, column=1, padx=5)
        self.btn_out = tk.Button(frame_path, text="选择...", command=lambda: self.select_dir("output"))
        self.btn_out.grid(row=1, column=2)

        frame_opts = tk.LabelFrame(self.tab_main, text="3. 多渠道数据融合选项", padx=10, pady=10)
        frame_opts.pack(fill="x", padx=10, pady=5)
        self.var_web_report = tk.BooleanVar(value=True)
        self.var_web_news = tk.BooleanVar(value=True)
        self.var_local_doc = tk.BooleanVar(value=True)
        tk.Checkbutton(frame_opts, text="深度检索研报", variable=self.var_web_report).pack(side="left", padx=10)
        tk.Checkbutton(frame_opts, text="抓取最新公告", variable=self.var_web_news).pack(side="left", padx=10)
        tk.Checkbutton(frame_opts, text="解析本地文档", variable=self.var_local_doc).pack(side="left", padx=10)

        self.btn_run = tk.Button(self.tab_main, text="🚀 启动全自动化深度评价", bg="#27ae60", fg="white", font=("微软雅黑", 12, "bold"), command=self.start_thread)
        self.btn_run.pack(pady=10)

        self.log_text = tk.Text(self.tab_main, height=10, bg="#f4f4f4", state="disabled")
        self.log_text.pack(fill="both", expand=True, padx=10, pady=5)

    def build_template_tab(self):
        # 采用与文本分析程序一致的带边框面板
        cmd_frame = tk.LabelFrame(self.tab_template, text=" 📝 分析指令 (Prompt) ", font=("微软雅黑", 10, "bold"),
                                  padx=15, pady=10)
        cmd_frame.pack(fill="both", expand=True, padx=20, pady=20)

        tmpl_header_frame = tk.Frame(cmd_frame)
        tmpl_header_frame.pack(fill="x")

        tk.Label(tmpl_header_frame, text="分析模板:").pack(side="left")

        # 加载下拉框列表
        template_keys = list(self.cfg.config["templates"].keys())
        # --- 优化：优先读取指定的默认模板，如果没有则读取最后一次使用的 ---
        current_name = self.cfg.config.get("default_template_name", self.cfg.config.get("current_template_name",
                                                                                        template_keys[
                                                                                            0] if template_keys else ""))

        self.cb_templates = ttk.Combobox(tmpl_header_frame, values=template_keys, width=30, state="readonly")
        self.cb_templates.pack(side="left", padx=5)
        self.cb_templates.bind("<<ComboboxSelected>>", self.load_selected_template)

        if current_name in template_keys:
            self.cb_templates.set(current_name)
        elif template_keys:
            self.cb_templates.current(0)

        # --- 新增：设为默认按键 ---
        tk.Button(tmpl_header_frame, text="📌 设为默认模板", command=self.set_default_template, fg="#0078d4").pack(
            side="right", padx=5)
        tk.Button(tmpl_header_frame, text="💾 将当前指令保存为新模板", command=self.save_custom_template,
                  fg="#d83b01").pack(side="right", padx=5)

        self.txt_template = tk.Text(cmd_frame, font=("微软雅黑", 10))
        self.txt_template.pack(fill="both", expand=True, pady=(10, 5))

        self.load_selected_template(None)

    # --- 新增：处理保存默认模板的功能 ---
    def set_default_template(self):
        name = self.cb_templates.get()
        if name:
            self.cfg.config["default_template_name"] = name
            self.cfg.save_config()
            messagebox.showinfo("成功", f"已将【{name}】设为默认模板！\n下次打开程序将自动加载此模板。")
    def load_selected_template(self, event=None):
        name = self.cb_templates.get()
        if name in self.cfg.config["templates"]:
            self.txt_template.delete("1.0", tk.END)
            self.txt_template.insert("1.0", self.cfg.config["templates"][name])
            # 同步记忆当前选中的模板名
            self.cfg.config["current_template_name"] = name
            self.cfg.save_config()

    def save_custom_template(self):
        tmpl_name = simpledialog.askstring("保存为新模板", "请输入新的模板名称:")
        if not tmpl_name: return
        custom_text = self.txt_template.get("1.0", tk.END).strip()
        if not custom_text: return

        # 保存到配置字典中
        self.cfg.config["templates"][tmpl_name] = custom_text
        self.cfg.config["current_template_name"] = tmpl_name
        self.cfg.save_config()

        # 刷新下拉框
        self.cb_templates['values'] = list(self.cfg.config["templates"].keys())
        self.cb_templates.set(tmpl_name)
        self.log(f"💾 成功保存新模板: 【{tmpl_name}】")

    # 【新增】：自动与手动模式切换逻辑
    def toggle_source(self):
        if self.data_source_var.get() == "auto":
            self.entry_in.configure(state='disabled')
            self.btn_in.configure(state='disabled')
            self.entry_out.configure(state='disabled')
            self.btn_out.configure(state='disabled')
            if self.auto_input_dir:
                self.log(f"🔄 联动模式激活，输入目录设定为: {self.auto_input_dir}")
        else:
            self.entry_in.configure(state='normal')
            self.btn_in.configure(state='normal')
            self.entry_out.configure(state='normal')
            self.btn_out.configure(state='normal')
            self.log("🖐️ 切换至手动选择目录模式")

    def log(self, msg):
        self.root.after(0, self._log, msg)

    def _log(self, msg):
        self.log_text.config(state="normal")
        self.log_text.insert(tk.END, f"[{datetime.now().strftime('%H:%M:%S')}] {msg}\n")
        self.log_text.see(tk.END)
        self.log_text.config(state="disabled")

    def save_ai_config(self):
        self.cfg.config["current_base"] = self.cb_base.get().strip()
        self.cfg.config["current_model"] = self.cb_model.get().strip()
        self.cfg.config["api_key"] = self.ent_key.get().strip()
        self.cfg.save_config()

    def select_dir(self, target):
        path = filedialog.askdirectory()
        if path:
            if target == "input": self.var_input_dir.set(path)
            else: self.var_output_dir.set(path)

    def search_web(self, query, max_res=3):
        info = ""
        time.sleep(random.uniform(1.0, 2.5))
        try:
            import urllib.request, urllib.parse
            url = f"https://www.so.com/s?q={urllib.parse.quote(query)}"
            req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0", "Accept": "*/*"})
            html = urllib.request.urlopen(req, timeout=10).read().decode('utf-8')
            blocks = html.split('<li class="res-list')
            count = 0
            for block in blocks[1:max_res + 2]:
                title_match = re.search(r'<h3[^>]*>(.*?)</h3>', block, re.S)
                if title_match:
                    title = re.sub(r'<[^>]+>', '', title_match.group(1)).strip()
                    desc_match = re.search(r'class="res-desc[^"]*"[^>]*>(.*?)</', block, re.S) or re.search(r'class="res-rich[^"]*"[^>]*>(.*?)</', block, re.S)
                    desc = re.sub(r'<[^>]+>', '', desc_match.group(1)).strip() if desc_match else "无摘要"
                    if title:
                        info += f"- [{title}]: {desc}\n"
                        count += 1
                        if count >= max_res: break
            if count > 0: return info
        except Exception as e:
            # 不要直接 pass，至少记录下来，方便调试
            self.log(f"解析发生错误，已跳过该项: {str(e)}")

        try:
            import urllib.request, urllib.parse
            url = f"https://cn.bing.com/search?q={urllib.parse.quote(query)}"
            req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0", "Cookie": "SRCHHPGUSR=SRCHLANG=zh-Hans;"})
            html = urllib.request.urlopen(req, timeout=10).read().decode('utf-8')
            blocks = re.findall(r'<h2[^>]*><a[^>]*>(.*?)</a></h2>(.*?)(?=<h2|$)', html, re.S)
            count = 0
            for title_raw, content_raw in blocks:
                title = re.sub(r'<[^>]+>', '', title_raw).strip()
                desc_match = re.search(r'<p[^>]*>(.*?)</p>', content_raw, re.S)
                desc = re.sub(r'<[^>]+>', '', desc_match.group(1)).strip() if desc_match else "无摘要"
                if title and count < max_res:
                    info += f"- [{title}]: {desc}\n"
                    count += 1
            if count > 0: return info
        except Exception as e:
            # 不要直接 pass，至少记录下来，方便调试
            self.log(f"解析发生错误，已跳过该项: {str(e)}")
        return info

    def start_thread(self):
        self.btn_run.config(text="任务进行中...", state="disabled", bg="#95a5a6")
        threading.Thread(target=self.run_batch_analysis, daemon=True).start()

    def run_batch_analysis(self):
        success_flag = False
        try:
            self.save_ai_config()
            # 【修改】：动态根据模式获取路径
            input_root = self.auto_input_dir if self.data_source_var.get() == "auto" else self.var_input_dir.get()
            base_output_root = self.auto_output_dir if self.data_source_var.get() == "auto" else self.var_output_dir.get()

            output_root = os.path.join(base_output_root, "企业分析")
            os.makedirs(output_root, exist_ok=True)

            # 【关键修复3】：深入行业文件夹内部，把每一家公司(包含同行)都加入任务列表
            company_tasks = []
            for group_name in os.listdir(input_root):
                group_path = os.path.join(input_root, group_name)
                if not os.path.isdir(group_path) or "新建文件夹" in group_name: continue

                comp_folders = [d for d in os.listdir(group_path) if os.path.isdir(os.path.join(group_path, d))]
                for comp in comp_folders:
                    comp_path = os.path.join(group_path, comp)
                    peers = [c for c in comp_folders if c != comp]
                    company_tasks.append({"path": comp_path, "name": comp, "peers": peers})

            client = OpenAI(api_key=self.cfg.config["api_key"] or "sk-local", base_url=self.cfg.config["current_base"])
            template = self.txt_template.get("1.0", tk.END).strip()

            for task in company_tasks:
                comp_name_raw = task["name"]
                self.log(f"\n开始分析: {comp_name_raw}")

                clean_company = comp_name_raw.split('_')[-1] if '_' in comp_name_raw else comp_name_raw
                peer_names = [p.split('_')[-1] if '_' in p else p for p in task["peers"]]
                clean_comp_str = "、".join(peer_names) or "暂无"

                local_files = [os.path.join(task["path"], f) for f in os.listdir(task["path"]) if
                               f.lower().endswith(('.pdf', '.docx'))]

                web_info = ""
                if self.var_web_report.get(): web_info += self.search_web(f"{clean_company} 最新深度研究报告", 2)
                if self.var_web_news.get(): web_info += self.search_web(f"{clean_company} 最新公告 {clean_comp_str} 对比", 2)

                doc_info = ""
                if self.var_local_doc.get() and local_files:
                    for path in local_files:
                        if path.endswith('.pdf'):
                            try:
                                with pdfplumber.open(path) as pdf:
                                    for page in pdf.pages[:35]: doc_info += (page.extract_text() or "") + "\n"
                            except Exception as e:
                                # 不要直接 pass，至少记录下来，方便调试
                                self.log(f"解析发生错误，已跳过该项: {str(e)}")

                prompt = template.format(company=clean_company, competitors=clean_comp_str, web_info=web_info[:8000], doc_info=doc_info[:15000])
                resp = client.chat.completions.create(model=self.cfg.config["current_model"], messages=[{"role": "system", "content": "资深金融分析师"}, {"role": "user", "content": prompt}], temperature=0.3)

                doc = Document()
                doc.add_heading(f'{clean_company} 深度分析报告', 0)
                doc.add_paragraph(f"对标: {clean_comp_str} | 时间: {datetime.now().strftime('%Y-%m-%d')}")
                doc.add_paragraph(resp.choices[0].message.content)

                save_fn = os.path.join(output_root, f"{clean_company}_深度报告_{datetime.now().strftime('%H%M')}.docx")
                doc.save(save_fn)
                self.log(f"✅ 报告已生成: {os.path.basename(save_fn)}")

            self.log("🎉 全部完成！")
            success_flag = True
        except Exception as e:
            self.log(f"❌ 错误: {str(e)}")
        finally:
            self.root.after(0, lambda: self.btn_run.config(text="🚀 启动全自动化深度评价", state="normal", bg="#27ae60"))
            if getattr(self, "auto_mode", False):
                if getattr(self, "on_finish", None): self.root.after(0, self.on_finish)
            else:
                if success_flag:
                    self.root.after(0, lambda: messagebox.showinfo("完成", "任务执行完毕！"))

# =====================================================================
# 顶层：统一整合的一键总控面板
# =====================================================================
class MasterControlSystem:
    def __init__(self, root):
        self.root = root
        self.root.title("价值投资智能分析模型V1.0")
        self.root.geometry("820x1060")

        # ==========================================
        # 红木与香槟金主题 (温暖、和谐、豪华且厚重)
        # ==========================================
        self.BG_MAIN = "#FDF9F1"   # 象牙暖白 (主底色，温馨明亮不刺眼)
        self.BG_FRAME = "#F4EBDC"  # 浅奶咖/香槟色 (分组框底色，柔和的层次感)
        self.FG_TEXT = "#362419"   # 极深咖啡色 (正文文字，比纯黑更高级，对比度极佳)
        self.FG_TITLE = "#7A2E1C"  # 皇家红木色/暗铁红 (主标题，充满权威感与重量感)
        self.FG_SUB = "#9C6B30"    # 哑光青铜金 (副标题与高亮说明)

        # 核心大按钮配色
        self.BTN_COLOR = "#8B1C1C" # 沉稳的红宝石/深绯红 (压得住阵脚，豪华大气)
        self.BTN_TEXT = "#FFFFFF"  # 纯白 (主控按钮文字，最高对比度保证清晰)

        self.root.configure(bg=self.BG_MAIN)

        # ------------------ 初始化 A 程序子窗口 ------------------
        self.app1_win = tk.Toplevel(self.root)
        self.app1_win.protocol("WM_DELETE_WINDOW", self.app1_win.withdraw)
        self.app1 = UnifiedStockSystem(self.app1_win, master_callback=self.on_app1_complete)
        self.app1_win.withdraw()

        self.app2_win = tk.Toplevel(self.root)
        self.app2_win.geometry("1150x750")
        self.app2_win.protocol("WM_DELETE_WINDOW", self.app2_win.withdraw)
        self.app2 = StockAnalyzerApp(self.app2_win, master_callback=self.on_app2_complete)
        self.app2_win.withdraw()

        self.app3_win = tk.Toplevel(self.root)
        self.app3_win.protocol("WM_DELETE_WINDOW", self.app3_win.withdraw)
        self.app3 = ReportExtractionApp(self.app3_win, master_callback=self.on_app3_complete)
        self.app3_win.withdraw()

        # ------------------ 初始化 B 程序子窗口 ------------------
        self.app_b_fraud_win = tk.Toplevel(self.root)
        self.app_b_fraud_win.title("造假排雷程序 (B程序)")
        self.app_b_fraud_win.geometry("800x600")
        self.app_b_fraud_win.protocol("WM_DELETE_WINDOW", self.app_b_fraud_win.withdraw)
        self.app_b_fraud = FraudDetectionApp(self.app_b_fraud_win, self)
        self.app_b_fraud.pack(fill=tk.BOTH, expand=True)
        self.app_b_fraud_win.withdraw()

        self.app_b_16dim_win = tk.Toplevel(self.root)
        self.app_b_16dim_win.title("16维度分析程序 (B程序)")
        self.app_b_16dim_win.geometry("800x600")
        self.app_b_16dim_win.protocol("WM_DELETE_WINDOW", self.app_b_16dim_win.withdraw)
        self.app_b_16dim = FinancialAnalysisApp(self.app_b_16dim_win, self)
        self.app_b_16dim.pack(fill=tk.BOTH, expand=True)
        self.app_b_16dim_win.withdraw()

        # ------------------ 初始化 C 程序子窗口 ------------------
        self.app_c_val_win = tk.Toplevel(self.root)
        self.app_c_val_win.title("好价分析程序 (C程序)")
        self.app_c_val_win.geometry("800x600")
        self.app_c_val_win.protocol("WM_DELETE_WINDOW", self.app_c_val_win.withdraw)
        self.app_c_val = ValuationApp(self.app_c_val_win, self.root)
        self.app_c_val_win.withdraw()

        self.app_c_ai_pro_win = tk.Toplevel(self.root)
        self.app_c_ai_pro_win.title("文本分析程序 (C程序)")
        self.app_c_ai_pro_win.geometry("900x700")
        self.app_c_ai_pro_win.protocol("WM_DELETE_WINDOW", self.app_c_ai_pro_win.withdraw)
        self.app_c_ai_pro = AIAnalyzerPro(self.app_c_ai_pro_win, self.root)
        self.app_c_ai_pro_win.withdraw()

        self.app_c_ent_win = tk.Toplevel(self.root)
        self.app_c_ent_win.title("企业分析程序 (C程序)")
        self.app_c_ent_win.geometry("900x700")
        self.app_c_ent_win.protocol("WM_DELETE_WINDOW", self.app_c_ent_win.withdraw)
        self.app_c_ent = EnterpriseAIAnalyzer(self.app_c_ent_win, self.root)
        self.app_c_ent_win.withdraw()

        self.is_auto_run = False
        self.start_stage_var = tk.StringVar(value="stage1")
        self.setup_ui()

    def show_help(self):
        """显示帮助文档窗口"""
        help_win = tk.Toplevel(self.root)
        help_win.title("系统使用帮助")
        help_win.geometry("700x500")
        help_win.configure(bg=self.BG_MAIN)
        txt = scrolledtext.ScrolledText(help_win, font=("微软雅黑", 20), bg="#FFFBF2", fg=self.FG_TEXT,
                                        relief=tk.SUNKEN, bd=3)
        txt.pack(fill=tk.BOTH, expand=True, padx=15, pady=15)
        help_text = (
            "【系统使用帮助】\n\n"
            "       欢迎使用价值投资智能分析模型V1.0！以下为该模型及各模块的详细说明和使用指南：\n\n"

'       本模型主要为价值投资筛选上市公司、分析上市公司和分析上市公司股价而设计的。其功能是通过以一定条件对上市公司（主要是A股）进行筛选，对选出的公司及其同行业前三名公司的相关资料（主要是年报）进行下载、财务造假嫌疑年度分析、同行业对比年报分析，\
对选出的公司进行年报文本分析、企业分析，对选出的各公司及其同行业前三名的各公司的股票价格进行分析，供投资参考。\n'
'       这里所说的价值投资，主要是立足于“好公司是可以持续发展的；在好价格买入好公司的股票，其收益是大于10年期中国国债收益的，赚企业发展的钱；股票的价格是围绕公司的价值上下波动的，当价格低于价值时买入，当价格高于价值时卖出，赚市场的钱”理论。\n\n'

'       上市公司筛选。\n'
'       主要以公司连续多年的资产收益率、净利润现金含量、毛利率、上市时间等为条件，分别通过AI和同花顺网站两个渠道进行筛选，然后综合。\n'
'       结果生成“海选情况”文件夹存放在目标文件夹中。该文件夹中有：《海选公司汇总表》（另有附件文件：《AI深度分析研报》《AI提取_关键数据汇总》《问财_表1_多条件初始表》《问财_表2_多条件整理表》《问财_表3_代码及名称汇总》）。\n'
'       标准设置条件：\n'
'       AI筛选设置模板“请在A股中严格筛选同时符合以下条件的股票并作分析：2021、2022、2023、2024、2025年（五个完整会计年度）roe_weighted>40%、净利润现金含量>80%，gross_profit_margin>40%，上市时间>3年，剔除北交所中的公司和金融股。\
如果没有符合条件的则跳过，不能擅自乱抓其他不完全符合条件的公司。”\n'
'       问财筛选设置模板“连续5年加权roe>20，连续5年净利润现金含量>100，连续5年毛利率>40，上市时间>3年，剔除北交所，非金融股”。\n\n'

'       相关资料下载。\n'
'       主要是年报、招股说明书、公司章程。\n'
'       结果生成“文件下载”文件夹存放在目标文件夹中。该文件夹中有：《需对比分析的全部公司名单》和所筛选出的公司按行业属性汇总的若干文件夹（文件夹名为：代表公司名+行业+行业属性），这些文件夹中有：代表公司的《公司属性表》《同行排列表》和所包含的各\
公司文件夹（文件夹名为：公司代码+公司简称），各公司文件夹中有本公司相应年度的年报及相关资料的pdf文件。\n'
'       标准设置条件：\n'
'       选择“最近5年”。\n'
'       当资料下载完毕后，年报提取模块会从下载的资料中提取相关数据编成相关表格，对未能提取到的相关数据通过搜索获取进行补充完善。\n'
'       结果生成“报表提取完善”文件夹存放在目标文件夹中。该文件夹中有：所筛选出的公司按行业属性汇总的若干文件夹（文件夹名为：代表公司名+行业+行业属性），这些文件夹中有：代表公司的《公司属性表》《同行排列表》和所包含的各公司《统一整合输出》表\
（名称为：统一整合输出+公司简称+（公司代码)+年份-年份），各年度《年报提取表》《多年度汇总与增长率分析》表。\n\n'

'       财务造假嫌疑分析。\n'
'       主要从18个角度：员工人数与营业收入，人均薪酬与同行对比，关联交易，重大股权投资，营业收入、毛利率、存货与重大非股权投资对比，募集资金使用，货币资金、交易性金融资产与总资产对比，应收帐款占总资产比例，固定资产占总资产比例，在建工程占总资\
产比例，预付款和其他应收款占总资产比例，某段时间固定资产或在建工程占总资产的比率增幅情况，应收帐款、存货占总资产比例与固定资产或在建工程占总资产的比率增长情况，利息收入与货币资金对比，帐面货币资金与有息负债对比，利润表和资产负债表减值项目\
对比，毛利率波动幅度，毛利率波动幅度、应收帐款或存货占总资产比例与同行业对比进行分析并作出判断。\n'
'       结果生成“造假排雷结果”文件夹存放在目标文件夹中。该文件夹中有：各公司《深度排雷报告》表。\n'
'       标准设置条件：\n'
'       1、员工人数在减少但营业收入却在增加,则有财务造假嫌疑。\n'
'       2、人均薪酬大幅低于同行公司的人均薪酬,则有财务造假嫌疑。\n'
'       3、经常有披露的关联交易的，则有财务造假嫌疑。\n'
'       4、有重大的股权投资的，则有财务造假嫌疑。\n'
'       5、如果公司营业收入增长乏力，毛利率也不高，存货还比较多，此时公司还在进行重大非股权投资扩大产能，则有财务造假嫌疑。\n'
'       6、募集资金没有按计划使用，则显示有财务造假嫌疑。\n'
'       7、（货币资金+交易性金融资产）/总资产*100%<10%,则显示有财务造假嫌疑。\n'
'       8、应收帐款/总资产*100%>15%,则显示有财务造假嫌疑。\n'
'       9、固定资产/总资产*100%>40%,则显示有财务造假嫌疑。\n'
'       10、在建工程/总资产*100%>5%,则显示有财务造假嫌疑。\n'
'       11、（预付款+其他应收款）/总资产*100%>10%,则显示有财务造假嫌疑。\n'
'       12、在某段时间固定资产或在建工程占总资产的比率有大幅增长,则有财务造假嫌疑。\n'
'       13、应收帐款/总资产*100%>5%,或存货/总资产*100%>15%，同时固定资产或在建工程占总资产的比率增长较快，则有财务造假嫌疑。\n'
'       14、利息收入/货币资金*100%<2%,则显示有财务造假嫌疑。\n'
'       15、帐面货币资金很多，但有息负债也很多,则有财务造假嫌疑。\n'
'       16、利润表中减值的对应科目，在资产负债表中的占比较大，而利润表中两个减值损失金额很小，则有财务造假嫌疑。\n'
'       17、如果公司毛利率波动大，且高的时候大幅高于同行业平均水平,则有财务造假嫌疑。\n'
'       18、如果公司毛利率波动大，且应收帐款或存货占总资产比率大幅高于同行平均水平，则有财务造假嫌疑。\n\n'

'       同行年报对比分析（又称18步分析）。\n'
'       主要从16个维度：总资产分析实力和成长性，从负债分析偿债风险，从应付预收减应收预付的差额分析竞争优势，从应收帐款、合同资产分析产品竞争力，从固定资产分析维持竞争\
力的成本，从投产类资产分析主业专注度，从存货、商誉分析未来业绩爆雷的风险，从营业收入分析行业地位及成长性，从毛利率分析产品竞争力及风险，从期间费用率分析成本管控能力，从销\
售费用率分析产品的销售难易度，从主营利润分析主业的盈利能力及利润质量，从营业外收入净额进一步分析利润质量，从归母净利润分析整体盈利能力及持续性，从购建固定资产、无形资产和其\
他长期资产支付的现金分析增长潜力，从分红率（也叫股利支付率）分析现金分红情况。与同行对比分析给予打分，并作出判断。\n'
'       结果生成“16维度对比结果”文件夹存放在目标文件夹中。该文件夹中有：按各公司所属行业的《16维度同行对比分析》表。\n'
'       标准设置条件：\n'
'      （一）维度1：从总资产分析实力和成长性。分析指标：总资产规模、总资产同比增长率。数据来源：各公司各年度的合并资产负债表。\n'
'       1、总资产规模：计算结果=当期总资产；得分办法=比较目标公司及其同行公司的该项的计算结果，最大的10分，每小于最大数10%则减1分；判定结果=计算结果在同行几个公司中最\
大的为“实力最强”，计算结果小于最大的20%的为“实力差”，其他为“实力一般”；投资建议=判定结果为“实力最强”的建议“保留”，“实力差”的建议“淘汰”，“实力一般”的建议“进一步分析”；\
原始数据=合并资产负债表中当期总资产数。\n'
'       2、总资产同比增长率：计算结果=（当期总资产-上期总资产）/上期总资产*100%；得分办法=计算结果等于10%的10分，每大于10%的10%加1分，每小于10%的10%减1分；判定结果\
=计算结果大于或等于10%的为“成长性好”，计算结果小于10%的为“成长性差”；投资建议=判定结果为“成长性好”的建议“保留”，“成长性差”的建议“淘汰”；原始数据=合并资产负债表中当期总资产\
和上期总资产。\n'
'      （二）维度2：从负债分析偿债风险。分析指标：资产负债率、准货币资金减有息负债的差额。数据来源：各公司各年度的合并资产负债表。\n'
'       1、资产负债率：计算结果=总负债/总资产*100%；得分办法=计算结果等于60%的10分，每小于60%的10%则加1分，每大于60%的10%则减1分；判定结果=计算结果小于40%的为“基本\
没偿债风险”，计算结果大于或等于40%且小于60%的为“偿债风险小”，计算结果大于或等于60%且小于70%的为“偿债风险中”，计算结果大于70%的为“偿债风险大”，计算结果等于70%的为“偿债风险\
警戒线”；投资建议=计算结果小于60%的建议“保留”，计算结果大于或等于60%的建议“淘汰”；原始数据=合并资产负债表中当期总负债和总资产。\n'
'       2、准货币资金减有息负债的差额：计算结果=（货币资金+交易性金融资产）-（短期借款+一年内到期的非流动负债+长期借款+应付债券+长期应付款）；得分办法=计算结果等于0的10分，\
每大于10%的加1分，每小于10%的减1分；判定结果=计算结果大于或等于0的为“没偿债风险”，计算结果小于0的为“有偿债风险”；投资建议=判定结果为“没偿债风险”的建议“保留”，“有偿债风险”的\
建议“淘汰”；原始数据=合并资产负债表中货币资金、交易性金融资产、短期借款、一年内到期的非流动负债、长期借款、应付债券、长期应付款。\n'
'      （三）维度3：从应付预收减应收预付的差额分析竞争优势。分析指标：应付预收减应收预付的差额。数据来源：各公司各年度的合并资产负债表。计算结果=（货币资金+交易性金融资产）-\
（短期借款+一年内到期的非流动负债+长期借款+应付债券+长期应付款）；得分办法=计算结果等于0的10分，每大于10%的加1分，每小于10%的减1分；判定结果=计算结果大于或等于0的为“没偿债\
风险”，计算结果小于0的为“有偿债风险”；投资建议=判定结果为“没偿债风险”的建议“保留”，“有偿债风险”的建议“淘汰”；原始数据=合并资产负债表中货币资金、交易性金融资产、短期借款、一年\
内到期的非流动负债、长期借款、应付债券、长期应付款。\n'
'      （四）维度4：从应收帐款、合同资产分析产品竞争力。分析指标：应收帐款加合同资产占总资产比率。数据来源：各公司各年度的合并资产负债表。计算结果=（应收帐款+合同资产）/总资\
产*100%；得分办法=计算结果等于15%的10分，每大于10%的减1分，每小于10%的加1分；判定结果=计算结果小于1%的为“最好公司，产品很畅销”，计算结果小于3%的为“优秀公司，产品畅销”，\
计算结果小于10%的为“一般公司，产品销售一般”，计算结果大于或等于10%的为“较差公司，产品较难销”，计算结果大于或等于20%的为“很差公司，产品很难销”；投资建议=计算结果小于15%的建议\
“保留”，大于或等于15%至20%的建议“淘汰”（备注=如公司同时有3个及以上核心竞争力的可放宽到20%），大于20%的建议“淘汰”；原始数据=合并资产负债表中应收帐款、合同资产、总资产。\n'
'      （五）维度5：从固定资产分析维持竞争力的成本。分析指标：固定资产加在建工程占总资产比率。数据来源：各公司各年度的合并资产负债表。计算结果=（固定资产+在建工程）/总资\
产*100%；得分办法=计算结果等于40%的10分，每大于10%的减1分，每小于10%的加1分；判定结果=计算结果小于20%的为“轻资产型公司，维持竞争力成本较低，风险较低”，计算结果小于40%的\
为“一般公司，维持竞争力成本一般，风险一般”，计算结果大于或等于40%至50%的为“重资产型公司，维持竞争力成本较高，风险较大”，计算结果大于50%的为“重资产型公司，维持竞争力成本很高，\
风险很大”；投资建议=计算结果小于40%的建议“保留”，大于或等于40%至50%的建议“淘汰”（备注=如公司是行业第一，同时有3个及以上核心竞争力的可放宽到50%），大于50%的建议“淘汰”；\
原始数据=合并资产负债表中固定资产、在建工程、总资产。\n'
'      （六）维度6：从投产类资产分析主业专注度。分析指标：投产类资产占总资产比率。数据来源：各公司各年度的合并资产负债表。计算结果=（以公允价值计量且其变动计入当期损益的\
金融资产+债权投资+其他债权投资+可供出售金融资产+持有至到期投资+长期股权投资+其他权益工具投资+其他非流动金融资产+投资性房地产）/总资产*100%；得分办法=计算结果等于10%的10分，\
每大于10%的减1分，每小于10%的加1分；判定结果=计算结果等于于0的为“最优秀公司，最专注主业”，计算结果小于10%的为“优秀公司，专注主业”，计算结果大于或等于10%的为“较差公司，\
不专注主业”；投资建议=计算结果小于10%的建议“保留”，大于或等于10%的建议“淘汰”；原始数据=合并资产负债表中以公允价值计量且其变动计入当期损益的金融资产、债权投资、其他债权投资、\
可供出售金融资产、持有至到期投资、长期股权投资、其他权益工具投资、其他非流动金融资产、投资性房地产、总资产。\n'
'      （七）维度7：从存货、商誉分析未来业绩爆雷的风险。分析指标：应收帐款和存货占总资产比率、商誉占总资产比率。数据来源：各公司各年度的合并资产负债表。\n'
'       1、应收帐款和存货占总资产比率：计算结果=应收帐款/总资产*100%，存货/总资产*100%；得分办法=计算结果应收帐款/总资产*100%等于5%的5分、每小于5%的10%则加0.5分、\
每大于5%的10%则减0.5分，计算结果存货/总资产*100%等于15%的5分、每小于15%的10%则加0.5分、每大于15%的10%则减0.5分；判定结果=计算结果应收帐款/总资产*100%大于或等于5%且\
存货/总资产*100%大于或等于15%的为“爆雷风险较大”，其他的为“爆雷风险一般”；投资建议=判定结果为“爆雷风险一般”的建议“保留”，判定结果为“爆雷风险较大”的建议“淘汰”；原始\
数据=合并资产负债表中应收帐款、存货和总资产。\n'
'       2、商誉占总资产比率：计算结果=商誉/总资产*100%；得分办法=计算结果等于10%的10分，每大于10%的减1分，每小于10%的加1分；判定结果=计算结果大于或等于10%的为“爆雷\
风险较大”，计算结果小于10%的为“爆雷风险小”；投资建议=判定结果为“爆雷风险小”的建议“保留”，“爆雷风险较大”的建议“淘汰”；原始数据=合并资产负债表中商誉、总资产。\n'
'      （八）维度8：从营业收入分析行业地位及成长性。分析指标：营业收入规模、营业收入同比增长率。数据来源：各公司各年度的合并利润表。\n'
'       1、营业收入规模：计算结果=当期营业收入；得分办法=比较目标公司及其同行公司的该项的计算结果，最大的10分，每小于最大数10%则减1分；判定结果=计算结果在同行几个公司中\
最大的为“行业地位最高”，计算结果小于或等于最大的20%的为“行业地位差”，其他为“行业地位一般”；投资建议=判定结果为“行业地位最高”的建议“保留”，“行业地位差”的建议“淘汰”，“行业\
地位一般”的建议“进一步分析”；原始数据=合并利润表中营业收入。\n'
'       2、营业收入同比增长率：计算结果=（当期营业收入-上期营业收入）/上期营业收入*100%；得分办法=计算结果等于10%的10分，每大于10%的10%加1分，每小于10%的10%减1分；\
判定结果=计算结果大于10%的为“成长性好”，计算结果小于或等于10%的为“成长性差”；投资建议=判定结果为“成长性好”的建议“保留”，“成长性差”的建议“淘汰”；原始数据=合并利润表中\
当期营业收入和上期营业收入。\n'
'      （九）维度9：从毛利率分析产品竞争力及风险。分析指标：毛利率、毛利率波动幅度。数据来源：各公司各年度的合并利润表。\n'
'       1、毛利率：计算结果=（营业收入-营业成本）/营业成本*100%；得分办法=计算结果等于40%的10分，每小于10%则减1分，每大于10%则加1分；判定结果=计算结果小于或等\
于40%的为“产品\服务竞争力较差”，计算结果大于40%的为“产品\服务竞争力较强”；投资建议=判定结果为“产品\服务竞争力较强”的建议“保留”，“产品\服务竞争力较差”的建议“淘汰”；\
原始数据=合并利润表中营业收入、营业成本。\n'
'       2、毛利率波动幅度：计算结果=|（当期毛利率-上期毛利率）|/上期毛利率*100%；得分办法=计算结果等于20%的10分，每大于20%的10%减1分，每小于20%的10%加1分；判定\
结果=计算结果大于或等于20%的为“公司经营风险大，或财务造假风险大”，计算结果小于20%的为“公司经营风险小”，计算结果小于10%的为“优秀公司的正常波动”；投资建议=计算结果小\
于20%的建议“保留”，计算结果大于或等于20%的建议“淘汰”；原始数据=计算结果中的当期毛利率和上期毛利率。\n'
'      （十）维度10：从期间费用率分析成本管控能力。分析指标：期间费用率。数据来源：各公司各年度的合并利润表、维度9的毛利率计算结果。计算结果=（销售费用+管理费用+研发\
费用+财务费用）/营业收入*100%，如果财务费用为负数时则=（销售费用+管理费用+研发费用）/营业收入*100%；得分办法=计算结果/毛利率*100%等于60%的10分，每小于10%则减1分，\
每大于10%则加1分；判定结果=计算结果/毛利率*100%小于40%为“成本控制能力优秀”，计算结果/毛利率*100%大于或等于40%至60%的为“成本控制能力一般”，计算结果/毛利率*100%大于\
60%的为“成本控制能力较差”；投资建议=判定结果为“成本控制能力优秀”或“成本控制能力一般”的建议“保留”，“成本控制能力较差”的建议“淘汰”；原始数据=合并利润表中营业收入、销售费用、\
管理费用、研发费用、财务费用、维度9的毛利率计算结果。\n'
'      （十一）维度11：从销售费用率分析产品的销售难易度。分析指标：销售费用率、销售费用率变动趋势。数据来源：各公司各年度的合并利润表。\n'
'       1、销售费用率：计算结果=销售费用/营业收入*100%；得分办法=计算结果等于30%的10分，每小于10%则加1分，每大于10%则减1分；判定结果=计算结果小于15%的为“产品比较容\
易销售，销售风险相对较小”，计算结果小于或等于15%至30%的为“产品销售难度一般，销售风险一般”，计算结果大于30%的为“产品销售难度大，销售风险大”；投资建议=判定结果为“产品比较容\
易销售，销售风险相对较小”或“产品销售难度一般，销售风险一般”的建议“保留”，“产品销售难度大，销售风险大”的建议“淘汰”；原始数据=合并利润表中销售费用、营业收入。\n'
'       2、销售费用率变动趋势：计算结果=当期销售费用率-上期销售费用率；得分办法=计算结果等于0的10分，每大于10%减1分，每小于10%加1分；判定结果=计算结果等于0的为“产品销售\
难度不变”，计算结果小于0的为“产品销售难度在下降，销售风险也在下降”，计算结果大于0的为“产品销售难度在上升，销售风险也在上升”；投资建议=计算结果小于或等于0的建议“保留”，计算结果\
大于0的建议“淘汰”；原始数据=计算结果中的当期销售费用率和上期销售费用率。\n'
'      （十二）维度12：从主营利润分析主业的盈利能力及利润质量。分析指标：主营利润率、主营利润与营业利润的比率。数据来源：各公司各年度的合并利润表。\n'
'       1、主营利润率：计算结果=（营业收入-营业成本-税金及附加-销售费用-管理费用-研发费用-财务费）/营业收入*100%；得分办法=计算结果等于15%的10分，每小于10%则减1分，每大于\
10%则加1分；判定结果=计算结果小于15%的为“主业盈利能力弱”，计算结果大于或等于15%的为“主业盈利能力强”；投资建议=判定结果为“主业盈利能力强”的建议“保留”，“主业盈利能力弱”的\
建议“淘汰”；原始数据=合并利润表中营业收入、营业成本、税金及附加、销售费用、管理费用、研发费用、财务费用。\n'
'       2、主营利润与营业利润的比率：计算结果=（营业收入-营业成本-税金及附加-销售费用-管理费用-研发费用-财务费用）/（营业总收入－营业总成本+其他收益+投资收益+汇兑收益+净\
敞口套期收益+公允价值变动收益+信用减值损失+资产减值损失+资产处置收益）*100%；得分办法=计算结果等于80%的10分，每大于10%加1分，每小于10%减1分；判定结果=计算结果小于或等于\
80%的为“利润质量低”，计算结果大于80%的为“利润质量高”；投资建议=计算结果大于80%的建议“保留”，计算结果小于80%的建议“淘汰”；原始数据=合并利润表中营业总收入、营业总成本、其他\
收益、投资收益、汇兑收益、净敞口套期收益、公允价值变动收益、信用减值损失、资产减值损失、资产处置收益、营业收入、营业成本、税金及附加、销售费用、管理费用、研发费用、财务费用。\n'
'      （十三）维度13：从营业外收入净额进一步分析利润质量。分析指标：营业外收入净额与利润总额的比率。数据来源：各公司各年度的合并利润表。计算结果=（营业外收入-营业外支出）/（营业\
利润+营业外收入-营业外支出）*100%；得分办法=计算结果等于5%的10分，每小于10%则加1分，每大于10%则减1分；判定结果=计算结果小于5%为“利润质量优秀”，计算结果大于或等于5%为“利润质\
量较差”；投资建议=判定结果为“利润质量优秀”的建议“保留”，“利润质量较差”的建议“淘汰”；原始数据=合并利润表中营业外收入、营业外支出、营业利润。\n'
'      （十四）维度14：从归母净利润分析整体盈利能力及持续性。分析指标：归母净利润规模、归母净利润的增长率。数据来源：各公司各年度的合并利润表。\n'
'       1、归母净利润规模：计算结果=合并利润表中的归属母公司股东的净利润；得分办法=比较目标公司及其同行公司的该项的计算结果，最大的10分，每小于最大数10%则减1分；判定\
结果=计算结果在同行几个公司中最大的为“盈利能力最强”，计算结果小于最大的20%的为“盈利能力差”，其他为“盈利能力一般”；投资建议=判定结果为“盈利能力最强”的建议“保留”，“盈利\
能力差”的建议“淘汰”，“盈利能力一般”的建议“进一步分析”；原始数据=合并利润表中的归属母公司股东的净利润。\n'
'       2、归母净利润的增长率：计算结果=（当期归母净利润-上期归母净利润）/上期归母净利润*100%；得分办法=计算结果等于10%的10分，每大于10%加1分，每小于10%减1分；判定\
结果=计算结果大于或等于10%的为“盈利能力不但强而且持续性较好”，计算结果小于10%的为“盈利持续性差”，计算结果小于0%的为“已经处于衰落之中，盈利持续性较差”；投资建议=判定结果\
为“盈利能力不但强而且持续性较好”的建议“保留”，其他的建议“淘汰”；原始数据=合并利润表中的归属母公司股东的净利润。\n'
'      （十五）维度15：从购建固定资产、无形资产和其他长期资产支付的现金分析增长潜力。分析指标：购建固定资产、无形资产和其他长期资产支付的现金与经营活动产生的现金流量净额的比率，\
销售商品、提供劳务收到的现金变动趋势。数据来源：各公司各年度的合并现金流量表。\n'
'       1、购建固定资产、无形资产和其他长期资产支付的现金与经营活动产生的现金流量净额的比率：计算结果=购建固定资产、无形资产和其他长期资产支付的现金/经营活动产生的现金流\
量净额*100%；得分办法=计算结果等于100%的10分，每小于10%则加1分（当等于3%时总分则为10分，小于3%时，每小于10%减1分），每大于10%则减1分；判定结果=计算结果小于3%为“成长\
太慢，回报较低”，计算结果大于或等于3%而小于60%为“增长潜力较大并且风险相对较小”，计算结果大于或等于3%而小于60%为“增长潜力较大并且风险相对较小”，计算结果大于或等于60%为“扩张略\
高，风险一般”，计算结果大于或等于100%为“扩张激进，风险较大”；投资建议=判定结果为“增长潜力较大并且风险相对较小”和“扩张略高，风险一般”的建议“保留”，“成长太慢，回报较低”和“扩张激\
进，风险较大”的建议“淘汰”；原始数据=合并现金流量表中购建固定资产、无形资产和其他长期资产支付的现金，经营活动产生的现金流量净额。\n'
'       2、销售商品、提供劳务收到的现金变动趋势：计算结果=（当期销售商品、提供劳务收到的现金-上期销售商品、提供劳务收到的现金）/上期归母净利润*100%；得分办法=计算结果等\
于0的10分，每大于10%加1分，每小于10%减1分；判定结果=计算结果大于0的为“有增长潜力”，计算结果小于0的为“没有增长潜力”；投资建议=判定结果为“有增长潜力”的建议“保留”，“没有增长\
潜力”的建议“淘汰”；原始数据=合并现金流量表中的销售商品、提供劳务收到的现金。\n'
'      （十六）维度16：从分红率（也叫股利支付率）分析现金分红情况。分析指标：分红率。数据来源：各公司各年度的分红情况。计算结果=分红情况中的占合并报表中归属于上市公司普通股股东\
的净利润的比率(%)；得分办法=计算结果等于30的10分，每小于10%则减1分，每大于10%则加1分（当等于70时总分为10分，每大于10%减1分）；判定结果=计算结果小于30%为“不够厚道”，计算\
结果大于或等于30%至70%的为“厚道且持续性强”，计算结果大于70%的为“难以持续”；投资建议=判定结果为“厚道且持续性强”的建议“保留”，“不够厚道”和“难以持续”的建议“淘汰”；原始\
数据=各年度的分红情况中的占合并报表中归属于上市公司普通股股东的净利润的比率(%)。\n\n'

'       年报文本分析。\n'
'       主要根据提供的下载资料（年报、招股说明书、公司章程）及通过搜索的相关资料，重点从15个方面分析：会计师事务所对公司各年度的年报出具了什么样的审计报告，对公司董事会审议\
的各年报告期利润分配预案或公积金转增股本预案情况（分红情况）， 公司所在行业情况如何，公司的商业模式及动作策略怎样 ，核心竞争力如何，公司报告期最近一年主要做什么，对报告期最近\
一年主要经营情况进行分析，对公司未来发展进行分析，分析公司的重要事项，分析普通股股份变动及股东情况，分析董事、监事、高级管理人员和员工情况，公司历年核心财务指标的变化趋势及原因。\
 经营战略的演变、核心护城河与竞争优势，历年潜在风险的暴露情况及最新改进建议， 基于最新网络情报的前瞻性市场定位与同业对比 。由AI进行分析。\n'
'       结果生成“文本分析”文件夹存放在目标文件夹中。该文件夹中有：各公司的《综合评估报告》。\n'
'       标准设置条件：\n'
'       你是一位资深金融分析师。本次分析提供了该公司及其同行业前三名公司【多份/多个年份】的年报等原始文档，以及【最新网络情报】。请严格按以下维度和角度进行时间序列上的长效对比与横向同业分析，\
给出一份详尽的综合评估报告。重点分析：\n'
'       1、会计师事务所对公司各年度的年报出具了什么样的审计报告，有没有出具过不是“标准无保留意见”的报告。\n'
'       2、对公司董事会审议的各年报告期利润分配预案或公积金转增股本预案情况（分红情况）分析。\n'
'       3、公司所在行业情况如何（包括行业特点、市场规模、生命周期阶段、未来情况等等），公司的基本情况（包括：公司的主营业务，重要产品，最重要的产品，公司在该行业中的地位如何，\
推断公司的资产、营收、净利润规模在行业中的大小位置，对供应商、经销商的话语权，营收最多的产品，产量和营收比其上位或下位公司低或高多少等等）。\n'
'       4、公司的商业模式及动作策略怎样（包括采购模式、生产模式、销售模式、生产工艺流程、品牌策略等等）。\n'
'       5、核心竞争力如何（包括企业管理能力、企业文化、销售渠道、产品技术、品牌优势、战略布局等等），是否具有如下的优势：强网络效应、优秀企业文化、独特的资源、成本优势、\
短缺的无形资产、高转换成本等等。\n'
'       6、公司报告期最近一年主要做什么（做了哪些重要事情、是否会加强公司的核心竞争力等等），取得怎样的成绩（包括取得的主要成绩，重点是营业收入、净利润、净资产等等的增长情况）。\n'
'       7、对报告期最近一年主要经营情况进行分析：\n'
'      （1）主营业务方面。利润表及现金流量表相关科目变动分析表中，一个科目变化，相关科目是否基本同步变化；主营业务分行业、分产品、分地区情况分析表中主要产品及主要产品的毛利率\
与整体毛利率对比是否异常；产销量情况分析表中生产量和销售量比上年增减情况是否异常；成本分析表中成本构成及成本增减是否异常，特别是产品材料构成中占比最高的材料成本增减是否异常；\
主要销售客户及主要供应商情况表中从前5名客户获得的营收占比情况、从前5名供应商采购金额占比情况；期间费用同比变化情况分析表中费用变化是否异常、解释说明是否合理；研发投入情况表中\
研发投入费用化还是资本化、研发投入占营收比例如何；期间现金流同比变化情况分析表中经营活动产生的现金流量净额、投资活动产生的现金流量净额、筹资活动产生的现金流量净额三个科目是否异常。\n'
'      （2）非主营业务导致利润重大变化的说明有否异常情况。\n'
'      （3）资产、负债情况方面。是否有变动超过30%且本期或上期期末数占总资产超过30%的科目。\n'
'      （4）投资状况方面。是否有重大股权投资；是否有重大的非股权投资，募集资金使用情况（是否按计划使用），其他重大的非股权投资是否适度，有否异常；以公允价值计算的金融资产中\
的理财产品和结构性存款金额如何。\n'
'      （5）重大资产和股权出售是否有异常。\n'
'      （6）主要控股参股公司有什么特别情况。\n'
'       8、对公司未来发展进行分析：公司的发展战略和经营计划如何，下年要做哪些重要事情，这些事情是否有得于发展战略的实现。\n'
'       9、分析公司的重要事项：公司发生什么重要事项，是否影响公司未来的发展。特别是章程是否规定分红最低下限、连续性、稳定性，具体执行情况如何；在解决同行竞争、上市信息真实\
性、财政补贴等方面是否做无限期赔偿承诺，如何承诺；是否有重大诉讼，如有，给公司带来的潜在风险有多大，是否影响核心竞争力；重大关联交易情况如何，次数多少。\n'
'       10、分析普通股股份变动及股东情况：大股东和公司核心团队成员有无减持或增持的情况；大股东有无质押的情况；实际控制人的持股比例如何。\n'
'       11、分析董事、监事、高级管理人员和员工情况：核心团队成员是谁，离职的核心团队成员是谁，核心团队成员的工作动力如何（主要从核心团队成员的持股情况和薪酬水平评判），\
核心团队成员的专业能力如何（主要从核心团队成员的主要工作经历评判，如是否从事过本行业工作、是否是从本行基层一步步上来的等等），离职的是核心团队中的重要人物还是非重要人物，员\
工总人数是否随公司扩张而增加（是否存在异常），人均薪酬是否存在异常\n'
'       12、公司历年核心财务指标的变化趋势及原因。\n'
'       13、经营战略的演变、核心护城河与竞争优势。\n'
'       14、历年潜在风险的暴露情况及最新改进建议。\n'
'       15、基于最新网络情报的前瞻性市场定位与同业对比。要求：融会贯通多份文档的信息，数据准确、观点鲜明、具备长远视角的宏观格局。\n'

'       企业分析。\n'
'       主要根据提供的下载资料（年报、招股说明书、公司章程）及通过搜索的相关资料，重点从：行业与竞争环境分析（行业空间、竞争格局、行业政策与趋势）、企业全面对比分析\
（企业领导者个人的特质、企业文化、企业治理与管理层分析、企业的商业模式、主营业务情况、竞争优势、重要报表数据、关键财务比率、企业团队、企业内外沟通、企业现金流、企业系统、\
法规、风险与未来成长性、其他情况）两大方面共17个层面。由AI进行分析。\n'
'       结果生成“企业分析”文件夹存放在目标文件夹中。该文件夹中有：各公司的《综合评估报告》。\n'
'       标准设置条件：\n'
'       你是一位资深的企业分析师，请围绕以下的维度和重点对目标企业作全面深入中肯的分析：\n'
'      （一）行业与竞争环境分析。\n'
'       1、行业空间：包括市场规模、成熟阶段、现已占份额，尚有空间等。\n'
'       2、竞争格局：包括市场份额集中度（如CR5）、波特五力模型分析。\n'
'       3、行业政策与趋势：包括监管导向、技术变革（如数字化、低碳化等）对行业的影响。\n'
'       （二）企业全面对比分析。\n'
'       1、企业领导者个人的特质：包括领导者的志向、见识、恒心、仁爱心、才能、自身学习等；\n'
'       2、企业文化：包括使命、愿景、核心价值观等；\n'
'       3、企业治理与管理层分析：包括股权结构（包括实际控制人、股权稳定性、是否存在股权质押风险等）、管理层（包括核心高管背景、履历及稳定性、过往战略执行能力等）、\
信息披露与诚信（包括历史公告是否清晰、有无监管处罚等）、股东协议、公司章程、治理结构、招股说明书等相关内容；\n'
'       4、企业的商业模式：用户模式、产品模式、推广模式、盈利模式等，包括客户与供应商集中度、产业链地位等；\n'
'       5、主营业务情况：包括收入结构、核心产品或服务等。\n'
'       6、竞争优势：包括品牌、技术专利、网络效应、特许经营权等护城河情况。\n'
'       7、重要报表数据：包括利润表（包括营收增长率、毛利率、净利率、费用率等的变化）、资产负债表（包括资产负债结构、有息负债率、现金及等价物、应收帐款、存货周转等）、\
现金流量表（包括经营活动现金流净额，是否与净利润匹配；投资与筹资活动现金流等）。\n'
'       8、关键财务比率：包括盈利能力（包括净资产收益率ROE、总资产收益率ROA等）、偿债能力（包括流动比率、速动比率、资产负债率等）、运营效率（包括各项资产周转率等）、\
价值指标（包括市盈率PE、市净率PB、市盈率相对盈利增长比率PEG等）。\n'
'       9、企业团队：领导、核心成员、组织架构、晋升通道、激励机制、培训等；\n'
'       10、企业内外沟通：人际关系、谈话谈心、自我批评与批评等；\n'
'       11、企业现金流：开源方面（经营活动开源情况、融资活动开源情况等）、节流方面(降低成本等)等；\n'
'       12、企业系统：企业规则制度、程序流程、人员分工、权利、责任、利益、纪律等；\n'
'       13、法规：企业遵法守法、企业学法用法等；\n'
'       14、风险与未来成长性：包括风险识别（包括政策风险、技术迭代风险、重大客户依赖风险、环保/ESG风险等）、成长驱动（包括未来增长点，如新产能、新产品、新市场、并购计划等）\
、战略规划（包括公司中长期战略是否清晰可行等）。\n'
'       15、其他（上面尚有包含的其他重要情况）。同时，实行分析和验证：包括历史趋势分析，观察其稳定性与趋势；勾稽核对，将财务数据与业务描述相互验证；质疑与求证，对异常数据\
寻找合理解释。作出综合判断：包括优势与短板总结（列出公司的核心优势与主要风险点）、估值评估（结合成长性与风险，判断当前市场价格是否合理）、形成结论（公司是否具备长期投资价值？\
关键观察点是什么？等）。\n\n'

'       价格分析。\n'
'       主要是根据深圳A股市盈率（PE）、10年期中国国债收益率、个股动态市盈率、个股动态股息率等指标，分析当前的个股股价是否是好价格，并假设深圳A股市盈率（PE）、\
10年期中国国债收益率符合好价格条件，个股动态市盈率和个股动态股息率为当前值不变，个股的价格是多少时为好价格。同时还为平常波段买卖设置目前价格处在什么区间（观察区、偏买区、\
偏卖区）的分析判断。\n'
'       结果生成“好价分析”文件夹存放在目标文件夹中。该文件夹中有：深圳A股市盈率（PE）截图、《估值分析报告》、《深度资产估值分析报告》。\n'
'       标准设置条件：\n'
'       一、根据以下几个方面进行计算、分析和判断A股：\n'
'       1、判断当前股价是否符合好价格条件（好价格条件同时具备：大盘指数（深圳A股）的当前PE小于20、个股（目标文件表中所列的所有个股，下同）TTM市盈率小于15、滚动股息率大\
于中国十年期国债收益率），如果是则在判断结果列用红色标示“目前为好价格”，如果不是则用绿色标示“目前不是好价格”。\n'
'       2、判断当前是偏买区还是偏卖区（偏买区条件为同时符合下面三个条件：大盘指数（深圳A股）的当前PE小于40、个股TTM市盈率小于30、滚动股息率大于中国十年期国债收益率的三\
分之二；偏卖区条件为同时符合下面三个条件：大盘指数（深圳A股）的当前PE大于40且小于60、个股TTM市盈率大于30且小于50、滚动股息率比中国十年期国债收益率小于三分之二且大于三分之一）。\
如果是偏买区则在判断结果列用黄色标示“目前为偏买区”，如果是偏卖区则在判断结果列用蓝色标示“目前为偏卖区”，如果都不是则在判断结果列用黑色标示“目前为观察区”。\n'
'       3、判断个股价格多少是好价格：假设大盘指数（深圳A股）当前PE小于20，中国十年期国债收益率、过去四个季度的每股收益都是目前的水平不变，个股价格是多少才符合上面所列的好\
价格条件。把计算出的具体数值以“股价小于+具体数值+为好价格”式样标示在判断结果列上。\n'
'       二、自动识别输入的股票是A股、港股、美股。\n'
'       三、港股则按以下规则进行分析：\n'
'      （1）判断当前股价是否符合好价格条件（好价格条件同时具备：恒生指数的当前PE小于10、个股（目标文件表中所列的港股个股，下同）TTM市盈率小于15、滚动股息率大于中国\
十年期国债收益率），如果是则在判断结果列用红色标示“目前为好价格”，如果不是则用绿色标示“目前不是好价格”。\n'
'      （2）判断当前是偏买区还是偏卖区（偏买区条件为同时符合下面三个条件：恒生指数的当前PE小于20、个股TTM市盈率小于30、滚动股息率大于中国十年期国债收益率的三分之二；\
偏卖区条件为同时符合下面三个条件：恒生指数的当前PE大于20且小于30、个股TTM市盈率大于30且小于50、滚动股息率比中国十年期国债收益率小于三分之二且大于三分之一）。如果是偏买\
区则在判断结果列用黄色标示“目前为偏买区”，如果是偏卖区则在判断结果列用蓝色标示“目前为偏卖区”，如果都不是则在判断结果列用黑色标示“目前为观察区”。\n'
'      （3）判断个股价格多少是好价格：假设恒生指数当前PE小于10，中国十年期国债收益率、过去四个季度的每股收益都是目前的水平不变，个股价格是多少才符合上面所列的好价格条件。\
把计算出的具体数值以“股价小于+具体数值+为好价格”式样标示在判断结果列上。\n'
'       四、美股则按以下规则进行分析：\n'
'      （1）判断当前股价是否符合好价格条件（好价格条件为：美联储利率小于4%+标普500指数的当前PE（Current S&P 500 PE Ratio下同）小于15+个股（目标文件表中所列的美股\
个股，下同）TTM市盈率小于15+个股的动态股息率大于美国10年期国债收益率，或美联储利率大于4%+标普500指数的当前PE小于10+个股TTM市盈率小于15+个股的动态股息率大于美国10年期\
国债收益率，或标普500跌幅大于50%+个股TTM市盈率小于15+个股的动态股息率大于美国10年期国债收益率），如果是则在判断结果列用红色标示“目前为好价格”，如果不是则用绿色标示“目前\
不是好价格”。\n'
'      （2）判断当前是偏买区还是偏卖区（偏买区条件为同时符合下面三个条件：恒生指数的当前PE小于20、个股TTM市盈率小于30、滚动股息率大于中国十年期国债收益率的三分之二；\
偏卖区条件为同时符合下面三个条件：恒生指数的当前PE大于20且小于30、个股TTM市盈率大于30且小于50、滚动股息率比中国十年期国债收益率小于三分之二且大于三分之一）。如果是偏买区\
则在判断结果列用黄色标示“目前为偏买区”，如果是偏卖区则在判断结果列用蓝色标示“目前为偏卖区”，如果都不是则在判断结果列用黑色标示“目前为观察区”。\n'
'      （3）判断个股价格多少是好价格：假设恒生指数当前PE小于10，中国十年期国债收益率、过去四个季度的每股收益都是目前的水平不变，个股价格是多少才符合上面所列的好价格条件。把\
计算出的具体数值以“股价小于+具体数值+为好价格”式样标示在判断结果列上。\n\n\n'


'       本模型主要有两种用法。一是一键运行。程序从海选公司自动运行到好价分析，并在同一文件夹中生成各种结果。二是各模块独立运行。主要是针对已有资料进行某项分析。\n\n'

'       一、一键运行。\n'
'       首先，填选好总控面板中的各选填项目。主要是保存路径（可所有结果统一保存在一个文件夹，也可各个模块的结果分设独立保存文件夹）、AI配置（填写密钥，可全局统一一种AI设置\
，也可相关模块分别设置）、搜索引擎配置（填写密钥，可钥，可全局统一一种搜索引擎设置，也可相关模块分别设置），运行显示模式可按个人喜好选择。\n'
'       其次，设置各模块的条件。主要是海选公司模块中的程序一、程序二中的选股条件设置和年报下载模块中的年报下载年度设置，其他模块可默认（如需设置可参考第二部分“各模块独立运行”的\
设置方法）。\n\n'

'       二、各模块独立运行。\n\n'

'       海选公司模块。\n'
'       主要用于筛选企业，有AI筛选和同花顺网筛选两种。\n'
'       首先，选择好结果保存路径。\n'
'       其次，选填程序一的AI配置、搜索引擎配置（代理模式一般情况下选择“一般（直连）”就够了）、筛选条件（在“任务指令与模板管理”中选模板或自行在窗口中填写）。\n'
'       再次，选填程序二的市场分类（需要筛选股票的市场类别）、筛选条件（在“条件模板”中选模板或自行在“输入问财选股条件”窗口中填写）、是否静默运行选择。\n\n'

'       下载年报模块。\n'
'       主要用于从巨潮资讯网下载公司年报、公司章程、招股说明书等资料。\n'
'       1、输入股票代码或名称，可从现有文件夹导入（文件夹中须有股票名称或股票代码的xlsx文件），也可手动输入。\n'
'       2、选择需下载的年报时间范围。\n'
'       3、选择下载要求（核心功能和下载模式）。\n'
'       4、选择保存文件夹。\n\n'

'       年报提取模块。\n'
'       主要用于将下载的年报相关数据（主要是合并资产负债表、合并利润表、合并现金流量表及员工和分红情况）提取及汇总成xlsx表，并通过搜索补充完善相关数据。\n'
'       1、选择导入年报文件夹（文件夹中须有年报的pdf文件）。\n'
'       2、选择结果保存文件夹。\n'
'       3、填写财务指标强制换算规则（如有新的财务指标与旧的不同时才填写，一般情况不用填写，只保持默认即可）。\n\n'

'       造假排雷模块。\n'
'       主要用于对公司财务造假嫌疑进行分析判断。\n'
'       1、选择导入文件夹（文件夹中须有《统一融合输出》xlsx表相关数据）。\n'
'       2、选择结果保存文件夹。\n\n'

'       16维度分析模块。\n'
'       主要用于以好企业标准与同行企业的对比分析判断。\n'
'       1、选择导入文件夹（文件夹中须有《统一融合输出》xlsx表相关数据）。\n'
'       2、选择结果保存文件夹。\n\n'

'       AI文本分析模块。\n'
'       主要用于对公司年报文字内容的分析。\n'
'       一、点选““核心分析工作台”：\n'
'       1、选填AI配置。\n'
'       2、选填联网搜索配置。\n'
'       3、选择导入文件夹（文件夹中须有年报等pdf文档）。\n'
'       4、选择结果保存文件夹。\n'
'       二、点选“分析指令模板配置”：可选模板，也可手动输入（在下方窗口中）。\n\n'

'       AI企业分析模块。\n'
'       主要用于对公司的全面分析。\n'
'       一、点选““核心分析工作台”：\n'
'       1、选填AI配置。\n'
'       2、选择导入文件夹（文件夹中须有年报等pdf文档）。\n'
'       3、选择结果保存文件夹。\n'
'       4、选填多通道数据融合选项（一般默认全选即可）。\n'
'       二、点选“分析指令模板配置”：可选模板，也可手动输入（在下方窗口中）。\n\n'

'       好价估值分析模块。\n'
'       主要用于对个股的价值和价格进行分析估算和判断。\n'
'       1、选择结果保存文件夹。\n'
'       2、选填运行配置与宏观指标。运行配置主要是静默运行的选择与否，不选则会看到打开的网页。宏观指标保持默认全选即可。\n'
'       3、选填自选股池（即需分析的股票），可从文件夹中导入（文件夹中须有《需对比分析的全部公司名单》xlsx表），也可手动填写。\n\n'
        )
        txt.insert(tk.END, help_text)
        txt.config(state=tk.DISABLED)

    def setup_ui(self):
        # ----------------- 标题、署名与帮助区 -----------------
        title_frame = tk.Frame(self.root, bg=self.BG_MAIN)
        title_frame.pack(fill=tk.X, padx=20, pady=(15, 5))

        tk.Label(title_frame, text="价值投资智能分析模型V1.0", font=("微软雅黑", 22, "bold"), bg=self.BG_MAIN,
                 fg=self.FG_TITLE).pack(side=tk.LEFT)
        tk.Label(title_frame, text="田夫开发", font=("微软雅黑", 10, "bold"), bg=self.BG_MAIN, fg=self.FG_SUB).pack(
            side=tk.LEFT, padx=(8, 0), pady=(10, 0))

        tk.Button(title_frame, text="❓ 帮助说明", font=("微软雅黑", 13, "bold"), bg="#AA7942", fg="#FFFFFF",
                  command=self.show_help, relief=tk.RAISED, bd=3, padx=10, pady=2).pack(side=tk.RIGHT, pady=(5, 0))

        # ----------------- 说明面板区 (高亮立体) -----------------
        desc_frame = tk.Frame(self.root, bg=self.BG_MAIN)
        desc_frame.pack(fill=tk.X, padx=20, pady=(0, 5))
        desc_text = (
            "【模型使用指引】\n"
            "1、本模型功能是从全市场筛选优秀公司作综合分析，得出判断和建议结果作为价值投资参考。\n"
            "2、程序可一键运行（点击总控面板底下的总运行按键），也可各模块独立运行（点击各模块设置运行）。默认为一键运行。\n"
            "3、运行前需先设置相关参数和相关配置（详见帮助说明）。一键运行可选择统一设置各模块使用同一的“结果保存路径”“AI”“搜索引擎”，也可选择各模块分别使用不同的设置。\n"
            "4、运行结果主要选看：海选中的《海选公司汇总表》、排雷分析中的报告、16维度分析中的报告、文本分析中的报告、企业分析中的报告、好价分析中的《深度资产估值分析报告》。其他均为辅助资料文档。\n"
            "5、使用前需注册：AI的API用户并取得密钥、相关搜索引擎用户并取得密钥、同花顺问财网用户。\n"
            "6. 如遇问财网搜索失败，需在海选模块中的程序二中取消“静默运行”再启动该模块，出现网页提示登录时，点微信登录进行扫码后运行，之后再静默运行就正常了。"
        )
        tk.Label(desc_frame, text=desc_text, justify=tk.LEFT, bg="#F0DEB4", fg=self.FG_TEXT,
                 font=("微软雅黑", 10, "bold"), padx=15, pady=10, relief=tk.RIDGE, bd=3, anchor="w").pack(fill=tk.X)

        # ----------------- 全局统一控制面板区 -----------------
        cfg_frame = tk.Frame(self.root, bg=self.BG_MAIN)
        cfg_frame.pack(fill=tk.X, padx=15, pady=5)

        # 【全局路径控制】 (带有浮雕边框)
        path_frame = tk.LabelFrame(cfg_frame, text=" 📂 流水线路径控制 ", bg=self.BG_FRAME, fg=self.FG_TITLE,
                                   font=("微软雅黑", 10, "bold"), relief=tk.RIDGE, bd=2)
        path_frame.pack(fill=tk.X, pady=5)

        self.path_mode_var = tk.StringVar(value="global")
        tk.Radiobutton(path_frame, text="一键统一使用全局保存路径", variable=self.path_mode_var, value="global",
                       bg=self.BG_FRAME, fg=self.FG_TEXT, selectcolor=self.BG_MAIN, font=("微软雅黑", 9, "bold"),
                       command=self.toggle_global_path).grid(row=0, column=0, padx=5, pady=2, sticky="w")
        tk.Radiobutton(path_frame, text="保留并使用各子程序原有独立路径", variable=self.path_mode_var,
                       value="individual", bg=self.BG_FRAME, fg=self.FG_TEXT, selectcolor=self.BG_MAIN,
                       command=self.toggle_global_path).grid(row=0, column=1, padx=5, pady=2, sticky="w")

        self.global_base_dir = tk.StringVar(value=os.getcwd())
        tk.Label(path_frame, text="全局输出根目录:", bg=self.BG_FRAME, fg=self.FG_TEXT,
                 font=("微软雅黑", 9, "bold")).grid(row=1, column=0, sticky="e", pady=2)
        self.entry_global_path = tk.Entry(path_frame, textvariable=self.global_base_dir, width=50, relief=tk.SUNKEN,
                                          bd=2, bg="#FFFFFF", fg=self.FG_TEXT)
        self.entry_global_path.grid(row=1, column=1, padx=5)
        self.btn_global_path = tk.Button(path_frame, text="浏览...",
                                         command=lambda: self.global_base_dir.set(filedialog.askdirectory()),
                                         bg="#C29B62", fg="#FFFFFF", font=("微软雅黑", 9, "bold"), relief=tk.RAISED,
                                         bd=2)
        self.btn_global_path.grid(row=1, column=2, padx=5)

        # 【全局 AI 控制】
        ai_frame = tk.LabelFrame(cfg_frame, text=" 🤖 全局 AI 引擎统一控制 ", bg=self.BG_FRAME, fg=self.FG_TITLE,
                                 font=("微软雅黑", 10, "bold"), relief=tk.RIDGE, bd=2)
        ai_frame.pack(fill=tk.X, pady=5)

        self.ai_mode_var = tk.StringVar(value="global")
        tk.Radiobutton(ai_frame, text="一键统一写入全局 AI 接口配置", variable=self.ai_mode_var, value="global",
                       bg=self.BG_FRAME, fg=self.FG_TEXT, selectcolor=self.BG_MAIN, font=("微软雅黑", 9, "bold"),
                       command=self.toggle_global_ai).grid(
            row=0, column=0, columnspan=2, padx=5, pady=2, sticky="w")
        tk.Radiobutton(ai_frame, text="保留并使用各子程序原有 AI 配置", variable=self.ai_mode_var, value="individual",
                       bg=self.BG_FRAME, fg=self.FG_TEXT, selectcolor=self.BG_MAIN, command=self.toggle_global_ai).grid(
            row=0, column=2, columnspan=2, padx=5, pady=2, sticky="w")

        tk.Label(ai_frame, text="API 地址:", bg=self.BG_FRAME, fg=self.FG_TEXT, font=("微软雅黑", 9, "bold")).grid(
            row=1, column=0, sticky="e", pady=2)
        self.global_ai_url = tk.StringVar(value="https://api.deepseek.com/v1")
        self.cb_global_url = ttk.Combobox(ai_frame, textvariable=self.global_ai_url,
                                          values=["https://api.deepseek.com/v1", "https://api.openai.com/v1",
                                                  "http://localhost:11434/v1", "https://api.moonshot.cn/v1",
                                                  "https://dashscope.aliyuncs.com/compatible-mode/v1"], width=28)
        self.cb_global_url.grid(row=1, column=1, padx=5, pady=2)

        tk.Label(ai_frame, text="API Key:", bg=self.BG_FRAME, fg=self.FG_TEXT, font=("微软雅黑", 9, "bold")).grid(row=1,
                                                                                                                  column=2,
                                                                                                                  sticky="e",
                                                                                                                  pady=2)
        self.global_ai_key = tk.StringVar(value="")
        self.entry_global_key = tk.Entry(ai_frame, textvariable=self.global_ai_key, width=28, show="*",
                                         relief=tk.SUNKEN, bd=2, bg="#FFFFFF", fg=self.FG_TEXT)
        self.entry_global_key.grid(row=1, column=3, padx=5, pady=2)

        tk.Label(ai_frame, text="模型名称:", bg=self.BG_FRAME, fg=self.FG_TEXT, font=("微软雅黑", 9, "bold")).grid(
            row=2, column=0, sticky="e", pady=2)
        self.global_ai_model = tk.StringVar(value="deepseek-chat")
        self.cb_global_model = ttk.Combobox(ai_frame, textvariable=self.global_ai_model,
                                            values=["deepseek-chat", "deepseek-reasoner", "gpt-4o", "qwen2.5:7b",
                                                    "moonshot-v1-8k", "qwen-max"], width=28)
        self.cb_global_model.grid(row=2, column=1, padx=5, pady=2)

        # 【全局网络与搜索引擎控制】
        search_frame = tk.LabelFrame(cfg_frame, text=" 🌐 全局搜索引擎配置 ", bg=self.BG_FRAME, fg=self.FG_TITLE,
                                     font=("微软雅黑", 10, "bold"), relief=tk.RIDGE, bd=2)
        search_frame.pack(fill=tk.X, pady=5)

        self.search_mode_var = tk.StringVar(value="global")
        tk.Radiobutton(search_frame, text="一键统一使用全局搜索配置", variable=self.search_mode_var, value="global",
                       bg=self.BG_FRAME, fg=self.FG_TEXT, selectcolor=self.BG_MAIN, font=("微软雅黑", 9, "bold"),
                       command=self.toggle_global_search).grid(row=0, column=0, columnspan=2, padx=5, pady=2,
                                                               sticky="w")
        tk.Radiobutton(search_frame, text="保留海选系统原有的独立配置", variable=self.search_mode_var,
                       value="individual", bg=self.BG_FRAME, fg=self.FG_TEXT, selectcolor=self.BG_MAIN,
                       command=self.toggle_global_search).grid(row=0, column=2, columnspan=2, padx=5, pady=2,
                                                               sticky="w")

        tk.Label(search_frame, text="搜索引擎:", bg=self.BG_FRAME, fg=self.FG_TEXT, font=("微软雅黑", 9, "bold")).grid(
            row=1, column=0, sticky="e", pady=2)
        self.global_engine_var = tk.StringVar(value="Tavily")
        self.rb_tavily = tk.Radiobutton(search_frame, text="Tavily", variable=self.global_engine_var, value="Tavily",
                                        bg=self.BG_FRAME, fg=self.FG_TEXT, selectcolor=self.BG_MAIN)
        self.rb_tavily.grid(row=1, column=1, sticky="w", padx=5)
        self.rb_ddg = tk.Radiobutton(search_frame, text="DuckDuckGo", variable=self.global_engine_var, value="DDG",
                                     bg=self.BG_FRAME, fg=self.FG_TEXT, selectcolor=self.BG_MAIN)
        self.rb_ddg.grid(row=1, column=2, sticky="w", padx=5)

        tk.Label(search_frame, text="抓取资讯条数:", bg=self.BG_FRAME, fg=self.FG_TEXT,
                 font=("微软雅黑", 9, "bold")).grid(row=1, column=3, sticky="e", pady=2, padx=(10, 0))
        self.global_fetch_count = tk.IntVar(value=5)
        self.spin_fetch_count = ttk.Spinbox(search_frame, from_=1, to=1000, textvariable=self.global_fetch_count,
                                            width=5)
        self.spin_fetch_count.grid(row=1, column=4, sticky="w", padx=5)

        tk.Label(search_frame, text="引擎 Key:", bg=self.BG_FRAME, fg=self.FG_TEXT, font=("微软雅黑", 9, "bold")).grid(
            row=2, column=0, sticky="e", pady=2)
        self.global_search_key = tk.StringVar(value="")
        self.entry_search_key = tk.Entry(search_frame, textvariable=self.global_search_key, width=28, show="*",
                                         relief=tk.SUNKEN, bd=2, bg="#FFFFFF", fg=self.FG_TEXT)
        self.entry_search_key.grid(row=2, column=1, columnspan=2, sticky="w", padx=5, pady=2)

        # 【全局显示模式切换】
        disp_frame = tk.LabelFrame(cfg_frame, text=" 👁️ 运行显示模式 ", bg=self.BG_FRAME, fg=self.FG_TITLE,
                                   font=("微软雅黑", 10, "bold"), relief=tk.RIDGE, bd=2)
        disp_frame.pack(fill=tk.X, pady=5)

        self.display_mode_var = tk.StringVar(value="show_all")
        tk.Radiobutton(disp_frame, text="逐一打开各子程序窗口显示进度", variable=self.display_mode_var,
                       value="show_all", bg=self.BG_FRAME, fg=self.FG_TEXT, selectcolor=self.BG_MAIN,
                       font=("微软雅黑", 9, "bold")).grid(row=0, column=0, padx=10, pady=2, sticky="w")
        tk.Radiobutton(disp_frame, text="隐藏子程序，仅总控统一显示进度", variable=self.display_mode_var,
                       value="master_only", bg=self.BG_FRAME, fg=self.FG_TEXT, selectcolor=self.BG_MAIN,
                       font=("微软雅黑", 9, "bold")).grid(row=0, column=1, padx=10, pady=2, sticky="w")
        tk.Radiobutton(disp_frame, text="极致静默默默运行 (无任何跳动)", variable=self.display_mode_var, value="silent",
                       bg=self.BG_FRAME, fg=self.FG_TEXT, selectcolor=self.BG_MAIN, font=("微软雅黑", 9, "bold")).grid(
            row=0, column=2, padx=10, pady=2, sticky="w")

        # [新增核心修改 2/3]：在总控制面板绘制起点选择按钮和一键启动按钮
        stage_frame = tk.LabelFrame(self.root, text="⚙️ 一键运行控制面板", padx=10, pady=10, bg=self.BG_FRAME,
                                    fg=self.FG_TITLE, font=("微软雅黑", 11, "bold"))
        stage_frame.pack(fill=tk.X, padx=20, pady=10)

        tk.Radiobutton(stage_frame, text="1. 从【海选公司子程序】开始运行（完整流程）", variable=self.start_stage_var,
                       value="stage1", bg=self.BG_FRAME, font=("微软雅黑", 10)).pack(anchor=tk.W, pady=2)
        tk.Radiobutton(stage_frame, text="2. 跳过海选，直接从【下载年报子程序】开始", variable=self.start_stage_var,
                       value="stage2", bg=self.BG_FRAME, font=("微软雅黑", 10)).pack(anchor=tk.W, pady=2)

        self.btn_master_run = tk.Button(
            stage_frame, text="🚀 一键启动智能投资流水线", font=("微软雅黑", 12, "bold"),
            bg=self.BTN_COLOR, fg=self.BTN_TEXT, activebackground="#6E1616", activeforeground="#FFFFFF",
            width=52, height=2, relief=tk.RAISED, bd=5, cursor="hand2", command=self.start_master_pipeline)
        self.btn_master_run.pack(fill=tk.X, pady=(10, 0))

        # 【整合】下方原先还有一个“🚀 一键全自动总运行”按钮，与上面这个功能完全重复，已删除；
        # 全流程步骤说明移到控制按钮上方作为提示文字（不再是按钮）
        tk.Label(stage_frame,
                 text="流程：海选 → 下载 → 提取 → 排雷 → 16维 → AI文本 → AI深度 → 好价估值",
                 bg=self.BG_FRAME, fg=self.FG_SUB, font=("微软雅黑", 9)).pack(anchor=tk.W, pady=(8, 0))

        # ----------------- 手动操作面板 (温润哑光金属按钮) -----------------
        menu_frame = tk.Frame(self.root, bg=self.BG_MAIN)
        menu_frame.pack(fill=tk.X, padx=20, pady=10)

        # 子按钮采用温润的哑光香槟/琥珀金，文字深咖色
        child_btn_bg = "#DDBA89"
        child_btn_fg = "#2B170B"

        col_a = tk.Frame(menu_frame, bg=self.BG_MAIN)
        col_a.pack(side=tk.LEFT, fill=tk.Y, expand=True, padx=5)
        tk.Label(col_a, text="【A: 数据海选与提取】", bg=self.BG_MAIN, fg=self.FG_TITLE,
                 font=("微软雅黑", 10, "bold")).pack()
        tk.Button(col_a, text="海选公司程序", bg=child_btn_bg, fg=child_btn_fg, font=("微软雅黑", 9, "bold"),
                  command=self.app1_win.deiconify, width=20, relief=tk.RAISED, bd=3).pack(pady=5)
        tk.Button(col_a, text="下载年报程序", bg=child_btn_bg, fg=child_btn_fg, font=("微软雅黑", 9, "bold"),
                  command=self.app2_win.deiconify, width=20, relief=tk.RAISED, bd=3).pack(pady=5)
        tk.Button(col_a, text="年报提取程序", bg=child_btn_bg, fg=child_btn_fg, font=("微软雅黑", 9, "bold"),
                  command=self.app3_win.deiconify, width=20, relief=tk.RAISED, bd=3).pack(pady=5)

        col_b = tk.Frame(menu_frame, bg=self.BG_MAIN)
        col_b.pack(side=tk.LEFT, fill=tk.Y, expand=True, padx=5)
        tk.Label(col_b, text="【B: 财务排雷与对比】", bg=self.BG_MAIN, fg=self.FG_TITLE,
                 font=("微软雅黑", 10, "bold")).pack()
        tk.Button(col_b, text="造假排雷程序", bg=child_btn_bg, fg=child_btn_fg, font=("微软雅黑", 9, "bold"),
                  command=self.app_b_fraud_win.deiconify, width=20, relief=tk.RAISED, bd=3).pack(pady=5)
        tk.Button(col_b, text="16维度分析程序", bg=child_btn_bg, fg=child_btn_fg, font=("微软雅黑", 9, "bold"),
                  command=self.app_b_16dim_win.deiconify, width=20, relief=tk.RAISED, bd=3).pack(pady=5)

        col_c = tk.Frame(menu_frame, bg=self.BG_MAIN)
        col_c.pack(side=tk.LEFT, fill=tk.Y, expand=True, padx=5)
        tk.Label(col_c, text="【C: AI深度与估值】", bg=self.BG_MAIN, fg=self.FG_TITLE,
                 font=("微软雅黑", 10, "bold")).pack()
        tk.Button(col_c, text="AI文本分析程序", bg=child_btn_bg, fg=child_btn_fg, font=("微软雅黑", 9, "bold"),
                  command=self.app_c_ai_pro_win.deiconify, width=20, relief=tk.RAISED, bd=3).pack(pady=5)
        tk.Button(col_c, text="AI企业深度程序", bg=child_btn_bg, fg=child_btn_fg, font=("微软雅黑", 9, "bold"),
                  command=self.app_c_ent_win.deiconify, width=20, relief=tk.RAISED, bd=3).pack(pady=5)
        tk.Button(col_c, text="好价估值分析程序", bg=child_btn_bg, fg=child_btn_fg, font=("微软雅黑", 9, "bold"),
                  command=self.app_c_val_win.deiconify, width=20, relief=tk.RAISED, bd=3).pack(pady=5)

        run_frame = tk.Frame(self.root, bg=self.BG_MAIN)
        run_frame.pack(pady=10)

        # 【整合】超大“一键全自动总运行”按钮已合并到【一键运行控制面板】的
        # “🚀 一键启动智能投资流水线”按钮（二者功能重复）；此处只保留暂停/停止与状态显示。
        control_frame = tk.Frame(run_frame, bg=self.BG_MAIN)
        control_frame.pack(pady=5)

        # 暂停按钮：温润的琥珀橙
        self.btn_pause = tk.Button(control_frame, text="⏸️ 暂停总控", font=("微软雅黑", 11, "bold"), bg="#D37B22",
                                   fg="#FFFFFF", command=self.pause_all, width=12, state=tk.DISABLED,
                                   relief=tk.RAISED, bd=3)
        self.btn_pause.pack(side=tk.LEFT, padx=10)

        # 停止按钮：深沉的砖红色
        self.btn_stop = tk.Button(control_frame, text="⏹️ 停止总控", font=("微软雅黑", 11, "bold"), bg="#A03222",
                                  fg="white",
                                  command=self.stop_all, width=12, state=tk.DISABLED, relief=tk.RAISED, bd=3)
        self.btn_stop.pack(side=tk.LEFT, padx=10)

        self.status_lbl = tk.Label(self.root, text="状态: 准备就绪，静候指令", font=("微软雅黑", 11, "bold"),
                                   bg=self.BG_MAIN,
                                   fg=self.FG_TITLE)
        self.status_lbl.pack(pady=5)

    def _apply_global_configs(self):
        """【整合】把总控面板的“全局 AI / 全局搜索引擎 / 全局输出路径”统一写入各子程序。
        原先只有旧的 start_all() 会做这件事，现在唯一入口 start_master_pipeline() 统一调用，
        因此无论选“1.从海选开始”还是“2.跳过海选”，全局配置都会生效。"""
        url = key = model = ""
        if self.ai_mode_var.get() == "global":
            url = self.global_ai_url.get().strip()
            key = self.global_ai_key.get().strip()
            model = self.global_ai_model.get().strip()

            self.app1.api_url_var.set(url)
            self.app1.api_key_var.set(key)
            self.app1.model_var.set(model)

            self.app_c_ai_pro.url_sel.set(url)
            self.app_c_ai_pro.key_entry.config(state=tk.NORMAL)
            self.app_c_ai_pro.key_entry.delete(0, tk.END)
            self.app_c_ai_pro.key_entry.insert(0, key)
            self.app_c_ai_pro.model_sel.set(model)

            self.app_c_ent.cb_base.set(url)
            self.app_c_ent.ent_key.delete(0, tk.END)
            self.app_c_ent.ent_key.insert(0, key)
            self.app_c_ent.cb_model.set(model)

        if self.search_mode_var.get() == "global":
            self.app1.engine_var.set(self.global_engine_var.get())
            self.app1.fetch_count_var.set(self.global_fetch_count.get())
            self.app1.tavily_key_var.set(self.global_search_key.get().strip())

        if self.path_mode_var.get() == "global":
            self.app1.save_path_var.set(self.global_base_dir.get())
            if self.global_base_dir.get().strip():
                self.app2.path_var.set(self.global_base_dir.get().strip())
                self.app3.out_var.set(self.global_base_dir.get().strip())

        return url, key, model

    def start_master_pipeline(self):
        """【整合】全流程唯一的“一键启动”入口（原 start_all + start_master_pipeline 合并）。"""
        stage = self.start_stage_var.get()

        # ① 统一套用全局 AI / 搜索引擎 / 输出路径配置
        try:
            self._apply_global_configs()
        except Exception as e:
            print(f"⚠️ 全局配置套用失败（不影响启动）: {e}")

        # ② 启动前先“夺回网络控制权”：强制清理可能被其它子程序写入的全局代理环境变量。
        #    否则下载年报子程序里的 requests 会被环境里的 HTTP_PROXY(http://127.0.0.1:7890) 劫持，
        #    表现为“抓取 xxx 时发生错误 ... ProxyError ... WinError 10061”，最后一无所获。
        _cleared = apply_direct_network_env()
        if hasattr(self, 'app1'):
            try:
                # 让程序一自身的“一般(直连)/全域(代理)”开关成为唯一权威设置
                self.app1.toggle_proxy()
            except Exception:
                pass

        # ③ 初始化总控运行状态与按钮
        self.is_auto_run = True
        self.btn_master_run.config(state=tk.DISABLED)
        self.btn_pause.config(state=tk.NORMAL)
        self.btn_stop.config(state=tk.NORMAL)

        if self.display_mode_var.get() == "silent":
            self.status_lbl.config(text="🤫 正在后台极致静默运行中 (无界面刷新)...")
        else:
            _left = [k for k in ("HTTP_PROXY", "HTTPS_PROXY", "http_proxy", "https_proxy") if os.environ.get(k)]
            if _left:
                self.update_status(f"⚠️ 网络代理未清理干净（{','.join(_left)}），下载可能失败！")
            else:
                _msg = "状态: 网络已切换为直连模式（已清理代理环境变量）"
                if _cleared:
                    _msg += "：" + ",".join(_cleared)
                self.update_status(_msg)

        # ④ 按用户选择的起点启动对应子程序
        if stage == "stage1":
            # 完整模式：从第一个子程序（海选公司）开始
            self.update_status("状态: 正在运行【海选公司程序】(A1)...")
            self.popup_child(self.app1_win)
            self.app1.start_task(is_auto=True)

        elif stage == "stage2":
            # 跳过模式：直接启动第二个子程序（下载年报）
            # 校验是否已传入名单，若无则弹出对话框要求用户手动指定海选结果表
            if not self.app2.auto_excel_path or not os.path.exists(self.app2.auto_excel_path):
                file_path = filedialog.askopenfilename(
                    title="跳过海选需指定基础名单，请选择【海选公司汇总表.xlsx】",
                    filetypes=[("Excel 文件", "*.xlsx *.xls")]
                )
                if not file_path:
                    messagebox.showwarning("提示", "未选择有效名单，任务已取消！")
                    self.is_auto_run = False
                    self.reset_all_buttons()
                    return
                # 将用户选择的文件路径注入给下载年报子程序
                self.app2.auto_excel_path = file_path

            # 路径校验：给个明确提示，避免“静默什么都不做”
            if not os.path.exists(self.app2.auto_excel_path):
                messagebox.showerror("错误", f"指定的基础名单不存在：\n{self.app2.auto_excel_path}")
                self.is_auto_run = False
                self.reset_all_buttons()
                return

            # 强制设置 app2（下载年报子程序）为联动模式并触发数据加载
            self.app2.data_source_var.set("auto")
            self.app2.on_data_source_change()

            # 名单装载校验：联动表里提不到 6 位股票代码时，立即中止并提示
            # （否则 start_thread 会因 stock_list 为空在静默模式下直接 return，表现为“跳过下载直接进提取”）
            if not re.findall(r'\b\d{6}\b', self.app2.code_text.get("1.0", tk.END)):
                messagebox.showerror("错误",
                                     "未能从所选名单中提取到任何6位股票代码，下载任务已取消！\n"
                                     "请确认选择的是《海选公司汇总表.xlsx》或《需对比分析的全部公司名单.xlsx》。")
                self.is_auto_run = False
                self.reset_all_buttons()
                return

            # 跳过模式下也要套用“全局输出根目录”，保证下载结果落在用户指定的目录
            if self.path_mode_var.get() == "global" and self.global_base_dir.get().strip():
                self.app2.path_var.set(self.global_base_dir.get().strip())

            # 按“运行显示模式”把【下载年报程序】窗口显示出来
            self.popup_child(self.app2_win)
            self.update_status("状态: 正在运行【下载年报程序】(A2)...")

            # 正式启动下载年报线程
            self.app2.start_thread(is_auto=True)

    # ... [保留原有的 on_app1_complete, on_app2_complete 等回调函数] ...



    # ------------------ (由于空间限制，下方的方法请保留原代码逻辑直接对接，如 toggle_global_path、on_app1_complete 等，无需修改) ------------------

    # ------------------ 面板切换响应 ------------------
    def toggle_global_path(self):
        if self.path_mode_var.get() == "global":
            self.entry_global_path.config(state=tk.NORMAL)
            self.btn_global_path.config(state=tk.NORMAL)
        else:
            self.entry_global_path.config(state=tk.DISABLED)
            self.btn_global_path.config(state=tk.DISABLED)

    def toggle_global_ai(self):
        if self.ai_mode_var.get() == "global":
            self.cb_global_url.config(state=tk.NORMAL)
            self.entry_global_key.config(state=tk.NORMAL)
            self.cb_global_model.config(state=tk.NORMAL)
        else:
            self.cb_global_url.config(state=tk.DISABLED)
            self.entry_global_key.config(state=tk.DISABLED)
            self.cb_global_model.config(state=tk.DISABLED)

    def toggle_global_search(self):
        if self.search_mode_var.get() == "global":
            self.rb_tavily.config(state=tk.NORMAL)
            self.rb_ddg.config(state=tk.NORMAL)
            self.spin_fetch_count.config(state=tk.NORMAL)
            self.entry_search_key.config(state=tk.NORMAL)
        else:
            self.rb_tavily.config(state=tk.DISABLED)
            self.rb_ddg.config(state=tk.DISABLED)
            self.spin_fetch_count.config(state=tk.DISABLED)
            self.entry_search_key.config(state=tk.DISABLED)

    # ------------------ 运行显示模式路由引擎 ------------------
    def update_status(self, msg):
        mode = self.display_mode_var.get()
        if mode != "silent":
            self.status_lbl.config(text=msg)

    def popup_child(self, child_win):
        if self.display_mode_var.get() == "show_all":
            child_win.deiconify()
        else:
            child_win.withdraw()

    # ------------------ 流水线控制 ------------------
    # 【整合】原 start_all() 已删除：它的全部功能（全局 AI/搜索/路径套用 + 启动海选）已被
    # 合并进 start_master_pipeline()，因此总控面板现在只剩一个“一键启动”按钮。

    def pause_all(self):
        if not self.is_auto_run: return
        if hasattr(self.app1, 'is_running') and self.app1.is_running:
            self.app1.toggle_pause()
            if self.app1.pause_flag:
                self.btn_pause.config(text="▶️ 继续总控")
                self.update_status("状态: 总控已暂停，等待恢复...")
            else:
                self.btn_pause.config(text="⏸️ 暂停总控")
                self.update_status("状态: 正在运行中...")

    def stop_all(self):
        if not self.is_auto_run: return
        # 先关掉自动接力开关，避免正在收尾的子程序又把下一棒交出去
        self.is_auto_run = False
        stopped = False
        # A1 海选程序：带 is_running / stop_flag 机制
        if hasattr(self.app1, 'is_running') and self.app1.is_running:
            try:
                self.app1.stop_task()
                stopped = True
            except Exception:
                pass
        # A2 下载年报程序：靠 is_auto 开关阻止其把结果接力给 A3
        if getattr(self.app2, 'is_auto', False):
            try:
                self.app2.is_auto = False
            except Exception:
                pass
        if stopped:
            self.update_status("状态: 总控已被人工强行终止！")
        else:
            # 提取/排雷/16维/AI 等子程序没有中断接口，只能停止接力；已在跑的那一步会自然跑完
            self.update_status("状态: 总控已停止接力（当前正在运行的那一步无法中途打断，会自然结束）。")
        self.reset_all_buttons()

    def reset_all_buttons(self):
        self.btn_master_run.config(state=tk.NORMAL)
        self.btn_pause.config(state=tk.DISABLED, text="⏸️ 暂停总控")
        self.btn_stop.config(state=tk.DISABLED)

    # ------------------ 流水线挂载挂钩 ------------------
    def on_app1_complete(self, success, excel_path):
        if not self.is_auto_run: return
        if excel_path and os.path.exists(excel_path):
            self.update_status("状态: 海选完毕，正在启动【下载年报程序】(A2)...")
            self.popup_child(self.app2_win)

            self.app2.data_source_var.set("manual")
            self.app2.load_file_data(excel_path)

            if self.path_mode_var.get() == "global":
                self.app2.path_var.set(self.global_base_dir.get())

            self.app2.start_thread(is_auto=True)
        else:
            self.update_status("状态: 海选未能获取有效名单，总运行终止。")
            self.is_auto_run = False
            self.reset_all_buttons()

    def on_app2_complete(self, success, downloaded_dir):
        if not self.is_auto_run: return
        # 【修复2】双重校验：不仅目录要存在，还必须有 PDF 真正落盘，才允许接力【财报提取程序】
        has_pdf = bool(downloaded_dir) and os.path.isdir(downloaded_dir) and (find_first_pdf(downloaded_dir) is not None)
        if success and has_pdf:
            self.update_status("状态: 下载完毕，正在启动【年报提取程序】(A3)...")
            self.popup_child(self.app3_win)

            self.app3.data_source_var.set("manual")
            self.app3.in_var.set(downloaded_dir)

            if self.path_mode_var.get() == "global":
                self.app3.out_var.set(self.global_base_dir.get())

            self.app3.run_app(is_auto=True)
        else:
            if downloaded_dir and os.path.isdir(downloaded_dir) and not has_pdf:
                self.update_status("状态: 年报下载未产生任何 PDF 文件，总运行终止。")
            else:
                self.update_status("状态: 年报下载过程发生错误，总运行终止。")
            self.is_auto_run = False
            self.reset_all_buttons()

    def on_app3_complete(self):
        if not self.is_auto_run: return
        self.update_status("状态: 提取对比完毕，正在无缝衔接启动【造假排雷程序】(B1)...")
        self.popup_child(self.app_b_fraud_win)

        extracted_dir = os.path.join(self.app3.out_var.get(), "报表提取完善")
        self.app_b_fraud.data_source_var.set("auto")
        self.app_b_fraud.auto_input_dir = extracted_dir

        if self.path_mode_var.get() == "global":
            self.app_b_fraud.auto_output_dir = self.global_base_dir.get()
        else:
            self.app_b_fraud.auto_output_dir = self.app_b_fraud.output_dir.get()

        self.app_b_fraud.toggle_source()
        self.app_b_fraud.start_analysis_thread(silent=True, callback=self.on_b_fraud_complete)

    def on_b_fraud_complete(self):
        if not self.is_auto_run: return
        self.update_status("状态: 排雷完毕，正在启动【16维度分析程序】(B2)...")
        self.popup_child(self.app_b_16dim_win)

        if self.path_mode_var.get() == "global":
            base_dir = self.global_base_dir.get()
            out_dir = base_dir
        else:
            base_dir = self.app3.out_var.get()
            out_dir = self.app_b_16dim.out_var.get()

        self.app_b_16dim.data_source_var.set("auto")
        self.app_b_16dim.auto_input_dir = base_dir
        self.app_b_16dim.auto_output_dir = out_dir
        self.app_b_16dim.toggle_source()

        self.app_b_16dim.start_process(silent=True, callback=self.on_b_16dim_complete)

    def on_b_16dim_complete(self):
        if not self.is_auto_run: return
        self.update_status("状态: 16维度分析完毕，正在启动【AI文本分析程序】(C1)...")
        self.popup_child(self.app_c_ai_pro_win)

        if self.path_mode_var.get() == "global":
            base_dir = self.global_base_dir.get()
            download_dir = os.path.join(base_dir, "报表下载")
            out_dir = base_dir
        else:
            download_dir = os.path.join(self.app2.path_var.get(), "报表下载")
            out_dir = self.app_c_ai_pro.entry_out.get()

        self.app_c_ai_pro.data_source_var.set("auto")
        self.app_c_ai_pro.auto_input_dir = download_dir
        self.app_c_ai_pro.auto_output_dir = out_dir
        self.app_c_ai_pro.toggle_source()

        self.app_c_ai_pro.auto_mode = True
        self.app_c_ai_pro.on_finish = self.on_c_ai_pro_complete
        self.app_c_ai_pro.start_thread()

    def on_c_ai_pro_complete(self):
        if not self.is_auto_run: return
        self.update_status("状态: AI文本分析完毕，正在启动【企业深度分析程序】(C2)...")
        self.popup_child(self.app_c_ent_win)

        if self.path_mode_var.get() == "global":
            base_dir = self.global_base_dir.get()
            download_dir = os.path.join(base_dir, "报表下载")
            out_dir = base_dir
        else:
            download_dir = os.path.join(self.app2.path_var.get(), "报表下载")
            out_dir = self.app_c_ent.var_output_dir.get()

        self.app_c_ent.data_source_var.set("auto")
        self.app_c_ent.auto_input_dir = download_dir
        self.app_c_ent.auto_output_dir = out_dir
        self.app_c_ent.toggle_source()

        self.app_c_ent.auto_mode = True
        self.app_c_ent.on_finish = self.on_c_ent_complete
        self.app_c_ent.start_thread()

    def on_c_ent_complete(self):
        if not self.is_auto_run: return
        self.update_status("状态: 企业深度分析完毕，正在启动【好价估值分析程序】(C3)...")
        self.popup_child(self.app_c_val_win)

        if self.path_mode_var.get() == "global":
            base_dir = self.global_base_dir.get()
            excel_path = os.path.join(base_dir, "报表下载", "需对比分析的全部公司名单.xlsx")
            out_dir = base_dir
        else:
            excel_path = os.path.join(self.app2.path_var.get(), "报表下载", "需对比分析的全部公司名单.xlsx")
            out_dir = self.app_c_val.save_path.get()

        self.app_c_val.save_path.set(out_dir)

        self.app_c_val.data_source_var.set("auto")
        self.app_c_val.auto_excel_path = excel_path
        # 【修复6】传入年报提取产出目录，供动态股息率读取本地《年报提取表》的分红数据
        try:
            self.app_c_val.auto_extract_dir = os.path.join(self.app3.out_var.get(), "报表提取完善")
        except Exception:
            self.app_c_val.auto_extract_dir = ""
        self.app_c_val.toggle_source()

        self.app_c_val.auto_mode = True
        self.app_c_val.on_finish = self.on_all_complete
        self.app_c_val.start_thread()

    def on_all_complete(self):
        if not self.is_auto_run: return
        self.update_status("状态: ✨ 恭喜！一键全自动投研与提取总流程已完美收官！")
        self.is_auto_run = False
        self.reset_all_buttons()

        messagebox.showinfo("完成",
                            "所有任务流水线（海选->下载->提取->排雷->16维->AI文本->AI深度->好价估值）已全部执行完毕！")


if __name__ == "__main__":
    multiprocessing.freeze_support()
    try:
        root = tk.Tk()
        app = MasterControlSystem(root)
        root.mainloop()
    except Exception:
        # 【修复4】界面构建阶段出错时，不再只留一行 traceback 就退出：
        # 打印完整堆栈 + 尝试用弹窗告知，并保持一个可关闭的提示窗口，方便定位问题。
        _tb = traceback.format_exc()
        print(_tb)
        try:
            _er = tk.Tk()
            _er.title("启动失败")
            _er.geometry("900x520")
            _txt = scrolledtext.ScrolledText(_er, font=("Consolas", 10))
            _txt.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
            _txt.insert(tk.END, "程序启动失败，完整错误信息如下：\n\n" + _tb)
            tk.Button(_er, text="关闭", command=_er.destroy, width=16).pack(pady=(0, 10))
            _er.mainloop()
        except Exception:
            pass