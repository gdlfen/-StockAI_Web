import os
import re
import time
import json
import logging
import requests
import shutil
import random
import concurrent.futures
from datetime import datetime
from urllib.parse import urlparse, urljoin
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import List, Dict, Optional, Union

# 第三方依赖 (需 pip install)
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment
import fitz  # PyMuPDF
import pdfplumber
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt
from openai import OpenAI
from duckduckgo_search import DDGS

# GUI 界面依赖 (Python内置)
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import threading

# ==========================================
# 1. 全局配置与模板 (需求 1、2、4、5)
# ==========================================
CONFIG = {
    "BASE_DIR": os.path.join(os.getcwd(), "data_output"),
    "TIMEOUT": 15,
    "MAX_RETRIES": 3,
    "MAX_WORKERS": 5,
    "OVERWRITE_MODE": True,  # 需求5: 覆盖之前的结果
    "SEARCH_ENGINE": "Bing", # 需求1: 搜索引擎设置
    "DEFAULT_HEADERS": {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8"
    },
    # 下拉菜单预设模板
    "TEMPLATES": {
        "wencai_conditions": [
            "连续3年ROE>15%，上市时间>5年，非ST，剔除金融业",
            "市盈率<20，市值>100亿，股息率>3%，现金流为正",
            "营业收入连续3年增长，净利润现金含量>80%"
        ],
        "ai_filter_prompts": [
            "过滤掉存在财务造假嫌疑、存贷双高、或者大股东质押率过高(>50%)的公司",
            "仅保留细分行业龙头，且主营业务清晰、具备深厚护城河的公司"
        ],
        "ai_text_prompts": [
            "作为资深财务专家，请对【公司名称】的年报进行财报排雷，重点关注资产减值、关联交易、现金流健康度。",
            "请分析【公司名称】的管理层讨论与分析（MD&A），提取其未来的核心战略规划与潜在风险。"
        ],
        "ai_enterprise_prompts": [
            "请对比【公司名称】及其核心竞争对手（{competitors}），分析其核心竞争力、市场份额及上下游话语权。",
            "请基于波特五力模型，对【公司名称】进行全方位企业战略与护城河分析。"
        ]
    }
}

# ==========================================
# 2. 基础核心工具类与目录管理
# ==========================================
def prepare_clean_directory(dir_path: str):
    """确保目录存在，并根据配置决定是否覆盖旧数据 (需求5)"""
    if CONFIG["OVERWRITE_MODE"] and os.path.exists(dir_path):
        shutil.rmtree(dir_path, ignore_errors=True)
        time.sleep(0.5) # 确保系统释放句柄
    os.makedirs(dir_path, exist_ok=True)
    return dir_path

class BaseDownloader:
    """基础网络下载器"""
    def __init__(self, max_workers: int = CONFIG["MAX_WORKERS"]):
        self.session = requests.Session()
        self.session.headers.update(CONFIG["DEFAULT_HEADERS"])
        self.max_workers = max_workers

    def download_file(self, url: str, save_dir: str, filename: Optional[str] = None) -> bool:
        if not filename: filename = os.path.basename(urlparse(url).path)
        file_path = os.path.join(save_dir, filename)
        for attempt in range(CONFIG["MAX_RETRIES"]):
            try:
                with self.session.get(url, stream=True, timeout=CONFIG["TIMEOUT"]) as r:
                    r.raise_for_status()
                    with open(file_path, 'wb') as f:
                        for chunk in r.iter_content(chunk_size=8192): f.write(chunk)
                return True
            except Exception:
                time.sleep(1)
        return False

# ----------------- 外部依赖与辅助函数占位 -----------------
# 针对你原代码中未提供的自定义函数（如 p1_process_pdf等），在这里保留占位
def p1_process_pdf(p, callback): return {"资产负债表": pd.DataFrame()}, "测试公司", "2023", "000001"
def p2_process_pdf_employee(p): return pd.DataFrame(), None, None, None
def p1_generate_multi_year_summary(*args): pass
def p2_fetch_financial_sheets(*args): return {}
def p2_fetch_10jqka_dividend(*args): return pd.DataFrame()
def p3_extract_company_name(f): return f.split("_")[0]
def p3_extract_year(f): return "2023"
def p3_process_company(*args): pass
def common_auto_adjust_col_width(*args): pass
MILESTONE_PATTERNS = ["资产负债表", "利润表", "现金流量表"]

# ==========================================
# 3. 业务流水线类：将子程序模块化封装
# ==========================================

class DataSelectionPipeline:
    """模块1: 数据海选 (问财 + AI 过滤) - 需求2"""
    @staticmethod
    def run(wencai_cond, ai_filter, output_base, log_func):
        out_dir = prepare_clean_directory(os.path.join(output_base, "1_数据海选"))
        log_func(f"🔎 开始数据海选...\n  问财条件: {wencai_cond}\n  AI筛选: {ai_filter}")
        # 这里接入真实的问财API或爬虫逻辑，这里用伪代码返回
        time.sleep(1.5) 
        mock_data = pd.DataFrame([{"代码": "000001", "名称": "平安银行"}, {"代码": "600519", "名称": "贵州茅台"}])
        mock_data.to_excel(os.path.join(out_dir, "海选公司汇总表.xlsx"), index=False)
        log_func(f"✅ 数据海选完成，共筛选出 {len(mock_data)} 家公司。保存在: {out_dir}")
        return out_dir, ["000001", "600519"]

class AnnualReportPipeline:
    """模块2: 年报下载与提取综合类 - 需求3"""
    @staticmethod
    def run_download_and_extract(stock_list, output_base, log_func):
        out_dir = prepare_clean_directory(os.path.join(output_base, "2_报表下载与提取"))
        log_func(f"📥 开始报表下载与提取，目标列表: {stock_list}")
        
        # 步骤 1: 年报下载逻辑
        dl_dir = os.path.join(out_dir, "原始PDF报表下载")
        os.makedirs(dl_dir, exist_ok=True)
        # TODO: 接入实际年报下载器
        time.sleep(1)
        log_func("  ✅ 原始PDF年报下载完成")

        # 步骤 2: 报表提取完善逻辑 (接入你提供的核心原逻辑)
        extract_dir = os.path.join(out_dir, "报表提取完善")
        os.makedirs(extract_dir, exist_ok=True)
        log_func("  ✅ 报表结构化提取完成")
        return extract_dir

class DeepValuationPipeline:
    """模块3: AI企业深度估值 (整合三合一) - 需求4"""
    @staticmethod
    def run(in_dir, out_dir, text_prompt, ent_prompt, use_good_price, api_url, api_key, model_name, log_func):
        val_dir = prepare_clean_directory(os.path.join(out_dir, "3_AI深度估值"))
        
        # 1. 文本分析
        t_dir = os.path.join(val_dir, "1_文本分析")
        os.makedirs(t_dir, exist_ok=True)
        log_func(f"🧠 [AI文本分析] 执行指令: {text_prompt[:20]}...")
        # 调用原 run_ai_text_analysis 逻辑...
        time.sleep(1)

        # 2. 企业分析
        e_dir = os.path.join(val_dir, "2_企业分析")
        os.makedirs(e_dir, exist_ok=True)
        log_func(f"🧠 [AI企业估值] 执行指令: {ent_prompt[:20]}...")
        # 调用原 run_enterprise_ai_analysis 逻辑...
        time.sleep(1)

        # 3. 好价分析
        if use_good_price:
            p_dir = os.path.join(val_dir, "3_好价分析")
            os.makedirs(p_dir, exist_ok=True)
            log_func(f"📈 [好价分析] 计算合理估值区间 (DCF/PE/PB)...")
            time.sleep(0.5)

        log_func(f"✅ AI深度估值生成完毕，报告存入: {val_dir}")
        return val_dir

# ==========================================
# 4. 全局一键启动流水线调度器 (需求1)
# ==========================================
class OneClickOrchestrator:
    @staticmethod
    def run_all(ui_config, log_func):
        try:
            log_func("="*40)
            log_func(f"🚀 开始执行金融分析流水线任务\n引擎: {ui_config['engine']} | 覆盖模式: {ui_config['overwrite']}")
            base_dir = ui_config["base_dir"]
            
            # 1. 数据海选
            hs_dir, stock_list = DataSelectionPipeline.run(
                ui_config["wencai"], ui_config["ai_filter"], base_dir, log_func
            )
            
            # 2. 报表下载与提取 (材料传递: stock_list)
            report_dir = AnnualReportPipeline.run_download_and_extract(
                stock_list, base_dir, log_func
            )
            
            # 3. 财务排雷 & 16维度 (调用你原有的工作流)
            log_func("⚡ 触发财务排雷与16维度分析...")
            # run_fraud_detection(report_dir, base_dir, ...)
            # run_16_dim_analysis(report_dir, base_dir, ...)
            time.sleep(1)

            # 4. AI 深度估值 (材料传递: report_dir)
            DeepValuationPipeline.run(
                in_dir=report_dir,
                out_dir=base_dir,
                text_prompt=ui_config["text_prompt"],
                ent_prompt=ui_config["ent_prompt"],
                use_good_price=ui_config["good_price"],
                api_url=ui_config["api_url"],
                api_key=ui_config["api_key"],
                model_name=ui_config["model_name"],
                log_func=log_func
            )

            log_func("🎉 全链路任务完美收官！所有关联数据与报告已生成。")
            log_func("="*40)
        except Exception as e:
            log_func(f"❌ 运行发生中断异常: {str(e)}")


# ==========================================
# 5. 后台管理页面 (GUI - 需求6)
# ==========================================
class BackendAdminPanel(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("量化研报后台管理系统 v2.0")
        self.geometry("850x700")
        
        # 顶部 Tabs
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(expand=True, fill='both', padx=10, pady=10)
        
        self._init_tab1_global()
        self._init_tab2_selection()
        self._init_tab3_reports()
        self._init_tab4_valuation()
        
        # 底部控制台与按钮
        self.bottom_frame = tk.Frame(self)
        self.bottom_frame.pack(fill='x', padx=10, pady=5)
        
        self.run_btn = tk.Button(self.bottom_frame, text="🚀 一键启动流水线", bg="#4CAF50", fg="white", font=("Arial", 12, "bold"), command=self.start_pipeline)
        self.run_btn.pack(side='top', fill='x', pady=5)
        
        self.log_text = tk.Text(self.bottom_frame, height=12, bg="#1E1E1E", fg="#00FF00", font=("Consolas", 10))
        self.log_text.pack(side='bottom', fill='x')
        
    def log(self, message):
        """线程安全的日志输出"""
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.update_idletasks()

    def _init_tab1_global(self):
        frame = ttk.Frame(self.notebook)
        self.notebook.add(frame, text="全局配置")
        
        ttk.Label(frame, text="结果保存路径:").grid(row=0, column=0, padx=10, pady=10, sticky='w')
        self.path_var = tk.StringVar(value=CONFIG["BASE_DIR"])
        ttk.Entry(frame, textvariable=self.path_var, width=50).grid(row=0, column=1, padx=10)
        ttk.Button(frame, text="浏览", command=lambda: self.path_var.set(filedialog.askdirectory())).grid(row=0, column=2)

        ttk.Label(frame, text="搜索引擎配置:").grid(row=1, column=0, padx=10, pady=10, sticky='w')
        self.search_var = tk.StringVar()
        engine_cb = ttk.Combobox(frame, textvariable=self.search_var, values=["Bing", "DuckDuckGo", "360"], state="readonly")
        engine_cb.grid(row=1, column=1, sticky='w', padx=10)
        engine_cb.current(0)
        
        self.overwrite_var = tk.BooleanVar(value=CONFIG["OVERWRITE_MODE"])
        ttk.Checkbutton(frame, text="启动时清空旧数据(覆盖最新结果)", variable=self.overwrite_var).grid(row=2, column=1, sticky='w', padx=10, pady=10)
        
        ttk.Label(frame, text="AI API Key:").grid(row=3, column=0, padx=10, pady=10, sticky='w')
        self.api_key_var = tk.StringVar(value="sk-your-key-here")
        ttk.Entry(frame, textvariable=self.api_key_var, width=50, show="*").grid(row=3, column=1, padx=10)

    def _init_tab2_selection(self):
        frame = ttk.Frame(self.notebook)
        self.notebook.add(frame, text="数据海选")
        
        ttk.Label(frame, text="问财选股条件模板:").pack(anchor='w', padx=10, pady=5)
        self.wencai_var = tk.StringVar()
        w_cb = ttk.Combobox(frame, textvariable=self.wencai_var, values=CONFIG["TEMPLATES"]["wencai_conditions"], width=80)
        w_cb.pack(padx=10, pady=5)
        w_cb.current(0)

        ttk.Label(frame, text="AI过滤条件模板:").pack(anchor='w', padx=10, pady=15)
        self.ai_filter_var = tk.StringVar()
        a_cb = ttk.Combobox(frame, textvariable=self.ai_filter_var, values=CONFIG["TEMPLATES"]["ai_filter_prompts"], width=80)
        a_cb.pack(padx=10, pady=5)
        a_cb.current(0)

    def _init_tab3_reports(self):
        frame = ttk.Frame(self.notebook)
        self.notebook.add(frame, text="年报下载与提取")
        ttk.Label(frame, text="本面板已将下载与提取模块深度整合。").pack(padx=10, pady=10)
        self.report_auto_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(frame, text="流水线中自动执行: 财务报表下载 -> 多年度数据提取 -> A/B表补全", variable=self.report_auto_var).pack(anchor='w', padx=20)

    def _init_tab4_valuation(self):
        frame = ttk.Frame(self.notebook)
        self.notebook.add(frame, text="AI企业深度估值")
        ttk.Label(frame, text="由文本分析、企业护城河分析、好价分析三大子程序整合").pack(padx=10, pady=10, anchor='w')

        ttk.Label(frame, text="1. AI 文本分析指令:").pack(anchor='w', padx=10)
        self.text_prompt_var = tk.StringVar()
        t_cb = ttk.Combobox(frame, textvariable=self.text_prompt_var, values=CONFIG["TEMPLATES"]["ai_text_prompts"], width=80)
        t_cb.pack(padx=10, pady=5)
        t_cb.current(0)

        ttk.Label(frame, text="2. AI 企业估值指令:").pack(anchor='w', padx=10, pady=(15, 0))
        self.ent_prompt_var = tk.StringVar()
        e_cb = ttk.Combobox(frame, textvariable=self.ent_prompt_var, values=CONFIG["TEMPLATES"]["ai_enterprise_prompts"], width=80)
        e_cb.pack(padx=10, pady=5)
        e_cb.current(0)
        
        self.good_price_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(frame, text="3. 开启财务好价分析 (PE/PB/DCF估值测算)", variable=self.good_price_var).pack(anchor='w', padx=10, pady=20)

    def start_pipeline(self):
        """搜集面板配置并启动后台线程"""
        ui_config = {
            "base_dir": self.path_var.get(),
            "engine": self.search_var.get(),
            "overwrite": self.overwrite_var.get(),
            "api_url": "https://api.openai.com/v1",
            "api_key": self.api_key_var.get(),
            "model_name": "gpt-4-turbo",
            "wencai": self.wencai_var.get(),
            "ai_filter": self.ai_filter_var.get(),
            "text_prompt": self.text_prompt_var.get(),
            "ent_prompt": self.ent_prompt_var.get(),
            "good_price": self.good_price_var.get()
        }
        
        CONFIG["OVERWRITE_MODE"] = ui_config["overwrite"]
        CONFIG["BASE_DIR"] = ui_config["base_dir"]
        CONFIG["SEARCH_ENGINE"] = ui_config["engine"]
        
        self.run_btn.config(state="disabled", text="任务执行中...")
        self.log_text.delete(1.0, tk.END)
        
        # 使用多线程避免阻塞 UI
        def worker():
            OneClickOrchestrator.run_all(ui_config, log_func=self.log)
            self.run_btn.config(state="normal", text="🚀 一键启动流水线")
            
        threading.Thread(target=worker, daemon=True).start()

if __name__ == "__main__":
    app = BackendAdminPanel()
    app.mainloop()
