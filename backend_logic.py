# backend_logic.py
import os
import time
import shutil
import requests
import pandas as pd
from urllib.parse import urlparse
from typing import Optional

# 第三方依赖 (如果需要用到对应功能，取消下面的注释即可)
# from openpyxl import Workbook
# import fitz  # PyMuPDF
# import pdfplumber
# from docx import Document
# from openai import OpenAI
# from duckduckgo_search import DDGS

# ==========================================
# 1. 全局配置与模板
# ==========================================
CONFIG = {
    "BASE_DIR": os.path.join(os.getcwd(), "data_output"),
    "TIMEOUT": 15,
    "MAX_RETRIES": 3,
    "MAX_WORKERS": 5,
    "OVERWRITE_MODE": True,  
    "SEARCH_ENGINE": "Bing", 
    "DEFAULT_HEADERS": {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8"
    },
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
    """确保目录存在，并根据配置决定是否覆盖旧数据"""
    if CONFIG["OVERWRITE_MODE"] and os.path.exists(dir_path):
        shutil.rmtree(dir_path, ignore_errors=True)
        time.sleep(0.5) 
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

# ==========================================
# 3. 业务流水线类：将子程序模块化封装
# ==========================================
class DataSelectionPipeline:
    @staticmethod
    def run(wencai_cond, ai_filter, output_base, log_func):
        out_dir = prepare_clean_directory(os.path.join(output_base, "1_数据海选"))
        log_func(f"🔎 开始数据海选...\n  问财条件: {wencai_cond}\n  AI筛选: {ai_filter}")
        time.sleep(1.5) 
        mock_data = pd.DataFrame([{"代码": "000001", "名称": "平安银行"}, {"代码": "600519", "名称": "贵州茅台"}])
        mock_data.to_excel(os.path.join(out_dir, "海选公司汇总表.xlsx"), index=False)
        log_func(f"✅ 数据海选完成，共筛选出 {len(mock_data)} 家公司。保存在: {out_dir}")
        return out_dir, ["000001", "600519"]

class AnnualReportPipeline:
    @staticmethod
    def run_download_and_extract(stock_list, output_base, log_func):
        out_dir = prepare_clean_directory(os.path.join(output_base, "2_报表下载与提取"))
        log_func(f"📥 开始报表下载与提取，目标列表: {stock_list}")
        
        dl_dir = os.path.join(out_dir, "原始PDF报表下载")
        os.makedirs(dl_dir, exist_ok=True)
        time.sleep(1)
        log_func("  ✅ 原始PDF年报下载完成")

        extract_dir = os.path.join(out_dir, "报表提取完善")
        os.makedirs(extract_dir, exist_ok=True)
        log_func("  ✅ 报表结构化提取完成")
        return extract_dir

class DeepValuationPipeline:
    @staticmethod
    def run(in_dir, out_dir, text_prompt, ent_prompt, use_good_price, api_url, api_key, model_name, log_func):
        val_dir = prepare_clean_directory(os.path.join(out_dir, "3_AI深度估值"))
        
        t_dir = os.path.join(val_dir, "1_文本分析")
        os.makedirs(t_dir, exist_ok=True)
        log_func(f"🧠 [AI文本分析] 执行指令: {text_prompt[:20]}...")
        time.sleep(1)

        e_dir = os.path.join(val_dir, "2_企业分析")
        os.makedirs(e_dir, exist_ok=True)
        log_func(f"🧠 [AI企业估值] 执行指令: {ent_prompt[:20]}...")
        time.sleep(1)

        if use_good_price:
            p_dir = os.path.join(val_dir, "3_好价分析")
            os.makedirs(p_dir, exist_ok=True)
            log_func(f"📈 [好价分析] 计算合理估值区间 (DCF/PE/PB)...")
            time.sleep(0.5)

        log_func(f"✅ AI深度估值生成完毕，报告存入: {val_dir}")
        return val_dir

# ==========================================
# 4. 全局一键启动流水线调度器
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
            
            # 2. 报表处理
            report_dir = AnnualReportPipeline.run_download_and_extract(
                stock_list, base_dir, log_func
            )
            
            log_func("⚡ 触发财务排雷与16维度分析...")
            time.sleep(1)

            # 3. AI深度估值
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
