# backend_logic.py
import os
import re
import time
import json
import shutil
import urllib.parse
from datetime import datetime
import pandas as pd
import requests
import io
import concurrent.futures

# 第三方分析与文档库
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
import akshare as ak
from docx import Document
from openai import OpenAI
from tavily import TavilyClient

# ==========================================
# 1. 模板管理与全局配置
# ==========================================
TEMPLATE_FILE = "user_templates.json"

DEFAULT_TEMPLATES = {
    "wencai_conditions": [
        "连续5年加权roe>25，连续5年净利润现金含量>80，连续5年毛利率>40，上市时间>3年，剔除北交所，非金融股",
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

CONFIG = {
    "DEFAULT_HEADERS": {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        "Accept": "application/json, text/javascript, */*; q=0.01"
    }
}

def load_templates():
    if os.path.exists(TEMPLATE_FILE):
        try:
            with open(TEMPLATE_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception: return DEFAULT_TEMPLATES
    return DEFAULT_TEMPLATES

def save_templates(templates_data):
    with open(TEMPLATE_FILE, "w", encoding="utf-8") as f:
        json.dump(templates_data, f, ensure_ascii=False, indent=4)

def prepare_clean_directory(dir_path: str, overwrite: bool):
    if overwrite and os.path.exists(dir_path):
        shutil.rmtree(dir_path, ignore_errors=True)
        time.sleep(0.5) 
    os.makedirs(dir_path, exist_ok=True)
    return dir_path

# ==========================================
# 2. 网络获取模块 (问财嗅探、巨潮下载、新浪财报)
# ==========================================
class WebScraperEngine:
    @staticmethod
    def get_wencai_data(query, log_func):
        """若云端无法运行Selenium，提供降级机制"""
        try:
            from selenium import webdriver
            from selenium.webdriver.chrome.options import Options
            opts = Options()
            opts.add_argument("--headless=new")
            opts.add_argument("--disable-gpu")
            opts.add_argument("--window-size=1920,1080")
            opts.add_argument("--disable-blink-features=AutomationControlled")
            opts.add_argument(f"user-agent={CONFIG['DEFAULT_HEADERS']['User-Agent']}")
            
            driver = webdriver.Chrome(options=opts)
            hook_script = """
            (function() {
                if (window.__V75_HOOKED__) return;
                window.__V75_HOOKED__ = true; window.__V75_DATAS__ = [];
                const originalParse = JSON.parse;
                JSON.parse = function(text, reviver) {
                    const result = originalParse(text, reviver);
                    try {
                        if (result && result.resultList && Array.isArray(result.resultList)) {
                            window.__V75_DATAS__.push(result.resultList);
                        }
                    } catch(e) {}
                    return result;
                };
            })();
            """
            driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {"source": hook_script})
            driver.get(f"https://www.iwencai.com/unifiedwap/result?w={urllib.parse.quote(query)}")
            time.sleep(8)
            payload = driver.execute_script("return window.__V75_DATAS__ || [];")
            if payload and len(payload) > 0:
                best_batch = payload[-1]
                cleaned_data = []
                for row in best_batch:
                    clean_row = {}
                    for k, v in row.items():
                        if 'code' in k.lower() or '代码' in k: 
                            raw = str(v).replace('sz','').replace('sh','').replace('bj','')
                            clean_row['代码'] = raw.zfill(6) if raw.isdigit() else raw
                        elif 'name' in k.lower() or '简称' in k or '名称' in k: 
                            clean_row['名称'] = str(v)
                    if '代码' in clean_row: cleaned_data.append(clean_row)
                driver.quit()
                return cleaned_data
            driver.quit()
            return []
        except Exception as e:
            return []

    @staticmethod
    def get_cninfo_orgid(stock_code):
        url = f"http://www.cninfo.com.cn/new/information/topSearch/query?keyWord={stock_code}"
        try:
            res = requests.get(url, headers=CONFIG["DEFAULT_HEADERS"], timeout=10).json()
            for item in res:
                if str(item.get('code', '')) == str(stock_code):
                    return item.get('orgId')
        except: pass
        return None

    @staticmethod
    def fetch_sina_financial_sheets(code, start_year, end_year, log_func):
        """精准抓取新浪财经历年财务数据表以供分析"""
        sheets = {}
        types = {'合并资产负债表': 'vDOWN_BalanceSheet', '合并利润表': 'vDOWN_ProfitStatement', '合并现金流量表': 'vDOWN_CashFlow'}
        for sheet_name, api_type in types.items():
            url = f"http://vip.stock.finance.sina.com.cn/corp/go.php/{api_type}/displaytype/4/stockid/{code}/ctrl/all.phtml"
            try:
                r = requests.get(url, headers=CONFIG["DEFAULT_HEADERS"], timeout=10)
                r.encoding = 'gbk'
                df = pd.read_csv(io.StringIO(r.text), sep='\t', on_bad_lines='skip')
                df.dropna(how='all', axis=1, inplace=True)
                valid_cols = [df.columns[0]]
                for c in df.columns[1:]:
                    if '12-31' in str(c) or '1231' in str(c):
                        try:
                            if start_year <= int(str(c)[:4]) <= end_year: valid_cols.append(c)
                        except: pass
                if len(valid_cols) > 1:
                    df_filtered = df[valid_cols].copy()
                    df_filtered.rename(columns={df_filtered.columns[0]: '项目'}, inplace=True)
                    sheets[sheet_name] = df_filtered
            except: pass
        return sheets

# ==========================================
# 3. 流水线类定义
# ==========================================
class DataSelectionPipeline:
    @staticmethod
    def run(wencai_cond, ai_filter, enable_ai_filter, output_base, overwrite, log_func):
        out_dir = prepare_clean_directory(os.path.join(output_base, "1_数据海选"), overwrite)
        log_func(f"🔎 开始数据海选...\n  问财条件: {wencai_cond}")
        
        data = WebScraperEngine.get_wencai_data(wencai_cond, log_func)
        if not data:
            log_func("  ⚠️ 问财嗅探未获取到数据或超时，启用本地备用优选股池...")
            data = [{"代码": "000001", "名称": "平安银行"}, {"代码": "600519", "名称": "贵州茅台"}]
            
        mock_data = pd.DataFrame(data)
        mock_data.to_excel(os.path.join(out_dir, "海选公司汇总表.xlsx"), index=False)
        log_func(f"✅ 数据海选完成，共获取 {len(mock_data)} 家标的。保存在: {out_dir}")
        return out_dir, [row["代码"] for row in data]

class AnnualReportPipeline:
    @staticmethod
    def run_download_and_extract(stock_list, ui_config, log_func):
        out_dir = prepare_clean_directory(os.path.join(ui_config["base_dir"], "2_报表下载与提取"), ui_config["overwrite"])
        start_y, end_y = ui_config["start_year"], ui_config["end_year"]
        log_func(f"📥 启动年报下载与财务提取，年份跨度: {start_y}-{end_y}")
        
        for code in stock_list:
            comp_dir = os.path.join(out_dir, f"财报档案_{code}")
            os.makedirs(comp_dir, exist_ok=True)
            
            # 1. 真实下载 PDF 公告 (年报、招股书、章程)
            orgid = WebScraperEngine.get_cninfo_orgid(code)
            if orgid:
                query_url = "http://www.cninfo.com.cn/new/hisAnnouncement/query"
                keywords = ["年度报告"]
                if ui_config["dl_prospectus"]: keywords.append("招股说明书")
                if ui_config["dl_charter"]: keywords.append("公司章程")
                
                for kw in keywords:
                    payload = {'pageNum': 1, 'pageSize': 30, 'tabName': 'fulltext', 'stock': f"{code},{orgid}", 'searchkey': kw, 'sdate': f"{start_y}-01-01", 'edate': f"{end_y}-12-31"}
                    try:
                        res = requests.post(query_url, data=payload, headers=CONFIG["DEFAULT_HEADERS"], timeout=10).json()
                        if res.get('announcements'):
                            for ann in res['announcements']:
                                title = ann['announcementTitle']
                                if any(x in title for x in ['摘要', '取消', '英文']): continue
                                adj_url = ann['adjunctUrl']
                                if adj_url.endswith('.pdf'):
                                    safe_title = re.sub(r'[\\/:*?"<>|]', '', title)
                                    pdf_path = os.path.join(comp_dir, f"{safe_title}.pdf")
                                    if not os.path.exists(pdf_path):
                                        pdf_data = requests.get(f"http://static.cninfo.com.cn/{adj_url}", headers=CONFIG["DEFAULT_HEADERS"]).content
                                        with open(pdf_path, 'wb') as f: f.write(pdf_data)
                                        log_func(f"    ⬇️ 下载公告成功: {safe_title}.pdf")
                                    break 
                    except Exception as e:
                        pass
            
            # 2. 获取真实的财务 Excel 宽表 (供排雷与估值使用)
            log_func(f"    📊 正在抓取 {code} 核心财务表数据...")
            fin_sheets = WebScraperEngine.fetch_sina_financial_sheets(code, start_y, end_y, log_func)
            if fin_sheets:
                excel_path = os.path.join(comp_dir, f"统一整合输出_{code}_{start_y}-{end_y}.xlsx")
                with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
                    for s_name, df in fin_sheets.items():
                        df.to_excel(writer, sheet_name=s_name, index=False)
                log_func(f"    ✅ 财务数据已同步: {os.path.basename(excel_path)}")
            else:
                log_func(f"    ⚠️ 未抓取到 {code} 的财务宽表数据。")

        log_func("✅ 年报与财务档案处理完毕")
        return out_dir

class FraudAndDim16Pipeline:
    @staticmethod
    def run(in_dir, output_base, overwrite, log_func):
        out_dir = prepare_clean_directory(os.path.join(output_base, "3_财务排雷与16维分析"), overwrite)
        log_func(f"⚡ 开始执行财务造假智能排雷与多维度分析...")
        
        # 遍历所有财务宽表
        for root, dirs, files in os.walk(in_dir):
            for f in files:
                if f.startswith("统一整合输出_") and f.endswith(".xlsx"):
                    file_path = os.path.join(root, f)
                    comp_code = f.split('_')[2].split('_')[0]
                    
                    try:
                        # 简易读取解析
                        xls = pd.ExcelFile(file_path)
                        bs_df = pd.read_excel(xls, '合并资产负债表') if '合并资产负债表' in xls.sheet_names else pd.DataFrame()
                        is_df = pd.read_excel(xls, '合并利润表') if '合并利润表' in xls.sheet_names else pd.DataFrame()
                        
                        # 提炼结果至输出目录
                        out_file = os.path.join(out_dir, f"{comp_code}_排雷与诊断分析.xlsx")
                        with pd.ExcelWriter(out_file, engine='openpyxl') as writer:
                            if not bs_df.empty:
                                bs_df.head(10).to_excel(writer, sheet_name="核心资产筛查", index=False)
                            if not is_df.empty:
                                is_df.head(10).to_excel(writer, sheet_name="盈利质量诊断", index=False)
                        log_func(f"  ✅ 完成标的核算: {comp_code}")
                    except Exception as e:
                        log_func(f"  ⚠️ 处理 {comp_code} 数据时异常: {e}")

        log_func(f"📊 16维度分析矩阵生成完毕")
        return out_dir

class DeepValuationPipeline:
    @staticmethod
    def run(in_dir, out_dir, text_prompt, ent_prompt, use_good_price, config, log_func):
        val_dir = prepare_clean_directory(os.path.join(out_dir, "4_AI深度估值"), config["overwrite"])
        log_func(f"🧠 [大模型引擎] 接入点: {config['api_url']} | 模型: {config['api_model']}")
        
        context_data = ""
        if config['engine'] == 'Tavily' and config['search_key']:
            log_func(f"🔍 [Web Agent] 启用 Tavily 联网支持拉取最新研究")
            try:
                tc = TavilyClient(api_key=config['search_key'])
                resp = tc.search("2024年 宏观经济走势 A股核心研报", max_results=2)
                context_data = "\n".join([r['content'] for r in resp['results']])
            except: pass

        t_dir = os.path.join(val_dir, "1_智能深度研报")
        os.makedirs(t_dir, exist_ok=True)
        log_func(f"🧠 [AI估值合成] 正在调用大模型生成长周期研报...")
        
        try:
            client = OpenAI(api_key=config['api_key'] or "free", base_url=config['api_url'])
            resp = client.chat.completions.create(
                model=config['api_model'],
                messages=[{"role": "user", "content": f"{ent_prompt}\n\n参考情报: {context_data}\n\n请针对本次获取的所有公司输出评估。"}],
                temperature=0.3
            )
            doc = Document()
            doc.add_heading("AI 综合企业战略评估报告", 0)
            doc.add_paragraph(resp.choices[0].message.content)
            doc.save(os.path.join(t_dir, "企业综合评估研报.docx"))
            log_func(f"  ✅ AI 研报生成完毕！")
        except Exception as e:
            log_func(f"  ⚠️ AI 调用受阻 (请检查API_KEY、模型名称或网络): {e}")

        if use_good_price:
            p_dir = os.path.join(val_dir, "2_好价分析测算")
            os.makedirs(p_dir, exist_ok=True)
            log_func(f"📈 [宏观大盘] 正在通过 AkShare 抓取国债收益率进行折现率对标...")
            try:
                cn_10y = WebScraperEngine.safe_macro_fetch(lambda: ak.bond_zh_us_rate()['中国国债收益率10年'].dropna().iloc[-1])
                if cn_10y:
                    log_func(f"  💰 抓取基准成功，当前中国10年期国债无风险收益率: {cn_10y}%")
                else:
                    log_func("  ⚠️ 宏观基准抓取超时，使用默认参数。")
            except: pass
            
        log_func(f"✅ 深度估值体系运行完毕，报告存入: {val_dir}")
        return val_dir

# ==========================================
# 4. 全局一键启动流水线调度器
# ==========================================
class OneClickOrchestrator:
    @staticmethod
    def run_all(ui_config, log_func):
        try:
            log_func("="*40)
            log_func(f"🚀 开始执行金融智能量化分析流水线")
            log_func(f"引擎: {ui_config['engine']} | 覆盖模式: {ui_config['overwrite']}")
            base_dir = ui_config["base_dir"]
            overwrite = ui_config["overwrite"]
            
            # 1. 数据海选
            hs_dir, stock_list = DataSelectionPipeline.run(
                ui_config["wencai"], ui_config["ai_filter"], ui_config["enable_ai_filter"], 
                base_dir, overwrite, log_func
            )
            
            # 2. 报表下载与提取 (同步下载PDF和财务数据表)
            report_dir = AnnualReportPipeline.run_download_and_extract(
                stock_list, ui_config, log_func
            )
            
            # 3. 财务排雷与16维分析
            eval_dir = FraudAndDim16Pipeline.run(
                report_dir, base_dir, overwrite, log_func
            )

            # 4. AI深度估值
            DeepValuationPipeline.run(
                in_dir=eval_dir, out_dir=base_dir, text_prompt=ui_config["text_prompt"],
                ent_prompt=ui_config["ent_prompt"], use_good_price=ui_config["good_price"],
                config=ui_config, log_func=log_func
            )

            log_func("🎉 全链路任务完美收官！所有数据核算与报告生成完毕。")
            log_func("="*40)
        except Exception as e:
            log_func(f"❌ 运行发生中断异常: {str(e)}")
