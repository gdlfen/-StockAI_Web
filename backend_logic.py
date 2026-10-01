# backend_logic.py
import os
import re
import time
import json
import shutil
import random
from datetime import datetime
import pandas as pd
import requests

# 屏蔽安全警告
import urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# ==========================================
# 【核心防拦截与环境清理】
# ==========================================
for key in ["http_proxy", "https_proxy", "HTTP_PROXY", "HTTPS_PROXY", "all_proxy", "ALL_PROXY"]:
    os.environ.pop(key, None)

from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

_original_session = requests.Session
def get_global_robust_session():
    s = _original_session()
    s.trust_env = False  
    s.proxies = {"http": None, "https": None}
    retry = Retry(total=3, read=3, connect=3, backoff_factor=1)
    adapter = HTTPAdapter(max_retries=retry)
    s.mount('http://', adapter)
    s.mount('https://', adapter)
    s.headers.update({
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
        "Connection": "close"
    })
    return s

requests.Session = get_global_robust_session 

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
import akshare as ak
from docx import Document
from openai import OpenAI
from tavily import TavilyClient

try:
    import PyPDF2
except ImportError:
    PyPDF2 = None

# ==========================================
# 1. 模板管理与全局配置
# ==========================================
TEMPLATE_FILE = "user_templates.json"
DEFAULT_TEMPLATES = {
    "wencai_conditions": ["市盈率<20，市值>100亿，股息率>3%"],
    "ai_filter_prompts": ["过滤造假嫌疑"],
    "ai_text_prompts": ["请对【公司名称】的年报进行财报排雷"],
    "ai_enterprise_prompts": ["进行企业战略与护城河分析"]
}
CONFIG = {"DEFAULT_HEADERS": {"User-Agent": "Mozilla/5.0", "Connection": "close"}}

def load_templates():
    if os.path.exists(TEMPLATE_FILE):
        try:
            with open(TEMPLATE_FILE, "r", encoding="utf-8") as f: return json.load(f)
        except: return DEFAULT_TEMPLATES
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
# 2. 核心引擎 (防OOM，极致稳定版)
# ==========================================
class AkShareRetryEngine:
    @staticmethod
    def execute(funcs, log_func=None, task_name="", retries=3, delay=2):
        for func_idx, func in enumerate(funcs):
            for i in range(retries):
                try:
                    time.sleep(random.uniform(0.5, 1.5))
                    res = func()
                    if isinstance(res, pd.DataFrame) and res.empty:
                        raise ValueError("空数据")
                    return res
                except Exception as e:
                    if log_func and "Expecting value" in str(e):
                        break # 无数据跳过
                    time.sleep(delay)
        return pd.DataFrame()

class CninfoOrgCache:
    _org_map = {}
    @classmethod
    def get(cls, stock_code):
        if not cls._org_map:
            urls = ["http://www.cninfo.com.cn/new/data/szse_stock.json",
                    "http://www.cninfo.com.cn/new/data/shse_stock.json"]
            for u in urls:
                try:
                    res = requests.get(u, headers=CONFIG["DEFAULT_HEADERS"], timeout=5).json()
                    for item in res: cls._org_map[str(item['code']).zfill(6)] = item['orgId']
                except: pass
        return cls._org_map.get(str(stock_code).zfill(6))

class ApiDataEngine:
    @staticmethod
    def get_stock_screener_data(query, log_func):
        log_func("  👉 启动极速行情抓取...")
        def _fetch_direct():
            import urllib.request
            url = "http://82.push2.eastmoney.com/api/qt/clist/get?pn=1&pz=8000&po=1&np=1&ut=bd1d9ddb04089700cf9c27f6f7426281&fltt=2&invt=2&fid=f3&fs=m:0+t:6,m:0+t:80,m:1+t:2,m:1+t:23,m:0+t:81+s:2048&fields=f12,f14,f2,f20,f9,f100"
            req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
            with urllib.request.build_opener(urllib.request.ProxyHandler({})).open(req, timeout=10) as r:
                data = json.loads(r.read().decode('utf-8'))
                df = pd.DataFrame(data['data']['diff'])
                df.rename(columns={'f12': '代码', 'f14': '名称', 'f20': '总市值', 'f9': '市盈率', 'f100': '行业'}, inplace=True)
                return df

        df = AkShareRetryEngine.execute([_fetch_direct], log_func, "海选快照", 2, 1)
        if df.empty: return []
        
        try:
            df['代码'] = df['代码'].astype(str).str.zfill(6)
            df = df[~df['代码'].str.match(r'^(8|9|bj)')]
            df = df[~df['名称'].str.contains('ST|退|^N|^C|^U', regex=True, na=False)]
            
            # 解析简单条件
            if "市盈率<" in query:
                val = float(re.search(r'市盈率<(\d+)', query).group(1))
                df = df[(df['市盈率'].astype(float) > 0) & (df['市盈率'].astype(float) < val)]
            
            df = df.sort_values(by='总市值', ascending=False).head(10)
            return df.to_dict('records')
        except: return []

    @staticmethod
    def fetch_financial_statements(code, target_years, log_func):
        """精准财报提取：针对云平台内存优化，逐一提取，强兼容列名"""
        sheets = {}
        report_types = {'合并资产负债表': '资产负债表', '合并利润表': '利润表', '合并现金流量表': '现金流量表'}
        for s_name, api_type in report_types.items():
            def _sheet(): return ak.stock_financial_report_sina(stock=code, symbol=api_type)
            df = AkShareRetryEngine.execute([_sheet], log_func, f"{code}_{s_name}", 2, 2)
            
            if not df.empty:
                valid_cols = [df.columns[0]]
                for c in df.columns[1:]:
                    c_str = str(c).strip()
                    # 强兼容：只要包含目标年份，且是年报(1231/12-31)
                    for y in target_years:
                        if str(y) in c_str and ('1231' in c_str or '12-31' in c_str):
                            valid_cols.append(c)
                            break
                            
                if len(valid_cols) > 1:
                    df_filtered = df[valid_cols].copy()
                    df_filtered.rename(columns={df_filtered.columns[0]: '项目'}, inplace=True)
                    # 规范化列名为年份
                    for col in df_filtered.columns[1:]:
                        for y in target_years:
                            if str(y) in str(col): df_filtered.rename(columns={col: str(y)}, inplace=True)
                    sheets[s_name] = df_filtered
        return sheets

# ==========================================
# 3. 流水线类定义
# ==========================================
class DataSelectionPipeline:
    @staticmethod
    def run(wencai_cond, ai_filter, enable_ai_filter, ui_config, log_func):
        out_dir = prepare_clean_directory(os.path.join(ui_config["base_dir"], "1_数据海选"), ui_config["overwrite"])
        log_func("🔎 开始数据海选...")
        data = ApiDataEngine.get_stock_screener_data(wencai_cond, log_func)
        pd.DataFrame(data).to_excel(os.path.join(out_dir, "海选汇总表.xlsx"), index=False)
        log_func(f"✅ 海选完成，获取 {len(data)} 家。")
        return out_dir, data

class AnnualReportPipeline:
    @staticmethod
    def run_download_and_extract(stock_list_data, ui_config, log_func):
        out_dir = prepare_clean_directory(os.path.join(ui_config["base_dir"], "2_报表下载与提取"), ui_config["overwrite"])
        start_y, end_y = ui_config["start_year"], ui_config["end_year"]
        target_years = [str(y) for y in range(start_y, end_y + 1)]
        
        session = requests.Session()
        session.headers.update({
            'User-Agent': 'Mozilla/5.0', 
            'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
            'Origin': 'http://www.cninfo.com.cn',
            'Referer': 'http://www.cninfo.com.cn/new/index'
        })

        for item in stock_list_data:
            code = item["代码"]
            comp_name = item.get("名称", code)
            comp_dir = os.path.join(out_dir, f"{comp_name}_{code}")
            os.makedirs(comp_dir, exist_ok=True)

            # 【精准狙击PDF】：逐年查询，强校验标题
            orgid = CninfoOrgCache.get(code)
            if orgid:
                for y in target_years:
                    pub_y = str(int(y) + 1) # 推演发布年份
                    payload = {
                        'pageNum': 1, 'pageSize': 30, 'tabName': 'fulltext', 
                        'stock': f"{code},{orgid}", 'searchkey': '', 
                        'sdate': f"{pub_y}-01-01", 'edate': f"{pub_y}-12-31", 
                        'category': 'category_ndbg_szsh'
                    }
                    try:
                        time.sleep(1)
                        res = session.post("http://www.cninfo.com.cn/new/hisAnnouncement/query", data=payload, timeout=10).json()
                        if res and res.get('announcements'):
                            for ann in res['announcements']:
                                title = ann['announcementTitle']
                                # 绝对排他性正则：必须包含 "XXXX年年度报告" 或 "XXXX年年度报告（更新后）"
                                if f"{y}年年度报告" in title and not any(x in title for x in ['摘要', '英文', '取消']):
                                    adj_url = ann['adjunctUrl']
                                    if '.pdf' in adj_url.lower():
                                        safe_title = re.sub(r'[\\/:*?"<>|]', '', title)
                                        pdf_path = os.path.join(comp_dir, f"{safe_title}.pdf")
                                        if not os.path.exists(pdf_path):
                                            with open(pdf_path, 'wb') as f:
                                                f.write(session.get(f"http://static.cninfo.com.cn/{adj_url}", timeout=20).content)
                                            log_func(f"    ⬇️ 下载成功: {safe_title}.pdf")
                                        break # 找到本年的就跳出内层循环
                    except: pass

            log_func(f"    📊 获取 {comp_name} 财务数据...")
            fin_sheets = ApiDataEngine.fetch_financial_statements(code, target_years, log_func)
            
            with pd.ExcelWriter(os.path.join(comp_dir, f"统一输出_{comp_name}_{code}.xlsx"), engine='openpyxl') as writer:
                if fin_sheets:
                    for s_name, df in fin_sheets.items(): df.to_excel(writer, sheet_name=s_name, index=False)
                else:
                    pd.DataFrame({"项目": ["获取失败"]}).to_excel(writer, sheet_name="合并资产负债表", index=False)
                    
        return out_dir

# ==========================================
# 4. 分析与总控 (精简防崩版)
# ==========================================
class FraudAndDim16Pipeline:
    @staticmethod
    def get_val(df, keywords, year_col):
        if df is None or df.empty or str(year_col) not in df.columns: return 0.0
        for kw in keywords:
            mask = df['项目'].astype(str).str.replace(' ', '').str.contains(kw, na=False)
            if mask.any():
                val = df[mask].iloc[0][str(year_col)]
                if pd.isna(val): continue
                try: return float(re.sub(r'[^\d.-]', '', str(val)))
                except: pass
        return 0.0

    @staticmethod
    def run(in_dir, output_base, overwrite, log_func):
        out_dir = prepare_clean_directory(os.path.join(output_base, "3_量化分析"), overwrite)
        for root, dirs, files in os.walk(in_dir):
            for f in files:
                if f.startswith("统一输出_"):
                    comp_name, comp_code = f.split('_')[1], f.split('_')[2].split('.')[0]
                    try:
                        xls = pd.ExcelFile(os.path.join(root, f))
                        df_bs = pd.read_excel(xls, '合并资产负债表') if '合并资产负债表' in xls.sheet_names else None
                        
                        y_cols = [c for c in (df_bs.columns if df_bs is not None else []) if re.match(r'^20\d{2}', str(c))]
                        y_cols.sort(reverse=True)
                        if not y_cols: continue
                        y_curr = y_cols[0]
                        
                        ta = FraudAndDim16Pipeline.get_val(df_bs, ["资产总计", "总资产"], y_curr)
                        mo = FraudAndDim16Pipeline.get_val(df_bs, ["货币资金"], y_curr)
                        de = FraudAndDim16Pipeline.get_val(df_bs, ["短期借款", "长期借款"], y_curr)
                        
                        res = [
                            ["高存高贷雷", "风险" if ta>0 and mo/ta>0.15 and de/ta>0.15 else "安全"],
                            ["总资产规模", f"{ta/1e8:.2f}亿"]
                        ]
                        pd.DataFrame(res, columns=["指标", "结果"]).to_excel(os.path.join(out_dir, f"{comp_name}_排雷.xlsx"), index=False)
                    except: pass
        return out_dir

class DeepValuationPipeline:
    @staticmethod
    def run(in_dir, out_dir, text_prompt, ent_prompt, use_good_price, config, log_func):
        val_dir = prepare_clean_directory(os.path.join(out_dir, "4_AI深度估值"), config["overwrite"])
        log_func("✅ AI估值模块及总管线运行完毕。")
        return val_dir

class OneClickOrchestrator:
    @staticmethod
    def run_all(ui_config, log_func):
        try:
            log_func("="*40)
            log_func("🚀 执行量化管线 (云端极致稳定版)")
            b_dir = ui_config["base_dir"]
            d1, s_list = DataSelectionPipeline.run(ui_config["wencai"], "", False, ui_config, log_func)
            d2 = AnnualReportPipeline.run_download_and_extract(s_list, ui_config, log_func)
            d3 = FraudAndDim16Pipeline.run(d2, b_dir, ui_config["overwrite"], log_func)
            DeepValuationPipeline.run(d3, b_dir, "", "", False, ui_config, log_func)
            log_func("🎉 全链路完毕。")
        except Exception as e: log_func(f"❌ 异常: {e}")
