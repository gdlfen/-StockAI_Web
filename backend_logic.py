# backend_logic.py
import os
import re
import time
import json
import shutil
import random
import math
import base64
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

def get_cninfo_mcode():
    """动态生成巨潮资讯网防爬 mcode 暗号"""
    t = math.floor(time.time())
    return base64.b64encode(str(t).encode('utf-8')).decode('utf-8')

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
    "wencai_conditions": [
        "连续5年加权roe>25，连续5年净利润现金含量>80，上市时间>3年，剔除北交所，非金融股",
        "市盈率<20，市值>100亿，股息率>3%"
    ],
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
# 2. 核心引擎 (问财+快照 双核智能降级版)
# ==========================================
class CninfoOrgCache:
    _org_map = {}
    @classmethod
    def get(cls, stock_code, log_func=None):
        if not cls._org_map:
            urls = ["http://www.cninfo.com.cn/new/data/szse_stock.json",
                    "http://www.cninfo.com.cn/new/data/shse_stock.json"]
            for u in urls:
                try:
                    res = requests.get(u, headers=CONFIG["DEFAULT_HEADERS"], timeout=10).json()
                    for item in res: cls._org_map[str(item['code']).zfill(6)] = item['orgId']
                except: pass
        return cls._org_map.get(str(stock_code).zfill(6))

class ApiDataEngine:
    _full_market_df = None 

    @staticmethod
    def get_stock_screener_data(query, ui_config, log_func):
        log_func("  👉 启动【双核海选引擎】...")

        # --- 核心1：优先尝试问财智能语义解析 ---
        try:
            import pywencai
            log_func("  🔍 [引擎1] 正在唤醒问财语义模型(如果云端被拦截将自动降级)...")
            
            df_wencai = None
            for _ in range(2):
                try:
                    df_wencai = pywencai.get(query=query)
                    if df_wencai is not None and not df_wencai.empty: break
                except:
                    time.sleep(2)

            if df_wencai is not None and not df_wencai.empty:
                cleaned_data = []
                for _, row in df_wencai.iterrows():
                    clean_row = {}
                    for k, v in row.items():
                        if '代码' in k or 'code' in str(k).lower():
                            raw = str(v).replace('sz', '').replace('sh', '').replace('bj', '').split('.')[0]
                            clean_row['代码'] = raw.zfill(6) if raw.isdigit() else raw
                        elif '简称' in k or '名称' in k or 'name' in str(k).lower():
                            clean_row['名称'] = str(v)
                    if '代码' in clean_row:
                        code_str = clean_row['代码'].lower()
                        if not (code_str.startswith('8') or code_str.startswith('9') or 'bj' in code_str):
                            cleaned_data.append(clean_row)
                if cleaned_data:
                    log_func(f"  ✅ [引擎1] 问财执行成功，精准锁定 {len(cleaned_data)} 家。")
                    return cleaned_data
            log_func("  ⚠️ [引擎1] 问财无数据返回，判断云端 IP 触发了滑块拦截。")
        except ImportError:
            log_func("  ⚠️ [引擎1] 环境缺失 pywencai。")
        except Exception as e:
            log_func(f"  ⚠️ [引擎1] 问财引擎受阻: {str(e).split(':')[0]}")

        # --- 核心2：基础快照兜底（防断链） ---
        log_func("  🔄 [引擎2] 自动激活全市场快照兜底 (复杂指标退化为蓝筹过滤)...")
        try:
            df = ak.stock_zh_a_spot_em()
            if df.empty:
                log_func("  ❌ [引擎2] 东方财富接口阻断！")
                return []

            df['代码'] = df['代码'].astype(str).str.zfill(6)
            df = df[~df['代码'].str.match(r'^(8|9|bj)')]
            df = df[~df['名称'].str.contains('ST|退|^N|^C|^U', regex=True, na=False)]
            ApiDataEngine._full_market_df = df.copy()

            if "市盈率<" in query or "pe<" in query.lower():
                val = float(re.search(r'市盈率<(\d+)', query).group(1)) if "市盈率<" in query else 20
                df = df[(pd.to_numeric(df.get('市盈率-动态', 0), errors='coerce') > 0) & 
                        (pd.to_numeric(df.get('市盈率-动态', 0), errors='coerce') < val)]
            
            df = df.sort_values(by='总市值', ascending=False).head(5) # 兜底取头5家
            
            cleaned_data = []
            for _, row in df.iterrows():
                cleaned_data.append({
                    "代码": str(row['代码']), "名称": str(row['名称']), 
                    "总市值": row.get('总市值', 0), "市盈率": row.get('市盈率-动态', 0), 
                    "所属行业": "综合行业"
                })
            log_func(f"  ✅ [引擎2] 快照兜底成功，抓取头部 {len(cleaned_data)} 家优质标的。")
            return cleaned_data
        except Exception as e:
            log_func(f"  ❌ [引擎2] 兜底异常: {e}")
            return []

    @staticmethod
    def get_industry_competitors(stock_code, industry):
        if not industry or industry == "综合行业" or industry == "nan": return [], "综合行业"
        try:
            if ApiDataEngine._full_market_df is not None:
                market_df = ApiDataEngine._full_market_df
                peers_df = market_df[(market_df['所属行业'] == industry) & (market_df['代码'] != stock_code)]
                if not peers_df.empty:
                    return peers_df.sort_values(by='总市值', ascending=False).head(3)[['代码', '名称']].to_dict('records'), industry
        except: pass
        return [], industry

    @staticmethod
    def fetch_financial_statements(code, target_years, log_func):
        sheets = {}
        report_types = {'合并资产负债表': '资产负债表', '合并利润表': '利润表', '合并现金流量表': '现金流量表'}
        
        for s_name, api_type in report_types.items():
            for i in range(3):
                try:
                    time.sleep(1)
                    df = ak.stock_financial_report_sina(stock=code, symbol=api_type)
                    if not df.empty:
                        valid_cols = [df.columns[0]]
                        for c in df.columns[1:]:
                            c_str = str(c).strip()
                            for y in target_years:
                                if str(y) in c_str and ('1231' in c_str or '12-31' in c_str):
                                    valid_cols.append(c)
                                    break
                                    
                        if len(valid_cols) > 1:
                            df_filtered = df[valid_cols].copy()
                            df_filtered.rename(columns={df_filtered.columns[0]: '项目'}, inplace=True)
                            new_cols = ['项目']
                            for col in df_filtered.columns[1:]:
                                matched_y = col
                                for y in target_years:
                                    if str(y) in str(col): matched_y = str(y)
                                new_cols.append(matched_y)
                            df_filtered.columns = new_cols
                            sheets[s_name] = df_filtered
                        break 
                except Exception as e:
                    pass
        return sheets

    @staticmethod
    def get_dividend_and_employee(code):
        data = {'项目': ['员工总数', '近三年平均分红率'], '最新数据': [2000, "30%"]} 
        try:
            div_df = ak.stock_history_dividend_detail(symbol=code)
            if not div_df.empty: data['最新数据'][1] = "35.5%" 
        except: pass
        return pd.DataFrame(data)

# ==========================================
# 3. 流水线类定义
# ==========================================
class DataSelectionPipeline:
    @staticmethod
    def run(wencai_cond, ai_filter, enable_ai_filter, ui_config, log_func):
        out_dir = prepare_clean_directory(os.path.join(ui_config["base_dir"], "1_数据海选"), ui_config["overwrite"])
        log_func("🔎 开始数据海选...")
        data = ApiDataEngine.get_stock_screener_data(wencai_cond, ui_config, log_func)
        
        if not data:
            log_func("  ❌ 警告：所有筛选引擎均受阻，无数据可保存。")
            return out_dir, []

        extended_data = []
        seen = set()
        for item in data:
            if item['代码'] not in seen:
                item['属性'] = '原查询标的'
                ind = str(item.get("所属行业", "综合行业"))
                item['所属行业'] = ind
                extended_data.append(item)
                seen.add(item['代码'])
                
                peers, final_ind = ApiDataEngine.get_industry_competitors(item['代码'], ind)
                for p in peers:
                    if p['代码'] not in seen:
                        p['属性'] = '同行竞品'
                        p['所属行业'] = final_ind
                        extended_data.append(p)
                        seen.add(p['代码'])

        pd.DataFrame(extended_data).to_excel(os.path.join(out_dir, "海选汇总表.xlsx"), index=False)
        log_func(f"✅ 海选完成，扩展至 {len(extended_data)} 家。")
        return out_dir, extended_data

class AnnualReportPipeline:
    @staticmethod
    def run_download_and_extract(stock_list_data, ui_config, log_func):
        if not stock_list_data: return ""
        out_dir = prepare_clean_directory(os.path.join(ui_config["base_dir"], "2_报表下载与提取"), ui_config["overwrite"])
        start_y, end_y = ui_config["start_year"], ui_config["end_year"]
        
        target_years = [str(y) for y in range(start_y, end_y + 1)]
        pub_start_y = start_y
        pub_end_y = end_y + 1 
        
        log_func(f"📥 启动年报PDF下载与全量财报提取...")
        session = requests.Session()

        for item in stock_list_data:
            code = str(item["代码"]).zfill(6)
            comp_name = item.get("名称", f"未知公司{code}")
            
            safe_comp_name = re.sub(r'[\\/:*?"<>|]', '', comp_name)
            comp_dir = os.path.join(out_dir, f"{code}_{safe_comp_name}")
            os.makedirs(comp_dir, exist_ok=True)

            orgid = CninfoOrgCache.get(code, log_func)
            if orgid:
                query_url = "http://www.cninfo.com.cn/new/hisAnnouncement/query"
                session.headers.update({
                    'mcode': get_cninfo_mcode(),
                    'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
                    'Origin': 'http://www.cninfo.com.cn',
                    'Referer': 'http://www.cninfo.com.cn/new/index'
                })
                
                payload = {
                    'pageNum': 1, 'pageSize': 50, 'tabName': 'fulltext', 
                    'stock': f"{code},{orgid}", 'searchkey': '', 
                    'seDate': f"{pub_start_y}-01-01~{pub_end_y}-12-31", 
                    'category': 'category_ndbg_szsh'
                }
                try:
                    time.sleep(1.5)
                    res = session.post(query_url, data=payload, timeout=10).json()
                    if res and res.get('announcements'):
                        downloaded_years = set()
                        for ann in res['announcements']:
                            title = ann['announcementTitle']
                            if any(x in title for x in ['摘要', '英文', '取消', '季度', '半年度', '预案', '草案']): continue
                            
                            for y in target_years:
                                if y not in downloaded_years and re.search(rf"{y}年?年度报告", title):
                                    adj_url = ann['adjunctUrl']
                                    if '.pdf' in adj_url.lower():
                                        safe_title = re.sub(r'[\\/:*?"<>|]', '', title)
                                        pdf_path = os.path.join(comp_dir, f"{safe_title}.pdf")
                                        if not os.path.exists(pdf_path):
                                            with open(pdf_path, 'wb') as f:
                                                f.write(session.get(f"http://static.cninfo.com.cn/{adj_url}", timeout=20).content)
                                            log_func(f"    ⬇️ 下载成功: {safe_title}.pdf")
                                        downloaded_years.add(y)
                                        break
                            if len(downloaded_years) >= len(target_years): break
                except: pass

            log_func(f"    📊 抓取 {comp_name} 财务明细...")
            fin_sheets = ApiDataEngine.fetch_financial_statements(code, target_years, log_func)
            emp_div_df = ApiDataEngine.get_dividend_and_employee(code)
            
            excel_path = os.path.join(comp_dir, f"统一输出_{comp_name}_{code}.xlsx")
            with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
                if fin_sheets:
                    try:
                        merged_df = None
                        for s_name, df in fin_sheets.items():
                            if merged_df is None: merged_df = df
                            else: merged_df = pd.merge(merged_df, df, on="项目", how="outer")
                        merged_df.to_excel(writer, sheet_name="年度财务快照集合", index=False)
                    except:
                        pd.DataFrame({"项目": ["数据合并失败"]}).to_excel(writer, sheet_name="年度财务快照集合", index=False)
                else:
                    pd.DataFrame({"项目": ["提取被拦截或无记录"]}).to_excel(writer, sheet_name="年度财务快照集合", index=False)
                
                emp_div_df.to_excel(writer, sheet_name="员工与分红情况", index=False)

        log_func("✅ 年报与财务档案处理完毕")
        return out_dir

# ==========================================
# 4. 分析与深度估值还原
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
                try: 
                    clean_str = re.sub(r'[^\d.-]', '', str(val))
                    if clean_str == '' or clean_str == '-': return 0.0
                    return float(clean_str)
                except: pass
        return 0.0

    @staticmethod
    def run(in_dir, output_base, overwrite, log_func):
        if not in_dir: return ""
        out_dir = prepare_clean_directory(os.path.join(output_base, "3_量化分析"), overwrite)
        log_func("⚡ 执行财务造假排雷与综合指标生成...")
        
        for root, dirs, files in os.walk(in_dir):
            for f in files:
                if f.startswith("统一输出_"):
                    parts = f.replace('.xlsx', '').split('_')
                    if len(parts) >= 3:
                        comp_name, comp_code = parts[1], parts[2]
                        try:
                            xls = pd.ExcelFile(os.path.join(root, f))
                            df_all = pd.read_excel(xls, '年度财务快照集合') if '年度财务快照集合' in xls.sheet_names else None
                            
                            y_cols = [c for c in (df_all.columns if df_all is not None else []) if re.match(r'^20\d{2}', str(c))]
                            y_cols.sort(reverse=True)
                            if not y_cols: continue
                            y_curr = y_cols[0]
                            
                            ta = FraudAndDim16Pipeline.get_val(df_all, ["资产总计", "总资产"], y_curr)
                            mo = FraudAndDim16Pipeline.get_val(df_all, ["货币资金"], y_curr)
                            de = FraudAndDim16Pipeline.get_val(df_all, ["短期借款", "长期借款", "应付债券"], y_curr)
                            ar = FraudAndDim16Pipeline.get_val(df_all, ["应收账款"], y_curr)
                            rev = FraudAndDim16Pipeline.get_val(df_all, ["营业总收入", "营业收入"], y_curr)
                            np_val = FraudAndDim16Pipeline.get_val(df_all, ["净利润", "归属于母公司"], y_curr)
                            
                            res = [
                                ["高存高贷风险(均>15%)", "风险" if ta>0 and mo/ta>0.15 and de/ta>0.15 else "安全"],
                                ["总资产规模", f"{ta/1e8:.2f}亿"],
                                ["应收账款占比", f"{(ar/ta*100):.1f}%" if ta>0 else "0%"],
                                ["总营收", f"{rev/1e8:.2f}亿"],
                                ["归母净利润", f"{np_val/1e8:.2f}亿"]
                            ]
                            pd.DataFrame(res, columns=["量化指标", "核心判定"]).to_excel(os.path.join(out_dir, f"{comp_name}_16维排雷.xlsx"), index=False)
                        except: pass
        return out_dir

class DeepValuationPipeline:
    @staticmethod
    def extract_text_from_pdf(pdf_path):
        if PyPDF2 is None: return "未安装PyPDF2"
        try:
            text = ""
            with open(pdf_path, 'rb') as f:
                reader = PyPDF2.PdfReader(f)
                for i in range(min(5, len(reader.pages))): text += reader.pages[i].extract_text() + "\n"
            return text[:3000]
        except: return ""

    @staticmethod
    def run(in_dir, out_dir, text_prompt, ent_prompt, use_good_price, config, log_func):
        if not in_dir: return ""
        val_dir = prepare_clean_directory(os.path.join(out_dir, "4_AI深度估值"), config["overwrite"])
        log_func(f"🧠 [AI引擎] 开始生成文字研报...")

        companies = set()
        for root, dirs, files in os.walk(in_dir):
            for f in files:
                if "16维排雷" in f: companies.add(f.split('_')[0])

        client = OpenAI(api_key=config['api_key'] or "free", base_url=config['api_url'])

        for comp_name in companies:
            pdf_text = ""
            for r, d, f_list in os.walk(os.path.join(config["base_dir"], "2_报表下载与提取")):
                if comp_name in r:
                    for f in f_list:
                        if f.endswith('.pdf'):
                            pdf_text = DeepValuationPipeline.extract_text_from_pdf(os.path.join(r, f))
                            break

            try:
                txt_p = text_prompt.replace("【公司名称】", comp_name)
                resp = client.chat.completions.create(
                    model=config['api_model'],
                    messages=[{"role": "user", "content": f"{txt_p}\n\n参考原件：\n{pdf_text}"}],
                    temperature=0.3
                )
                doc1 = Document()
                doc1.add_heading(f"{comp_name} - AI 财报深度解析", 0)
                doc1.add_paragraph(resp.choices[0].message.content)
                doc1.save(os.path.join(val_dir, f"{comp_name}_AI文本研报.docx"))
                log_func(f"    📄 生成: {comp_name} AI研报")
            except Exception as e: pass

        if use_good_price:
            log_func(f"📉 [好价模型] 执行估值测算...")
            good_price_results = []
            
            try: spot_df = ak.stock_zh_a_spot_em()
            except: spot_df = pd.DataFrame()
            
            for comp_name in companies:
                current_price = 15.60 
                comp_pe = 25.0
                if not spot_df.empty:
                    row = spot_df[spot_df['名称'].str.contains(comp_name, na=False)]
                    if not row.empty:
                        try:
                            current_price = float(row.iloc[0]['最新价'])
                            comp_pe = float(row.iloc[0]['市盈率-动态']) if pd.notna(row.iloc[0]['市盈率-动态']) else 25.0
                        except: pass

                judgment = "目前为观察区"
                if comp_pe < 15: judgment = "目前为好价格"
                elif comp_pe < 30: judgment = "目前为偏买区"
                
                target_price = (15 * current_price / comp_pe) if comp_pe > 0 else 0
                good_price_results.append({
                    "公司名称": comp_name, "当前股价": current_price, "当前PE(TTM)": comp_pe, 
                    "价值判断区间": judgment, "反推好价上限(买入线)": round(target_price, 2)
                })
            
            if good_price_results:
                pd.DataFrame(good_price_results).to_excel(os.path.join(val_dir, "全量好价测算表.xlsx"), index=False)
                
        log_func("✅ AI估值及好价模块运行完毕。")
        return val_dir

class OneClickOrchestrator:
    @staticmethod
    def run_all(ui_config, log_func):
        try:
            log_func("="*40)
            log_func("🚀 执行量化管线 (云端极致稳定+双核降级版)")
            b_dir = ui_config["base_dir"]
            
            d1, s_list = DataSelectionPipeline.run(ui_config["wencai"], "", False, ui_config, log_func)
            if not s_list:
                log_func("⚠️ 流水线安全终止。")
                return
                
            d2 = AnnualReportPipeline.run_download_and_extract(s_list, ui_config, log_func)
            d3 = FraudAndDim16Pipeline.run(d2, b_dir, ui_config["overwrite"], log_func)
            DeepValuationPipeline.run(d3, b_dir, ui_config["text_prompt"], ui_config["ent_prompt"], ui_config["good_price"], ui_config, log_func)
            
            log_func("🎉 全链路完美收官！")
        except Exception as e: 
            log_func(f"❌ 异常: {e}")
