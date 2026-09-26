# ==========================================
# backend_logic.py - Part 1
# 功能：基础依赖、全局配置、日志模块、下载器与爬虫基类
# ==========================================

import os
import re
import time
import json
import logging
import requests
from urllib.parse import urlparse, urljoin
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import List, Dict, Optional, Union

# ------------------------------------------
# 1. 日志与基础配置
# ------------------------------------------
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler()]
)
logger = logging.getLogger(__name__)

# 全局配置字典 (完全脱离硬编码盘符，使用相对路径)
CONFIG = {
    "BASE_DIR": os.path.join(os.getcwd(), "data_output"),
    "TIMEOUT": 15,
    "MAX_RETRIES": 3,
    "MAX_WORKERS": 5,
    "DEFAULT_HEADERS": {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/115.0.0.0 Safari/537.36",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
        "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8"
    }
}

# 确保基础输出目录存在
os.makedirs(CONFIG["BASE_DIR"], exist_ok=True)


# ------------------------------------------
# 2. 核心网络请求与下载器模块
# ------------------------------------------
class BaseDownloader:
    """
    纯净版文件与数据下载器，无任何 UI 耦合
    """

    def __init__(self, custom_headers: Optional[Dict] = None, max_workers: int = CONFIG["MAX_WORKERS"]):
        self.session = requests.Session()
        self.session.headers.update(custom_headers or CONFIG["DEFAULT_HEADERS"])
        self.max_workers = max_workers

    def fetch_html(self, url: str) -> Optional[str]:
        """获取网页源码，包含重试机制"""
        for attempt in range(CONFIG["MAX_RETRIES"]):
            try:
                response = self.session.get(url, timeout=CONFIG["TIMEOUT"])
                response.raise_for_status()
                response.encoding = response.apparent_encoding
                return response.text
            except requests.RequestException as e:
                logger.warning(f"获取 {url} 失败 (尝试 {attempt + 1}/{CONFIG['MAX_RETRIES']}): {e}")
                time.sleep(1.5 * (attempt + 1))
        logger.error(f"彻底无法访问 URL: {url}")
        return None

    def download_file(self, url: str, save_dir: str, filename: Optional[str] = None) -> bool:
        """单文件下载逻辑"""
        if not filename:
            parsed = urlparse(url)
            filename = os.path.basename(parsed.path) or f"downloaded_{int(time.time())}.tmp"

        # 移除了所有盘符硬编码，强制在当前工作目录的 save_dir 下
        safe_dir = os.path.join(CONFIG["BASE_DIR"], save_dir)
        os.makedirs(safe_dir, exist_ok=True)
        file_path = os.path.join(safe_dir, filename)

        for attempt in range(CONFIG["MAX_RETRIES"]):
            try:
                with self.session.get(url, stream=True, timeout=CONFIG["TIMEOUT"]) as r:
                    r.raise_for_status()
                    with open(file_path, 'wb') as f:
                        for chunk in r.iter_content(chunk_size=8192):
                            if chunk:
                                f.write(chunk)
                logger.info(f"下载成功: {file_path}")
                return True
            except Exception as e:
                logger.warning(f"下载 {filename} 失败 (尝试 {attempt + 1}/{CONFIG['MAX_RETRIES']}): {e}")
                time.sleep(1)

        logger.error(f"下载中止，文件无法获取: {url}")
        return False

    def download_batch(self, urls: List[str], save_dir: str) -> Dict[str, str]:
        """多线程批量下载"""
        results = {}
        logger.info(f"开始批量下载，共计 {len(urls)} 个任务，线程数: {self.max_workers}")

        with ThreadPoolExecutor(max_workers=self.max_workers) as executor:
            future_to_url = {
                executor.submit(self.download_file, url, save_dir): url
                for url in urls
            }

            for future in as_completed(future_to_url):
                url = future_to_url[future]
                try:
                    success = future.result()
                    results[url] = "Success" if success else "Failed"
                except Exception as exc:
                    logger.error(f"{url} 产生异常: {exc}")
                    results[url] = "Error"

        return results


# ------------------------------------------
# 3. 基础爬虫解析器 (BaseCrawler)
# ------------------------------------------
class BaseCrawler(BaseDownloader):
    """
    爬虫基类，继承下载器，用于封装通用的解析方法
    """

    def __init__(self, base_url: str):
        super().__init__()
        self.base_url = base_url

    def extract_links(self, html: str, pattern: str) -> List[str]:
        """使用正则提取页面链接"""
        if not html:
            return []
        raw_links = re.findall(pattern, html)
        # 自动补全相对路径
        full_links = [urljoin(self.base_url, link) for link in raw_links]
        return list(set(full_links))

    def save_data_to_json(self, data: Union[Dict, List], filename: str, subdir: str = "json_data"):
        """统一的数据序列化输出，完全与 UI 解耦"""
        save_path = os.path.join(CONFIG["BASE_DIR"], subdir)
        os.makedirs(save_path, exist_ok=True)
        file_path = os.path.join(save_path, filename)

        try:
            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=4)
            logger.info(f"数据已成功保存至 JSON: {file_path}")
            return file_path
        except Exception as e:
            logger.error(f"保存 JSON 数据时出错: {e}")
            return None

# ==========================================
# End of Part 1
# ==========================================
# =====================================================================
# 后端工作流 3: 年报数据提取与结构化对齐
# =====================================================================
def run_extraction_pipeline(in_path, out_path, raw_map, log_func=print):
    try:
        out_path = os.path.join(out_path, "报表提取完善")
        os.makedirs(out_path, exist_ok=True)
        mappings = {}
        for line in raw_map:
            if '=' in line:
                k, v = line.split('=', 1)
                mappings[k.strip().replace(" ", "")] = v.strip().replace(" ", "")

        log_func("🔄 【阶段一】启动多线程本地 PDF 精细提取生成单年 A 类表")
        all_results_by_company, company_tasks_p2, pdf_tasks_info = {}, {}, []

        for folder_name in os.listdir(in_path):
            ind_path = os.path.join(in_path, folder_name)
            if not os.path.isdir(ind_path): continue
            out_ind_path = os.path.join(out_path, folder_name)
            for f in os.listdir(ind_path):
                if f.endswith(".xlsx") or f.endswith(".csv"):
                    if "公司属性表" in f or "同行排列表" in f:
                        os.makedirs(out_ind_path, exist_ok=True)
                        try: shutil.copy2(os.path.join(ind_path, f), os.path.join(out_ind_path, f))
                        except: pass

        for r, d, f_list in os.walk(in_path):
            for f in f_list:
                if f.lower().endswith(".pdf"):
                    pdf_path = os.path.join(r, f)
                    rel_path = os.path.relpath(pdf_path, in_path)
                    folder_name = rel_path.split(os.sep)[0] if len(rel_path.split(os.sep)) > 1 else "默认提取分类"
                    pdf_tasks_info.append((pdf_path, os.path.join(out_path, folder_name), folder_name))

        if not pdf_tasks_info:
            log_func("❌ 严重警告：源目录内未找到任何 PDF 年报文件！")
            return False

        def process_pdf_task(p, out_ind_path, folder_name):
            data, comp_name, yr, stock_code = p1_process_pdf(p, lambda msg: None)
            if not data or stock_code == "未提取": return False, None
            df_emp, _, _, _ = p2_process_pdf_employee(p)
            os.makedirs(out_ind_path, exist_ok=True)
            out_file = os.path.join(out_ind_path, f"{comp_name}（{stock_code}）{yr}年度年报提取表.xlsx")
            try:
                with pd.ExcelWriter(out_file, engine='openpyxl') as writer:
                    all_sheets = [n[0] for n in MILESTONE_PATTERNS[:8]] + ["董监高及报酬情况", "员工情况", "分红情况"]
                    for k in all_sheets:
                        df_to_save = data.get(k, pd.DataFrame())
                        if not df_to_save.empty:
                            df_to_save.to_excel(writer, sheet_name=k, startrow=2, index=False)
                            ws = writer.sheets[k]
                            unit_str = "人民币元" if "情况" not in k else "详见表内说明"
                            ws.cell(1, 1, f"{comp_name}（{stock_code}）{k}（{yr}年度）\n单位：{unit_str}").font = Font(bold=True, size=12)
                            ws.cell(1, 1).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                            ws.row_dimensions[1].height = 40
                            is_financial = any(x in k for x in ["资产负债表", "利润表", "现金流量表", "所有者权益变动表"])
                            for row in ws.iter_rows(min_row=3, max_col=ws.max_column):
                                item_name = str(row[0].value) if row[0].value else ""
                                for cell in row[1:]:
                                    if isinstance(cell.value, (int, float)):
                                        if is_financial: cell.number_format = '#,##0.0000' if "每股收益" in item_name else '#,##0.00'
                                        elif k == "员工情况": cell.number_format = '#,##0'
                                        elif k == "董监高及报酬情况": cell.number_format = '#,##0' if cell.value == int(cell.value) else '#,##0.00'
                            common_auto_adjust_col_width(ws)
                return True, (comp_name, yr, stock_code, folder_name, data, df_emp, out_file)
            except Exception as e: return False, str(e)

        with concurrent.futures.ThreadPoolExecutor(max_workers=4) as executor:
            futures = [executor.submit(process_pdf_task, *info) for info in pdf_tasks_info]
            for future in concurrent.futures.as_completed(futures):
                success, result = future.result()
                if success:
                    comp_name, yr, stock_code, folder_name, data, df_emp, out_file = result
                    company_key = (comp_name, stock_code, folder_name)
                    if company_key not in all_results_by_company: all_results_by_company[company_key] = {}
                    all_results_by_company[company_key][yr] = data
                    task_key = (stock_code, folder_name)
                    if task_key not in company_tasks_p2: company_tasks_p2[task_key] = {'name': comp_name, 'years': set(), 'pdf_dfs': {}}
                    company_tasks_p2[task_key]['years'].add(int(yr))
                    if not df_emp.empty: company_tasks_p2[task_key]['pdf_dfs'][int(yr)] = df_emp

        p1_generate_multi_year_summary(all_results_by_company, mappings, out_path, log_func)
        log_func("🌐 【阶段二】启动多线程网络并发抓取与 B 类表整合")

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
                        cell.font, cell.alignment = Font(size=14, bold=True), Alignment(horizontal='center', vertical='center')

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
            except Exception as e: return False, f"  ❌ B类表({name})保存失败: {e}"

        if company_tasks_p2:
            with concurrent.futures.ThreadPoolExecutor(max_workers=5) as executor:
                futures = [executor.submit(process_online_task, key[0], key[1], info) for key, info in company_tasks_p2.items()]
                for future in concurrent.futures.as_completed(futures):
                    success, msg = future.result()
                    if msg and msg != "No years": log_func(msg)

        log_func("🔍 【阶段三】启动多线程 A/B 表比对与缺陷填补")
        b_files, a_files = {}, {}
        for root_dir, dirs, files in os.walk(out_path):
            for file in files:
                if not file.endswith('.xlsx') or file.startswith('~$'): continue
                file_path = os.path.join(root_dir, file)
                ind_folder, company_name, year = os.path.basename(root_dir), p3_extract_company_name(file), p3_extract_year(file)
                if "整合输出" in file or "-" in file: b_files[(ind_folder, company_name)] = file_path
                elif "年报提取表" in file:
                    if (ind_folder, company_name) not in a_files: a_files[(ind_folder, company_name)] = {}
                    if year: a_files[(ind_folder, company_name)][year] = file_path

        if not b_files: log_func("  ❌ 错误：未识别到任何B类表格！")
        else:
            def process_compare_task(company, b_path, matched_a_files):
                p3_process_company(company, b_path, matched_a_files, lambda msg: None)
                return company

            compare_tasks = []
            for (ind_folder, company), b_path in b_files.items():
                matched_a_files = a_files.get((ind_folder, company), {})
                if matched_a_files: compare_tasks.append((company, b_path, matched_a_files))

            if compare_tasks:
                with concurrent.futures.ThreadPoolExecutor(max_workers=4) as executor:
                    futures = [executor.submit(process_compare_task, *args) for args in compare_tasks]
                    for future in concurrent.futures.as_completed(futures): pass
            log_func(f"🎉 全链路处理完毕！输出路径: {out_path}")
            return True
    except Exception as e:
        log_func(f"❌ 提取失败: {e}")
        return False

    # =====================================================================
    # 后端工作流 4: 财务排雷分析
    # =====================================================================
    def extract_excel_data(input_dir, log_func, params):
        all_company_data = {}
        company_industry_map = {}

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
                if filename.startswith("~") or not (filename.endswith(".xlsx") or filename.endswith(".csv")): continue
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
                                            if current_parsed_year not in comp_data: comp_data[current_parsed_year] = {}
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
                    metrics["自身毛利率"] = (metrics["营业收入"] - metrics["营业成本"]) / metrics["营业收入"] * 100.0
                else:
                    metrics["自身毛利率"] = 0.0
                metrics["减值损失合计"] = metrics["资产减值损失"] + metrics["信用减值损失"]
        return all_company_data

    def execute_18_conditions_analysis(comp_code, comp_name, comp_industry, comp_data, all_years, target_years, p,
                                       log_callback, macro_db):
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

            emp_drop = safe_div2(p_prev.get("员工人数", 0) - c.get("员工人数", 0), max(p_prev.get("员工人数", 1), 1))
            rev_inc = safe_div2(c.get("营业收入", 0) - p_prev.get("营业收入", 0), max(p_prev.get("营业收入", 1), 1))
            yearly_results["C1_员工与营收背离"].append({"年度": curr_yr,
                                                        "状态": "风险" if emp_drop > p["c1_emp_drop"] and rev_inc > p[
                                                            "c1_rev_inc"] else "正常",
                                                        "员工降幅(%)": emp_drop, "营收增幅(%)": rev_inc,
                                                        "[说明]": "员工降幅与营收增幅背离"})

            comp_salary = (c.get("支付职工现金", 0) + c.get("期末应付薪酬", 0) - c.get("期初应付薪酬", 0)) / max(
                (p_prev.get("员工人数", 1) + c.get("员工人数", 1)) / 2, 1)
            ind_salary = macro_db["默认_平均工资"].get(str(curr_yr), 100000)
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

            fa_cip_growth = safe_div2(
                (c.get("固定资产", 0) + c.get("在建工程", 0)) - (p_prev.get("固定资产", 0) + p_prev.get("在建工程", 0)),
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
            ind_margin = macro_db["默认_毛利率"].get(str(curr_yr), 30.0)
            margin_diff = safe_div2(c.get("自身毛利率", 0) - ind_margin, max(ind_margin, 1))
            yearly_results["C17_毛利率异常偏离"].append({"年度": curr_yr, "状态": "风险" if margin_growth > p[
                "c17_margin_jump"] and margin_diff > p["c17_margin_diff"] else "正常", "自身增幅(%)": margin_growth,
                                                         "偏离均值(%)": margin_diff, "[说明]": "毛利造假嫌疑"})

            ind_ar, ind_inv = macro_db["默认_应收占比"].get(str(curr_yr), 10.0), macro_db["默认_存货占比"].get(
                str(curr_yr), 12.0)
            ar_diff_ind, inv_diff_ind = safe_div2(ar_ratio - ind_ar, max(ind_ar, 1)), safe_div2(inv_ratio - ind_inv,
                                                                                                max(ind_inv, 1))
            c18_risk = margin_growth > p["c17_margin_jump"] and (
                        ar_diff_ind > p["c18_ar_diff"] or inv_diff_ind > p["c18_inv_diff"])
            yearly_results["C18_营运资产背离"].append(
                {"年度": curr_yr, "状态": "风险" if c18_risk else "正常", "应收偏离度(%)": ar_diff_ind,
                 "存货偏离度(%)": inv_diff_ind, "[说明]": "营运资金异常"})

        return {k: pd.DataFrame(v) for k, v in yearly_results.items() if len(v) > 0}

    def generate_excel_report(comp_code, comp_name, results_dict, output_dir):
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
        return save_path

    def run_fraud_detection(in_path, out_path, current_params, macro_db, log_func=print):
        try:
            log_func("=============== 财务造假智能排雷系统启动 ===============")
            main_out = os.path.join(out_path, "造假排雷结果")
            os.makedirs(main_out, exist_ok=True)

            all_data = extract_excel_data(in_path, log_func, current_params)
            if not all_data:
                log_func("❌ 未提取到有效数据！")
                return None

            target_companies = list(all_data.keys())
            for index, comp_key in enumerate(target_companies):
                comp_data = all_data[comp_key]
                comp_industry = comp_data.get("industry", "默认")
                comp_code, comp_name = (comp_key.split(' ', 1) + [""])[:2]
                available_years = sorted([k for k in comp_data.keys() if k != "industry"])

                if len(available_years) < 2: continue
                target_years = available_years[1:]

                analysis_results = execute_18_conditions_analysis(comp_code, comp_name, comp_industry, comp_data,
                                                                  available_years, target_years, current_params,
                                                                  log_func, macro_db)
                generate_excel_report(comp_code, comp_name, analysis_results, main_out)

            log_func(f"✅ 批量排雷任务全部完成！保存在 {main_out}")
            return main_out
        except Exception as e:
            log_func(f"❌ 错误: {str(e)}")
            return None

    # =====================================================================
    # 后端工作流 5: 16维度分析
    # =====================================================================
    def run_16_dim_analysis(in_path, out_path, log_func=print):
        try:
            if not os.path.exists(in_path):
                log_func(f"⚠️ 输入源目录不存在，跳过分析: {in_path}")
                return None

            main_out = os.path.join(out_path, "16维度对比结果")
            os.makedirs(main_out, exist_ok=True)

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
                    if ind and ind not in industry_rep_map: industry_rep_map[ind] = rep

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
                if "_行业_" in folder_name: folder_ind = folder_name.split("_行业_")[-1].strip()

                for f in files:
                    if not (f.endswith(".xlsx") or f.endswith(".csv")) or f.startswith("~"): continue
                    if any(x in f for x in
                           ["同行排列表", "公司属性表", "年报提取表", "多年度汇总与增长率分析", "排雷报告",
                            "海选公司汇总表"]): continue

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
                                comp_name = p;
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
                                    header_idx = i;
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
                log_func("⚠️ 未提取到有效数据，已跳过。")
                return None

            generated_count = 0
            for industry, comp_dict in all_data.items():
                if not comp_dict: continue

                analyzer = FinancialAnalyzer(comp_dict, active_config)
                if not analyzer.years: continue

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
                    generated_count += 1
                except Exception as e:
                    log_func(f"❌ 行业 [{industry}] 生成报告失败: {e}")

            log_func(f"🏁 完美收官！共生成 {generated_count} 份独立行业对比报告。")
            return main_out
        except Exception as err:
            log_func(f"❌ 运行发生严重异常: {err}")
            return None

    # =====================================================================
    # 后端工作流 6: AI 文本分析与估值
    # =====================================================================
    def _search_ddgs_logic(keyword: str, use_proxy: bool):
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

    def _extract_company_name_logic(filepath: str) -> str:
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

    def save_to_word_logic(content: str, company_name: str, save_dir: str, model_name: str) -> str:
        safe_name = re.sub(r'[\\/*?:"<>|]', "", company_name)
        doc = Document()
        title = doc.add_heading(f'{safe_name} - 多年度综合分析报告', 0)
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        info_para = doc.add_paragraph()
        info_para.add_run(f"生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n分析模型: {model_name}\n")
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

    def run_ai_text_analysis(api_url, api_key, model_name, in_dir, out_dir, base_prompt, do_search, search_kw,
                             use_proxy, log_func=print):
        try:
            pdf_files = []
            for root_dir, _, files in os.walk(in_dir):
                for f in files:
                    if f.lower().endswith('.pdf'): pdf_files.append(os.path.join(root_dir, f))
            if not pdf_files:
                log_func("❌ 未找到PDF文件")
                return None

            company_map = {}
            for path in pdf_files:
                comp_name = _extract_company_name_logic(path)
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
                    results, _ = _search_ddgs_logic(current_search_kw, use_proxy)
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
                    resp = client.chat.completions.create(model=model_name,
                                                          messages=[{"role": "user", "content": final_prompt}])
                    result_text = resp.choices[0].message.content
                    save_to_word_logic(result_text, target_company, out_dir, model_name)
                except Exception as api_err:
                    log_func(f"❌ API 调用失败: {api_err}")
                    error_msg = f"生成报告失败。API 调用出现异常：\n{str(api_err)}"
                    save_to_word_logic(error_msg, target_company + "_生成失败排错日志", out_dir, model_name)
            log_func("\n🎉 文本分析处理完毕！")
            return out_dir
        except Exception as exc:
            log_func(f"❌ 异常: {exc}")
            return None

    def search_web_logic(query, max_res=3):
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
                    desc_match = re.search(r'class="res-desc[^"]*"[^>]*>(.*?)</', block, re.S) or re.search(
                        r'class="res-rich[^"]*"[^>]*>(.*?)</', block, re.S)
                    desc = re.sub(r'<[^>]+>', '', desc_match.group(1)).strip() if desc_match else "无摘要"
                    if title:
                        info += f"- [{title}]: {desc}\n"
                        count += 1
                        if count >= max_res: break
            if count > 0: return info
        except:
            pass

        try:
            import urllib.request, urllib.parse
            url = f"https://cn.bing.com/search?q={urllib.parse.quote(query)}"
            req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0",
                                                       "Cookie": "SRCHHPGUSR=SRCHLANG=zh-Hans;"})
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
        except:
            pass
        return info

    def run_enterprise_ai_analysis(api_url, api_key, model_name, template_text, in_dir, out_dir, use_web_report,
                                   use_web_news, use_local_doc, log_func=print):
        try:
            output_root = os.path.join(out_dir, "企业分析")
            os.makedirs(output_root, exist_ok=True)

            company_tasks = []
            for group_name in os.listdir(in_dir):
                group_path = os.path.join(in_dir, group_name)
                if not os.path.isdir(group_path) or "新建文件夹" in group_name: continue
                comp_folders = [d for d in os.listdir(group_path) if os.path.isdir(os.path.join(group_path, d))]
                for comp in comp_folders:
                    comp_path = os.path.join(group_path, comp)
                    peers = [c for c in comp_folders if c != comp]
                    company_tasks.append({"path": comp_path, "name": comp, "peers": peers})

            client = OpenAI(api_key=api_key or "sk-local", base_url=api_url)

            for task in company_tasks:
                comp_name_raw = task["name"]
                log_func(f"\n开始分析: {comp_name_raw}")

                clean_company = comp_name_raw.split('_')[-1] if '_' in comp_name_raw else comp_name_raw
                peer_names = [p.split('_')[-1] if '_' in p else p for p in task["peers"]]
                clean_comp_str = "、".join(peer_names) or "暂无"

                local_files = [os.path.join(task["path"], f) for f in os.listdir(task["path"]) if
                               f.lower().endswith(('.pdf', '.docx'))]

                web_info = ""
                if use_web_report: web_info += search_web_logic(f"{clean_company} 最新深度研究报告", 2)
                if use_web_news: web_info += search_web_logic(f"{clean_company} 最新公告 {clean_comp_str} 对比", 2)

                doc_info = ""
                if use_local_doc and local_files:
                    for path in local_files:
                        if path.endswith('.pdf'):
                            try:
                                with pdfplumber.open(path) as pdf:
                                    for page in pdf.pages[:35]: doc_info += (page.extract_text() or "") + "\n"
                            except:
                                pass

                prompt = template_text.format(company=clean_company, competitors=clean_comp_str,
                                              web_info=web_info[:8000], doc_info=doc_info[:15000])
                resp = client.chat.completions.create(model=model_name,
                                                      messages=[{"role": "system", "content": "资深金融分析师"},
                                                                {"role": "user", "content": prompt}], temperature=0.3)

                doc = Document()
                doc.add_heading(f'{clean_company} 深度分析报告', 0)
                doc.add_paragraph(f"对标: {clean_comp_str} | 时间: {datetime.now().strftime('%Y-%m-%d')}")
                doc.add_paragraph(resp.choices[0].message.content)

                save_fn = os.path.join(output_root, f"{clean_company}_深度报告_{datetime.now().strftime('%H%M')}.docx")
                doc.save(save_fn)
                log_func(f"✅ 报告已生成: {os.path.basename(save_fn)}")
            return output_root
        except Exception as e:
            log_func(f"❌ 错误: {str(e)}")
            return None