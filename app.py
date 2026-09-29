# app.py
import streamlit as st
import os
import shutil
import time
from datetime import datetime

# 导入底层核心逻辑
from backend_logic import load_templates, save_templates, CONFIG, OneClickOrchestrator

# ==========================================
# 0. 状态初始化与辅助函数
# ==========================================
if "templates" not in st.session_state:
    st.session_state.templates = load_templates()

def save_new_template(category, new_text):
    if new_text and new_text not in st.session_state.templates[category]:
        st.session_state.templates[category].append(new_text)
        save_templates(st.session_state.templates)
        st.toast("✅ 新模板已永久保存！")
    elif new_text in st.session_state.templates[category]:
        st.toast("⚠️ 该模板已存在，无需重复保存。")

# ==========================================
# 1. 页面配置与侧边栏
# ==========================================
st.set_page_config(page_title="纯API量化研报中枢", page_icon="📈", layout="wide")

with st.sidebar:
    st.title("📚 帮助与指引")
    with st.expander("📖 系统重构说明", expanded=True):
        st.markdown("""
        **1. 底层引擎彻底革新**：全面放弃 Selenium 网页爬虫，替换为开源稳定合法的 `AkShare` 接口体系！  
        **2. 接口高可用机制**：内置三阶智能防断联与重连设计，完美解决 `Remote end closed connection` 异常。  
        **3. 行业提取精准度升级**：修复平安银行变白酒 BUG，直接调取权威版块成分股。
        """)

st.title("📈 价值投资智能分析系统 V6.0 (开源 API 纯享版)")
st.caption("全维度扩展 | API直连穿透提取 | 同业智能分析 | 好价贴现估值")

# 顶部核心业务流 Tabs
tab1, tab2, tab3, tab4, tab5 = st.tabs(["⚙️全局配置", "🔍数据海选", "📥年报处理", "📊排雷与16维", "🧠AI深度估值"])

# ==========================================
# 2. 面板配置区域
# ==========================================
with tab1:
    st.header("⚙️ 引擎底座设置")
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("网络与 AkShare 引擎设置")
        engine = st.selectbox("搜寻网络环境", ["Bing", "DuckDuckGo", "360", "Tavily"], index=0)
        search_key = st.text_input("网络密钥 (仅Tavily等需要)", type="password")
        overwrite = st.toggle("系统自净模式 (每次运行前抹除旧档)", value=True)
        
        st.markdown("---")
        ak_retries = st.number_input("AkShare 接口防断联最大重试次数", min_value=1, max_value=10, value=3)
        ak_delay = st.number_input("接口重试等待延迟 (秒)", min_value=1, max_value=10, value=2)

    with col2:
        st.subheader("AI算力接入层")
        api_urls = [
            "https://api.deepseek.com",
            "https://api.openai.com/v1", 
            "https://api.moonshot.cn/v1", 
            "https://api.siliconflow.cn/v1", 
            "http://localhost:11434/v1"
        ]
        api_url = st.selectbox("API 接口路由 (Base URL)", api_urls, index=0)
        api_key = st.text_input("大模型秘钥 (API Key)", type="password", value="sk-your-key")
        api_model = st.selectbox("核心调用模型", ["deepseek-chat", "gpt-4o", "gpt-4-turbo", "claude-3-5-sonnet"], index=0)

with tab2:
    st.header("🔍 数据海选策源地 (基于 API 计算池)")

    st.subheader("1. 基础硬指标初筛")
    wencai_sel = st.selectbox("调用预设策略（输入数值如: 市盈率<20，市值>100亿）", st.session_state.templates["wencai_conditions"])
    wencai_text = st.text_area("策略修改区", value=wencai_sel, height=80)
    if st.button("💾 将上方指标存入策略库"):
        save_new_template("wencai_conditions", wencai_text)

    st.divider()

    st.subheader("2. AI 软逻辑复核")
    enable_ai_filter = st.toggle("激活 LLM 深度过滤体系", value=False)
    if enable_ai_filter:
        ai_filter_sel = st.selectbox("调用预设风控指令", st.session_state.templates["ai_filter_prompts"])
        ai_filter_text = st.text_area("风控指令修改区", value=ai_filter_sel, height=80)
        if st.button("💾 将上方指令固化为新模板"):
            save_new_template("ai_filter_prompts", ai_filter_text)
    else:
        ai_filter_text = ""

with tab3:
    st.header("📥 A/B 表融合作业层")
    st.info("系统将直接通过 AkShare 金融接口安全抓取新浪/东财底层原表（资产负债、利润、现金流、员工分红），并调用巨潮合规 API 提取原件 PDF 供 AI 阅读。")

    st.subheader("时间跨度与公告选择")
    current_year = datetime.now().year
    col_y1, col_y2 = st.columns(2)
    with col_y1:
        start_year = st.number_input("财报提取起点 (年份)", min_value=2000, max_value=current_year, value=current_year-3)
    with col_y2:
        end_year = st.number_input("财报提取终点 (年份)", min_value=2000, max_value=current_year, value=current_year)

    dl_prospectus = st.checkbox("同源下载该公司的【招股说明书】(PDF)", value=True)
    dl_charter = st.checkbox("同源下载该公司的【公司章程】(PDF)", value=True)

with tab4:
    st.header("📊 量化评级与排雷中枢")
    st.info("已内嵌【18大类财务造假识别模型】(存贷双高/营收背离等) 及 【16维度基本面量化精算】体系。")

with tab5:
    st.header("🧠 AI 战略级研报合成")

    st.subheader("1. 长期财报穿透分析指令 (AI文本分析)")
    txt_sel = st.selectbox("载入预设提词", st.session_state.templates["ai_text_prompts"])
    txt_prompt = st.text_area("提词器编辑区 (MD&A精析)", value=txt_sel, height=120)
    if st.button("💾 收录为新文本分析模板"): save_new_template("ai_text_prompts", txt_prompt)

    st.divider()

    st.subheader("2. 企业护城河与竞争格局研判指令 (AI企业分析)")
    ent_sel = st.selectbox("载入预设提词", st.session_state.templates["ai_enterprise_prompts"])
    ent_prompt = st.text_area("企业提词器编辑区 (护城河穿透)", value=ent_sel, height=120)
    if st.button("💾 收录为新企业估值模板"): save_new_template("ai_enterprise_prompts", ent_prompt)

    st.divider()
    good_price = st.checkbox("开启自动【好价分析】(实时对标国债与市盈率计算目标建仓价)", value=True)

st.divider()

# ==========================================
# 3. 总控台启动器
# ==========================================
if st.button("🚀 激活全链路 API 量化作业矩阵", type="primary", use_container_width=True):
    ui_config = {
        "base_dir": "./data_output",
        "engine": engine, "search_key": search_key, "overwrite": overwrite,
        "ak_retries": ak_retries, "ak_delay": ak_delay,
        "api_url": api_url, "api_key": api_key, "api_model": api_model,
        "wencai": wencai_text, "enable_ai_filter": enable_ai_filter, "ai_filter": ai_filter_text,
        "start_year": start_year, "end_year": end_year, 
        "dl_prospectus": dl_prospectus, "dl_charter": dl_charter,
        "text_prompt": txt_prompt, "ent_prompt": ent_prompt, "good_price": good_price
    }

    with st.spinner('接口并行调度中，速度极快，切勿刷新...'):
        log_box = st.empty()
        logs = []

        def web_log(msg):
            logs.append(msg)
            log_box.text_area("节点实时作业监控 (近30条记录)", "\n".join(logs[-30:]), height=350)

        try:
            OneClickOrchestrator.run_all(ui_config, log_func=web_log)
            st.success("🎉 数据闭环彻底打通！全部独立研报与安全边际精算表均已生成！")

            output_dir = ui_config["base_dir"]
            if os.path.exists(output_dir):
                zip_path = shutil.make_archive(base_name="量化研报结果汇总", format="zip", root_dir=output_dir)
                with open(zip_path, "rb") as f:
                    st.download_button(
                        label="📦 [获取成果] 下载全部核心脱水数据与研报库 (ZIP)",
                        data=f, file_name=f"量化研报汇总_{int(time.time())}.zip",
                        mime="application/zip", use_container_width=True
                    )
        except Exception as e:
            st.error(f"❌ 运行发生中断异常: {str(e)}")
