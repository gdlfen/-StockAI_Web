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
st.set_page_config(page_title="量化研报中枢", page_icon="📈", layout="wide")

with st.sidebar:
    st.title("📚 帮助与指引")
    with st.expander("📖 使用指引", expanded=False):
        st.markdown("""
        **系统流转机制**：
        1. **全局配置**：设定网络节点与底层 LLM 模型参数。
        2. **数据海选**：触发底层数据抓取（若遇网络阻断自动降级为备用股池）。
        3. **一键启动**：系统将自动贯穿年报下载、真实财务数据提取、排雷精算与 AI 估值。
        """)
        
    with st.expander("❓ 疑难解答", expanded=False):
        st.markdown("""
        - **什么是Tavily？** 专为 AI 赋予实时检索能力的搜索引擎。
        - **API地址怎么选？** 系统预设了常用的大模型节点，直接在下拉框选择即可。
        - **保存模板失败？** 请确认直接在文本框修改后，点击下方按钮保存。
        """)

st.title("📈 量化投研管理系统 v4.2")
st.caption("集成真实年报自动下载与深度提取 | 移动端全自动自适应")

# 顶部核心业务流 Tabs
tab1, tab2, tab3, tab4, tab5 = st.tabs(["⚙️全局配置", "🔍数据海选", "📥年报处理", "📊排雷与16维", "🧠AI深度估值"])

# ==========================================
# 2. 面板配置区域
# ==========================================
with tab1:
    st.header("⚙️ 引擎底座设置")
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("网络与检索层")
        engine = st.selectbox("搜寻网络环境", ["Bing", "DuckDuckGo", "360", "Tavily"], index=0)
        search_key = st.text_input("网络密钥 (仅Tavily等需要)", type="password")
        overwrite = st.toggle("系统自净模式 (每次运行前抹除旧档)", value=True)
        
    with col2:
        st.subheader("AI算力接入层")
        # 需求2：将全局配置中的API地址设为下拉菜单
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
    st.header("🔍 数据海选策源地")
    
    st.subheader("1. 基础硬指标初筛")
    wencai_sel = st.selectbox("调用预设策略（问财指标）", st.session_state.templates["wencai_conditions"])
    wencai_text = st.text_area("策略修改区", value=wencai_sel, height=80)
    if st.button("💾 将上方指标存入策略库"):
        save_new_template("wencai_conditions", wencai_text)
        
    st.divider()
    
    st.subheader("2. AI 软逻辑复核")
    # 需求3：数据海选中的深度过滤切换键设为默认不过滤 (value=False)
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
    st.info("底层将前往巨潮资讯自动下载财报原件 PDF，并同步前往新浪财经抓取标准版财务宽表用于后续运算。")
    
    # 需求4：年报处理中增加年报下载年份跨度，并增加招股说明书、公司章程复选框
    st.subheader("时间跨度与公告选择")
    current_year = datetime.now().year
    col_y1, col_y2 = st.columns(2)
    with col_y1:
        start_year = st.number_input("年报下载起点 (年份)", min_value=2000, max_value=current_year, value=current_year-3)
    with col_y2:
        end_year = st.number_input("年报下载终点 (年份)", min_value=2000, max_value=current_year, value=current_year)
        
    dl_prospectus = st.checkbox("同时下载该公司的【招股说明书】(PDF)", value=True)
    dl_charter = st.checkbox("同时下载该公司的【公司章程】(PDF)", value=True)

with tab4:
    st.header("📊 量化评级与排雷中枢")
    st.info("将读取阶段三获取到的财务数据，自动执行智能财务排雷检测，并进行核心指标梳理。")

with tab5:
    st.header("🧠 AI 战略级研报合成")
    
    st.subheader("1. 长期财报穿透分析指令")
    txt_sel = st.selectbox("载入预设提词", st.session_state.templates["ai_text_prompts"])
    txt_prompt = st.text_area("提词器编辑区", value=txt_sel, height=120)
    if st.button("💾 收录为新文本分析模板"): save_new_template("ai_text_prompts", txt_prompt)

    st.divider()
    
    st.subheader("2. 企业护城河与竞争格局研判指令")
    ent_sel = st.selectbox("载入预设提词", st.session_state.templates["ai_enterprise_prompts"])
    ent_prompt = st.text_area("企业提词器编辑区", value=ent_sel, height=120)
    if st.button("💾 收录为新企业估值模板"): save_new_template("ai_enterprise_prompts", ent_prompt)
        
    st.divider()
    good_price = st.checkbox("融合 AkShare 宏观抓取：测算大盘基准收益率 (好价分析)", value=True)

st.divider()

# ==========================================
# 3. 总控台启动器
# ==========================================
if st.button("🚀 激活全链路量化作业矩阵", type="primary", use_container_width=True):
    # 将UI面板获取到的全部新参数打包进字典
    ui_config = {
        "base_dir": "./data_output",
        "engine": engine, "search_key": search_key, "overwrite": overwrite,
        "api_url": api_url, "api_key": api_key, "api_model": api_model,
        "wencai": wencai_text, "enable_ai_filter": enable_ai_filter, "ai_filter": ai_filter_text,
        "start_year": start_year, "end_year": end_year, 
        "dl_prospectus": dl_prospectus, "dl_charter": dl_charter,
        "text_prompt": txt_prompt, "ent_prompt": ent_prompt, "good_price": good_price
    }
    
    with st.spinner('集群运算中，请保持网络环境稳定，切勿刷新...'):
        log_box = st.empty()
        logs = []
        
        def web_log(msg):
            logs.append(msg)
            log_box.text_area("节点实时作业监控 (近20条记录)", "\n".join(logs[-20:]), height=300)

        try:
            # 传递配置开启流水线
            OneClickOrchestrator.run_all(ui_config, log_func=web_log)
            st.success("🎉 数据闭环打通，核心资产报表及报告已生成！")
            
            output_dir = ui_config["base_dir"]
            if os.path.exists(output_dir):
                zip_path = shutil.make_archive(base_name="量化研报结果汇总", format="zip", root_dir=output_dir)
                with open(zip_path, "rb") as f:
                    st.download_button(
                        label="📦 [获取成果] 下载全部脱水数据与研报库 (ZIP)",
                        data=f, file_name=f"量化研报汇总_{int(time.time())}.zip",
                        mime="application/zip", use_container_width=True
                    )
        except Exception as e:
            st.error(f"❌ 运行发生中断异常: {str(e)}")
