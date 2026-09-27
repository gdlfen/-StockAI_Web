# app.py
import streamlit as st
import os
import shutil
import time

# 导入后端逻辑
from backend_logic import load_templates, save_templates, CONFIG, OneClickOrchestrator

# ==========================================
# 0. 状态初始化与辅助函数
# ==========================================
if "templates" not in st.session_state:
    st.session_state.templates = load_templates()

def save_new_template(category, new_text):
    """保存新模板逻辑"""
    if new_text and new_text not in st.session_state.templates[category]:
        st.session_state.templates[category].append(new_text)
        save_templates(st.session_state.templates)
        st.toast("✅ 新模板已永久保存！")
    elif new_text in st.session_state.templates[category]:
        st.toast("⚠️ 该模板已存在，无需重复保存。")

# ==========================================
# 1. 页面配置与侧边栏 (需求5)
# ==========================================
st.set_page_config(page_title="量化研报后台", page_icon="📈", layout="wide")

# 左侧边栏：放置指引和帮助
with st.sidebar:
    st.title("📚 帮助与指引")
    with st.expander("📖 使用指引", expanded=False):
        st.markdown("""
        **系统操作流程**：
        1. 在**全局配置**填写搜索引擎与大模型的秘钥。
        2. 在**数据海选**调整初筛与过滤条件。
        3. 核实其他板块配置后，点击页面底部的**一键启动流水线**。
        4. 等待日志输出完成，点击生成的**ZIP下载按钮**获取文件。
        """)
        
    with st.expander("❓ 帮助说明", expanded=False):
        st.markdown("""
        **常见问题**：
        - **什么是Tavily？** 专为AI设计的搜索引擎，抓取研报比Bing更准。
        - **如何保存新模板？** 在输入框内修改内容，点击下方的“保存为新模板”按钮即可持久化。
        - **运行报错怎么办？** 请检查 API Key 和 API 地址是否填写正确。
        """)

st.title("📈 量化研报管理系统 v3.0")
st.caption("手机端完美适配版 | 支持自定义指令模板与完整分析流")

# 顶部 Tabs 
tab1, tab2, tab3, tab4, tab5 = st.tabs(["⚙️全局配置", "🔍数据海选", "📥年报处理", "📊排雷与16维", "🧠AI深度估值"])

# ==========================================
# 2. 各个面板配置区域
# ==========================================
with tab1:
    st.header("⚙️ 全局设置")
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("搜索引擎配置")
        engine = st.selectbox("选择搜索引擎", ["Bing", "DuckDuckGo", "360", "Tavily"], index=0)
        search_key = st.text_input("搜索引擎 Key (选填，仅Tavily等需要)", type="password")
        overwrite = st.toggle("启动时清空旧数据 (覆盖最新结果)", value=CONFIG.get("OVERWRITE_MODE", True))
        
    with col2:
        st.subheader("大模型 API 配置")
        api_url = st.text_input("API 请求地址", value="https://api.openai.com/v1")
        api_key = st.text_input("API Key", type="password", value="sk-your-key-here")
        api_model = st.selectbox("API 模型选择", ["gpt-4-turbo", "gpt-4o", "gpt-3.5-turbo", "claude-3-5-sonnet", "deepseek-chat"], index=0)

with tab2:
    st.header("🔍 数据海选条件")
    st.info("💡 提示：在输入框内修改条件后，点击对应下方的按钮即可将新条件存入模板库。")
    
    # 问财条件设置
    st.subheader("1. 问财初筛条件")
    wencai_sel = st.selectbox("快速选择预设（问财）", st.session_state.templates["wencai_conditions"])
    wencai_text = st.text_area("问财选股条件（可自由修改）", value=wencai_sel, height=100)
    if st.button("💾 将上方问财条件保存为新模板"):
        save_new_template("wencai_conditions", wencai_text)
        
    st.divider()
    
    # AI过滤设置
    st.subheader("2. AI 深度过滤")
    enable_ai_filter = st.toggle("启用 AI 深度过滤", value=True)
    if enable_ai_filter:
        ai_filter_sel = st.selectbox("快速选择预设（AI过滤）", st.session_state.templates["ai_filter_prompts"])
        ai_filter_text = st.text_area("AI 过滤条件（可自由修改）", value=ai_filter_sel, height=100)
        if st.button("💾 将上方AI过滤条件保存为新模板"):
            save_new_template("ai_filter_prompts", ai_filter_text)
    else:
        ai_filter_text = ""

with tab3:
    st.header("📥 年报处理")
    st.info("流水线将自动执行: 财务报表下载 -> 多年度数据提取 -> A/B表补全，生成标准的本地Excel数据源。")

with tab4:
    st.header("📊 财务排雷与16维分析")
    st.info("基于年报数据，全自动进行：智能造假排雷 -> 16维度同行对比，为你筛选绝对安全的标的。")

with tab5:
    st.header("🧠 AI深度估值指令设置")
    
    # 文本分析指令
    st.subheader("1. AI 文本分析指令")
    txt_sel = st.selectbox("快速选择预设（文本分析）", st.session_state.templates["ai_text_prompts"])
    txt_prompt = st.text_area("文本分析指令内容（可自由修改）", value=txt_sel, height=150)
    if st.button("💾 保存为新的文本分析模板"):
        save_new_template("ai_text_prompts", txt_prompt)

    st.divider()
    
    # 企业估值指令
    st.subheader("2. AI 企业估值与护城河指令")
    ent_sel = st.selectbox("快速选择预设（企业估值）", st.session_state.templates["ai_enterprise_prompts"])
    ent_prompt = st.text_area("企业估值指令内容（可自由修改）", value=ent_sel, height=150)
    if st.button("💾 保存为新的企业估值模板"):
        save_new_template("ai_enterprise_prompts", ent_prompt)
        
    st.divider()
    good_price = st.checkbox("开启财务好价分析 (自动测算 PE/PB/DCF 合理区间)", value=True)

st.divider()

# ==========================================
# 3. 核心运行逻辑与日志输出
# ==========================================
if st.button("🚀 一键启动完整流水线", type="primary", use_container_width=True):
    ui_config = {
        "base_dir": "./data_output",
        "engine": engine,
        "search_key": search_key,
        "overwrite": overwrite,
        "api_url": api_url,
        "api_key": api_key,
        "api_model": api_model,
        "wencai": wencai_text,
        "enable_ai_filter": enable_ai_filter,
        "ai_filter": ai_filter_text,
        "text_prompt": txt_prompt,
        "ent_prompt": ent_prompt,
        "good_price": good_price
    }
    
    with st.spinner('任务正在云端执行中，期间请勿刷新或关闭浏览器...'):
        log_box = st.empty()
        logs = []
        
        def web_log(msg):
            logs.append(msg)
            log_box.text_area("实时执行日志 (最近20条)", "\n".join(logs[-20:]), height=300)

        try:
            # 启动后端分析逻辑
            OneClickOrchestrator.run_all(ui_config, log_func=web_log)
            st.success("🎉 全链路任务完美收官！请点击下方按钮下载数据。")
            
            output_dir = ui_config["base_dir"]
            if os.path.exists(output_dir):
                zip_path = shutil.make_archive(
                    base_name="量化研报结果汇总", 
                    format="zip", 
                    root_dir=output_dir
                )
                
                with open(zip_path, "rb") as f:
                    st.download_button(
                        label="📦 点击下载全部分析报告及Excel数据 (ZIP)",
                        data=f,
                        file_name=f"量化研报汇总_{int(time.time())}.zip",
                        mime="application/zip",
                        use_container_width=True
                    )
            else:
                st.warning("⚠️ 任务执行完成，但未检测到输出数据文件夹。")
                
        except Exception as e:
            st.error(f"❌ 运行发生中断异常: {str(e)}")
