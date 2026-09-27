import streamlit as st
import os
import shutil
import time

# 核心解耦：从后端文件中导入配置和启动调度器
# 确保你的仓库中有一个名为 backend_logic.py 的文件
from backend_logic import CONFIG, OneClickOrchestrator

# ==========================================
# 1. 页面与UI初始化
# ==========================================
st.set_page_config(page_title="量化研报后台", page_icon="📈", layout="wide")

st.title("📈 量化研报管理系统 v2.0")
st.caption("手机端完美适配版 | 数据与研报自动打包下载")

# 顶部 Tabs (手机端支持滑动)
tab1, tab2, tab3, tab4 = st.tabs(["全局配置", "数据海选", "年报处理", "AI深度估值"])

# ==========================================
# 2. 面板配置区域
# ==========================================
with tab1:
    st.header("⚙️ 全局配置")
    engine = st.selectbox("搜索引擎配置", ["Bing", "DuckDuckGo", "360"], index=0)
    overwrite = st.toggle("启动时清空旧数据 (覆盖最新结果)", value=CONFIG.get("OVERWRITE_MODE", True))
    api_key = st.text_input("AI API Key", type="password", value="sk-your-key-here")

with tab2:
    st.header("🔍 数据海选")
    wencai = st.selectbox("问财选股条件模板", CONFIG["TEMPLATES"]["wencai_conditions"])
    ai_filter = st.selectbox("AI过滤条件模板", CONFIG["TEMPLATES"]["ai_filter_prompts"])

with tab3:
    st.header("📥 年报处理")
    st.info("流水线将自动执行: 财务报表下载 -> 多年度数据提取 -> A/B表补全")

with tab4:
    st.header("🧠 AI深度估值")
    text_prompt = st.selectbox("AI 文本分析指令", CONFIG["TEMPLATES"]["ai_text_prompts"])
    ent_prompt = st.selectbox("AI 企业估值指令", CONFIG["TEMPLATES"]["ai_enterprise_prompts"])
    good_price = st.checkbox("开启财务好价分析 (PE/PB/DCF估值测算)", value=True)

st.divider()

# ==========================================
# 3. 核心运行逻辑与日志输出
# ==========================================
if st.button("🚀 一键启动流水线", type="primary", use_container_width=True):
    # 收集UI界面上的所有配置
    ui_config = {
        "base_dir": "./data_output",
        "engine": engine,
        "overwrite": overwrite,
        "api_url": "https://api.openai.com/v1",
        "api_key": api_key,
        "model_name": "gpt-4-turbo",
        "wencai": wencai,
        "ai_filter": ai_filter,
        "text_prompt": text_prompt,
        "ent_prompt": ent_prompt,
        "good_price": good_price
    }
    
    # 同步到全局配置
    CONFIG["OVERWRITE_MODE"] = ui_config["overwrite"]
    CONFIG["BASE_DIR"] = ui_config["base_dir"]
    CONFIG["SEARCH_ENGINE"] = ui_config["engine"]
    
    with st.spinner('任务正在云端执行中，请耐心等待（请勿关闭浏览器）...'):
        log_box = st.empty()
        logs = []
        
        # 实时日志回调函数
        def web_log(msg):
            logs.append(msg)
            # 在网页上实时刷新最新 15 条日志
            log_box.text_area("实时执行日志", "\n".join(logs[-15:]), height=250)

        try:
            # 启动后端分析逻辑
            OneClickOrchestrator.run_all(ui_config, log_func=web_log)
            st.success("🎉 全链路任务完美收官！")
            
            # 任务完成后，打包生成的文件供手机端下载
            output_dir = ui_config["base_dir"]
            if os.path.exists(output_dir):
                zip_path = shutil.make_archive(
                    base_name="量化研报结果汇总", 
                    format="zip", 
                    root_dir=output_dir
                )
                
                with open(zip_path, "rb") as f:
                    st.download_button(
                        label="📦 点击下载全部分析报告及数据 (ZIP)",
                        data=f,
                        file_name=f"量化研报汇总_{int(time.time())}.zip",
                        mime="application/zip",
                        use_container_width=True
                    )
            else:
                st.warning("⚠️ 任务执行完成，但未检测到输出数据。")
                
        except Exception as e:
            st.error(f"❌ 运行发生中断异常: {str(e)}")
