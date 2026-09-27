import streamlit as st
import os
# 导入你的后台流水线代码
from backend_integrated_system import OneClickOrchestrator, CONFIG

# 设置页面在手机端自适应
st.set_page_config(page_title="量化研报后台", page_icon="📈", layout="wide")

st.title("📈 量化研报后台管理系统")

# 使用选项卡（Tab）代替 tkinter 的 Notebook，在手机上可以滑动切换
tab1, tab2, tab3, tab4 = st.tabs(["全局配置", "数据海选", "年报处理", "AI深度估值"])

with tab1:
    st.header("⚙️ 全局配置")
    engine = st.selectbox("搜索引擎配置", ["Bing", "DuckDuckGo", "360"])
    overwrite = st.toggle("启动时清空旧数据(覆盖最新结果)", value=CONFIG["OVERWRITE_MODE"])
    api_key = st.text_input("AI API Key", type="password", value="sk-your-key")

with tab2:
    st.header("🔍 数据海选")
    wencai = st.selectbox("问财选股条件模板", CONFIG["TEMPLATES"]["wencai_conditions"])
    ai_filter = st.selectbox("AI过滤条件模板", CONFIG["TEMPLATES"]["ai_filter_prompts"])

with tab3:
    st.header("📥 年报下载与提取")
    st.info("流水线中自动执行: 财务报表下载 -> 多年度数据提取 -> A/B表补全")

with tab4:
    st.header("🧠 AI企业深度估值")
    text_prompt = st.selectbox("AI 文本分析指令", CONFIG["TEMPLATES"]["ai_text_prompts"])
    ent_prompt = st.selectbox("AI 企业估值指令", CONFIG["TEMPLATES"]["ai_enterprise_prompts"])
    good_price = st.checkbox("开启财务好价分析 (PE/PB/DCF估值测算)", value=True)

# 底部：一键启动按钮
st.divider()
if st.button("🚀 一键启动流水线", use_container_width=True): # use_container_width 让按钮在手机上占满全宽
    # 构造传给后端的配置字典
    ui_config = {
        "base_dir": "./data_output", # 在服务器端的相对路径
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
    
    # 进度提示
    with st.spinner('任务正在执行中，请耐心等待（请勿关闭浏览器）...'):
        # 创建一个占位符用来显示后端的日志
        log_box = st.empty()
        
        # 改造一下原有的 log_func，让它输出到网页上
        logs = []
        def web_log(msg):
            logs.append(msg)
            # 在网页上实时刷新日志
            log_box.text_area("实时日志", "\n".join(logs[-10:]), height=200)

        # 执行后端代码
        try:
            OneClickOrchestrator.run_all(ui_config, log_func=web_log)
            st.success("🎉 全链路任务完美收官！")
            
            # TODO: 任务完成后，提供打包下载按钮 (见下文)
            
        except Exception as e:
            st.error(f"❌ 运行发生中断异常: {str(e)}")
