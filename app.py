import streamlit as st
import tempfile
import os
import shutil
from datetime import datetime

# 导入你的后端逻辑
from backend_logic import *

# 1. 页面基础配置 (适配手机端)
st.set_page_config(
    page_title="价值投资 AI 投研平台",
    page_icon="📈",
    layout="centered",
    initial_sidebar_state="collapsed"  # 手机上默认折叠侧边栏
)

# 2. 初始化 Session State (保存 API Key 等配置)
if 'api_key' not in st.session_state:
    st.session_state.api_key = ""
if 'api_url' not in st.session_state:
    st.session_state.api_url = "https://api.deepseek.com/v1"

st.title("📈 价值投资智能投研系统")
st.caption("手机端自适应版 | 全自动云端运算")

# 3. 侧边栏导航菜单
menu = st.sidebar.radio(
    "🧭 导航菜单",
    ["⚙️ 全局配置", "🔍 A: 数据海选", "📥 A: 年报下载与提取", "💣 B: 财务排雷与16维分析", "🧠 C: AI企业深度估值"]
)

# ==========================================
# 模块 0：全局配置
# ==========================================
if menu == "⚙️ 全局配置":
    st.header("⚙️ AI 模型全局配置")
    with st.form("config_form"):
        engine = st.selectbox("AI 引擎", ["DeepSeek", "OpenAI", "Kimi", "本地 Ollama"])
        url = st.text_input("API URL", value=st.session_state.api_url)
        key = st.text_input("API Key (数据加密，仅当前会话有效)", type="password", value=st.session_state.api_key)

        if st.form_submit_button("保存配置", use_container_width=True):
            st.session_state.api_url = url
            st.session_state.api_key = key
            st.success("✅ 配置已保存在当前云端会话中！")

# ==========================================
# 模块 1：问财海选
# ==========================================
elif menu == "🔍 A: 数据海选":
    st.header("🔍 问财海选与筛选")
    with st.form("scraper_form"):
        market = st.radio("市场分类", ["A股", "港股", "美股"], horizontal=True)
        query = st.text_area("问财选股条件", "连续5年加权roe>20\n连续5年净利润现金含量>100\n非金融股")
        submitted = st.form_submit_button("🚀 开始云端无头抓取", use_container_width=True)

    if submitted:
        with st.spinner("⏳ 正在云端启动无头浏览器，请不要锁屏... (约需1-3分钟)"):
            # 在云端创建临时文件夹
            with tempfile.TemporaryDirectory() as tmp_dir:
                try:
                    # TODO: 调用你修改后的 backend_logic 里的问财抓取函数
                    # 比如: result_file = run_wencai_scraper(query, market, tmp_dir)

                    st.success("✅ 数据抓取完成！")
                    # 提供下载按钮 (伪代码展示)
                    # with open(result_file, "rb") as f:
                    #     st.download_button("📥 下载海选名单.xlsx", data=f, file_name="海选结果.xlsx", use_container_width=True)
                except Exception as e:
                    st.error(f"抓取失败: {e}")

# ==========================================
# 模块 2：排雷与16维度分析 (核心：文件上传下载)
# ==========================================
elif menu == "💣 B: 财务排雷与16维分析":
    st.header("💣 财务排雷与16维分析")
    st.info("💡 手机端操作指南：请点击下方按钮，上传由年报提取生成的《统一整合输出》Excel文件。")

    # 手机端文件上传组件
    uploaded_files = st.file_uploader("上传财务数据表 (.xlsx/.csv)", type=["xlsx", "csv"], accept_multiple_files=True)

    if st.button("⚙️ 开始执行综合诊断", type="primary", use_container_width=True):
        if not uploaded_files:
            st.warning("⚠️ 请先上传至少一个财务表格！")
        else:
            with st.spinner("⏳ 服务器正在高速运算排雷指标与16维打分..."):
                # 核心机制：用临时文件夹替代原本的 C:\ 盘本地路径
                with tempfile.TemporaryDirectory() as tmp_in_dir, tempfile.TemporaryDirectory() as tmp_out_dir:

                    # 1. 把手机上传的文件写入服务器临时目录
                    for uf in uploaded_files:
                        with open(os.path.join(tmp_in_dir, uf.name), "wb") as f:
                            f.write(uf.getbuffer())

                    # 2. 调用纯净的后端逻辑
                    try:
                        # TODO: 这里调用你后端的排雷和16维计算函数
                        # report_path = run_fraud_detection(tmp_in_dir, tmp_out_dir)

                        st.success("✅ 综合诊断完成！")

                        # 3. 将服务器生成的文件转回给手机端下载
                        # with open(report_path, "rb") as f:
                        #     st.download_button(
                        #         label="📥 下载深度排雷与16维报告",
                        #         data=f,
                        #         file_name="综合财务分析总汇.xlsx",
                        #         mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        #         use_container_width=True
                        #     )
                    except Exception as e:
                        st.error(f"分析过程中发生错误: {e}")

# ==========================================
# 模块 3：AI 深度企业估值
# ==========================================
elif menu == "🧠 C: AI企业深度估值":
    st.header("🧠 AI 文本与估值分析")

    if not st.session_state.api_key:
        st.warning("⚠️ 请先在【全局配置】中填写您的 API Key！")

    with st.form("ai_form"):
        company = st.text_input("目标公司", "贵州茅台")
        peers = st.text_input("对标同行", "五粮液, 泸州老窖")
        prompt = st.text_area("分析指令", "请作为资深金融分析师，从竞争格局、护城河、财务健康度三个维度进行分析。")
        submitted = st.form_submit_button("🚀 生成万字研报", use_container_width=True)

    if submitted and st.session_state.api_key:
        with st.spinner("⏳ AI 正在撰写万字长文研报，并抓取实时宏观数据..."):
            with tempfile.TemporaryDirectory() as tmp_out_dir:
                try:
                    # TODO: 调用你后端的 AI 生成和 Document(Word) 保存逻辑
                    # report_path = generate_ai_report(st.session_state.api_url, st.session_state.api_key, company, peers, prompt, tmp_out_dir)

                    st.success("✅ AI 研报生成完毕！")

                    # with open(report_path, "rb") as f:
                    #     st.download_button(
                    #         label=f"📥 下载《{company}深度研报.docx》",
                    #         data=f,
                    #         file_name=f"{company}_AI深度研报.docx",
                    #         use_container_width=True
                    #     )
                except Exception as e:
                    st.error(f"调用 API 失败: {e}")