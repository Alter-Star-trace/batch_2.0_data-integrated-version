# app.py
import base64  # 顶部导入base64模块
import streamlit as st
from core_excel import process_excel_core
import datetime
import os
import tempfile

# -------------------------- 页面基础配置（极致移动端适配） --------------------------
st.set_page_config(
    page_title="舟山Excel处理工具（移动端）",
    page_icon="📊",
    layout="centered",  # 紧凑布局，适配手机窄屏，杜绝横向滚动
    initial_sidebar_state="collapsed"  # 永久隐藏侧边栏，避免移动端误触
)

# -------------------------- 自定义CSS（优化移动端触控/显示体验） --------------------------
st.markdown("""
    <style>
    /* 全局样式：放大字体、优化行高，适配手机阅读 */
    * {
        font-size: 14px !important;
        line-height: 1.6 !important;
    }
    /* 标题样式：适度放大，居中 */
    h1 { font-size: 20px !important; text-align: center; margin-bottom: 20px !important; }
    h2 { font-size: 16px !important; margin-top: 15px !important; margin-bottom: 10px !important; }
    /* 按钮：占满整行、放大触控区域、圆角，适配手机点击 */
    div.stButton > button {
        width: 100% !important;
        padding: 12px 0 !important;
        border-radius: 8px !important;
        font-size: 16px !important;
    }
    /* 上传组件：放大，适配手机选择文件 */
    div.stFileUploader > div {
        padding: 15px !important;
        border-radius: 8px !important;
    }
    /* 输入框：放大，适配手机输入 */
    div.stTextInput > div > input {
        padding: 10px !important;
        font-size: 16px !important;
    }
    /* 日志区域：灰色背景、圆角、固定最大高度、滚动条，避免页面过长 */
    .log-container {
        background-color: #f5f7fa !important;
        padding: 12px !important;
        border-radius: 8px !important;
        max-height: 300px !important;
        overflow-y: auto !important;
        white-space: pre-wrap !important;
    }
    /* 隐藏Streamlit默认页脚、菜单，净化界面 */
    footer { visibility: hidden !important; }
    div[data-testid="stToolbar"] { visibility: hidden !important; }
    div[data-testid="stDecoration"] { visibility: hidden !important; }
    </style>
""", unsafe_allow_html=True)

# -------------------------- 初始化Streamlit会话状态（保存日志/结果，避免刷新丢失） --------------------------
if "log_list" not in st.session_state:
    st.session_state.log_list = []  # 保存日志列表
if "process_success" not in st.session_state:
    st.session_state.process_success = False  # 处理是否成功
if "save_path" not in st.session_state:
    st.session_state.save_path = ""  # 结果文件路径（本地测试用）

# -------------------------- 日志回调函数（适配Streamlit，实时更新日志区域） --------------------------
def streamlit_log_callback(msg):
    """自定义日志回调，将日志存入会话状态，实现实时更新"""
    # 拼接时间戳，和原GUI/核心模块日志格式一致
    timestamp = datetime.datetime.now().strftime("[%Y-%m-%d %H:%M:%S]")
    log_msg = f"{timestamp} {msg}"
    st.session_state.log_list.append(log_msg)
    # 实时更新日志区域（只保留最新50条，避免内存溢出）
    if len(st.session_state.log_list) > 50:
        st.session_state.log_list = st.session_state.log_list[-50:]

# -------------------------- 页面主体布局（移动端友好，从上到下流式布局） --------------------------
st.title("📊 舟山Excel数据处理工具")
st.divider()

# 1. 模板文件上传（移动端适配）
st.subheader("📋 上传模板文件", divider="gray")
template_file = st.file_uploader(
    label="选择【舟山达成追踪表】模板（仅.xlsx格式）",
    type=["xlsx"],
    accept_multiple_files=False,
    help="请上传Excel模板文件，处理后将保留模板原有格式/公式"
)

st.divider()

# 2. 数据文件上传（移动端适配）
st.subheader("📈 上传数据文件", divider="gray")
data_file = st.file_uploader(
    label="选择【浙沪发货滞留】数据文件（仅.xlsx格式）",
    type=["xlsx"],
    accept_multiple_files=False,
    help="请上传包含发货/滞留表的Excel数据文件，将自动提取舟山区数据"
)

st.divider()

# 3. 结果文件名称配置（移动端适配，自动生成时间戳，支持自定义）
st.subheader("📝 结果文件配置", divider="gray")
current_time = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
default_filename = f"舟山达成追踪表_处理结果_{current_time}.xlsx"
result_filename = st.text_input(
    label="结果文件名称（自动添加.xlsx，无需手动输入）",
    value=default_filename.replace(".xlsx", ""),
    help="直接输入名称即可，系统会自动补充.xlsx后缀"
)
# 自动处理文件名，确保以.xlsx结尾
if not result_filename.endswith(".xlsx"):
    result_filename += ".xlsx"

st.divider()

# 4. 日志输出区域（实时更新，移动端适配）
st.subheader("📜 处理日志（实时更新）", divider="gray")
log_placeholder = st.empty()
# 渲染日志区域
with log_placeholder.container():
    log_content = "\n".join(st.session_state.log_list)
    st.markdown(f'<div class="log-container">{log_content}</div>', unsafe_allow_html=True)

# 初始化日志（首次加载时）
if len(st.session_state.log_list) == 0:
    streamlit_log_callback("🔍 程序已就绪，请上传模板和数据文件后点击【开始处理】")

st.divider()

# 5. 开始处理按钮 + 核心逻辑调用（移动端友好，加载状态提示）
st.subheader("🚀 开始处理", divider="gray")
if st.button("开始处理数据", type="primary"):
    # 重置会话状态
    st.session_state.log_list = []
    st.session_state.process_success = False
    st.session_state.save_path = ""
    streamlit_log_callback("🔍 开始校验上传文件，准备处理...")

    # 第一步：校验文件是否上传
    if not template_file or not data_file:
        streamlit_log_callback("❌ 错误：请先上传模板文件和数据文件，缺一不可！")
    else:
        # 第二步：将Streamlit上传的内存文件保存为临时文件（适配core_excel.py的文件路径入参）
        try:
            # 创建临时目录，自动清理
            with tempfile.TemporaryDirectory() as temp_dir:
                # 保存模板临时文件
                template_temp_path = os.path.join(temp_dir, "template_temp.xlsx")
                with open(template_temp_path, "wb") as f:
                    f.write(template_file.getbuffer())
                # 保存数据临时文件
                data_temp_path = os.path.join(temp_dir, "data_temp.xlsx")
                with open(data_temp_path, "wb") as f:
                    f.write(data_file.getbuffer())
                # 结果文件保存路径（项目根目录，方便用户查找）
                result_path = os.path.join(os.getcwd(), result_filename)
                st.session_state.save_path = result_path

                # 第三步：调用封装好的核心Excel处理函数（传入日志回调）
                streamlit_log_callback("⚙️ 开始调用核心处理逻辑，正在处理数据...")
                with st.spinner("处理中，请稍候（请勿刷新页面，避免中断）..."):
                    success, error_msg = process_excel_core(
                        template_path=template_temp_path,
                        data_path=data_temp_path,
                        save_path=result_path,
                        log_callback=streamlit_log_callback
                    )

                # 第四步：处理结果反馈
                st.session_state.process_success = success
                if success:
                    streamlit_log_callback(f"🎉 处理成功！结果文件已保存至项目根目录：{result_path}")
                else:
                    streamlit_log_callback(f"❌ 处理失败：{error_msg}")
        except Exception as e:
            streamlit_log_callback(f"❌ 临时文件处理失败：{str(e)}")

st.divider()

# 6. 结果下载区域（处理成功后显示，移动端直接下载）
st.subheader("📁 结果下载", divider="gray")
if st.session_state.process_success and os.path.exists(st.session_state.save_path):
    # 读取结果文件为字节流，支持移动端下载


    # 加固：将Excel文件转换为base64编码，强制指定下载格式
    with open(st.session_state.save_path, "rb") as f:
        result_bytes = f.read()
        b64 = base64.b64encode(result_bytes).decode()

    # 构建下载链接（强制Excel格式，避免浏览器误判）
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="{result_filename}" style="display:block;width:100%;padding:12px 0;text-align:center;background-color:#0e1117;color:white;border-radius:8px;text-decoration:none;font-size:16px;">点击下载处理结果Excel文件</a>'

    # 显示自定义下载按钮（替代原st.download_button，兼容性更强）
    st.markdown(href, unsafe_allow_html=True)
    st.info(f"💡 结果文件同时保存在本地：{st.session_state.save_path}", icon="ℹ️")
elif st.session_state.log_list and "处理失败" in st.session_state.log_list[-1]:
    st.error("❌ 处理失败，请查看上方日志排查问题！", icon="⚠️")
else:
    st.info("ℹ️ 请先上传文件并点击【开始处理】，处理成功后将显示下载按钮", icon="💡")

# -------------------------- 移动端使用提示 --------------------------
st.divider()
st.markdown(""" 
    ### 📱 移动端使用提示
    1.  推荐使用**Chrome/Safari/华为浏览器**打开，兼容性最佳；
    2.  上传文件时可选择手机本地/微信/QQ中的Excel文件；
    3.  下载的文件默认保存在手机「下载」文件夹，可在文件管理器中查找；
    4.  处理大文件时建议连接WiFi，避免移动数据消耗过大；
    5.  处理过程中请勿刷新页面，否则会中断处理并需要重新上传。
""", unsafe_allow_html=True)