import streamlit as st
import json
import time

# --- 1. 初始化状态 (Session State) ---
if 'status' not in st.session_state:
    st.session_state.status = 'idle'  # 状态: idle, running, success, error
if 'progress' not in st.session_state:
    st.session_state.progress = 0
if 'error_msg' not in st.session_state:
    st.session_state.error_msg = ''
if 'error_line' not in st.session_state:
    st.session_state.error_line = None

# --- 2. 核心逻辑函数 ---

def run_validation():
    """点击运行按钮后的逻辑"""
    # 重置状态
    st.session_state.status = 'running'
    st.session_state.error_msg = ''
    st.session_state.error_line = None
    st.session_state.progress = 0
    
    # 模拟加载进度条动画 (黑色背景 -> 绿色增长)
    bar_slot = st.empty()
    for i in range(101):
        st.session_state.progress = i
        # 强制刷新UI来实现动画效果 (在Streamlit中通常自动处理，这里为了模拟进度)
        time.sleep(0.01) 
    
    # 执行校验逻辑
    input_text = st.session_state.get('code_input', '')
    
    try:
        # A. 尝试解析 JSON
        data = json.loads(input_text)
        
        # B. 业务逻辑校验：比如必须包含 'target' 字段且不为空
        if not data.get('target'):
            raise ValueError("业务规则校验失败: 缺少 'target' 字段或值为空。")
            
        # C. 通过
        st.session_state.status = 'success'
        
    except json.JSONDecodeError as e:
        # JSON 格式错误，定位行号
        st.session_state.status = 'error'
        st.session_state.error_msg = f"语法错误: {e.msg}"
        st.session_state.error_line = e.lineno # 获取错误行号
        
    except Exception as e:
        # 其他业务逻辑错误
        st.session_state.status = 'error'
        st.session_state.error_msg = str(e)

def reset():
    st.session_state.status = 'idle'
    st.session_state.progress = 0
    st.session_state.error_msg = ''
    st.session_state.error_line = None

# --- 3. 页面布局与自定义 CSS (关键部分) ---

st.title("轩辕数据湖 - 任务执行控制台")

# 这里使用 CSS 注入来完美还原你要求的“黑色背景进度条”
# Streamlit 原生进度条是蓝色的，所以我们手写一段 HTML/CSS
status = st.session_state.status
progress_width = st.session_state.progress if status != 'error' else 0
bar_color = "#22c55e" if status == 'success' else "#22c55e" # 绿色

# 定义图标和颜色逻辑
if status == 'idle':
    icon = "▶" # 播放
    btn_color = "white"
    msg = "点击运行"
elif status == 'running':
    icon = "⏳" 
    btn_color = "white"
    msg = "执行中..."
elif status == 'success':
    icon = "✔"
    btn_color = "white"
    msg = "执行成功"
elif status == 'error':
    icon = "❗" # 警告感叹号
    btn_color = "#ef4444" # 红色
    msg = "校验未通过"

# 渲染自定义进度条 HTML
st.markdown(f"""
<style>
    .custom-bar-container {{
        position: relative;
        width: 100%;
        height: 50px;
        background-color: black; /* 默认黑色背景 */
        border-radius: 8px;
        overflow: hidden;
        margin-bottom: 20px;
        display: flex;
        align-items: center;
        cursor: pointer;
        border: 1px solid #333;
    }}
    .progress-fill {{
        position: absolute;
        height: 100%;
        width: {progress_width}%; 
        background-color: {bar_color}; /* 校验通过变绿 */
        transition: width 0.3s ease;
        z-index: 1;
    }}
    .content-layer {{
        position: relative;
        z-index: 2;
        display: flex;
        align-items: center;
        padding-left: 20px;
        color: white;
        font-family: monospace;
        font-weight: bold;
        width: 100%;
    }}
    .icon-box {{
        width: 30px;
        height: 30px;
        border-radius: 50%;
        background: rgba(255,255,255,0.2);
        display: flex;
        justify-content: center;
        align-items: center;
        margin-right: 10px;
        color: {btn_color};
    }}
    /* 失败时的红色光晕动画 */
    .error-pulse {{
        box-shadow: 0 0 0 0 rgba(239, 68, 68, 0.7);
        animation: pulse-red 1.5s infinite;
        background: #ef4444; 
    }}
    @keyframes pulse-red {{
        0% {{ transform: scale(0.95); box-shadow: 0 0 0 0 rgba(239, 68, 68, 0.7); }}
        70% {{ transform: scale(1); box-shadow: 0 0 0 10px rgba(239, 68, 68, 0); }}
        100% {{ transform: scale(0.95); box-shadow: 0 0 0 0 rgba(239, 68, 68, 0); }}
    }}
</style>

<div class="custom-bar-container">
    <div class="progress-fill"></div>
    <div class="content-layer">
        <div class="icon-box {'error-pulse' if status == 'error' else ''}">
            {icon}
        </div>
        <span>{msg}</span>
    </div>
</div>
""", unsafe_allow_html=True)

# 真实的触发按钮 (隐藏在逻辑中或放在上方，为了交互方便，我们用 Streamlit 原生按钮来触发逻辑)
col1, col2 = st.columns([1, 4])
with col1:
    # 这里的按钮用于触发 Python 逻辑
    if st.button("开始执行 (Run)", disabled=(status=='running')):
        run_validation()
        st.rerun() # 强制刷新以显示进度条变化
with col2:
    if status in ['success', 'error']:
        if st.button("重置 (Reset)"):
            reset()
            st.rerun()

# --- 4. 错误信息展示 ---
if status == 'error':
    st.error(f"❌ 失败原因: {st.session_state.error_msg}")

# --- 5. 在线编辑数据 & 可视化定位 ---
st.subheader("在线数据编辑")

# 默认数据
default_json = """{
    "task_id": 1001,
    "task_name": "数据清洗_V1",
    "target": "" 
}"""

# 计算 Text Area 的高度和样式
# 如果出错，我们尝试在 label 处提示
label_text = "JSON 配置"
if st.session_state.error_line:
    label_text += f" (👉 错误可能在第 {st.session_state.error_line} 行附近)"

# 输入框
code = st.text_area(
    label=label_text,
    value=st.session_state.get('code_input', default_json),
    height=200,
    key='code_input',
    help="在这里直接修改 JSON，然后点击上方运行。"
)

# 可视化定位：如果出错，我们在下方显示一个带行号的“代码快照”，高亮错误行
if status == 'error' and st.session_state.error_line:
    st.warning("🔍 错误定位分析：")
    lines = code.split('\n')
    
    # 简单的可视化：打印出每一行，并在错误行加箭头
    for idx, line in enumerate(lines):
        line_num = idx + 1
        if line_num == st.session_state.error_line:
            st.markdown(f"**Line {line_num}:** `{line}` 👈 <span style='color:red'>**HERE**</span>", unsafe_allow_html=True)
        else:
            st.markdown(f"<span style='color:gray'>Line {line_num}: {line}</span>", unsafe_allow_html=True)
