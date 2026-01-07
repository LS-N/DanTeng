import streamlit as st
import pandas as pd
import io
import time
import zipfile
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode

# ==========================================
# 0. 基础配置 & CSS 魔法
# ==========================================
st.set_page_config(page_title="淡藤财务报表 Pro", page_icon="😈", layout="wide", initial_sidebar_state="collapsed")

st.markdown("""
<style>
    /* 全局黑绿配色 */
    :root { --bg-color: #0d1117; --card-bg: #161b22; --accent: #238636; --text: #c9d1d9; --red: #da3633; }
    .stApp { background-color: var(--bg-color); color: var(--text); }
    
    /* 幽灵按钮特定样式 (放在右上角的重置) */
    .ghost-btn button {
        border: 1px dashed #444 !important;
        color: #666 !important;
        background: transparent !important;
        font-size: 0.8rem !important;
        padding: 0.2rem 0.5rem !important;
        min-height: 0px !important;
        height: 32px !important;
    }
    .ghost-btn button:hover {
        border-color: var(--red) !important;
        color: var(--red) !important;
    }

    /* 错误提示卡片 */
    .error-card { 
        border: 1px solid var(--red); 
        background: rgba(218, 54, 51, 0.1); 
        padding: 1rem; 
        border-radius: 6px; 
        margin-bottom: 1rem;
        display: flex;
        align-items: center;
        gap: 1rem;
    }
    
    /* Dialog 宽度优化 */
    div[data-testid="stDialog"] > div { width: 90vw !important; max-width: 1200px !important; }
</style>
""", unsafe_allow_html=True)

# ==========================================
# 1. 状态管理 (Session State)
# ==========================================
if 'df_a' not in st.session_state: st.session_state.df_a = None
if 'df_b' not in st.session_state: st.session_state.df_b = None
if 'file_a_id' not in st.session_state: st.session_state.file_a_id = None
if 'file_b_id' not in st.session_state: st.session_state.file_b_id = None

# 核心控制状态
if 'calc_result_zip' not in st.session_state: st.session_state.calc_result_zip = None
if 'error_df' not in st.session_state: st.session_state.error_df = None
if 'need_manual_trigger' not in st.session_state: st.session_state.need_manual_trigger = False # 自动/手动模式锁

# ==========================================
# 2. 侧边栏参数
# ==========================================
with st.sidebar:
    st.header("⚙️ 参数设置")
    PRICE_PER_DAY = st.number_input("人力单价 (元/天)", value=1500, step=100)
    MIN_HOURS_THRESHOLD = st.number_input("工时阈值 (小时)", value=500, help="低于此值报错")
    SUBSIDY_TAG = st.text_input("补助关键词", "差旅补助")

# ==========================================
# 3. 数据控制台 (带右上角幽灵按钮)
# ==========================================
with st.container(border=True):
    # 布局技巧：标题占左边，按钮占右边
    c_head_L, c_head_R = st.columns([8, 1]) 
    
    c_head_L.markdown("### 📂 数据源控制台")
    
    # 右上角幽灵按钮逻辑
    with c_head_R:
        st.markdown('<div class="ghost-btn">', unsafe_allow_html=True) # 注入 CSS 类
        if st.button("🗑️ 重置", help="清空所有数据和缓存，恢复自动模式"):
            st.session_state.clear()
            st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)

    # 文件上传区
    c_up1, c_up2 = st.columns(2)
    f_a = c_up1.file_uploader("Source A: 交付明细", type=['xlsx', 'csv'], key='u_a')
    f_b = c_up2.file_uploader("Source B: 差旅明细", type=['xlsx', 'csv'], key='u_b')

    # 文件加载逻辑 (检测变动)
    def load_file(file, old_id):
        if file is not None and file.file_id != old_id:
            try:
                if file.name.endswith('.csv'): df = pd.read_csv(file)
                else: df = pd.read_excel(file)
                df.columns = [str(c).strip() for c in df.columns]
                # 文件变动逻辑：
                # 1. 清除旧结果
                st.session_state.calc_result_zip = None 
                st.session_state.error_df = None
                # 2. 注意：不改变 need_manual_trigger 状态。
                # 如果之前报错(True)，换了文件依然保持 True，等待用户点"手动重算"
                return df, file.file_id
            except Exception as e:
                st.error(f"读取失败: {e}")
                return None, old_id
        return None, old_id

    new_df_a, new_id_a = load_file(f_a, st.session_state.file_a_id)
    if new_df_a is not None: 
        st.session_state.df_a = new_df_a
        st.session_state.file_a_id = new_id_a

    new_df_b, new_id_b = load_file(f_b, st.session_state.file_b_id)
    if new_df_b is not None: 
        st.session_state.df_b = new_df_b
        st.session_state.file_b_id = new_id_b

# ==========================================
# 4. 逻辑控制中枢
# ==========================================

# 准备工作
has_data = (st.session_state.df_a is not None and st.session_state.df_b is not None)
should_run = False # 本次刷新是否执行计算

# 场景判断
if has_data:
    # 场景 A: 已有计算结果 -> 不跑，直接显示结果区
    if st.session_state.calc_result_zip is not None:
        pass 
        
    # 场景 B: 之前报错了 (Need Manual) -> 显示手动按钮，不自动跑
    elif st.session_state.need_manual_trigger:
        st.divider()
        col_m1, col_m2 = st.columns([3, 1])
        col_m1.info("⚠️ 上次校验发现错误，已暂停自动计算。请修复源文件重新上传，或在线修复。")
        if col_m2.button("🔄 手动重新计算", type="primary", use_container_width=True):
            should_run = True
            st.session_state.need_manual_trigger = False # 解锁
            st.session_state.error_df = None # 清除旧错误
            
    # 场景 C: 默认自动模式 -> 自动跑
    else:
        should_run = True

# ==========================================
# 5. 执行引擎 (校验 -> 计算)
# ==========================================
if should_run:
    # --- 1. 校验逻辑 ---
    errors = []
    df_a = st.session_state.df_a
    df_b = st.session_state.df_b
    
    # 简单模拟校验
    # 必填列检查 (简化)
    if '工时' not in "".join(df_a.columns) and '交付工时' not in "".join(df_a.columns):
        errors.append({'严重级': 'High', '信息': '表A缺少[工时]列'})
    
    # 阈值检查 (模拟)
    # 假设如果工时列存在，检查总和
    try:
        # 这里为了演示，假设有任意空值就报错
        if df_a.isnull().values.any() or df_b.isnull().values.any():
             # 收集具体空值位置
             for c in df_a.columns:
                 if df_a[c].isnull().any(): errors.append({'严重级': 'Med', '位置': '表A', '列': c, '信息': '存在空值'})
             for c in df_b.columns:
                 if df_b[c].isnull().any(): errors.append({'严重级': 'Med', '位置': '表B', '列': c, '信息': '存在空值'})
    except: pass

    # --- 2. 校验结果分支 ---
    if errors:
        # 🔴 发现错误 -> 阻断
        st.session_state.error_df = pd.DataFrame(errors)
        st.session_state.need_manual_trigger = True # 开启手动锁
        st.session_state.calc_result_zip = None
        st.rerun() # 立即刷新，进入报错界面
    
    else:
        # 🟢 校验通过 -> 计算
        progress = st.progress(0, "启动计算引擎...")
        try:
            time.sleep(0.3)
            progress.progress(30, "数据清洗与映射...")
            time.sleep(0.3)
            progress.progress(60, "费率计算与合并...")
            
            # --- 模拟生成结果 ---
            buffer = io.BytesIO()
            with zipfile.ZipFile(buffer, 'w') as zf:
                with pd.ExcelWriter(io.BytesIO(), engine='xlsxwriter') as w:
                    df_a.to_excel(w, sheet_name='结果')
                    w.book.io.seek(0)
                    zf.writestr("结算报表.xlsx", w.book.io.read())
            
            st.session_state.calc_result_zip = buffer.getvalue()
            progress.progress(100, "完成")
            time.sleep(0.2)
            st.rerun() # 刷新进入结果页
            
        except Exception as e:
            st.error(f"计算崩溃: {e}")

# ==========================================
# 6. 界面渲染：报错态 / 结果态
# ==========================================

# A. 报错态界面 (当有 error_df 时)
if st.session_state.error_df is not None:
    st.markdown("""<div class="error-card"><h3 style="margin:0; color:#da3633">🚨 校验失败</h3></div>""", unsafe_allow_html=True)
    
    # 错误详情
    st.dataframe(st.session_state.error_df, use_container_width=True)
    
    # 操作区：下载报告 或 在线修复
    c_err1, c_err2 = st.columns(2)
    c_err1.download_button("📥 下载错误报告", st.session_state.error_df.to_csv(), "error_report.csv", use_container_width=True)
    
    @st.dialog("🛠️ 在线修复数据", width="large")
    def open_fix_dialog():
        t1, t2 = st.tabs(["表A (交付)", "表B (差旅)"])
        with t1:
            ga = AgGrid(st.session_state.df_a, editable=True, height=400, fit_columns_on_grid_load=False)
        with t2:
            gb = AgGrid(st.session_state.df_b, editable=True, height=400, fit_columns_on_grid_load=False)
            
        if st.button("💾 保存并重试", type="primary"):
            st.session_state.df_a = pd.DataFrame(ga['data'])
            st.session_state.df_b = pd.DataFrame(gb['data'])
            # 关键：在线修完，视为手动触发一次
            st.session_state.error_df = None
            st.session_state.need_manual_trigger = False
            st.rerun()

    if c_err2.button("🛠️ 在线修复 (弹窗)", type="primary", use_container_width=True):
        open_fix_dialog()

# B. 结果态界面 (当有 result_zip 时)
if st.session_state.calc_result_zip is not None:
    with st.container(border=True):
        st.success("✅ 报表已生成")
        st.download_button(
            "📦 下载结算报表 (ZIP)", 
            st.session_state.calc_result_zip, 
            "结算报表.zip", 
            "application/zip", 
            type="primary", 
            use_container_width=True
        )
