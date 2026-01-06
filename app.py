import streamlit as st
import pandas as pd
import io
import time
import zipfile
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode, DataReturnMode

# ==========================================
# 1. 系统配置
# ==========================================
st.set_page_config(page_title="淡藤财务报表 Pro", page_icon="😈", layout="wide", initial_sidebar_state="collapsed")

# --- 样式：保留原本的深色极简风，并适配 AgGrid ---
st.markdown("""
<style>
    :root { --bg-color: #000000; --accent-color: #4ade80; --error-color: #ef4444; }
    .stApp { background-color: var(--bg-color); color: #ffffff; }
    /* 按钮样式 */
    .stButton > button { border: 1px solid #ffffff; color: #000000; background: #ffffff; font-weight: bold; }
    .stButton > button:hover { box-shadow: 0 0 8px rgba(74, 222, 128, 0.6); }
    /* 错误提示区 */
    .error-box { border: 1px solid var(--error-color); background: #1a0505; padding: 1rem; border-radius: 6px; margin-bottom: 1rem; }
</style>
""", unsafe_allow_html=True)

st.title("😈 淡藤财务报表 Pro")
st.caption("可视化错误追踪 (Error Traceability) & 在线修正 (In-App Correction)")
st.divider()

# ==========================================
# 2. 状态管理 (Session State)
# ==========================================
# 我们需要把数据存在 Session 里，这样用户编辑后才不会丢失
if 'df_a_state' not in st.session_state: st.session_state.df_a_state = None
if 'df_b_state' not in st.session_state: st.session_state.df_b_state = None
if 'file_a_id' not in st.session_state: st.session_state.file_a_id = None
if 'file_b_id' not in st.session_state: st.session_state.file_b_id = None

# ==========================================
# 3. 数据上传与初始化
# ==========================================
with st.sidebar:
    st.header("⚙️ 参数")
    PRICE_PER_DAY = st.number_input("人力单价", 1500)
    SUBSIDY_TAG = st.text_input("补助关键词", "差旅补助")

col_up1, col_up2 = st.columns(2)
with col_up1:
    f_a = st.file_uploader("Source A: 交付明细", type=['xlsx', 'csv'], key='f_a')
with col_up2:
    f_b = st.file_uploader("Source B: 差旅明细", type=['xlsx', 'csv'], key='f_b')

def init_df(file, key_id, state_key):
    """如果上传了新文件，重置 Session 中的数据"""
    if file and file.file_id != st.session_state[key_id]:
        try:
            if file.name.endswith('.csv'): df = pd.read_csv(file)
            else: df = pd.read_excel(file)
            # 统一转字符串去空格，防止匹配错误
            df.columns = [str(c).strip() for c in df.columns]
            st.session_state[state_key] = df
            st.session_state[key_id] = file.file_id
        except Exception as e:
            st.error(f"读取失败: {e}")

init_df(f_a, 'file_a_id', 'df_a_state')
init_df(f_b, 'file_b_id', 'df_b_state')

# ==========================================
# 4. 核心逻辑：校验与修复
# ==========================================

# 字段映射配置（简化版，实际可做成 SelectBox）
# 假设这是默认映射，实际项目中可以结合之前的 SelectBox 逻辑
MAP_CFG = {
    'A': {'spm': 'SPM', 'hours': '交付工时', 'contract': '合同主体', 'user': '姓名'},
    'B': {'spm': 'SPM', 'amount': '金额', 'user': '出差人', 'type': '费用类型'}
}

def validate_and_trace(df, source_name, required_cols):
    """
    核心校验函数
    返回: (is_valid, error_list_df)
    """
    if df is None: return False, None
    
    errors = []
    # 检查必须列是否存在
    missing_cols = [c for c in required_cols if c not in df.columns]
    if missing_cols:
        return False, pd.DataFrame([{'错误类型': '列缺失', '详情': f"缺少列: {missing_cols}"}])

    # 检查空值
    for col in required_cols:
        # 找出空值行
        null_rows = df[df[col].isnull() | (df[col].astype(str).str.strip() == '')]
        if not null_rows.empty:
            for idx, row in null_rows.iterrows():
                # 生成错误快照：取这一行的所有数据转成字符串，方便查看上下文
                snapshot = " | ".join(row.astype(str).values[:5]) # 只取前5列作为预览
                errors.append({
                    '来源表': source_name,
                    'Excel行号': idx + 2, # pandas索引从0开始，Excel有表头，所以+2
                    '错误列': col,
                    '错误原因': '数值为空',
                    '行数据快照': snapshot
                })
    
    if errors:
        return False, pd.DataFrame(errors)
    return True, None

# 只有当两个表都加载后才开始处理
if st.session_state.df_a_state is not None and st.session_state.df_b_state is not None:
    
    # 1. 执行校验
    req_a = [MAP_CFG['A']['spm'], MAP_CFG['A']['hours'], MAP_CFG['A']['contract']]
    req_b = [MAP_CFG['B']['spm']]
    
    valid_a, err_a = validate_and_trace(st.session_state.df_a_state, "表A", req_a)
    valid_b, err_b = validate_and_trace(st.session_state.df_b_state, "表B", req_b)

    all_errors = pd.concat([err_a, err_b], ignore_index=True) if (err_a is not None or err_b is not None) else None

    # 2. 如果有错误 -> 进入“修复模式”
    if all_errors is not None and not all_errors.empty:
        st.markdown(f"""
        <div class="error-box">
            <h3>🚨 发现 {len(all_errors)} 处数据异常</h3>
            <p>无法继续计算。请选择：下载错误清单回源文件修改，或在下方直接修改。</p>
        </div>
        """, unsafe_allow_html=True)

        # [功能 A] 下载错误清单
        csv_buffer = io.BytesIO()
        all_errors.to_excel(csv_buffer, index=False)
        st.download_button(
            label="📥 下载详细错误清单 (Excel)",
            data=csv_buffer.getvalue(),
            file_name="错误定位清单.xlsx",
            mime="application/vnd.ms-excel",
        )

        st.divider()
        st.info("🛠️ 方式二：在线快速修复 (Ag-Grid)")
        
        # [功能 B] 在线 Ag-Grid 编辑
        # 封装一个显示编辑器的函数
        def show_editor(df, key, title):
            st.markdown(f"**{title}** (双击单元格修改，红色列为必填)")
            gb = GridOptionsBuilder.from_dataframe(df)
            gb.configure_default_column(editable=True, resizable=True)
            gb.configure_selection('single')
            # 高亮关键列
            gb.configure_columns(required_cols, header_name=f"*{required_cols}", cellStyle={'background-color': '#2a1a1a'})
            grid_options = gb.build()
            
            return AgGrid(
                df, 
                gridOptions=grid_options, 
                update_mode=GridUpdateMode.MANUAL, # 只有手动触发才更新，防止卡顿
                data_return_mode=DataReturnMode.FILTERED_AND_SORTED,
                height=300,
                theme='balham', # 适合紧凑数据
                key=key
            )

        # 根据哪里有错显示哪里
        new_df_a = None
        new_df_b = None

        col_edit1, col_edit2 = st.columns(2)
        
        # 编辑表A
        with col_edit1:
            if not valid_a:
                st.error("表A 存在错误")
                # 传入 Session 中的数据
                grid_response_a = show_editor(st.session_state.df_a_state, 'grid_a', 'Source A 编辑区')
                new_df_a = grid_response_a['data']

        # 编辑表B
        with col_edit2:
            if not valid_b:
                st.error("表B 存在错误")
                grid_response_b = show_editor(st.session_state.df_b_state, 'grid_b', 'Source B 编辑区')
                new_df_b = grid_response_b['data']

        # 保存按钮
        st.markdown("---")
        if st.button("✅ 保存修正并重新校验", type="primary", use_container_width=True):
            # 更新 Session State
            if new_df_a is not None:
                st.session_state.df_a_state = pd.DataFrame(new_df_a)
            if new_df_b is not None:
                st.session_state.df_b_state = pd.DataFrame(new_df_b)
            
            st.rerun() # 强制刷新页面，重新跑校验逻辑
            
        st.stop() # 阻断后续代码执行，直到校验通过

    # ==========================================
    # 5. 校验通过，执行清洗与计算 (Happy Path)
    # ==========================================
    st.success("✅ 数据校验通过！正在生成报表...")
    
    # 这里的代码和之前类似，直接使用 cleaned session data
    df_a = st.session_state.df_a_state.copy()
    df_b = st.session_state.df_b_state.copy()
    
    # 模拟计算延迟
    progress = st.progress(0)
    for i in range(100):
        time.sleep(0.005)
        progress.progress(i+1)
    
    # ... (此处省略具体的 groupby / merge 计算代码，逻辑同上一版) ...
    # 简单演示结果
    st.info("此处执行具体的清洗逻辑 (代码复用上一版)...")
    
    # 模拟结果
    st.metric("处理完成", "数据已就绪")
    st.balloons()
