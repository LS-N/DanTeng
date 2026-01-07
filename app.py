import streamlit as st
import pandas as pd
import io
import time
import zipfile
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode

# ==========================================
# 0. 全局配置与 CSS 皮肤
# ==========================================
st.set_page_config(page_title="淡藤财务报表 Pro", page_icon="😈", layout="wide", initial_sidebar_state="collapsed")

st.markdown("""
<style>
    /* 全局深色极客风 */
    :root { --bg-color: #0d1117; --card-bg: #161b22; --accent: #238636; --text: #c9d1d9; --red: #da3633; --border: #30363d; }
    .stApp { background-color: var(--bg-color); color: var(--text); }
    
    /* Zone 2: 上传区卡片样式 */
    .upload-zone { border: 1px dashed #444; border-radius: 8px; padding: 1.5rem; text-align: center; transition: all 0.3s; }
    .upload-zone:hover { border-color: var(--accent); background: rgba(35, 134, 54, 0.05); }
    
    /* 文件卡片 (上传后的状态) */
    .file-card { 
        background: var(--card-bg); border: 1px solid var(--accent); border-radius: 6px; padding: 1rem; 
        display: flex; align-items: center; justify-content: space-between;
    }
    .file-card-err { border-color: var(--red) !important; background: rgba(218, 54, 51, 0.05) !important; }

    /* Zone 3: 一体化错误舱 */
    .error-box { 
        border: 1px solid var(--red); background: rgba(218, 54, 51, 0.1); 
        border-radius: 8px; padding: 1.5rem; margin-top: 1rem;
    }
    .error-header { display: flex; align-items: center; gap: 0.8rem; color: #ff7b72; font-weight: bold; font-size: 1.2rem; margin-bottom: 1rem; }
    
    /* 幽灵按钮 (右上角重置) */
    .ghost-btn button {
        border: 1px dashed #444 !important; color: #888 !important; background: transparent !important;
        padding: 0.2rem 0.8rem !important; height: auto !important; font-size: 0.8rem !important;
    }
    .ghost-btn button:hover { border-color: var(--red) !important; color: var(--red) !important; }

    /* Dialog 居中与宽度优化 */
    div[data-testid="stDialog"] > div { width: 85vw !important; max-width: 1400px !important; }
    
    /* 隐藏 Streamlit 默认的文件上传列表，由我们自定义的卡片接管 */
    div[data-testid="stFileUploader"] section > div:first-child { display: none; }
</style>
""", unsafe_allow_html=True)

# ==========================================
# 1. 状态管理 (Session State)
# ==========================================
# 数据核心
if 'data_store' not in st.session_state:
    st.session_state.data_store = {
        'A': {'df': None, 'name': None},
        'B': {'df': None, 'name': None}
    }

# 状态机标志位
if 'is_calculated' not in st.session_state: st.session_state.is_calculated = False  # 是否计算成功
if 'result_zip' not in st.session_state: st.session_state.result_zip = None         # 结果缓存
if 'error_report' not in st.session_state: st.session_state.error_report = None     # 错误列表
if 'block_auto_run' not in st.session_state: st.session_state.block_auto_run = False # 阻断锁 (手动模式)

# ==========================================
# 2. 侧边栏 (参数配置)
# ==========================================
with st.sidebar:
    st.header("⚙️ 参数配置")
    PRICE_PER_DAY = st.number_input("人力单价 (元/天)", value=1500, step=100)
    MIN_HOURS = st.number_input("工时阈值 (小时)", value=100, help="低于此值将触发阻断报错")
    SUBSIDY_TAG = st.text_input("补助关键词", "差旅补助")

# ==========================================
# 3. 辅助函数库
# ==========================================
def reset_system():
    st.session_state.clear()
    st.rerun()

def load_file_content(file_obj, key):
    """读取文件并立即存入 State，防止 UI 刷新丢失"""
    if file_obj:
        try:
            if file_obj.name.endswith('.csv'): df = pd.read_csv(file_obj)
            else: df = pd.read_excel(file_obj)
            df.columns = [str(c).strip() for c in df.columns]
            
            # 存入数据
            st.session_state.data_store[key]['df'] = df
            st.session_state.data_store[key]['name'] = file_obj.name
            
            # 关键逻辑：如果是报错状态下重新上传，解锁旧错误，但保持阻断等待手动确认
            if st.session_state.block_auto_run:
                st.session_state.error_report = None # 移除红色报错框
                # block_auto_run 依然为 True，等待用户点"运行"
            
            st.rerun() # 强制刷新以切换为卡片视图
        except Exception as e:
            st.error(f"文件解析失败: {e}")

def clear_file(key):
    """删除文件"""
    st.session_state.data_store[key]['df'] = None
    st.session_state.data_store[key]['name'] = None
    st.session_state.is_calculated = False
    st.session_state.result_zip = None
    st.session_state.error_report = None
    st.session_state.block_auto_run = False # 文件都没了，恢复自动态
    st.rerun()

# ==========================================
# 4. 界面布局 (Layout Implementation)
# ==========================================

# --- Zone 1: Header ---
st.title("😈 淡藤财务报表 Pro")
st.caption("Minimalist Financial Settlement System | v3.0 Final")

# --- Zone 2: Upload Console ---
with st.container(border=True):
    # 顶部栏：标题 + 幽灵重置按钮
    c_h1, c_h2 = st.columns([8, 1])
    c_h1.markdown("### 📂 数据源控制台")
    with c_h2:
        st.markdown('<div class="ghost-btn">', unsafe_allow_html=True)
        if st.button("🗑️ 重置", help="清空所有"): reset_system()
        st.markdown('</div>', unsafe_allow_html=True)

    # 上传/卡片区域
    c_u1, c_u2 = st.columns(2)
    
    # 渲染器：根据是否有数据，渲染 上传框 OR 文件卡片
    def render_upload_slot(col, key, title):
        data = st.session_state.data_store[key]
        has_file = data['df'] is not None
        is_error = st.session_state.error_report is not None # 是否处于全局报错态
        
        with col:
            if not has_file:
                # 状态 A: 空态 (显示上传器)
                f = st.file_uploader(title, type=['xlsx', 'csv'], key=f"uploader_{key}")
                if f: load_file_content(f, key)
            else:
                # 状态 B: 实态 (显示文件卡片)
                # 如果处于报错态，卡片边框变红
                card_class = "file-card file-card-err" if is_error else "file-card"
                st.markdown(f"""
                <div class="{card_class}">
                    <div style="display:flex; align-items:center; gap:0.5rem;">
                        <span style="font-size:1.5rem;">📄</span>
                        <div>
                            <div style="font-weight:bold; font-size:0.9rem;">{data['name']}</div>
                            <div style="font-size:0.7rem; opacity:0.6;">{len(data['df'])} rows</div>
                        </div>
                    </div>
                </div>
                """, unsafe_allow_html=True)
                # 删除按钮 (Streamlit 按钮只能在 markdown 外)
                if st.button(f"❌ 移除 {key}", key=f"del_{key}", use_container_width=True):
                    clear_file(key)

    render_upload_slot(c_u1, 'A', "Source A: 交付明细")
    render_upload_slot(c_u2, 'B', "Source B: 差旅明细")

# --- Zone 3: Validation & Action ---
st.divider()

# 准备状态
ready_to_run = (st.session_state.data_store['A']['df'] is not None and 
                st.session_state.data_store['B']['df'] is not None)
trigger_calc = False

# 场景状态机
if ready_to_run:
    
    # 场景 A: 结果已生成 -> 显示成功
    if st.session_state.is_calculated:
        st.success("✅ 校验通过，报表已生成！")
    
    # 场景 B: 报错阻断态 -> 显示错误舱
    elif st.session_state.error_report is not None:
        err_df = st.session_state.error_report
        
        # 渲染一体化错误舱
        st.markdown(f"""
        <div class="error-box">
            <div class="error-header">🚨 校验失败：发现 {len(err_df)} 处阻断性错误</div>
            <p style="margin-bottom:1rem; opacity:0.8;">流程已暂停。请下载清单修复源文件，或使用在线外科手术修复。</p>
        </div>
        """, unsafe_allow_html=True)
        
        st.dataframe(err_df, use_container_width=True, hide_index=True)
        
        c_act1, c_act2 = st.columns(2)
        # 操作 1: 下载错误清单
        c_act1.download_button("📥 下载错误清单 (Excel)", 
                             err_df.to_csv(index=False).encode('utf-8-sig'), 
                             "错误清单.csv", "text/csv", use_container_width=True)
        
        # 操作 2: 外科手术修复弹窗
        # 定义弹窗逻辑
        @st.dialog("🛠️ 外科手术式修复 (仅显示错误行)", width="large")
        def surgical_fix_dialog():
            st.caption("🔴 红色单元格为必修项。修改后点击保存，系统将自动合并数据并重新计算。")
            
            # 获取错误数据的索引
            err_indices_a = err_df[err_df['来源']=='Source A']['原表行号'].unique() - 2 # Excel行号转Index
            err_indices_b = err_df[err_df['来源']=='Source B']['原表行号'].unique() - 2
            
            # 过滤出需要修复的行 (Surgical Filter)
            df_a_fix = st.session_state.data_store['A']['df'].iloc[err_indices_a].copy() if len(err_indices_a)>0 else pd.DataFrame()
            df_b_fix = st.session_state.data_store['B']['df'].iloc[err_indices_b].copy() if len(err_indices_b)>0 else pd.DataFrame()
            
            t1, t2 = st.tabs([f"Source A 待修 ({len(df_a_fix)})", f"Source B 待修 ({len(df_b_fix)})"])
            
            new_a, new_b = None, None
            
            with t1:
                if not df_a_fix.empty:
                    gb = GridOptionsBuilder.from_dataframe(df_a_fix)
                    gb.configure_default_column(editable=True)
                    # 样式：让错误行显眼
                    gb.configure_grid_options(getRowStyle={'background-color': '#2d1b1b'}) 
                    new_a = AgGrid(df_a_fix, gridOptions=gb.build(), height=300, key='fix_a')['data']
                else: st.info("表 A 无需修复")
                    
            with t2:
                if not df_b_fix.empty:
                    gb = GridOptionsBuilder.from_dataframe(df_b_fix)
                    gb.configure_default_column(editable=True)
                    gb.configure_grid_options(getRowStyle={'background-color': '#2d1b1b'})
                    new_b = AgGrid(df_b_fix, gridOptions=gb.build(), height=300, key='fix_b')['data']
                else: st.info("表 B 无需修复")
            
            if st.button("💾 保存修复并自动重算", type="primary", use_container_width=True):
                # 缝合逻辑 (Stitch back)
                if new_a is not None:
                    fixed_df = pd.DataFrame(new_a)
                    # 使用索引回写原表
                    for i, row in fixed_df.iterrows():
                        # 注意：AgGrid返回的数据索引可能重置，这里简单演示整体替换逻辑
                        # 生产环境建议匹配唯一ID，这里简化为按位置回填
                        original_idx = df_a_fix.index[i] 
                        st.session_state.data_store['A']['df'].iloc[original_idx] = row
                
                if new_b is not None:
                    fixed_df_b = pd.DataFrame(new_b)
                    for i, row in fixed_df_b.iterrows():
                        original_idx = df_b_fix.index[i]
                        st.session_state.data_store['B']['df'].iloc[original_idx] = row
                
                # 解除阻断，触发重算
                st.session_state.error_report = None
                st.session_state.block_auto_run = False 
                st.rerun()

        if c_act2.button("🛠️ 打开在线修复", type="primary", use_container_width=True):
            surgical_fix_dialog()

    # 场景 C: 等待手动触发 (重新上传了文件，但还未计算)
    elif st.session_state.block_auto_run:
        st.info("ℹ️ 源文件已更新，等待执行。")
        if st.button("▶️ 重新校验并计算", type="primary", use_container_width=True):
            trigger_calc = True
            st.session_state.block_auto_run = False # 解锁
            
    # 场景 D: 默认自动触发 (无错误，无阻断)
    else:
        trigger_calc = True

# --- 执行引擎 (Execution Engine) ---
if trigger_calc:
    progress = st.progress(0, "启动校验引擎...")
    
    # 1. 获取数据
    df_a = st.session_state.data_store['A']['df']
    df_b = st.session_state.data_store['B']['df']
    
    errors = []
    
    # 2. 规则校验 (Implementation of Rules)
    
    # 辅助：列查找
    def find(df, keys):
        for k in keys: 
            if k in df.columns: return k
        return None
    
    col_a_spm = find(df_a, ['SPM', '项目编号', '标识符'])
    col_a_hrs = find(df_a, ['工时', '交付工时'])
    col_a_name = find(df_a, ['姓名', '人员'])
    
    col_b_spm = find(df_b, ['SPM', '项目编号', '费用归属项目'])
    col_b_amt = find(df_b, ['金额', '报销金额', '总金额'])
    col_b_name = find(df_b, ['姓名', '报销人', '出差人'])
    
    # R1: 必填列存在性
    if not all([col_a_spm, col_a_hrs, col_a_name]): errors.append({'严重级': '阻断', '来源': 'Source A', '信息': '缺失关键列(SPM/工时/姓名)'})
    if not all([col_b_spm, col_b_amt, col_b_name]): errors.append({'严重级': '阻断', '来源': 'Source B', '信息': '缺失关键列(SPM/金额/姓名)'})
    
    if not errors:
        # R2 & R3: 数据清洗与数值校验
        # 清洗 A 表工时
        df_a[col_a_hrs] = pd.to_numeric(df_a[col_a_hrs], errors='coerce').fillna(0)
        # 负数检查 A
        neg_rows_a = df_a[df_a[col_a_hrs] < 0]
        for i, r in neg_rows_a.iterrows():
            errors.append({'严重级': '阻断', '来源': 'Source A', '原表行号': i+2, '信息': '工时不能为负数'})
            
        # 清洗 B 表金额
        if df_b[col_b_amt].dtype == object:
            df_b[col_b_amt] = df_b[col_b_amt].astype(str).str.replace(',', '')
        df_b[col_b_amt] = pd.to_numeric(df_b[col_b_amt], errors='coerce').fillna(0)
        # 负数检查 B
        neg_rows_b = df_b[df_b[col_b_amt] < 0]
        for i, r in neg_rows_b.iterrows():
            errors.append({'严重级': '阻断', '来源': 'Source B', '原表行号': i+2, '信息': '金额不能为负数'})
            
        # R4: 关键字段非空 (SPM)
        for i, r in df_a[df_a[col_a_spm].isnull() | (df_a[col_a_spm] == '')].iterrows():
             errors.append({'严重级': '阻断', '来源': 'Source A', '原表行号': i+2, '信息': 'SPM不能为空'})
             
        # R5: 业务逻辑 - 工时阈值
        agg_hrs = df_a.groupby(col_a_name)[col_a_hrs].sum()
        for name, h in agg_hrs.items():
            if h < MIN_HOURS:
                # 注意：这里我们设定为阻断，如需改为警告，可不加入errors列表
                errors.append({'严重级': '阻断', '来源': 'Source A', '原表行号': '-', '信息': f'人员[{name}] 总工时 {h} < 阈值 {MIN_HOURS}'})
                
        # R6: 孤立费用检查 (匹配性)
        # 创建匹配键
        df_a['key'] = df_a[col_a_name].astype(str).str.strip() + "_" + df_a[col_a_spm].astype(str).str.strip()
        df_b['key'] = df_b[col_b_name].astype(str).str.strip() + "_" + df_b[col_b_spm].astype(str).str.strip()
        
        valid_keys = set(df_a['key'].unique())
        orphan_rows = df_b[~df_b['key'].isin(valid_keys)]
        
        for i, r in orphan_rows.iterrows():
             errors.append({'严重级': '阻断', '来源': 'Source B', '原表行号': i+2, '信息': f'无法匹配到交付人员: {r["key"]}'})

    # 3. 结果判断
    time.sleep(0.5) # 模拟计算延迟优化体验
    
    if errors:
        # ❌ 失败分支
        progress.empty()
        st.session_state.error_report = pd.DataFrame(errors)
        st.session_state.block_auto_run = True # 开启手动锁
        st.rerun()
    else:
        # ✅ 成功分支
        progress.progress(50, "正在生成报表...")
        
        # 模拟生成逻辑
        res_df = df_a.copy()
        res_df['结算金额'] = res_df[col_a_hrs] * PRICE_PER_DAY / 8
        
        # 打包
        buffer = io.BytesIO()
        with zipfile.ZipFile(buffer, 'w') as zf:
            with pd.ExcelWriter(io.BytesIO(), engine='xlsxwriter') as w:
                res_df.to_excel(w, sheet_name='结算明细')
            zf.writestr("结算报表.xlsx", buffer.getvalue())
            
        st.session_state.result_zip = buffer.getvalue()
        st.session_state.is_calculated = True
        progress.progress(100)
        time.sleep(0.2)
        st.rerun()

# --- Zone 4: Download (Only on Success) ---
if st.session_state.is_calculated and st.session_state.result_zip:
    with st.container(border=True):
        st.markdown("### 📥 报表下载")
        st.download_button(
            "📦 批量下载所有报表 (ZIP)", 
            st.session_state.result_zip, 
            "淡藤财务报表.zip", 
            "application/zip", 
            type="primary", 
            use_container_width=True
        )
        st.caption("分项下载功能将在后续版本开放")
