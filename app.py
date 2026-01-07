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
    div[data-testid="stDialog"] > div[role="dialog"] { 
        width: 80vw !important; 
        max-width: 1200px !important; 
        margin: auto !important;
    }
    
    /* 隐藏 Streamlit 默认的文件上传列表 */
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
if 'is_calculated' not in st.session_state: st.session_state.is_calculated = False
if 'result_zip' not in st.session_state: st.session_state.result_zip = None
if 'result_files' not in st.session_state: st.session_state.result_files = {} # 存储分项文件
if 'error_report' not in st.session_state: st.session_state.error_report = None
if 'block_auto_run' not in st.session_state: st.session_state.block_auto_run = False

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
    if file_obj:
        try:
            if file_obj.name.endswith('.csv'): df = pd.read_csv(file_obj)
            else: df = pd.read_excel(file_obj)
            df.columns = [str(c).strip() for c in df.columns]
            st.session_state.data_store[key]['df'] = df
            st.session_state.data_store[key]['name'] = file_obj.name
            
            if st.session_state.block_auto_run:
                st.session_state.error_report = None
            
            st.rerun()
        except Exception as e:
            st.error(f"文件解析失败: {e}")

def clear_file(key):
    st.session_state.data_store[key]['df'] = None
    st.session_state.data_store[key]['name'] = None
    st.session_state.is_calculated = False
    st.session_state.result_zip = None
    st.session_state.result_files = {}
    st.session_state.error_report = None
    st.session_state.block_auto_run = False
    st.rerun()

# ==========================================
# 4. 界面布局
# ==========================================

# --- Zone 1: Header ---
st.title("😈 淡藤财务报表 Pro")
st.caption("Minimalist Financial Settlement System | v3.2 Full")

# --- Zone 2: Upload Console ---
with st.container(border=True):
    c_h1, c_h2 = st.columns([8, 1])
    c_h1.markdown("### 📂 数据源控制台")
    with c_h2:
        st.markdown('<div class="ghost-btn">', unsafe_allow_html=True)
        if st.button("🗑️ 重置", help="清空所有"): reset_system()
        st.markdown('</div>', unsafe_allow_html=True)

    c_u1, c_u2 = st.columns(2)
    
    def render_upload_slot(col, key, title):
        data = st.session_state.data_store[key]
        has_file = data['df'] is not None
        is_error = st.session_state.error_report is not None
        
        with col:
            if not has_file:
                f = st.file_uploader(title, type=['xlsx', 'csv'], key=f"uploader_{key}")
                if f: load_file_content(f, key)
            else:
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
                if st.button(f"❌ 移除 {key}", key=f"del_{key}", use_container_width=True):
                    clear_file(key)

    render_upload_slot(c_u1, 'A', "Source A: 交付明细")
    render_upload_slot(c_u2, 'B', "Source B: 差旅明细")

# --- Zone 3: Validation & Action ---
st.divider()

ready_to_run = (st.session_state.data_store['A']['df'] is not None and 
                st.session_state.data_store['B']['df'] is not None)
trigger_calc = False

if ready_to_run:
    if st.session_state.is_calculated:
        st.success("✅ 校验通过，报表已生成！")
    
    elif st.session_state.error_report is not None:
        err_df = st.session_state.error_report
        st.markdown(f"""
        <div class="error-box">
            <div class="error-header">🚨 校验失败：发现 {len(err_df)} 处阻断性错误</div>
            <p style="margin-bottom:1rem; opacity:0.8;">流程已暂停。请下载清单修复源文件，或使用在线外科手术修复。</p>
        </div>
        """, unsafe_allow_html=True)
        st.dataframe(err_df, use_container_width=True, hide_index=True)
        
        c_act1, c_act2 = st.columns(2)
        c_act1.download_button("📥 下载错误清单 (Excel)", 
                             err_df.to_csv(index=False).encode('utf-8-sig'), 
                             "错误清单.csv", "text/csv", use_container_width=True)
        
        @st.dialog("🛠️ 外科手术式修复 (仅显示错误行)", width="large")
        def surgical_fix_dialog():
            st.caption("🔴 红色单元格为必修项。修改后点击保存，系统将自动合并数据并重新计算。")
            
            def get_safe_indices(source_name):
                rows = err_df[err_df['来源'] == source_name]['原表行号']
                rows_numeric = pd.to_numeric(rows, errors='coerce').dropna()
                if rows_numeric.empty: return []
                return rows_numeric.unique().astype(int) - 2

            err_indices_a = get_safe_indices('Source A')
            err_indices_b = get_safe_indices('Source B')
            
            df_a_fix = st.session_state.data_store['A']['df'].iloc[err_indices_a].copy() if len(err_indices_a)>0 else pd.DataFrame()
            df_b_fix = st.session_state.data_store['B']['df'].iloc[err_indices_b].copy() if len(err_indices_b)>0 else pd.DataFrame()
            
            t1, t2 = st.tabs([f"Source A 待修 ({len(df_a_fix)})", f"Source B 待修 ({len(df_b_fix)})"])
            new_a, new_b = None, None
            
            with t1:
                if not df_a_fix.empty:
                    gb = GridOptionsBuilder.from_dataframe(df_a_fix)
                    gb.configure_default_column(editable=True)
                    gb.configure_grid_options(getRowStyle={'background-color': '#2d1b1b'}) 
                    new_a = AgGrid(df_a_fix, gridOptions=gb.build(), height=300, key='fix_a')['data']
                else: st.info("表 A 无需特定行修复（可能是全局性错误，如工时不足，请检查原始数据）")
            with t2:
                if not df_b_fix.empty:
                    gb = GridOptionsBuilder.from_dataframe(df_b_fix)
                    gb.configure_default_column(editable=True)
                    gb.configure_grid_options(getRowStyle={'background-color': '#2d1b1b'})
                    new_b = AgGrid(df_b_fix, gridOptions=gb.build(), height=300, key='fix_b')['data']
                else: st.info("表 B 无需特定行修复")
            
            if st.button("💾 保存修复并自动重算", type="primary", use_container_width=True):
                if new_a is not None:
                    fixed_df = pd.DataFrame(new_a)
                    for i, row in fixed_df.iterrows():
                        original_idx = df_a_fix.index[i] 
                        st.session_state.data_store['A']['df'].iloc[original_idx] = row
                if new_b is not None:
                    fixed_df_b = pd.DataFrame(new_b)
                    for i, row in fixed_df_b.iterrows():
                        original_idx = df_b_fix.index[i]
                        st.session_state.data_store['B']['df'].iloc[original_idx] = row
                
                st.session_state.error_report = None
                st.session_state.block_auto_run = False 
                st.rerun()

        if c_act2.button("🛠️ 打开在线修复", type="primary", use_container_width=True):
            surgical_fix_dialog()

    elif st.session_state.block_auto_run:
        st.info("ℹ️ 源文件已更新，等待执行。")
        if st.button("▶️ 重新校验并计算", type="primary", use_container_width=True):
            trigger_calc = True
            st.session_state.block_auto_run = False
            
    else:
        trigger_calc = True

# --- 执行引擎 (Execution Engine) ---
if trigger_calc:
    progress = st.progress(0, "启动校验引擎...")
    
    # 1. 获取数据
    df_a = st.session_state.data_store['A']['df']
    df_b = st.session_state.data_store['B']['df']
    
    errors = []
    
    # 2. 列名映射与查找
    def find_col(df, candidates):
        for c in candidates: 
            if c in df.columns: return c
        return None
    
    col_a_spm = find_col(df_a, ['SPM', '项目编号', '标识符'])
    col_a_hrs = find_col(df_a, ['工时', '交付工时', '投入工时'])
    col_a_name = find_col(df_a, ['姓名', '人员', '员工姓名'])
    col_a_proj = find_col(df_a, ['项目', '所属项目']) or df_a.columns[0]
    col_a_range = find_col(df_a, ['人事范围']) or df_a.columns[0]
    col_a_contract = find_col(df_a, ['合同主体']) or df_a.columns[0]
    col_a_sales = find_col(df_a, ['销售', '销售人员']) or df_a.columns[0]
    col_a_dept = find_col(df_a, ['销售部门', '部门']) or df_a.columns[0]
    
    col_b_spm = find_col(df_b, ['SPM', '项目编号', '费用归属项目'])
    col_b_amt = find_col(df_b, ['金额', '报销金额', '总金额'])
    col_b_name = find_col(df_b, ['姓名', '报销人', '出差人'])
    col_b_type = find_col(df_b, ['产品类型', '费用类型']) or df_b.columns[0]
    
    # R1: 必填列存在性
    if not all([col_a_spm, col_a_hrs, col_a_name]): errors.append({'严重级': '阻断', '来源': 'Source A', '信息': '缺失关键列(SPM/工时/姓名)'})
    if not all([col_b_spm, col_b_amt, col_b_name]): errors.append({'严重级': '阻断', '来源': 'Source B', '信息': '缺失关键列(SPM/金额/姓名)'})
    
    if not errors:
        # R2 & R3: 数据清洗与数值校验
        df_a[col_a_hrs] = pd.to_numeric(df_a[col_a_hrs], errors='coerce').fillna(0)
        neg_rows_a = df_a[df_a[col_a_hrs] < 0]
        for i, r in neg_rows_a.iterrows():
            errors.append({'严重级': '阻断', '来源': 'Source A', '原表行号': i+2, '信息': '工时不能为负数'})
            
        if df_b[col_b_amt].dtype == object:
            df_b[col_b_amt] = df_b[col_b_amt].astype(str).str.replace(',', '')
        df_b[col_b_amt] = pd.to_numeric(df_b[col_b_amt], errors='coerce').fillna(0)
        neg_rows_b = df_b[df_b[col_b_amt] < 0]
        for i, r in neg_rows_b.iterrows():
            errors.append({'严重级': '阻断', '来源': 'Source B', '原表行号': i+2, '信息': '金额不能为负数'})
            
        # R4: 关键字段非空 (SPM)
        for i, r in df_a[df_a[col_a_spm].isnull() | (df_a[col_a_spm] == '')].iterrows():
             errors.append({'严重级': '阻断', '来源': 'Source A', '原表行号': i+2, '信息': 'SPM不能为空'})
             
        # R5: 工时阈值
        agg_hrs = df_a.groupby(col_a_name)[col_a_hrs].sum()
        for name, h in agg_hrs.items():
            if h < MIN_HOURS:
                errors.append({'严重级': '阻断', '来源': 'Source A', '原表行号': '-', '信息': f'人员[{name}] 总工时 {h} < 阈值 {MIN_HOURS}'})
                
        # R6: 孤立费用检查
        df_a['key'] = df_a[col_a_name].astype(str).str.strip() + "_" + df_a[col_a_spm].astype(str).str.strip()
        df_b['key'] = df_b[col_b_name].astype(str).str.strip() + "_" + df_b[col_b_spm].astype(str).str.strip()
        
        valid_keys = set(df_a['key'].unique())
        orphan_rows = df_b[~df_b['key'].isin(valid_keys)]
        for i, r in orphan_rows.iterrows():
             errors.append({'严重级': '阻断', '来源': 'Source B', '原表行号': i+2, '信息': f'无法匹配到交付人员: {r["key"]}'})

    time.sleep(0.3)
    
    if errors:
        progress.empty()
        st.session_state.error_report = pd.DataFrame(errors)
        st.session_state.block_auto_run = True
        st.rerun()
    else:
        # ✅ 计算核心 (Calculation Logic Restored)
        progress.progress(30, "正在计算维度数据...")
        
        # 1. 聚合 A 表
        agg_rules = {col_a_hrs: 'sum'}
        for col in [col_a_proj, col_a_range, col_a_contract, col_a_sales, col_a_dept]:
            if col: agg_rules[col] = 'first'
            
        df_a_gp = df_a.groupby([col_a_name, col_a_spm], as_index=False).agg(agg_rules)
        
        # 2. 拆分 B 表 (补助 vs 费控)
        is_sub = df_b[col_b_type].astype(str).str.contains(SUBSIDY_TAG, na=False)
        grp_b = [col_b_name, col_b_spm]
        df_sub = df_b[is_sub].groupby(grp_b)[col_b_amt].sum().reset_index(name='差旅补助')
        df_fee = df_b[~is_sub].groupby(grp_b)[col_b_amt].sum().reset_index(name='差旅费控平台')
        
        # 3. 合并
        # 统一 key 类型
        for df in [df_a_gp, df_sub, df_fee]:
            df[col_a_spm if col_a_spm in df.columns else col_b_spm] = df[col_a_spm if col_a_spm in df.columns else col_b_spm].astype(str)
            
        res = pd.merge(df_a_gp, df_sub, left_on=[col_a_name, col_a_spm], right_on=[col_b_name, col_b_spm], how='left')
        res = pd.merge(res, df_fee, left_on=[col_a_name, col_a_spm], right_on=[col_b_name, col_b_spm], how='left')
        res = res.fillna(0)
        
        # 4. 算钱
        res['支持时间(人天)'] = res[col_a_hrs] / 8
        res['人力费用'] = res['支持时间(人天)'] * PRICE_PER_DAY
        res['结算费用合计'] = res['人力费用'] + res['差旅补助'] + res['差旅费控平台']
        
        progress.progress(70, "正在生成分项报表...")
        
        # === 生成表 3 (明细) ===
        t3_cols_map = {
            col_a_name: '人员', col_a_proj: '所属项目', col_a_range: '人事范围',
            col_a_spm: 'SPM', col_a_contract: '合同主体', col_a_sales: '销售人员',
            col_a_dept: '销售部门', col_a_hrs: '耗时(小时)'
        }
        t3 = res.rename(columns=t3_cols_map)
        final_cols = ['序号', '人员', '所属项目', '人事范围', 'SPM', '合同主体', '销售人员', '销售部门',
                      '差旅补助', '差旅费控平台', '耗时(小时)', '支持时间(人天)', '人力费用', '结算费用合计']
        t3.insert(0, '序号', range(1, len(t3)+1))
        # 仅保留存在的列
        t3 = t3[[c for c in final_cols if c in t3.columns]]
        
        # === 生成表 2 (结算汇总) ===
        # 维度：人事范围、合同主体、销售部门
        dim_cols = ['人事范围', '合同主体', '销售部门']
        # 确保这些列都在 t3 中
        valid_dims = [c for c in dim_cols if c in t3.columns]
        
        if valid_dims:
            t2 = t3.groupby(valid_dims).agg({'结算费用合计': 'sum', '支持时间(人天)': 'sum'}).reset_index()
            t2.columns = ['销售公司', '采购公司', '采购部门', '金额(含税,单位:元)', '工作量(人天)']
            t2.insert(0, '序号', range(1, len(t2)+1))
        else:
            t2 = pd.DataFrame({'提示': ['缺少维度字段，无法生成结算表']})

        # === 生成表 1 (工时统计) ===
        t1 = t3.groupby('人员')['耗时(小时)'].sum().reset_index()
        t1.rename(columns={'耗时(小时)': '项目工时'}, inplace=True)
        t1.insert(0, '序号', range(1, len(t1)+1))
        
        # 5. 打包结果
        def to_bytes(df):
            out = io.BytesIO()
            with pd.ExcelWriter(out, engine='xlsxwriter') as w: df.to_excel(w, index=False)
            return out.getvalue()

        b1, b2, b3 = to_bytes(t1), to_bytes(t2), to_bytes(t3)
        
        # 存入 Session State
        st.session_state.result_files = {
            't1': b1, 't2': b2, 't3': b3
        }
        
        buffer = io.BytesIO()
        with zipfile.ZipFile(buffer, 'w') as zf:
            zf.writestr("表1_工时统计.xlsx", b1)
            zf.writestr("表2_结算汇总.xlsx", b2)
            zf.writestr("表3_详细明细.xlsx", b3)
            
        st.session_state.result_zip = buffer.getvalue()
        st.session_state.is_calculated = True
        progress.progress(100)
        time.sleep(0.2)
        st.rerun()

# --- Zone 4: Download (Only on Success) ---
if st.session_state.is_calculated and st.session_state.result_zip:
    with st.container(border=True):
        st.markdown("### 📥 报表下载")
        
        # 批量下载
        st.download_button(
            "📦 批量下载所有报表 (ZIP)", 
            st.session_state.result_zip, 
            "淡藤财务报表_汇总.zip", 
            "application/zip", 
            type="primary", 
            use_container_width=True
        )
        
        st.markdown("---")
        
        # 分项下载 (回归！)
        cols_d = st.columns(3)
        files = st.session_state.result_files
        
        if 't1' in files:
            cols_d[0].download_button("📥 表1 (工时)", files['t1'], "表1_工时统计.xlsx", use_container_width=True)
        if 't2' in files:
            cols_d[1].download_button("📥 表2 (结算)", files['t2'], "表2_结算汇总.xlsx", use_container_width=True)
        if 't3' in files:
            cols_d[2].download_button("📥 表3 (明细)", files['t3'], "表3_详细明细.xlsx", use_container_width=True)
