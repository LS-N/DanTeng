import streamlit as st
import pandas as pd
import io
import time
import zipfile
from st_aggrid import AgGrid, GridOptionsBuilder

# ==========================================
# 0. 全局配置 & CSS
# ==========================================
st.set_page_config(page_title="淡藤财务报表 Pro", page_icon="😈", layout="wide", initial_sidebar_state="collapsed")

st.markdown("""
<style>
    :root { --bg-color: #0d1117; --card-bg: #161b22; --accent: #238636; --text: #c9d1d9; --red: #da3633; }
    .stApp { background-color: var(--bg-color); color: var(--text); }
    
    /* 统一上传槽位样式 */
    .upload-box {
        border: 1px dashed #444; border-radius: 8px; padding: 20px; min-height: 180px;
        display: flex; flex-direction: column; justify-content: center; align-items: center;
        background-color: rgba(22, 27, 34, 0.5); transition: all 0.3s;
    }
    .upload-box:hover { border-color: var(--accent); background-color: rgba(35, 134, 54, 0.05); }
    .upload-box-error { border-color: var(--red) !important; background-color: rgba(218, 54, 51, 0.05) !important; }
    
    /* 文件卡片 */
    .file-card-styled { 
        background: #21262d; border-left: 4px solid #238636; border-radius: 4px; padding: 15px; width: 100%;
        display: flex; align-items: center; justify-content: space-between; box-shadow: 0 4px 6px rgba(0,0,0,0.3);
    }
    .small-close-btn button { border: none !important; background: transparent !important; color: #666 !important; font-size: 1.2rem !important; padding: 0 !important; line-height: 1 !important; min-height: 0px !important; }
    .small-close-btn button:hover { color: var(--red) !important; }

    /* 错误舱 */
    .error-box { border: 1px solid var(--red); background: rgba(218, 54, 51, 0.1); border-radius: 8px; padding: 1.5rem; margin-top: 1rem; }
    
    /* 幽灵按钮 */
    .ghost-btn button { border: 1px dashed #444 !important; color: #888 !important; background: transparent !important; padding: 0.2rem 0.8rem !important; height: auto !important; font-size: 0.8rem !important; }
    .ghost-btn button:hover { border-color: var(--red) !important; color: var(--red) !important; }

    /* Dialog 修正 */
    div[data-testid="stDialog"] > div[role="dialog"] { width: 80vw !important; max-width: 1200px !important; margin: auto !important; }
    
    /* 隐藏默认上传组件 */
    div[data-testid="stFileUploader"] section > div:first-child { display: none; }
    div[data-testid="stFileUploader"] { padding-top: 0px; }
    div[data-testid="stVerticalBlock"] > div { overflow: visible !important; }
</style>
""", unsafe_allow_html=True)

# ==========================================
# 1. 状态管理
# ==========================================
if 'data_store' not in st.session_state:
    st.session_state.data_store = {
        'A': {'df': None, 'name': None},
        'B': {'df': None, 'name': None}
    }
if 'is_calculated' not in st.session_state: st.session_state.is_calculated = False
if 'result_zip' not in st.session_state: st.session_state.result_zip = None
if 'error_report' not in st.session_state: st.session_state.error_report = None
if 'block_auto_run' not in st.session_state: st.session_state.block_auto_run = False

# ==========================================
# 2. 侧边栏 & 辅助函数
# ==========================================
with st.sidebar:
    st.header("⚙️ 参数配置")
    PRICE_PER_DAY = st.number_input("人力单价 (元/天)", value=1500, step=100)
    MIN_HOURS = st.number_input("工时阈值 (小时)", value=100)
    SUBSIDY_TAG = st.text_input("补助关键词", "差旅补助")

def reset_system():
    st.session_state.clear()
    st.rerun()

def load_file_content(file_obj, key):
    if file_obj:
        try:
            if file_obj.name.endswith('.csv'): df = pd.read_csv(file_obj)
            else: df = pd.read_excel(file_obj)
            df.columns = [str(c).strip() for c in df.columns]
            
            # === 核心改动 1: 自动注入系统行号 ===
            # 从 1 开始计数，方便人类阅读
            df['_sys_id'] = range(1, len(df) + 1)
            
            st.session_state.data_store[key]['df'] = df
            st.session_state.data_store[key]['name'] = file_obj.name
            if st.session_state.block_auto_run: st.session_state.error_report = None
            st.rerun()
        except Exception as e:
            st.error(f"解析失败: {e}")

def clear_file(key):
    st.session_state.data_store[key]['df'] = None
    st.session_state.data_store[key]['name'] = None
    st.session_state.is_calculated = False
    st.session_state.result_zip = None
    st.session_state.error_report = None
    st.session_state.block_auto_run = False
    st.rerun()

# ==========================================
# 3. 界面布局
# ==========================================
st.title("😈 淡藤财务报表 Pro")

# --- Zone 2: 数据源控制台 ---
with st.container(border=True):
    c_h1, c_h2 = st.columns([8, 1])
    c_h1.markdown("### 📂 数据源控制台")
    with c_h2:
        st.markdown('<div class="ghost-btn">', unsafe_allow_html=True)
        if st.button("🗑️ 重置"): reset_system()
        st.markdown('</div>', unsafe_allow_html=True)

    c_u1, c_u2 = st.columns(2)
    
    def render_beauty_slot(col, key, title):
        data = st.session_state.data_store[key]
        has_file = data['df'] is not None
        is_error = st.session_state.error_report is not None
        
        with col:
            box_class = "upload-box"
            if is_error and has_file: box_class += " upload-box-error"
            st.markdown(f'<div class="{box_class}">', unsafe_allow_html=True)
            
            if not has_file:
                st.markdown(f"<div style='text-align:center; color:#8b949e; font-weight:600; margin-bottom:10px;'>{title}</div>", unsafe_allow_html=True)
                f = st.file_uploader(title, type=['xlsx', 'csv'], key=f"uploader_{key}", label_visibility="collapsed")
                if f: load_file_content(f, key)
            else:
                st.markdown(f"""
                <div class="file-card-styled">
                    <div style="display:flex; flex-direction:column;">
                        <span style="font-size:0.7rem; color:#8b949e; margin-bottom:4px;">{title.split(':')[0]}</span>
                        <span style="font-weight:bold; font-size:1rem; color:#fff; word-break:break-all;">{data['name']}</span>
                        <span style="font-size:0.8rem; color:#238636; margin-top:5px;">✓ {len(data['df'])} 行数据已加载</span>
                    </div>
                </div>
                """, unsafe_allow_html=True)
                c_pad, c_del = st.columns([8, 1])
                with c_del:
                    st.markdown('<div class="small-close-btn">', unsafe_allow_html=True)
                    if st.button("✕", key=f"del_{key}", help="移除此文件"): clear_file(key)
                    st.markdown('</div>', unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)

    render_beauty_slot(c_u1, 'A', "Source A: 投入明细")
    render_beauty_slot(c_u2, 'B', "Source B: 差旅明细")

# --- Zone 3: 校验与执行 ---
st.divider()

ready_to_run = (st.session_state.data_store['A']['df'] is not None and 
                st.session_state.data_store['B']['df'] is not None)
trigger_calc = False

if ready_to_run:
    if st.session_state.is_calculated:
        st.success("✅ 校验通过，报表已生成！")
    
    elif st.session_state.error_report is not None:
        err_df = st.session_state.error_report
        # 区分 可修复(数据行错误) 和 不可修复(逻辑错误)
        fixable_df = err_df[err_df['类型'] == '数据错误']
        logic_df = err_df[err_df['类型'] == '逻辑错误']
        
        st.markdown(f"""
        <div class="error-box">
            <h3 style="color:#ff7b72; margin:0">🚨 校验失败</h3>
            <p>发现 <b>{len(fixable_df)}</b> 个数据项错误（可修复），<b>{len(logic_df)}</b> 个计算逻辑错误（请检查源文件）。</p>
        </div>
        """, unsafe_allow_html=True)
        
        # 展示错误列表
        st.dataframe(err_df[['类型', '来源', '行号', '信息']], use_container_width=True, hide_index=True)
        
        c_act1, c_act2 = st.columns(2)
        c_act1.download_button("📥 下载错误清单", err_df.to_csv(index=False).encode('utf-8-sig'), "error.csv", "text/csv", use_container_width=True)
        
        @st.dialog("🛠️ 在线修复 (仅展示可修复项)", width="large")
        def surgical_fix_dialog():
            if len(logic_df) > 0:
                st.warning(f"⚠️ 注意：存在 {len(logic_df)} 个逻辑错误（如工时不足），这些无法在线修复，请在本地调整数据后重新上传。")
            
            # === 核心逻辑 2: 仅提取"数据错误"类型的行 ===
            # 使用 _sys_id 来定位，稳如泰山
            def get_fix_df(src):
                # 找到该来源下，类型为'数据错误'的行号
                target_ids = fixable_df[fixable_df['来源'] == src]['_sys_id'].unique()
                if len(target_ids) == 0: return pd.DataFrame()
                # 过滤出这些行
                full_df = st.session_state.data_store['src_map'][src]['df']
                return full_df[full_df['_sys_id'].isin(target_ids)].copy()

            # 临时构建映射方便取数
            st.session_state.data_store['src_map'] = {
                'Source A': st.session_state.data_store['A'],
                'Source B': st.session_state.data_store['B']
            }
            
            df_a_fix = get_fix_df('Source A')
            df_b_fix = get_fix_df('Source B')
            
            t1, t2 = st.tabs([f"A ({len(df_a_fix)})", f"B ({len(df_b_fix)})"])
            new_a, new_b = None, None
            
            with t1:
                if not df_a_fix.empty:
                    # 隐藏 _sys_id 列，不让用户改
                    gb = GridOptionsBuilder.from_dataframe(df_a_fix.fillna(""))
                    gb.configure_column("_sys_id", hide=True) 
                    gb.configure_default_column(editable=True)
                    new_a = AgGrid(df_a_fix.fillna(""), gridOptions=gb.build(), height=300, key='fa')['data']
                else: st.info("无数据级错误")
            with t2:
                if not df_b_fix.empty:
                    gb = GridOptionsBuilder.from_dataframe(df_b_fix.fillna(""))
                    gb.configure_column("_sys_id", hide=True)
                    gb.configure_default_column(editable=True)
                    new_b = AgGrid(df_b_fix.fillna(""), gridOptions=gb.build(), height=300, key='fb')['data']
                else: st.info("无数据级错误")
            
            if st.button("💾 保存并重算", type="primary"):
                # 回写逻辑：根据 _sys_id 精准缝合
                if new_a is not None:
                    res = pd.DataFrame(new_a)
                    # 将 _sys_id 设为索引以便匹配
                    origin_df = st.session_state.data_store['A']['df'].set_index('_sys_id')
                    update_df = res.set_index('_sys_id')
                    origin_df.update(update_df)
                    st.session_state.data_store['A']['df'] = origin_df.reset_index()
                
                if new_b is not None:
                    res = pd.DataFrame(new_b)
                    origin_df = st.session_state.data_store['B']['df'].set_index('_sys_id')
                    update_df = res.set_index('_sys_id')
                    origin_df.update(update_df)
                    st.session_state.data_store['B']['df'] = origin_df.reset_index()
                
                st.session_state.error_report = None
                st.session_state.block_auto_run = False
                st.rerun()

        # 只有存在可修复错误时，才显示修复按钮
        if not fixable_df.empty:
            if c_act2.button("🛠️ 打开在线修复", type="primary", use_container_width=True):
                surgical_fix_dialog()
        else:
            c_act2.info("⚠️ 当前错误属于逻辑/统计错误，请检查源文件逻辑。")

    elif st.session_state.block_auto_run:
        st.info("ℹ️ 源文件已更新，等待确认...")
        if st.button("▶️ 重新校验并计算", type="primary", use_container_width=True):
            trigger_calc = True
            st.session_state.block_auto_run = False
    else:
        trigger_calc = True

# --- Zone 4: 计算引擎 ---
if trigger_calc:
    progress = st.progress(0, "启动...")
    df_a = st.session_state.data_store['A']['df']
    df_b = st.session_state.data_store['B']['df']
    errors = []
    
    def fc(df, ks): 
        for k in ks: 
            if k in df.columns: return k
    
    ca_spm, ca_hrs, ca_name = fc(df_a, ['SPM','项目编号']), fc(df_a, ['工时','交付工时']), fc(df_a, ['姓名','人员'])
    cb_spm, cb_amt, cb_name = fc(df_b, ['SPM','费用归属项目']), fc(df_b, ['金额','报销金额']), fc(df_b, ['姓名','报销人'])
    
    # 统一添加错误，增加 type 参数
    def add_err(err_type, src, row_id, msg):
        errors.append({
            '类型': err_type, 
            '来源': src, 
            '_sys_id': row_id, # 用于内部定位
            '行号': row_id if isinstance(row_id, int) else '-', # 用于展示
            '信息': msg
        })

    # R1: 缺列 (逻辑错误，直接阻断，无法行级修复)
    if not all([ca_spm, ca_hrs, ca_name]): 
        add_err('逻辑错误', 'Source A', '-', '缺失关键列(SPM/工时/姓名)')
    if not all([cb_spm, cb_amt, cb_name]): 
        add_err('逻辑错误', 'Source B', '-', '缺失关键列(SPM/金额/姓名)')
    
    if not errors: # 只有列存在才继续
        df_a[ca_hrs] = pd.to_numeric(df_a[ca_hrs], errors='coerce').fillna(0)
        df_b[cb_amt] = pd.to_numeric(df_b[cb_amt].astype(str).str.replace(',',''), errors='coerce').fillna(0)
        
        # === 数据级错误 (Data Errors) - 可修复 ===
        # 1. 负数
        for i,r in df_a[df_a[ca_hrs]<0].iterrows(): add_err('数据错误','Source A', r['_sys_id'], '工时为负')
        for i,r in df_b[df_b[cb_amt]<0].iterrows(): add_err('数据错误','Source B', r['_sys_id'], '金额为负')
        # 2. 空值 (SPM)
        for i,r in df_a[df_a[ca_spm].isnull() | (df_a[ca_spm]=='')].iterrows(): 
            add_err('数据错误','Source A', r['_sys_id'], 'SPM为空')
            
        # === 逻辑级错误 (Logic Errors) - 不可修复 ===
        # 3. 阈值不足
        for n,h in df_a.groupby(ca_name)[ca_hrs].sum().items():
            if h < MIN_HOURS: 
                add_err('逻辑错误','Source A', '-', f'人员[{n}]总工时({h})低于阈值({MIN_HOURS})')
            
    time.sleep(0.3)
    if errors:
        progress.empty()
        # 补全列防止 KeyError
        err_df_raw = pd.DataFrame(errors)
        st.session_state.error_report = err_df_raw
        st.session_state.block_auto_run = True
        st.rerun()
    else:
        # 计算
        progress.progress(50, "计算中...")
        df_a['key'] = df_a[ca_name].astype(str)+"_"+df_a[ca_spm].astype(str)
        df_b['key'] = df_b[cb_name].astype(str)+"_"+df_b[cb_spm].astype(str)
        
        res = df_a.copy()
        
        def to_excel(d):
            b = io.BytesIO()
            # 导出时去掉系统行号
            out_d = d.drop(columns=['_sys_id'], errors='ignore')
            out_d.to_excel(b, index=False)
            return b.getvalue()
            
        st.session_state.result_zip = to_excel(res)
        st.session_state.is_calculated = True
        progress.progress(100)
        st.rerun()

if st.session_state.is_calculated:
    with st.container(border=True):
        st.success("✅ 生成完毕")
        st.download_button("📦 下载报表", st.session_state.result_zip, "report.xlsx", use_container_width=True)
