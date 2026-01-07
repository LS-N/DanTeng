import streamlit as st
import pandas as pd
import io
import time
import zipfile
from st_aggrid import AgGrid, GridOptionsBuilder

# ==========================================
# 0. 全局配置 & CSS 美学重构
# ==========================================
st.set_page_config(page_title="淡藤财务报表 Pro", page_icon="😈", layout="wide", initial_sidebar_state="collapsed")

st.markdown("""
<style>
    :root { --bg-color: #0d1117; --card-bg: #161b22; --accent: #238636; --text: #c9d1d9; --red: #da3633; }
    .stApp { background-color: var(--bg-color); color: var(--text); }
    
    /* Zone 2: 统一上传槽位样式 (增高，无滚动条) */
    .upload-box {
        border: 1px dashed #444;
        border-radius: 8px;
        padding: 20px;
        min-height: 180px; /* 增高高度 */
        display: flex;
        flex-direction: column;
        justify-content: center;
        align-items: center;
        background-color: rgba(22, 27, 34, 0.5);
        transition: all 0.3s;
        position: relative;
    }
    .upload-box:hover {
        border-color: var(--accent);
        background-color: rgba(35, 134, 54, 0.05);
    }
    .upload-box-error {
        border-color: var(--red) !important;
        background-color: rgba(218, 54, 51, 0.05) !important;
    }
    
    /* 文件卡片美化 (深色磨砂质感) */
    .file-card-styled { 
        background: #21262d; 
        border-left: 4px solid #238636;
        border-radius: 4px; 
        padding: 15px; 
        width: 100%;
        display: flex;
        align-items: center;
        justify-content: space-between;
        box-shadow: 0 4px 6px rgba(0,0,0,0.3);
    }
    
    /* 弱化显示的删除按钮 X */
    .small-close-btn button {
        border: none !important;
        background: transparent !important;
        color: #666 !important;
        font-size: 1.2rem !important;
        padding: 0 !important;
        line-height: 1 !important;
        min-height: 0px !important;
    }
    .small-close-btn button:hover {
        color: var(--red) !important;
    }

    /* 错误舱 */
    .error-box { 
        border: 1px solid var(--red); background: rgba(218, 54, 51, 0.1); 
        border-radius: 8px; padding: 1.5rem; margin-top: 1rem;
    }
    
    /* 幽灵按钮 */
    .ghost-btn button {
        border: 1px dashed #444 !important; color: #888 !important; background: transparent !important;
        padding: 0.2rem 0.8rem !important; height: auto !important; font-size: 0.8rem !important;
    }
    .ghost-btn button:hover { border-color: var(--red) !important; color: var(--red) !important; }

    /* Dialog 修正 */
    div[data-testid="stDialog"] > div[role="dialog"] { 
        width: 80vw !important; max-width: 1200px !important; margin: auto !important;
    }
    
    /* 隐藏默认上传组件列表 */
    div[data-testid="stFileUploader"] section > div:first-child { display: none; }
    div[data-testid="stFileUploader"] { padding-top: 0px; }
    
    /* 强制隐藏 st.container 的滚动条 (如果有) */
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
if 'result_files' not in st.session_state: st.session_state.result_files = {}
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

# --- Zone 2: 数据源控制台 (修复版) ---
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
            # 动态计算 CSS 类名
            box_class = "upload-box"
            if is_error and has_file: box_class += " upload-box-error"
            
            # 使用 markdown 构建容器 div
            st.markdown(f'<div class="{box_class}">', unsafe_allow_html=True)
            
            if not has_file:
                # 状态 A: 待上传 (居中显示)
                # 使用 transparent container 占位
                st.markdown(f"<div style='text-align:center; color:#8b949e; font-weight:600; margin-bottom:10px;'>{title}</div>", unsafe_allow_html=True)
                st.markdown(f"<div style='opacity:0.6; font-size:0.8rem; margin-bottom:15px;'>支持 .xlsx / .csv</div>", unsafe_allow_html=True)
                f = st.file_uploader(title, type=['xlsx', 'csv'], key=f"uploader_{key}", label_visibility="collapsed")
                if f: load_file_content(f, key)
            else:
                # 状态 B: 已上传 (美化卡片)
                # 这里的布局：左边是信息，右上角是 X
                st.markdown(f"""
                <div class="file-card-styled">
                    <div style="display:flex; flex-direction:column;">
                        <span style="font-size:0.7rem; color:#8b949e; margin-bottom:4px;">{title.split(':')[0]}</span>
                        <span style="font-weight:bold; font-size:1rem; color:#fff; word-break:break-all;">{data['name']}</span>
                        <span style="font-size:0.8rem; color:#238636; margin-top:5px;">✓ {len(data['df'])} 行数据已加载</span>
                    </div>
                </div>
                """, unsafe_allow_html=True)
                
                # 右上角弱化删除按钮 (使用 columns 布局欺骗视觉)
                # 注意：Streamlit 按钮很难绝对定位，这里我们在卡片下方放一个极简的"撤销/删除"链接样式按钮
                c_pad, c_del = st.columns([8, 1])
                with c_del:
                    st.markdown('<div class="small-close-btn">', unsafe_allow_html=True)
                    if st.button("✕", key=f"del_{key}", help="移除此文件"):
                        clear_file(key)
                    st.markdown('</div>', unsafe_allow_html=True)

            st.markdown('</div>', unsafe_allow_html=True) # 关闭 upload-box

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
        st.markdown(f"""
        <div class="error-box">
            <h3 style="color:#ff7b72; margin:0">🚨 校验失败：发现 {len(err_df)} 处错误</h3>
            <p>请修复以下问题。</p>
        </div>
        """, unsafe_allow_html=True)
        st.dataframe(err_df, use_container_width=True, hide_index=True)
        
        c_act1, c_act2 = st.columns(2)
        c_act1.download_button("📥 下载错误清单", err_df.to_csv(index=False).encode('utf-8-sig'), "error.csv", "text/csv", use_container_width=True)
        
        @st.dialog("🛠️ 在线修复", width="large")
        def surgical_fix_dialog():
            # === Bug 修复：安全获取索引 ===
            def get_indices(src):
                # 必须先确保列存在
                if '来源' not in err_df.columns or '原表行号' not in err_df.columns:
                    return []
                rows = pd.to_numeric(err_df[err_df['来源']==src]['原表行号'], errors='coerce').dropna()
                return rows.unique().astype(int) - 2 if not rows.empty else []

            idx_a, idx_b = get_indices('Source A'), get_indices('Source B')
            df_a_fix = st.session_state.data_store['A']['df'].iloc[idx_a].copy() if len(idx_a) else pd.DataFrame()
            df_b_fix = st.session_state.data_store['B']['df'].iloc[idx_b].copy() if len(idx_b) else pd.DataFrame()
            
            t1, t2 = st.tabs([f"A ({len(df_a_fix)})", f"B ({len(df_b_fix)})"])
            new_a, new_b = None, None
            
            with t1:
                if not df_a_fix.empty:
                    # fillna 防止前端崩溃
                    gb = GridOptionsBuilder.from_dataframe(df_a_fix.fillna(""))
                    gb.configure_default_column(editable=True)
                    new_a = AgGrid(df_a_fix.fillna(""), gridOptions=gb.build(), height=300, key='fa')['data']
                else: st.info("无行级错误 (可能是缺失列或全局错误)")
            with t2:
                if not df_b_fix.empty:
                    # fillna 防止前端崩溃
                    gb = GridOptionsBuilder.from_dataframe(df_b_fix.fillna(""))
                    gb.configure_default_column(editable=True)
                    new_b = AgGrid(df_b_fix.fillna(""), gridOptions=gb.build(), height=300, key='fb')['data']
                else: st.info("无行级错误")
            
            if st.button("💾 保存并重算", type="primary"):
                if new_a is not None:
                    res = pd.DataFrame(new_a)
                    for i, r in res.iterrows(): st.session_state.data_store['A']['df'].iloc[df_a_fix.index[i]] = r
                if new_b is not None:
                    res = pd.DataFrame(new_b)
                    for i, r in res.iterrows(): st.session_state.data_store['B']['df'].iloc[df_b_fix.index[i]] = r
                
                st.session_state.error_report = None
                st.session_state.block_auto_run = False
                st.rerun()

        if c_act2.button("🛠️ 打开在线修复", type="primary", use_container_width=True):
            surgical_fix_dialog()

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
    
    # 查找列
    def fc(df, ks): 
        for k in ks: 
            if k in df.columns: return k
    
    ca_spm, ca_hrs, ca_name = fc(df_a, ['SPM','项目编号']), fc(df_a, ['工时','交付工时']), fc(df_a, ['姓名','人员'])
    cb_spm, cb_amt, cb_name = fc(df_b, ['SPM','费用归属项目']), fc(df_b, ['金额','报销金额']), fc(df_b, ['姓名','报销人'])
    
    # 统一错误添加函数 (防止 KeyError)
    def add_err(lvl, src, row, msg):
        errors.append({'严重级': lvl, '来源': src, '原表行号': row, '信息': msg})

    # R1 校验
    if not all([ca_spm, ca_hrs, ca_name]): 
        add_err('阻断', 'Source A', '-', '缺失关键列(SPM/工时/姓名)')
    if not all([cb_spm, cb_amt, cb_name]): 
        add_err('阻断', 'Source B', '-', '缺失关键列(SPM/金额/姓名)')
    
    if not errors:
        df_a[ca_hrs] = pd.to_numeric(df_a[ca_hrs], errors='coerce').fillna(0)
        df_b[cb_amt] = pd.to_numeric(df_b[cb_amt].astype(str).str.replace(',',''), errors='coerce').fillna(0)
        
        # 负数检查
        for i,r in df_a[df_a[ca_hrs]<0].iterrows(): add_err('阻断','Source A', i+2, '工时负数')
        for i,r in df_b[df_b[cb_amt]<0].iterrows(): add_err('阻断','Source B', i+2, '金额负数')
        # 空值检查
        for i,r in df_a[df_a[ca_spm].isnull() | (df_a[ca_spm]=='')].iterrows(): add_err('阻断','Source A', i+2, 'SPM空')
        # 阈值
        for n,h in df_a.groupby(ca_name)[ca_hrs].sum().items():
            if h < MIN_HOURS: add_err('阻断','Source A', '-', f'{n}工时不足')
            
    time.sleep(0.3)
    if errors:
        progress.empty()
        # 强制标准化列名，防止 KeyError
        err_df_raw = pd.DataFrame(errors)
        # 确保关键列存在
        for c in ['严重级', '来源', '原表行号', '信息']:
            if c not in err_df_raw.columns: err_df_raw[c] = '-'
            
        st.session_state.error_report = err_df_raw
        st.session_state.block_auto_run = True
        st.rerun()
    else:
        # 计算
        progress.progress(50, "计算中...")
        df_a['key'] = df_a[ca_name].astype(str)+"_"+df_a[ca_spm].astype(str)
        df_b['key'] = df_b[cb_name].astype(str)+"_"+df_b[cb_spm].astype(str)
        
        res = df_a.copy()
        
        # 打包下载
        def to_excel(d):
            b = io.BytesIO()
            d.to_excel(b, index=False)
            return b.getvalue()
            
        st.session_state.result_zip = to_excel(res)
        st.session_state.is_calculated = True
        progress.progress(100)
        st.rerun()

if st.session_state.is_calculated:
    with st.container(border=True):
        st.success("✅ 生成完毕")
        st.download_button("📦 下载报表", st.session_state.result_zip, "report.xlsx", use_container_width=True)
