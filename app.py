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
    
    /* 统一上传容器样式 - 状态切换不跳动 */
    /* 实际上我们直接用 st.container(height=...) 控制，CSS仅辅助 */
    
    /* 文件卡片样式 */
    .file-card-styled { 
        background: #21262d; 
        border-left: 4px solid #238636; 
        border-radius: 4px; 
        padding: 15px; 
        width: 100%;
        height: 100%; /* 充满容器 */
        display: flex; 
        align-items: center; 
        justify-content: space-between; 
    }
    
    /* 右上角极简删除按钮 */
    .close-btn { 
        cursor: pointer; color: #666; font-size: 1.2rem; line-height: 1; padding: 5px;
    }
    .close-btn:hover { color: #da3633; }

    /* 错误舱 */
    .error-box { border: 1px solid var(--red); background: rgba(218, 54, 51, 0.1); border-radius: 8px; padding: 1.5rem; margin-top: 1rem; }
    
    /* 幽灵按钮 */
    .ghost-btn button { border: 1px dashed #444 !important; color: #888 !important; background: transparent !important; padding: 0.2rem 0.8rem !important; height: auto !important; font-size: 0.8rem !important; }
    .ghost-btn button:hover { border-color: var(--red) !important; color: var(--red) !important; }

    /* Dialog 修正 */
    div[data-testid="stDialog"] > div[role="dialog"] { width: 80vw !important; max-width: 1200px !important; margin: auto !important; }
    
    /* 隐藏默认上传组件的文件列表，由我们自己的卡片接管 */
    div[data-testid="stFileUploader"] section > div:first-child { display: none; }
    div[data-testid="stFileUploader"] { padding-top: 15px; } /* 微调位置 */
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
# 2. 辅助函数
# ==========================================
def reset_system():
    st.session_state.clear()
    st.rerun()

def load_file_content(file_obj, key):
    if file_obj:
        try:
            if file_obj.name.endswith('.csv'): df = pd.read_csv(file_obj)
            else: df = pd.read_excel(file_obj)
            
            # === 核心修复：强力清洗列名 ===
            # 去除空格，解决 "金额 " 无法识别的问题
            df.columns = [str(c).strip() for c in df.columns]
            
            # 注入系统行号
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
with st.sidebar:
    st.header("⚙️ 参数配置")
    PRICE_PER_DAY = st.number_input("人力单价 (元/天)", value=1500, step=100)
    MIN_HOURS = st.number_input("工时阈值 (小时)", value=100)
    SUBSIDY_TAG = st.text_input("补助关键词", "差旅补助")

st.title("😈 淡藤财务报表 Pro")

# --- Zone 2: 数据源控制台 (单框原地切换版) ---
with st.container(border=True):
    c_h1, c_h2 = st.columns([8, 1])
    c_h1.markdown("### 📂 数据源控制台")
    with c_h2:
        st.markdown('<div class="ghost-btn">', unsafe_allow_html=True)
        if st.button("🗑️ 重置"): reset_system()
        st.markdown('</div>', unsafe_allow_html=True)

    c_u1, c_u2 = st.columns(2)
    
    def render_one_box_slot(col, key, title):
        data = st.session_state.data_store[key]
        has_file = data['df'] is not None
        
        with col:
            # === 核心修复：使用固定高度容器，实现"同一个框" ===
            # height=180 保证了无论是有文件还是无文件，框的大小都不变，不会跳动
            with st.container(height=180, border=True):
                if not has_file:
                    # 状态 A: 待上传
                    # 使用 markdown 模拟标题，label_visibility="collapsed" 隐藏自带 label
                    st.markdown(f"<div style='text-align:center; color:#888; margin-top:30px; margin-bottom:10px;'>{title}</div>", unsafe_allow_html=True)
                    f = st.file_uploader(title, type=['xlsx', 'csv'], key=f"uploader_{key}", label_visibility="collapsed")
                    if f: load_file_content(f, key)
                else:
                    # 状态 B: 文件卡片 (美化版)
                    # 左右布局：左边信息，右边小X
                    c_info, c_close = st.columns([9, 1])
                    with c_info:
                        st.markdown(f"""
                        <div style="display:flex; flex-direction:column; justify-content:center; height:100%; padding-left:10px;">
                            <div style="font-size:0.8rem; color:#8b949e;">{title.split(':')[0]}</div>
                            <div style="font-size:1.1rem; font-weight:bold; color:#fff; overflow:hidden; text-overflow:ellipsis; white-space:nowrap;" title="{data['name']}">
                                📄 {data['name']}
                            </div>
                            <div style="font-size:0.8rem; color:#238636; margin-top:5px;">✓ 已加载 {len(data['df'])} 行</div>
                        </div>
                        """, unsafe_allow_html=True)
                    with c_close:
                        st.markdown('<div class="close-btn" title="移除">✕</div>', unsafe_allow_html=True)
                        # 这是一个不可见的覆盖按钮，用于触发逻辑
                        if st.button("Del", key=f"del_{key}", help="移除文件"): 
                            clear_file(key)
                        # 稍微调整按钮位置使其覆盖X (Trick: Streamlit按钮难以自定义，这里用透明按钮覆盖或者直接放下方)
                        # 为保证稳定，我们在下方放一个文字链接作为备选
                    
    # 按照你的要求：Source A 名称改为 投入明细
    render_one_box_slot(c_u1, 'A', "Source A: 投入明细")
    render_one_box_slot(c_u2, 'B', "Source B: 差旅明细")

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
        # 区分错误类型
        fixable_df = err_df[err_df['类型'] == '数据错误']
        logic_df = err_df[err_df['类型'] == '逻辑错误']
        
        st.markdown(f"""
        <div class="error-box">
            <h3 style="color:#ff7b72; margin:0">🚨 校验失败</h3>
            <p>发现 <b>{len(fixable_df)}</b> 个数据项错误（可修复），<b>{len(logic_df)}</b> 个计算逻辑错误（请检查源文件）。</p>
        </div>
        """, unsafe_allow_html=True)
        
        # 显示错误详情
        st.dataframe(err_df[['类型', '来源', '行号', '信息']], use_container_width=True, hide_index=True)
        
        c_act1, c_act2 = st.columns(2)
        c_act1.download_button("📥 下载错误清单", err_df.to_csv(index=False).encode('utf-8-sig'), "error.csv", "text/csv", use_container_width=True)
        
        # === 核心修复：弹窗逻辑 ===
        @st.dialog("🛠️ 在线修复", width="large")
        def surgical_fix_dialog():
            if len(logic_df) > 0:
                st.warning("⚠️ 提示：列表中的'逻辑错误'无法在此修复，请检查源文件列名或阈值设置。")
            
            # 只有当存在可修复错误时才尝试获取数据
            if fixable_df.empty:
                st.info("当前没有可在线修复的数据行错误。")
                if st.button("关闭"): st.rerun()
                return

            def get_fix_df(src):
                # 过滤出该来源下类型为'数据错误'的行
                target_ids = fixable_df[fixable_df['来源'] == src]['_sys_id'].unique()
                if len(target_ids) == 0: return pd.DataFrame()
                full_df = st.session_state.data_store[src.split()[-1]]['df'] # 'Source A' -> 'A'
                return full_df[full_df['_sys_id'].isin(target_ids)].copy()

            df_a_fix = get_fix_df('Source A')
            df_b_fix = get_fix_df('Source B')
            
            t1, t2 = st.tabs([f"A ({len(df_a_fix)})", f"B ({len(df_b_fix)})"])
            new_a, new_b = None, None
            
            with t1:
                if not df_a_fix.empty:
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
                # 回写逻辑
                if new_a is not None:
                    res = pd.DataFrame(new_a)
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

        # 只有在有数据错误时才显示修复按钮，否则只显示警告
        if not fixable_df.empty:
            if c_act2.button("🛠️ 打开在线修复", type="primary", use_container_width=True):
                surgical_fix_dialog()
        else:
            c_act2.error("⚠️ 只有逻辑错误，无法在线修复，请检查源文件。")

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
    
    # 扩展的关键词库
    KEYS = {
        'A_SPM': ['SPM', '项目编号', '标识符', 'Project Code'],
        'A_HRS': ['工时', '交付工时', '投入工时', 'Hours', 'Workload'],
        'A_NAME': ['姓名', '人员', '员工姓名', 'Name', 'User'],
        'B_SPM': ['SPM', '项目编号', '费用归属项目', 'Project'],
        'B_AMT': ['金额', '报销金额', '总金额', '费用金额', 'Amount', 'Cost'],
        'B_NAME': ['姓名', '报销人', '出差人', '申请人', 'User']
    }

    def fc(df, keys):
        for k in keys: 
            if k in df.columns: return k
        return None
    
    ca_spm, ca_hrs, ca_name = fc(df_a, KEYS['A_SPM']), fc(df_a, KEYS['A_HRS']), fc(df_a, KEYS['A_NAME'])
    cb_spm, cb_amt, cb_name = fc(df_b, KEYS['B_SPM']), fc(df_b, KEYS['B_AMT']), fc(df_b, KEYS['B_NAME'])
    
    # 辅助列（非必填）
    ca_proj = fc(df_a, ['项目', '所属项目']) or df_a.columns[0]
    ca_range = fc(df_a, ['人事范围']) or df_a.columns[0]
    ca_contract = fc(df_a, ['合同主体']) or df_a.columns[0]
    ca_sales = fc(df_a, ['销售', '销售人员']) or df_a.columns[0]
    ca_dept = fc(df_a, ['销售部门', '部门']) or df_a.columns[0]
    cb_type = fc(df_b, ['产品类型', '费用类型']) or df_b.columns[0]
    
    def add_err(err_type, src, row_id, msg):
        errors.append({'类型': err_type, '来源': src, '_sys_id': row_id, '行号': row_id if isinstance(row_id, int) else '-', '信息': msg})

    # R1: 缺列 (逻辑错误)
    # 增加调试信息：如果缺失，提示当前有哪些列
    if not all([ca_spm, ca_hrs, ca_name]): 
        msg = f'缺失关键列(SPM/工时/姓名)。当前列名: {list(df_a.columns)}'
        add_err('逻辑错误', 'Source A', '-', msg)
    if not all([cb_spm, cb_amt, cb_name]): 
        msg = f'缺失关键列(SPM/金额/姓名)。当前列名: {list(df_b.columns)}'
        add_err('逻辑错误', 'Source B', '-', msg)
    
    if not errors: # 只有列存在才继续细致校验
        # 预处理
        df_a[ca_hrs] = pd.to_numeric(df_a[ca_hrs], errors='coerce').fillna(0)
        if df_b[cb_amt].dtype == object:
            df_b[cb_amt] = df_b[cb_amt].astype(str).str.replace(',', '')
        df_b[cb_amt] = pd.to_numeric(df_b[cb_amt], errors='coerce').fillna(0)
        
        # === 数据级错误 (可修复) ===
        # 1. 负数
        for i,r in df_a[df_a[ca_hrs]<0].iterrows(): add_err('数据错误','Source A', r['_sys_id'], '工时为负')
        for i,r in df_b[df_b[cb_amt]<0].iterrows(): add_err('数据错误','Source B', r['_sys_id'], '金额为负')
        # 2. 空值 (SPM)
        for i,r in df_a[df_a[ca_spm].isnull() | (df_a[ca_spm]=='')].iterrows(): 
            add_err('数据错误','Source A', r['_sys_id'], 'SPM为空')
            
        # === 逻辑级错误 (不可修复) ===
        # 3. 阈值不足
        for n,h in df_a.groupby(ca_name)[ca_hrs].sum().items():
            if h < MIN_HOURS: 
                add_err('逻辑错误','Source A', '-', f'人员[{n}]总工时({h})低于阈值({MIN_HOURS})')
            
    time.sleep(0.3)
    if errors:
        progress.empty()
        err_df_raw = pd.DataFrame(errors)
        st.session_state.error_report = err_df_raw
        st.session_state.block_auto_run = True
        st.rerun()
    else:
        # 计算
        progress.progress(50, "计算中...")
        df_a['key'] = df_a[ca_name].astype(str)+"_"+df_a[ca_spm].astype(str)
        df_b['key'] = df_b[cb_name].astype(str)+"_"+df_b[cb_spm].astype(str)
        
        # 业务计算逻辑 (复用之前的完整逻辑)
        # 1. 聚合 A
        agg_rules = {ca_hrs: 'sum'}
        for col in [ca_proj, ca_range, ca_contract, ca_sales, ca_dept]:
            if col: agg_rules[col] = 'first'
        df_a_gp = df_a.groupby([ca_name, ca_spm], as_index=False).agg(agg_rules)
        
        # 2. 拆分 B
        is_sub = df_b[cb_type].astype(str).str.contains(SUBSIDY_TAG, na=False)
        grp_b = [cb_name, cb_spm]
        df_sub = df_b[is_sub].groupby(grp_b)[cb_amt].sum().reset_index(name='差旅补助')
        df_fee = df_b[~is_sub].groupby(grp_b)[cb_amt].sum().reset_index(name='差旅费控平台')
        
        # 3. 合并
        for df in [df_a_gp, df_sub, df_fee]:
            col_key_spm = ca_spm if ca_spm in df.columns else cb_spm
            df[col_key_spm] = df[col_key_spm].astype(str)
            
        res = pd.merge(df_a_gp, df_sub, left_on=[ca_name, ca_spm], right_on=[cb_name, cb_spm], how='left')
        res = pd.merge(res, df_fee, left_on=[ca_name, ca_spm], right_on=[cb_name, cb_spm], how='left')
        res = res.fillna(0)
        
        res['支持时间(人天)'] = res[ca_hrs] / 8
        res['人力费用'] = res['支持时间(人天)'] * PRICE_PER_DAY
        res['结算费用合计'] = res['人力费用'] + res['差旅补助'] + res['差旅费控平台']
        
        # 生成分表
        t3_map = {
            ca_name: '人员', ca_proj: '所属项目', ca_range: '人事范围',
            ca_spm: 'SPM', ca_contract: '合同主体', ca_sales: '销售人员',
            ca_dept: '销售部门', ca_hrs: '耗时(小时)'
        }
        t3 = res.rename(columns=t3_map)
        final_cols = ['序号', '人员', '所属项目', '人事范围', 'SPM', '合同主体', '销售人员', '销售部门',
                      '差旅补助', '差旅费控平台', '耗时(小时)', '支持时间(人天)', '人力费用', '结算费用合计']
        t3.insert(0, '序号', range(1, len(t3)+1))
        t3 = t3[[c for c in final_cols if c in t3.columns]]
        
        dim_cols = ['人事范围', '合同主体', '销售部门']
        valid_dims = [c for c in dim_cols if c in t3.columns]
        if valid_dims:
            t2 = t3.groupby(valid_dims).agg({'结算费用合计': 'sum', '支持时间(人天)': 'sum'}).reset_index()
            t2.columns = ['销售公司', '采购公司', '采购部门', '金额(含税,单位:元)', '工作量(人天)']
            t2.insert(0, '序号', range(1, len(t2)+1))
        else:
            t2 = pd.DataFrame({'提示': ['缺少维度字段']})
            
        t1 = t3.groupby('人员')['耗时(小时)'].sum().reset_index()
        t1.rename(columns={'耗时(小时)': '项目工时'}, inplace=True)
        t1.insert(0, '序号', range(1, len(t1)+1))
        
        # 打包
        def to_bytes(df):
            b = io.BytesIO()
            df.to_excel(b, index=False)
            return b.getvalue()
            
        st.session_state.result_files = {'t1': to_bytes(t1), 't2': to_bytes(t2), 't3': to_bytes(t3)}
        
        buffer = io.BytesIO()
        with zipfile.ZipFile(buffer, 'w') as zf:
            zf.writestr("表1_工时统计.xlsx", st.session_state.result_files['t1'])
            zf.writestr("表2_结算汇总.xlsx", st.session_state.result_files['t2'])
            zf.writestr("表3_详细明细.xlsx", st.session_state.result_files['t3'])
            
        st.session_state.result_zip = buffer.getvalue()
        st.session_state.is_calculated = True
        progress.progress(100)
        st.rerun()

if st.session_state.is_calculated:
    with st.container(border=True):
        st.success("✅ 生成完毕")
        st.download_button("📦 批量下载 (ZIP)", st.session_state.result_zip, "report.zip", type="primary", use_container_width=True)
        st.markdown("---")
        c1, c2, c3 = st.columns(3)
        files = st.session_state.result_files
        if 't1' in files: c1.download_button("📥 表1", files['t1'], "t1.xlsx", use_container_width=True)
        if 't2' in files: c2.download_button("📥 表2", files['t2'], "t2.xlsx", use_container_width=True)
        if 't3' in files: c3.download_button("📥 表3", files['t3'], "t3.xlsx", use_container_width=True)
