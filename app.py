import streamlit as st
import pandas as pd
import io
import time
import zipfile
from st_aggrid import AgGrid, GridOptionsBuilder

# ==========================================
# 0. 全局配置 & CSS
# ==========================================
st.set_page_config(page_title="淡藤财务报表 Pro", page_icon="😈", layout="wide", initial_sidebar_state="expanded")

st.markdown("""
<style>
    :root { --bg-color: #0d1117; --card-bg: #161b22; --accent: #238636; --text: #c9d1d9; --red: #da3633; }
    .stApp { background-color: var(--bg-color); color: var(--text); }
    
    /* 上传容器 */
    .file-card-styled { 
        background: #21262d; border-left: 4px solid #238636; border-radius: 4px; padding: 15px; 
        width: 100%; height: 100%; display: flex; align-items: center; justify-content: space-between; 
    }
    .close-btn { cursor: pointer; color: #666; font-size: 1.2rem; line-height: 1; padding: 5px; }
    .close-btn:hover { color: #da3633; }

    /* 错误舱 */
    .error-box { border: 1px solid var(--red); background: rgba(218, 54, 51, 0.1); border-radius: 8px; padding: 1.5rem; margin-top: 1rem; }
    
    /* 幽灵按钮 */
    .ghost-btn button { border: 1px dashed #444 !important; color: #888 !important; background: transparent !important; padding: 0.2rem 0.8rem !important; height: auto !important; font-size: 0.8rem !important; }
    .ghost-btn button:hover { border-color: var(--red) !important; color: var(--red) !important; }

    /* 侧边栏猫猫按钮 */
    .cat-btn button {
        margin-top: 2rem; border: 1px solid #30363d !important; background: #161b22 !important; color: #c9d1d9 !important;
        border-radius: 8px !important; padding: 0.5rem 1rem !important; width: 100%; text-align: center;
    }
    .cat-btn button:hover { border-color: #a371f7 !important; color: #a371f7 !important; transform: scale(1.02); }

    /* 映射页样式 */
    .map-row { background: #1c2128; padding: 10px; border-radius: 6px; margin-bottom: 8px; border: 1px solid #30363d; display: flex; align-items: center; }
    .map-label { color: #8b949e; font-size: 0.85rem; width: 100px; }
    .map-val { font-weight: bold; color: #c9d1d9; }
    
    /* 隐藏默认上传组件 */
    div[data-testid="stFileUploader"] section > div:first-child { display: none; }
    div[data-testid="stFileUploader"] { padding-top: 15px; }
</style>
""", unsafe_allow_html=True)

# ==========================================
# 1. 状态管理 & 配置初始化
# ==========================================
if 'page' not in st.session_state: st.session_state.page = 'main'
if 'data_store' not in st.session_state:
    st.session_state.data_store = {
        'A': {'df': None, 'name': None},
        'B': {'df': None, 'name': None}
    }
if 'is_calculated' not in st.session_state: st.session_state.is_calculated = False
if 'result_zip' not in st.session_state: st.session_state.result_zip = None
if 'error_report' not in st.session_state: st.session_state.error_report = None
if 'block_auto_run' not in st.session_state: st.session_state.block_auto_run = False
if 'is_editing_mapping' not in st.session_state: st.session_state.is_editing_mapping = False # 编辑模式锁

# === 核心：初始化字段映射配置 (带所属表标记) ===
def init_mapping_config():
    if 'mapping_config' not in st.session_state:
        # 模板数据 (默认只读)
        data = [
            # --- 表1: 工时统计 ---
            {"所属表": "表1", "目标字段": "人员", "来源": "Source A", "匹配字段": "姓名", "计算逻辑": "主键"},
            {"所属表": "表1", "目标字段": "项目工时", "来源": "Source A", "匹配字段": "交付工时", "计算逻辑": "SUM聚合"},
            
            # --- 表2: 结算汇总 ---
            {"所属表": "表2", "目标字段": "销售公司", "来源": "Source A", "匹配字段": "人事范围", "计算逻辑": "维度展示"},
            {"所属表": "表2", "目标字段": "采购公司", "来源": "Source A", "匹配字段": "合同主体", "计算逻辑": "维度展示"},
            {"所属表": "表2", "目标字段": "采购部门", "来源": "Source A", "匹配字段": "销售部门", "计算逻辑": "维度展示"},
            
            # --- 表3: 详细明细 (A表部分) ---
            {"所属表": "表3", "目标字段": "SPM", "来源": "Source A", "匹配字段": "SPM", "计算逻辑": "主键 (匹配用)"},
            {"所属表": "表3", "目标字段": "工时", "来源": "Source A", "匹配字段": "交付工时", "计算逻辑": "核心计算"},
            {"所属表": "表3", "目标字段": "姓名", "来源": "Source A", "匹配字段": "姓名", "计算逻辑": "主键 (匹配用)"},
            {"所属表": "表3", "目标字段": "项目名称", "来源": "Source A", "匹配字段": "所属项目", "计算逻辑": "维度"},
            
            # --- 表3: 详细明细 (B表部分) ---
            {"所属表": "表3", "目标字段": "SPM (B)", "来源": "Source B", "匹配字段": "费用归属项目", "计算逻辑": "外键"},
            {"所属表": "表3", "目标字段": "金额", "来源": "Source B", "匹配字段": "报销金额", "计算逻辑": "SUM"},
            {"所属表": "表3", "目标字段": "姓名 (B)", "来源": "Source B", "匹配字段": "报销人", "计算逻辑": "外键"},
            {"所属表": "表3", "目标字段": "费用类型", "来源": "Source B", "匹配字段": "费用类型", "计算逻辑": "分类"},
        ]
        st.session_state.mapping_config = pd.DataFrame(data)

init_mapping_config()

# ==========================================
# 2. 辅助函数
# ==========================================
def switch_page(page_name):
    st.session_state.page = page_name
    st.rerun()

def reset_system():
    st.session_state.clear()
    st.session_state.page = 'main'
    init_mapping_config()
    st.rerun()

def load_file_content(file_obj, key):
    if file_obj:
        try:
            if file_obj.name.endswith('.csv'): df = pd.read_csv(file_obj)
            else: df = pd.read_excel(file_obj)
            df.columns = [str(c).strip() for c in df.columns]
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

# --- 核心：从配置中动态获取列名 (单一值) ---
def get_config_key(source_name, target_field):
    df_conf = st.session_state.mapping_config
    # 模糊查找目标字段 (防止表1/表2重复字段名干扰，优先匹配所属表，这里简化为全表搜)
    row = df_conf[(df_conf['来源'] == source_name) & (df_conf['目标字段'] == target_field)]
    if row.empty: return None
    return str(row.iloc[0]['匹配字段']).strip()

# ==========================================
# 3. 页面渲染逻辑
# ==========================================

# --- 页面 A: 主工作台 ---
def render_main_page():
    with st.sidebar:
        st.header("⚙️ 参数配置")
        PRICE_PER_DAY = st.number_input("人力单价 (元/天)", value=1500, step=100)
        MIN_HOURS = st.number_input("工时阈值 (小时)", value=100)
        SUBSIDY_TAG = st.text_input("补助关键词", "差旅补助")
        
        st.markdown("<br>"*5, unsafe_allow_html=True)
        st.markdown('<div class="cat-btn">', unsafe_allow_html=True)
        if st.button("🐱 字段映射配置", help="自定义匹配规则"):
            switch_page('mapping')
        st.markdown('</div>', unsafe_allow_html=True)

    st.title("😈 淡藤财务报表 Pro")

    with st.container(border=True):
        c_h1, c_h2 = st.columns([8, 1])
        c_h1.markdown("### 📂 数据源控制台")
        with c_h2:
            st.markdown('<div class="ghost-btn">', unsafe_allow_html=True)
            if st.button("🗑️ 重置"): reset_system()
            st.markdown('</div>', unsafe_allow_html=True)

        c_u1, c_u2 = st.columns(2)
        
        def render_slot(col, key, title):
            data = st.session_state.data_store[key]
            has_file = data['df'] is not None
            with col:
                with st.container(height=180, border=True):
                    if not has_file:
                        st.markdown(f"<div style='text-align:center; color:#888; margin-top:30px; margin-bottom:10px;'>{title}</div>", unsafe_allow_html=True)
                        f = st.file_uploader(title, type=['xlsx', 'csv'], key=f"uploader_{key}", label_visibility="collapsed")
                        if f: load_file_content(f, key)
                    else:
                        c_info, c_close = st.columns([9, 1])
                        with c_info:
                            st.markdown(f"""
                            <div class="file-card-styled">
                                <div>
                                    <div style="font-size:0.8rem; color:#8b949e;">{title.split(':')[0]}</div>
                                    <div style="font-weight:bold; color:#fff;">📄 {data['name']}</div>
                                    <div style="font-size:0.8rem; color:#238636;">✓ {len(data['df'])} 行</div>
                                </div>
                            </div>""", unsafe_allow_html=True)
                        with c_close:
                            if st.button("Del", key=f"del_{key}"): clear_file(key)

        render_slot(c_u1, 'A', "Source A: 投入明细")
        render_slot(c_u2, 'B', "Source B: 差旅明细")

    st.divider()
    ready_to_run = (st.session_state.data_store['A']['df'] is not None and 
                    st.session_state.data_store['B']['df'] is not None)
    trigger_calc = False

    if ready_to_run:
        if st.session_state.is_calculated:
            st.success("✅ 校验通过，报表已生成！")
        
        elif st.session_state.error_report is not None:
            err_df = st.session_state.error_report
            fixable = err_df[err_df['类型']=='数据错误']
            logic = err_df[err_df['类型']=='逻辑错误']
            
            st.markdown(f"""
            <div class="error-box">
                <h3 style="color:#ff7b72; margin:0">🚨 校验失败</h3>
                <p>发现 <b>{len(fixable)}</b> 个数据项错误，<b>{len(logic)}</b> 个配置/逻辑错误。</p>
            </div>""", unsafe_allow_html=True)
            
            st.dataframe(err_df[['类型','来源','行号','信息']], use_container_width=True, hide_index=True)
            
            c1, c2 = st.columns(2)
            c1.download_button("📥 下载错误清单", err_df.to_csv(index=False).encode('utf-8-sig'), "error.csv", "text/csv", use_container_width=True)
            
            @st.dialog("🛠️ 在线修复", width="large")
            def surgical_fix():
                if fixable.empty:
                    st.info("无数据错误，请检查配置或源文件。")
                    if st.button("关闭"): st.rerun()
                    return

                def get_fix(src):
                    if '_sys_id' not in fixable.columns: return pd.DataFrame()
                    t_ids = fixable[fixable['来源']==src]['_sys_id'].unique()
                    if len(t_ids)==0: return pd.DataFrame()
                    full = st.session_state.data_store[src.split()[-1]]['df']
                    return full[full['_sys_id'].isin(t_ids)].copy()

                da, db = get_fix('Source A'), get_fix('Source B')
                t1, t2 = st.tabs([f"A ({len(da)})", f"B ({len(db)})"])
                na, nb = None, None
                
                with t1:
                    if not da.empty:
                        gb = GridOptionsBuilder.from_dataframe(da.fillna(""))
                        gb.configure_column("_sys_id", hide=True)
                        gb.configure_default_column(editable=True)
                        na = AgGrid(da.fillna(""), gridOptions=gb.build(), height=300, key='fa')['data']
                    else: st.info("无数据错误")
                with t2:
                    if not db.empty:
                        gb = GridOptionsBuilder.from_dataframe(db.fillna(""))
                        gb.configure_column("_sys_id", hide=True)
                        gb.configure_default_column(editable=True)
                        nb = AgGrid(db.fillna(""), gridOptions=gb.build(), height=300, key='fb')['data']
                    else: st.info("无数据错误")
                
                if st.button("💾 保存并重算", type="primary"):
                    if na is not None:
                        res = pd.DataFrame(na)
                        od = st.session_state.data_store['A']['df'].set_index('_sys_id')
                        od.update(res.set_index('_sys_id'))
                        st.session_state.data_store['A']['df'] = od.reset_index()
                    if nb is not None:
                        res = pd.DataFrame(nb)
                        od = st.session_state.data_store['B']['df'].set_index('_sys_id')
                        od.update(res.set_index('_sys_id'))
                        st.session_state.data_store['B']['df'] = od.reset_index()
                    st.session_state.error_report = None
                    st.session_state.block_auto_run = False
                    st.rerun()

            if not fixable.empty:
                if c2.button("🛠️ 打开在线修复", type="primary", use_container_width=True): surgical_fix()
            else:
                c2.warning("⚠️ 存在配置错误或逻辑错误，无法在线修复，请去【猫猫配置页】检查列名匹配规则。")

        elif st.session_state.block_auto_run:
            st.info("ℹ️ 源文件已更新，等待确认...")
            if st.button("▶️ 重新校验并计算", type="primary", use_container_width=True):
                trigger_calc = True
                st.session_state.block_auto_run = False
        else:
            trigger_calc = True

    if trigger_calc:
        progress = st.progress(0, "启动...")
        df_a = st.session_state.data_store['A']['df']
        df_b = st.session_state.data_store['B']['df']
        errors = []

        # === 核心：从配置表读取单个匹配列 ===
        # 如果用户配置了多个(旧逻辑), 这里需要确保V7逻辑下我们只取一个
        # 在 V7 逻辑中，get_config_key 返回的是单一字符串
        
        ca_spm = get_config_key('Source A', 'SPM')
        ca_hrs = get_config_key('Source A', '工时')
        ca_name = get_config_key('Source A', '姓名')
        
        cb_spm = get_config_key('Source B', 'SPM (B)')
        cb_amt = get_config_key('Source B', '金额')
        cb_name = get_config_key('Source B', '姓名 (B)')

        # 维度字段
        ca_proj = get_config_key('Source A', '项目名称') or df_a.columns[0]
        ca_range = get_config_key('Source A', '人事范围') or df_a.columns[0]
        ca_contract = get_config_key('Source A', '合同主体') or df_a.columns[0]
        ca_sales = get_config_key('Source A', '销售部门') or df_a.columns[0]
        cb_type = get_config_key('Source B', '费用类型') or df_b.columns[0]

        def add_err(etype, src, rid, msg):
            errors.append({'类型': etype, '来源': src, '_sys_id': rid, '行号': rid if isinstance(rid, int) else '-', '信息': msg})

        # 1. 缺列检查
        def check_col(df, col, src, target_name):
            if col not in df.columns:
                add_err('逻辑错误', src, '-', f'未找到列[{col}] (对应目标:{target_name})。请去配置页检查。')
                return False
            return True

        if check_col(df_a, ca_spm, 'Source A', 'SPM') and \
           check_col(df_a, ca_hrs, 'Source A', '工时') and \
           check_col(df_a, ca_name, 'Source A', '姓名'):
           pass
        
        if check_col(df_b, cb_spm, 'Source B', 'SPM') and \
           check_col(df_b, cb_amt, 'Source B', '金额') and \
           check_col(df_b, cb_name, 'Source B', '姓名'):
           pass

        if not errors:
            # 预处理
            df_a[ca_hrs] = pd.to_numeric(df_a[ca_hrs], errors='coerce').fillna(0)
            if df_b[cb_amt].dtype == object: df_b[cb_amt] = df_b[cb_amt].astype(str).str.replace(',', '')
            df_b[cb_amt] = pd.to_numeric(df_b[cb_amt], errors='coerce').fillna(0)

            # 2. 数据错误
            for i,r in df_a[df_a[ca_hrs]<0].iterrows(): add_err('数据错误','Source A', r['_sys_id'], '工时为负')
            for i,r in df_b[df_b[cb_amt]<0].iterrows(): add_err('数据错误','Source B', r['_sys_id'], '金额为负')
            for i,r in df_a[df_a[ca_spm].isnull() | (df_a[ca_spm]=='')].iterrows(): add_err('数据错误','Source A', r['_sys_id'], 'SPM为空')

            # 3. 逻辑错误
            agg = df_a.groupby(ca_name)[ca_hrs].sum()
            for n,h in agg.items():
                if h < MIN_HOURS: add_err('逻辑错误','Source A', '-', f'人员[{n}]总工时({h}) < 阈值')
            
            df_a['key'] = df_a[ca_name].astype(str)+"_"+df_a[ca_spm].astype(str)
            df_b['key'] = df_b[cb_name].astype(str)+"_"+df_b[cb_spm].astype(str)
            orphans = df_b[~df_b['key'].isin(df_a['key'])]
            for i,r in orphans.iterrows(): add_err('逻辑错误','Source B', '-', f'无法匹配到交付人员: {r["key"]}')

        time.sleep(0.3)
        if errors:
            progress.empty()
            st.session_state.error_report = pd.DataFrame(errors)
            if '原表行号' not in st.session_state.error_report.columns:
                 st.session_state.error_report['原表行号'] = st.session_state.error_report['行号']
            st.session_state.block_auto_run = True
            st.rerun()
        else:
            progress.progress(50, "计算中...")
            agg_rules = {ca_hrs: 'sum'}
            for c in [ca_proj, ca_range, ca_contract, ca_sales]: agg_rules[c] = 'first'
            df_a_gp = df_a.groupby([ca_name, ca_spm], as_index=False).agg(agg_rules)
            
            is_sub = df_b[cb_type].astype(str).str.contains(SUBSIDY_TAG, na=False)
            grp_b = [cb_name, cb_spm]
            df_sub = df_b[is_sub].groupby(grp_b)[cb_amt].sum().reset_index(name='差旅补助')
            df_fee = df_b[~is_sub].groupby(grp_b)[cb_amt].sum().reset_index(name='差旅费控平台')
            
            for d in [df_a_gp, df_sub, df_fee]:
                k = ca_spm if ca_spm in d.columns else cb_spm
                d[k] = d[k].astype(str)
            
            res = pd.merge(df_a_gp, df_sub, left_on=[ca_name, ca_spm], right_on=[cb_name, cb_spm], how='left')
            res = pd.merge(res, df_fee, left_on=[ca_name, ca_spm], right_on=[cb_name, cb_spm], how='left')
            res = res.fillna(0)
            
            res['支持时间(人天)'] = res[ca_hrs] / 8
            res['人力费用'] = res['支持时间(人天)'] * PRICE_PER_DAY
            res['结算费用合计'] = res['人力费用'] + res['差旅补助'] + res['差旅费控平台']
            
            # 生成结果
            t3 = res.rename(columns={ca_name:'人员', ca_proj:'所属项目', ca_range:'人事范围', ca_spm:'SPM',
                                     ca_contract:'合同主体', ca_sales:'销售部门', ca_hrs:'耗时(小时)'})
            final_cols = ['序号','人员','所属项目','人事范围','SPM','合同主体','销售部门',
                          '差旅补助','差旅费控平台','耗时(小时)','支持时间(人天)','人力费用','结算费用合计']
            t3.insert(0, '序号', range(1, len(t3)+1))
            t3 = t3[[c for c in final_cols if c in t3.columns]]
            
            dims = [c for c in ['人事范围','合同主体','销售部门'] if c in t3.columns]
            if dims:
                t2 = t3.groupby(dims).agg({'结算费用合计':'sum', '支持时间(人天)':'sum'}).reset_index()
                t2.columns = ['销售公司','采购公司','采购部门','金额(含税,单位:元)','工作量(人天)']
                t2.insert(0, '序号', range(1, len(t2)+1))
            else: t2 = pd.DataFrame()
            
            t1 = t3.groupby('人员')['耗时(小时)'].sum().reset_index()
            t1.rename(columns={'耗时(小时)':'项目工时'}, inplace=True)
            t1.insert(0, '序号', range(1, len(t1)+1))
            
            def to_bytes(d):
                b = io.BytesIO()
                d.to_excel(b, index=False)
                return b.getvalue()
            
            st.session_state.result_zip = to_bytes(t3) 
            st.session_state.result_files = {'t1':to_bytes(t1), 't2':to_bytes(t2), 't3':to_bytes(t3)}
            
            buf = io.BytesIO()
            with zipfile.ZipFile(buf, 'w') as z:
                z.writestr("表1.xlsx", st.session_state.result_files['t1'])
                z.writestr("表2.xlsx", st.session_state.result_files['t2'])
                z.writestr("表3.xlsx", st.session_state.result_files['t3'])
            st.session_state.result_zip = buf.getvalue()
            
            st.session_state.is_calculated = True
            progress.progress(100)
            st.rerun()

    if st.session_state.is_calculated and st.session_state.result_zip:
        with st.container(border=True):
            st.success("✅ 生成完毕")
            st.download_button("📦 批量下载 (ZIP)", st.session_state.result_zip, "report.zip", type="primary", use_container_width=True)
            st.markdown("---")
            c1, c2, c3 = st.columns(3)
            fs = st.session_state.result_files
            if 't1' in fs: c1.download_button("📥 表1", fs['t1'], "t1.xlsx", use_container_width=True)
            if 't2' in fs: c2.download_button("📥 表2", fs['t2'], "t2.xlsx", use_container_width=True)
            if 't3' in fs: c3.download_button("📥 表3", fs['t3'], "t3.xlsx", use_container_width=True)

# --- 页面 B: 逻辑映射页 (猫猫按钮进入) ---
def render_mapping_page():
    st.markdown("### 🐱 字段映射 & 逻辑配置")
    
    col1, col2 = st.columns([1, 4])
    with col1:
        if st.button("⬅️ 返回主页", use_container_width=True): switch_page('main')
    
    # 状态栏: 编辑模式切换
    with col2:
        c_status, c_edit = st.columns([3, 1])
        with c_edit:
            # 只有当文件都上传了，才允许进入编辑模式
            has_files = st.session_state.data_store['A']['df'] is not None and st.session_state.data_store['B']['df'] is not None
            
            if not st.session_state.is_editing_mapping:
                if st.button("✏️ 编辑配置", type="primary", use_container_width=True):
                    if not has_files:
                        st.toast("⚠️ 请先在主页上传 A/B 表，以便系统获取真实列名。", icon="🚫")
                    else:
                        st.session_state.is_editing_mapping = True
                        st.rerun()
            else:
                if st.button("💾 保存生效", type="primary", use_container_width=True):
                    st.session_state.is_editing_mapping = False
                    st.session_state.is_calculated = False # 强制重算
                    st.session_state.block_auto_run = False
                    st.session_state.error_report = None
                    st.success("配置已更新！")
                    time.sleep(0.5)
                    st.rerun()

    st.divider()
    
    # 获取真实列名作为下拉选项
    cols_a = list(st.session_state.data_store['A']['df'].columns) if st.session_state.data_store['A']['df'] is not None else []
    cols_b = list(st.session_state.data_store['B']['df'].columns) if st.session_state.data_store['B']['df'] is not None else []

    df_conf = st.session_state.mapping_config
    
    # 分 Tab 展示
    tab1, tab2, tab3 = st.tabs(["表1: 工时统计", "表2: 结算汇总", "表3: 详细明细"])
    
    # 渲染器
    def render_table_config(table_name):
        subset = df_conf[df_conf['所属表'] == table_name]
        
        for idx, row in subset.iterrows():
            # 获取当前行的索引，用于更新
            real_idx = idx
            
            if not st.session_state.is_editing_mapping:
                # 只读模式
                st.markdown(f"""
                <div class="map-row">
                    <div class="map-label">{row['目标字段']}</div>
                    <div style="flex:1; font-family:monospace; color:#a5d6ff;">{row['匹配字段']}</div>
                    <div style="font-size:0.8rem; color:#666; width:100px;">{row['来源']}</div>
                    <div style="font-size:0.8rem; color:#238636; width:150px;">{row['计算逻辑']}</div>
                </div>
                """, unsafe_allow_html=True)
            else:
                # 编辑模式
                c1, c2, c3, c4 = st.columns([2, 3, 2, 2])
                c1.markdown(f"**{row['目标字段']}**")
                
                # 下拉菜单：根据来源动态选择列名列表
                options = cols_a if row['来源'] == 'Source A' else cols_b
                current_val = row['匹配字段']
                
                # 防止当前值不在选项中导致报错 (容错)
                index_val = 0
                if current_val in options:
                    index_val = options.index(current_val)
                
                new_val = c2.selectbox(
                    "匹配列", 
                    options=options, 
                    index=index_val, 
                    key=f"sel_{real_idx}", 
                    label_visibility="collapsed"
                )
                
                # 实时更新 session state
                st.session_state.mapping_config.at[real_idx, '匹配字段'] = new_val
                
                c3.caption(f"来源: {row['来源']}")
                c4.caption(f"逻辑: {row['计算逻辑']}")
                st.divider()

    with tab1: render_table_config("表1")
    with tab2: render_table_config("表2")
    with tab3: render_table_config("表3")

# ==========================================
# 4. 路由控制
# ==========================================
if st.session_state.page == 'main':
    render_main_page()
elif st.session_state.page == 'mapping':
    render_mapping_page()
