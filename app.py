import streamlit as st
import pandas as pd
import io
import time
import zipfile
from st_aggrid import AgGrid, GridOptionsBuilder

# ==============================================================================
# Zone 0: 全局配置 & 样式注入
# ==============================================================================
st.set_page_config(page_title="淡藤财务报表 Pro", page_icon="😈", layout="wide", initial_sidebar_state="expanded")

def inject_css():
    st.markdown("""
    <style>
        :root { --bg-color: #0d1117; --card-bg: #161b22; --accent: #238636; --text: #c9d1d9; --red: #da3633; --border: #30363d; }
        .stApp { background-color: var(--bg-color); color: var(--text); }
        
        /* [UI] 卡片 */
        .file-card-styled { 
            background: #21262d; border-left: 4px solid #238636; border-radius: 6px; padding: 15px; 
            width: 100%; display: flex; align-items: center; justify-content: space-between; margin-bottom: 10px;
        }
        /* [UI] 错误舱 */
        .error-box { border: 1px solid var(--red); background: rgba(218, 54, 51, 0.1); border-radius: 8px; padding: 1.5rem; margin-top: 1rem; }
        
        /* [UI] 映射控制台样式 */
        .mapping-container {
            background-color: #161b22;
            border: 1px solid #30363d;
            border-radius: 8px;
            padding: 20px;
            height: 100%;
        }
        .mapping-header {
            font-size: 1.1rem; font-weight: bold; color: #fff; margin-bottom: 15px; border-bottom: 1px solid #30363d; padding-bottom: 10px; display: flex; justify-content: space-between; align-items: center;
        }
        .group-title {
            font-size: 0.85rem; color: #8b949e; font-weight: 600; margin-top: 15px; margin-bottom: 8px; text-transform: uppercase; letter-spacing: 0.5px;
        }
        .field-row {
            display: flex; align-items: center; justify-content: space-between;
            padding: 8px 10px; background: #0d1117; border-radius: 4px; margin-bottom: 6px; border: 1px solid #21262d;
        }
        .field-label { font-size: 0.9rem; font-weight: 500; color: #c9d1d9; }
        .field-desc { font-size: 0.75rem; color: #484f58; margin-top: 2px; }
        .status-dot { height: 8px; width: 8px; border-radius: 50%; display: inline-block; margin-right: 6px; }
        .status-ok { background-color: #238636; box-shadow: 0 0 5px #238636; }
        .status-err { background-color: #da3633; box-shadow: 0 0 5px #da3633; }
        .status-wait { background-color: #484f58; }
        
        /* [UI] 按钮 */
        .ghost-btn button { border: 1px dashed #444 !important; color: #888 !important; background: transparent !important; }
        .cat-btn button { border: 1px solid #30363d !important; background: #161b22 !important; color: #c9d1d9 !important; width: 100%; margin-top: 10px; }
        .cat-btn button:hover { border-color: #a371f7 !important; color: #a371f7 !important; }
        
        /* 隐藏上传组件默认列表 */
        div[data-testid="stFileUploader"] section > div:first-child { display: none; }
        div[data-testid="stFileUploader"] { padding-top: 15px; }
        /* 紧凑下拉框 */
        div[data-testid="stSelectbox"] > div > div { min-height: 32px; font-size: 0.85rem; }
    </style>
    """, unsafe_allow_html=True)

# ==============================================================================
# Zone A: 纯逻辑层 (DataEngine)
# ==============================================================================
class DataEngine:
    @staticmethod
    def get_default_config():
        """
        配置表结构：增加 '分组' 字段用于 UI 分区展示
        """
        return pd.DataFrame([
            # --- Source A: 工时表 ---
            {"所属表": "表3", "目标字段": "人员", "源表": "Source A", "匹配字段": "人员", "分组": "🔑 关联主键", "逻辑": "Join Key"},
            {"所属表": "表3", "目标字段": "SPM", "源表": "Source A", "匹配字段": "SPM", "分组": "🔑 关联主键", "逻辑": "Join Key"},
            {"所属表": "表3", "目标字段": "工时", "源表": "Source A", "匹配字段": "交付工时", "分组": "📊 核心指标", "逻辑": "SUM聚合"},
            
            {"所属表": "表3", "目标字段": "所属项目", "源表": "Source A", "匹配字段": "所属项目", "分组": "🏷️ 维度属性", "逻辑": "维度"},
            {"所属表": "表3", "目标字段": "人事范围", "源表": "Source A", "匹配字段": "人事范围", "分组": "🏷️ 维度属性", "逻辑": "-> 销售公司"},
            {"所属表": "表3", "目标字段": "合同主体", "源表": "Source A", "匹配字段": "合同主体", "分组": "🏷️ 维度属性", "逻辑": "-> 采购公司"},
            {"所属表": "表3", "目标字段": "销售人员", "源表": "Source A", "匹配字段": "销售", "分组": "🏷️ 维度属性", "逻辑": "维度"},
            {"所属表": "表3", "目标字段": "销售部门", "源表": "Source A", "匹配字段": "销售部门", "分组": "🏷️ 维度属性", "逻辑": "-> 采购部门"},
            
            # --- Source B: 费用表 ---
            {"所属表": "表3", "目标字段": "人员 (B)", "源表": "Source B", "匹配字段": "出差人", "分组": "🔑 关联主键", "逻辑": "Join Key"},
            {"所属表": "表3", "目标字段": "SPM (B)", "源表": "Source B", "匹配字段": "SPM", "分组": "🔑 关联主键", "逻辑": "Join Key"},
            {"所属表": "表3", "目标字段": "金额", "源表": "Source B", "匹配字段": "金额", "分组": "📊 核心指标", "逻辑": "SUM聚合"},
            {"所属表": "表3", "目标字段": "费用类型", "源表": "Source B", "匹配字段": "产品类型", "分组": "🧬 分类依据", "逻辑": "补助/费控"},
        ])

    @staticmethod
    def get_col(config_df, target, source_table):
        row = config_df[(config_df['目标字段'] == target) & (config_df['源表'] == source_table)]
        if row.empty: return None
        return str(row.iloc[0]['匹配字段']).strip()

    @staticmethod
    def clean_num(df, col):
        if col not in df.columns: return 0
        return pd.to_numeric(df[col].astype(str).str.replace(',', ''), errors='coerce').fillna(0)

    @staticmethod
    def validate(df_a, df_b, config_df, min_hours):
        errors = []
        c = lambda t, s: DataEngine.get_col(config_df, t, s)
        
        # 1. 映射字典
        map_a = {
            '人员': c('人员', 'Source A'), 'SPM': c('SPM', 'Source A'), '耗时': c('工时', 'Source A'),
            '项目': c('所属项目', 'Source A'), '范围': c('人事范围', 'Source A'), '合同': c('合同主体', 'Source A'),
            '销售': c('销售人员', 'Source A'), '部门': c('销售部门', 'Source A')
        }
        map_b = {
            '人员': c('人员 (B)', 'Source B'), 'SPM': c('SPM (B)', 'Source B'),
            '金额': c('金额', 'Source B'), '类型': c('费用类型', 'Source B')
        }
        
        # 2. 存在性检查
        def check_exist(df, col, src, target):
            if col and col not in df.columns:
                errors.append({'类型':'逻辑错误', '来源':src, '_sys_id':'-', '行号':'-', '信息':f'缺列: {col} (目标:{target})'})
                return False
            return True

        valid_a = all([check_exist(df_a, col, 'Source A', t) for t, col in map_a.items() if col])
        valid_b = all([check_exist(df_b, col, 'Source B', t) for t, col in map_b.items() if col])
        
        if not (valid_a and valid_b): return errors, df_a, df_b

        # 3. 数值清洗
        df_a_clean = df_a.copy()
        df_b_clean = df_b.copy()
        df_a_clean[map_a['耗时']] = DataEngine.clean_num(df_a_clean, map_a['耗时'])
        df_b_clean[map_b['金额']] = DataEngine.clean_num(df_b_clean, map_b['金额'])

        # 4. 数据错误
        for i, r in df_a_clean[df_a_clean[map_a['耗时']] < 0].iterrows():
            errors.append({'类型':'数据错误', '来源':'Source A', '_sys_id':r['_sys_id'], '行号':r['_sys_id'], '信息':'工时为负'})
        for i, r in df_a[df_a[map_a['SPM']].isnull() | (df_a[map_a['SPM']] == '')].iterrows():
             errors.append({'类型':'数据错误', '来源':'Source A', '_sys_id':r['_sys_id'], '行号':r['_sys_id'], '信息':'SPM为空'})
        for i, r in df_b_clean[df_b_clean[map_b['金额']] < 0].iterrows():
            errors.append({'类型':'数据错误', '来源':'Source B', '_sys_id':r['_sys_id'], '行号':r['_sys_id'], '信息':'金额为负'})

        # 5. 业务逻辑
        agg = df_a_clean.groupby(map_a['人员'])[map_a['耗时']].sum()
        for n, h in agg.items():
            if h < min_hours:
                 errors.append({'类型':'逻辑错误', '来源':'Source A', '_sys_id':'-', '行号':'-', '信息':f'人员[{n}]总工时({h}) < 阈值'})
        
        df_a['key'] = df_a[map_a['人员']].astype(str) + "_" + df_a[map_a['SPM']].astype(str)
        df_b['key'] = df_b[map_b['人员']].astype(str) + "_" + df_b[map_b['SPM']].astype(str)
        
        orphans = df_b[~df_b['key'].isin(df_a['key'])]
        for key in orphans['key'].unique():
             errors.append({'类型':'逻辑错误', '来源':'Source B', '_sys_id':'-', '行号':'-', '信息':f'孤立费用，无对应工时: {key}'})

        return errors, df_a, df_b

    @staticmethod
    def calculate(df_a, df_b, config_df, price_per_day, subsidy_tag):
        c = lambda t, s: DataEngine.get_col(config_df, t, s)
        
        # 字段获取
        col_a_user = c('人员', 'Source A')
        col_a_spm = c('SPM', 'Source A')
        col_a_hrs = c('工时', 'Source A')
        dims = {
            'project': c('所属项目', 'Source A'), 'range': c('人事范围', 'Source A'), 'contract': c('合同主体', 'Source A'),
            'sales': c('销售人员', 'Source A'), 'dept': c('销售部门', 'Source A')
        }
        col_b_user = c('人员 (B)', 'Source B')
        col_b_spm = c('SPM (B)', 'Source B')
        col_b_amt = c('金额', 'Source B')
        col_b_type = c('费用类型', 'Source B')

        # 清洗
        df_a[col_a_hrs] = DataEngine.clean_num(df_a, col_a_hrs)
        df_b[col_b_amt] = DataEngine.clean_num(df_b, col_b_amt)

        # A表聚合
        agg_rules = {col_a_hrs: 'sum'}
        for _, col in dims.items():
            if col: agg_rules[col] = 'first'
        df_a_gp = df_a.groupby([col_a_user, col_a_spm], as_index=False).agg(agg_rules)

        # B表聚合
        is_sub = df_b[col_b_type].astype(str).str.contains(subsidy_tag, na=False)
        grp_b = [col_b_user, col_b_spm]
        df_sub = df_b[is_sub].groupby(grp_b)[col_b_amt].sum().reset_index(name='差旅补助')
        df_fee = df_b[~is_sub].groupby(grp_b)[col_b_amt].sum().reset_index(name='差旅费控平台')

        # 合并
        df_a_gp[col_a_spm] = df_a_gp[col_a_spm].astype(str)
        df_sub[col_b_spm] = df_sub[col_b_spm].astype(str)
        df_fee[col_b_spm] = df_fee[col_b_spm].astype(str)
        
        res = pd.merge(df_a_gp, df_sub, left_on=[col_a_user, col_a_spm], right_on=[col_b_user, col_b_spm], how='left')
        res = pd.merge(res, df_fee, left_on=[col_a_user, col_a_spm], right_on=[col_b_user, col_b_spm], how='left')
        res = res.fillna(0)

        # 计算
        res['支持时间(人天)'] = res[col_a_hrs] / 8
        res['人力费用'] = res['支持时间(人天)'] * price_per_day
        res['结算费用合计'] = res['人力费用'] + res['差旅补助'] + res['差旅费控平台']

        # 结果表3
        rename_map = {
            col_a_user: '人员', dims['project']: '所属项目', dims['range']: '人事范围',
            col_a_spm: 'SPM', dims['contract']: '合同主体', dims['sales']: '销售人员',
            dims['dept']: '销售部门', col_a_hrs: '耗时(小时)'
        }
        rename_map = {k:v for k,v in rename_map.items() if k in res.columns}
        t3 = res.rename(columns=rename_map)
        
        final_cols = ['序号','人员','所属项目','人事范围','SPM','合同主体','销售人员','销售部门',
                      '差旅补助','差旅费控平台','耗时(小时)','支持时间(人天)','人力费用','结算费用合计']
        t3.insert(0, '序号', range(1, len(t3)+1))
        t3 = t3[[c for c in final_cols if c in t3.columns]]

        # 结果表2
        dims_t2 = ['人事范围', '合同主体', '销售部门']
        if all(c in t3.columns for c in dims_t2):
            t2 = t3.groupby(dims_t2).agg({'结算费用合计': 'sum', '支持时间(人天)': 'sum'}).reset_index()
            t2.columns = ['销售公司', '采购公司', '采购部门', '金额(含税,单位:元)', '工作量(人天)']
            t2.insert(0, '序号', range(1, len(t2)+1))
        else:
            t2 = pd.DataFrame({'提示': ['缺少维度字段']})

        # 结果表1
        if '人员' in t3.columns and '耗时(小时)' in t3.columns:
            t1 = t3.groupby('人员')['耗时(小时)'].sum().reset_index()
            t1.rename(columns={'耗时(小时)': '项目工时'}, inplace=True)
            t1.insert(0, '序号', range(1, len(t1)+1))
        else:
            t1 = pd.DataFrame({'提示': ['缺少人员字段']})

        return {'t1': t1, 't2': t2, 't3': t3}

    @staticmethod
    def to_bytes(df):
        b = io.BytesIO()
        out = df.drop(columns=['_sys_id'], errors='ignore')
        out.to_excel(b, index=False)
        return b.getvalue()

# ==============================================================================
# Zone B: UI 组件层 (View)
# ==============================================================================
class UIComponents:
    @staticmethod
    def render_sidebar():
        with st.sidebar:
            st.header("⚙️ 参数配置")
            p = st.number_input("人力单价 (元/天)", value=1500, step=100)
            h = st.number_input("工时阈值 (小时)", value=100)
            s = st.text_input("补助关键词", "差旅补助")
            st.markdown("---")
            st.markdown('<div class="cat-btn">', unsafe_allow_html=True)
            if st.button("🐱 字段映射配置", help="进入映射控制台"):
                st.session_state.page = 'mapping'
                st.rerun()
            st.markdown('</div>', unsafe_allow_html=True)
            return p, h, s

    @staticmethod
    def render_file_slot(key, title, data_store):
        data = data_store[key]
        has_file = data['df'] is not None
        with st.container(height=180, border=True):
            if not has_file:
                st.markdown(f"<div style='text-align:center; color:#888; margin-top:30px;'>{title}</div>", unsafe_allow_html=True)
                return st.file_uploader(title, type=['xlsx', 'csv'], key=f"uploader_{key}", label_visibility="collapsed")
            else:
                c1, c2 = st.columns([9, 1])
                c1.markdown(f"<div class='file-card-styled'><div><div style='font-size:0.8rem; color:#8b949e;'>{title.split(':')[0]}</div><div style='font-weight:bold; color:#fff;'>📄 {data['name']}</div><div style='font-size:0.8rem; color:#238636;'>✓ {len(data['df'])} 行</div></div></div>", unsafe_allow_html=True)
                if c2.button("Del", key=f"del_{key}"): return "DELETE"
        return None

    @staticmethod
    def render_mapping_panel(source_name, title, df_config, file_cols, is_edit):
        """渲染单个源表的配置面板"""
        st.markdown(f"""
        <div class="mapping-container">
            <div class="mapping-header">
                <span>{title}</span>
                <span style="font-size:0.8rem; color:#238636; font-weight:normal;">
                    {'🟢 已连接' if file_cols else '⚪ 待上传'}
                </span>
            </div>
        """, unsafe_allow_html=True)
        
        subset = df_config[df_config['源表'] == source_name]
        
        # 按分组渲染
        for group in subset['分组'].unique():
            st.markdown(f"<div class='group-title'>{group}</div>", unsafe_allow_html=True)
            group_data = subset[subset['分组'] == group]
            
            for idx, row in group_data.iterrows():
                # 状态检测
                mapped_col = row['匹配字段']
                status_cls = "status-wait"
                if file_cols:
                    status_cls = "status-ok" if mapped_col in file_cols else "status-err"
                
                st.markdown(f'<div class="field-row">', unsafe_allow_html=True)
                
                # 左侧：Label + Desc
                c_left, c_right = st.columns([5, 5])
                with c_left:
                    st.markdown(f"""
                    <div style="display:flex; align-items:center;">
                        <span class="status-dot {status_cls}"></span>
                        <div>
                            <div class="field-label">{row['目标字段']}</div>
                            <div class="field-desc">{row['逻辑']}</div>
                        </div>
                    </div>
                    """, unsafe_allow_html=True)
                
                # 右侧：Editor / Display
                with c_right:
                    if is_edit:
                        # 如果没有文件列，提供一个空的或当前值的列表
                        opts = file_cols if file_cols else [mapped_col]
                        # 容错：如果当前配置的值不在文件列中，加进去
                        if mapped_col not in opts: opts = [mapped_col] + opts
                        
                        new_val = st.selectbox(
                            "s", opts, 
                            index=opts.index(mapped_col), 
                            key=f"sel_{idx}", 
                            label_visibility="collapsed"
                        )
                        st.session_state.mapping_config.at[idx, '匹配字段'] = new_val
                    else:
                        color = "#f85149" if status_cls == "status-err" else "#a5d6ff"
                        st.markdown(f"<div style='text-align:right; font-family:monospace; color:{color};'>{mapped_col}</div>", unsafe_allow_html=True)
                
                st.markdown('</div>', unsafe_allow_html=True) # close field-row
                
        st.markdown("</div>", unsafe_allow_html=True) # close container

    @staticmethod
    def render_error_report(err_df, on_fix):
        fixable = err_df[err_df['类型']=='数据错误']
        logic = err_df[err_df['类型']=='逻辑错误']
        st.markdown(f"<div class='error-box'><h3 style='color:#ff7b72; margin:0'>🚨 校验失败</h3><p>发现 <b>{len(fixable)}</b> 个数据错误，<b>{len(logic)}</b> 个逻辑错误。</p></div>", unsafe_allow_html=True)
        st.dataframe(err_df[['类型','来源','行号','信息']], use_container_width=True, hide_index=True)
        c1, c2 = st.columns(2)
        c1.download_button("📥 下载清单", err_df.to_csv(index=False).encode('utf-8-sig'), "err.csv", "text/csv", use_container_width=True)
        if not fixable.empty:
            if c2.button("🛠️ 在线修复", type="primary", use_container_width=True): on_fix()

    @staticmethod
    def render_download_zone(result_files, result_zip):
        with st.container(border=True):
            st.success("✅ 生成完毕")
            st.download_button("📦 批量下载 (ZIP)", result_zip, "report.zip", type="primary", use_container_width=True)
            st.markdown("---")
            c1, c2, c3 = st.columns(3)
            if 't1' in result_files: c1.download_button("📥 表1: 工时统计", result_files['t1'], "t1.xlsx", use_container_width=True)
            if 't2' in result_files: c2.download_button("📥 表2: 结算汇总", result_files['t2'], "t2.xlsx", use_container_width=True)
            if 't3' in result_files: c3.download_button("📥 表3: 详细明细", result_files['t3'], "t3.xlsx", use_container_width=True)

# ==============================================================================
# Zone C: 控制层 (Controller)
# ==============================================================================
if 'page' not in st.session_state: st.session_state.page = 'main'
if 'data_store' not in st.session_state: st.session_state.data_store = {'A': {'df': None, 'name': None}, 'B': {'df': None, 'name': None}}
if 'mapping_config' not in st.session_state: st.session_state.mapping_config = DataEngine.get_default_config()
if 'is_calculated' not in st.session_state: st.session_state.is_calculated = False
if 'error_report' not in st.session_state: st.session_state.error_report = None
if 'block_auto_run' not in st.session_state: st.session_state.block_auto_run = False
if 'is_editing_mapping' not in st.session_state: st.session_state.is_editing_mapping = False

inject_css()

if st.session_state.page == 'main':
    price, hours_limit, sub_tag = UIComponents.render_sidebar()
    st.title("😈 淡藤财务报表 Pro")
    
    with st.container(border=True):
        c_h1, c_h2 = st.columns([8, 1])
        c_h1.markdown("### 📂 数据源控制台")
        if c_h2.button("🗑️ 重置"): 
            st.session_state.clear()
            st.rerun()
        
        c1, c2 = st.columns(2)
        def handle(key, title):
            res = UIComponents.render_file_slot(key, title, st.session_state.data_store)
            if res == "DELETE":
                st.session_state.data_store[key] = {'df': None, 'name': None}
                st.session_state.is_calculated = False
                st.session_state.error_report = None
                st.session_state.block_auto_run = False
                st.rerun()
            elif res:
                try:
                    df = pd.read_csv(res) if res.name.endswith('.csv') else pd.read_excel(res)
                    df.columns = [str(c).strip() for c in df.columns]
                    df['_sys_id'] = range(1, len(df)+1)
                    st.session_state.data_store[key] = {'df': df, 'name': res.name}
                    if st.session_state.block_auto_run: st.session_state.error_report = None
                    st.rerun()
                except Exception as e: st.error(f"Error: {e}")
        handle('A', "Source A: 投入明细")
        handle('B', "Source B: 差旅明细")

    st.divider()
    has_files = st.session_state.data_store['A']['df'] is not None and st.session_state.data_store['B']['df'] is not None
    trigger = False

    if has_files:
        if st.session_state.is_calculated:
            UIComponents.render_download_zone(st.session_state.result_files, st.session_state.result_zip)
        elif st.session_state.error_report is not None:
            def fix_action():
                @st.dialog("🛠️ 在线修复", width="large")
                def show_fix():
                    fixable = st.session_state.error_report[st.session_state.error_report['类型']=='数据错误']
                    def get_df(src):
                        ids = fixable[fixable['来源']==src]['_sys_id'].unique()
                        if len(ids)==0: return pd.DataFrame()
                        full = st.session_state.data_store[src.split()[-1]]['df']
                        return full[full['_sys_id'].isin(ids)].copy()
                    
                    da, db = get_df('Source A'), get_df('Source B')
                    t1, t2 = st.tabs([f"A ({len(da)})", f"B ({len(db)})"])
                    na, nb = None, None
                    
                    with t1:
                        if not da.empty:
                            gb = GridOptionsBuilder.from_dataframe(da.fillna(""))
                            gb.configure_column("_sys_id", hide=True)
                            gb.configure_default_column(editable=True)
                            na = AgGrid(da.fillna(""), gridOptions=gb.build(), height=300)['data']
                        else: st.info("无数据错误")
                    with t2:
                        if not db.empty:
                            gb = GridOptionsBuilder.from_dataframe(db.fillna(""))
                            gb.configure_column("_sys_id", hide=True)
                            gb.configure_default_column(editable=True)
                            nb = AgGrid(db.fillna(""), gridOptions=gb.build(), height=300)['data']
                        else: st.info("无数据错误")
                    
                    if st.button("💾 保存并重算", type="primary"):
                        if na is not None:
                            od = st.session_state.data_store['A']['df'].set_index('_sys_id')
                            od.update(pd.DataFrame(na).set_index('_sys_id'))
                            st.session_state.data_store['A']['df'] = od.reset_index()
                        if nb is not None:
                            od = st.session_state.data_store['B']['df'].set_index('_sys_id')
                            od.update(pd.DataFrame(nb).set_index('_sys_id'))
                            st.session_state.data_store['B']['df'] = od.reset_index()
                        st.session_state.error_report = None
                        st.session_state.block_auto_run = False
                        st.rerun()
                show_fix()
            UIComponents.render_error_report(st.session_state.error_report, fix_action)
        elif st.session_state.block_auto_run:
            st.info("ℹ️ 源文件已更新，等待确认...")
            if st.button("▶️ 重新校验并计算", type="primary", use_container_width=True): trigger = True
        else:
            trigger = True

    if trigger:
        with st.spinner("计算中..."):
            errs, df_a, df_b = DataEngine.validate(
                st.session_state.data_store['A']['df'].copy(),
                st.session_state.data_store['B']['df'].copy(),
                st.session_state.mapping_config,
                hours_limit
            )
            time.sleep(0.3)
            if errs:
                st.session_state.error_report = pd.DataFrame(errs)
                st.session_state.block_auto_run = True
                st.rerun()
            else:
                res = DataEngine.calculate(df_a, df_b, st.session_state.mapping_config, price, sub_tag)
                st.session_state.result_files = {k: DataEngine.to_bytes(v) for k, v in res.items()}
                buf = io.BytesIO()
                with zipfile.ZipFile(buf, 'w') as z:
                    z.writestr("表1_工时.xlsx", st.session_state.result_files['t1'])
                    z.writestr("表2_结算.xlsx", st.session_state.result_files['t2'])
                    z.writestr("表3_明细.xlsx", st.session_state.result_files['t3'])
                st.session_state.result_zip = buf.getvalue()
                st.session_state.is_calculated = True
                st.rerun()

elif st.session_state.page == 'mapping':
    st.markdown("### 🐱 字段映射控制台")
    c1, c2 = st.columns([1, 4])
    if c1.button("⬅️ 返回主页", use_container_width=True): 
        st.session_state.page = 'main'
        st.rerun()
    
    with c2:
        c_status, c_edit = st.columns([3, 1])
        with c_edit:
            has_files = st.session_state.data_store['A']['df'] is not None
            if not st.session_state.is_editing_mapping:
                if st.button("✏️ 开启配置", type="primary", use_container_width=True):
                    if not has_files: st.toast("建议先在主页上传文件，以便自动获取列名", icon="⚠️")
                    st.session_state.is_editing_mapping = True
                    st.rerun()
            else:
                if st.button("💾 保存生效", type="primary", use_container_width=True):
                    st.session_state.is_editing_mapping = False
                    st.session_state.is_calculated = False
                    st.session_state.block_auto_run = False
                    st.session_state.error_report = None
                    st.rerun()
    
    st.divider()
    
    # 准备列名
    cols_a = list(st.session_state.data_store['A']['df'].columns) if st.session_state.data_store['A']['df'] is not None else []
    cols_b = list(st.session_state.data_store['B']['df'].columns) if st.session_state.data_store['B']['df'] is not None else []
    
    # 左右分栏布局
    c_left, c_right = st.columns(2)
    
    with c_left:
        UIComponents.render_mapping_panel(
            "Source A", "🏭 Source A: 投入明细 (工时流)", 
            st.session_state.mapping_config, 
            cols_a, st.session_state.is_editing_mapping
        )
        
    with c_right:
        UIComponents.render_mapping_panel(
            "Source B", "💸 Source B: 差旅明细 (费用流)", 
            st.session_state.mapping_config, 
            cols_b, st.session_state.is_editing_mapping
        )
