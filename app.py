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
            background: #21262d; border-left: 4px solid #238636; border-radius: 4px; padding: 15px; 
            width: 100%; height: 100%; display: flex; align-items: center; justify-content: space-between; 
        }
        /* [UI] 错误舱 */
        .error-box { border: 1px solid var(--red); background: rgba(218, 54, 51, 0.1); border-radius: 8px; padding: 1.5rem; margin-top: 1rem; }
        
        /* [UI] 顶部导航 */
        .nav-header { font-size: 1.2rem; font-weight: bold; margin-bottom: 20px; }
        
        /* [UI] 表格容器 */
        .table-container {
            border: 1px solid #30363d;
            border-radius: 6px;
            background-color: #0d1117;
            overflow: hidden;
            margin-bottom: 20px;
        }
        
        /* [UI] 表头单元格 - 增加右边框和居中 */
        div.header-cell {
            background-color: #161b22;
            color: #c9d1d9;
            font-weight: bold;
            padding: 15px 5px;
            border-bottom: 2px solid #30363d;
            border-right: 1px solid #30363d; /* 垂直分割线 */
            height: 100%;
            display: flex; 
            align-items: center; 
            justify-content: center; /* 水平居中 */
            text-align: center;
        }
        
        /* [UI] 内容行单元格 - 增加右边框和居中 */
        div.row-cell {
            padding: 12px 5px;
            border-bottom: 1px solid #21262d;
            border-right: 1px solid #21262d; /* 垂直分割线 */
            height: 100%;
            display: flex; 
            align-items: center; 
            justify-content: center; /* 水平居中 */
            text-align: center;
            font-size: 0.9rem;
            width: 100%;
        }
        
        /* 去掉最后一列的右边框 */
        .no-border-right { border-right: none !important; }
        
        /* 标签 */
        .source-tag { 
            padding: 2px 8px; border-radius: 4px; font-size: 0.75rem; font-weight: 500; display: inline-block;
        }
        .tag-source { background: rgba(56, 139, 253, 0.15); color: #58a6ff; border: 1px solid rgba(56, 139, 253, 0.4); }
        .tag-result { background: rgba(238, 138, 36, 0.15); color: #db8e37; border: 1px solid rgba(238, 138, 36, 0.4); }
        
        /* 说明条 */
        .info-bar {
            background-color: rgba(35, 134, 54, 0.1);
            border: 1px solid rgba(35, 134, 54, 0.3);
            color: #c9d1d9;
            padding: 10px 15px;
            border-radius: 6px;
            margin-bottom: 15px;
            font-size: 0.85rem;
            display: flex; align-items: center; gap: 10px;
        }

        /* 隐藏 Streamlit 默认组件杂项 */
        div[data-testid="stFileUploader"] section > div:first-child { display: none; }
        div[data-testid="stFileUploader"] { padding-top: 15px; }
        div[data-testid="stSelectbox"] { margin-bottom: 0px; }
        div[data-testid="stSelectbox"] > div > div { min-height: 32px; border-color: #30363d; background-color: #0d1117; justify-content: center; }
    </style>
    """, unsafe_allow_html=True)

# ==============================================================================
# Zone A: 纯逻辑层 (DataEngine) - V10.0 内核
# ==============================================================================
class DataEngine:
    @staticmethod
    def get_default_config():
        return pd.DataFrame([
            # === 结果表 3 (底表) ===
            {"所属表": "结果表3", "目标字段": "人员", "源表": "Source A", "匹配字段": "人员", "计算逻辑": "主键 (Join Key)"},
            {"所属表": "结果表3", "目标字段": "SPM", "源表": "Source A", "匹配字段": "SPM", "计算逻辑": "主键 (Join Key)"},
            {"所属表": "结果表3", "目标字段": "耗时(小时)", "源表": "Source A", "匹配字段": "交付工时", "计算逻辑": "SUM聚合 (清洗)"},
            {"所属表": "结果表3", "目标字段": "所属项目", "源表": "Source A", "匹配字段": "所属项目", "计算逻辑": "维度 (First)"},
            {"所属表": "结果表3", "目标字段": "人事范围", "源表": "Source A", "匹配字段": "人事范围", "计算逻辑": "维度 (->销售公司)"},
            {"所属表": "结果表3", "目标字段": "合同主体", "源表": "Source A", "匹配字段": "合同主体", "计算逻辑": "维度 (->采购公司)"},
            {"所属表": "结果表3", "目标字段": "销售人员", "源表": "Source A", "匹配字段": "销售", "计算逻辑": "维度 (First)"},
            {"所属表": "结果表3", "目标字段": "销售部门", "源表": "Source A", "匹配字段": "销售部门", "计算逻辑": "维度 (->采购部门)"},
            
            {"所属表": "结果表3", "目标字段": "人员 (B)", "源表": "Source B", "匹配字段": "出差人", "计算逻辑": "外键 (Join Key)"},
            {"所属表": "结果表3", "目标字段": "SPM (B)", "源表": "Source B", "匹配字段": "SPM", "计算逻辑": "外键 (Join Key)"},
            {"所属表": "结果表3", "目标字段": "金额", "源表": "Source B", "匹配字段": "金额", "计算逻辑": "SUM聚合 (清洗)"},
            {"所属表": "结果表3", "目标字段": "费用类型", "源表": "Source B", "匹配字段": "产品类型", "计算逻辑": "分类依据 (补助/费控)"},

            # === 结果表 2 (结算) ===
            {"所属表": "结果表2", "目标字段": "销售公司", "源表": "结果表3", "匹配字段": "人事范围", "计算逻辑": "维度分组"},
            {"所属表": "结果表2", "目标字段": "采购公司", "源表": "结果表3", "匹配字段": "合同主体", "计算逻辑": "维度分组"},
            {"所属表": "结果表2", "目标字段": "采购部门", "源表": "结果表3", "匹配字段": "销售部门", "计算逻辑": "维度分组"},
            {"所属表": "结果表2", "目标字段": "金额", "源表": "结果表3", "匹配字段": "结算费用合计", "计算逻辑": "SUM聚合"},
            {"所属表": "结果表2", "目标字段": "工作量", "源表": "结果表3", "匹配字段": "支持时间(人天)", "计算逻辑": "SUM聚合"},

            # === 结果表 1 (工时) ===
            {"所属表": "结果表1", "目标字段": "人员", "源表": "结果表3", "匹配字段": "人员", "计算逻辑": "维度分组"},
            {"所属表": "结果表1", "目标字段": "项目工时", "源表": "结果表3", "匹配字段": "耗时(小时)", "计算逻辑": "SUM聚合"},
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
        
        map_a = {
            '人员': c('人员', 'Source A'), 'SPM': c('SPM', 'Source A'), '耗时': c('耗时(小时)', 'Source A'),
            '项目': c('所属项目', 'Source A'), '范围': c('人事范围', 'Source A'), '合同': c('合同主体', 'Source A'),
            '销售': c('销售人员', 'Source A'), '部门': c('销售部门', 'Source A')
        }
        map_b = {
            '人员': c('人员 (B)', 'Source B'), 'SPM': c('SPM (B)', 'Source B'),
            '金额': c('金额', 'Source B'), '类型': c('费用类型', 'Source B')
        }
        
        def check_exist(df, col, src, target):
            if col and col not in df.columns:
                errors.append({'类型':'逻辑错误', '来源':src, '_sys_id':'-', '行号':'-', '信息':f'缺列: {col} (目标:{target})'})
                return False
            return True

        valid_a = all([check_exist(df_a, col, 'Source A', t) for t, col in map_a.items() if col])
        valid_b = all([check_exist(df_b, col, 'Source B', t) for t, col in map_b.items() if col])
        
        if not (valid_a and valid_b): return errors, df_a, df_b

        df_a_clean = df_a.copy()
        df_b_clean = df_b.copy()
        df_a_clean[map_a['耗时']] = DataEngine.clean_num(df_a_clean, map_a['耗时'])
        df_b_clean[map_b['金额']] = DataEngine.clean_num(df_b_clean, map_b['金额'])

        for i, r in df_a_clean[df_a_clean[map_a['耗时']] < 0].iterrows():
            errors.append({'类型':'数据错误', '来源':'Source A', '_sys_id':r['_sys_id'], '行号':r['_sys_id'], '信息':'工时为负'})
        for i, r in df_a[df_a[map_a['SPM']].isnull() | (df_a[map_a['SPM']] == '')].iterrows():
             errors.append({'类型':'数据错误', '来源':'Source A', '_sys_id':r['_sys_id'], '行号':r['_sys_id'], '信息':'SPM为空'})
        for i, r in df_b_clean[df_b_clean[map_b['金额']] < 0].iterrows():
            errors.append({'类型':'数据错误', '来源':'Source B', '_sys_id':r['_sys_id'], '行号':r['_sys_id'], '信息':'金额为负'})

        agg = df_a_clean.groupby(map_a['人员'])[map_a['耗时']].sum()
        for n, h in agg.items():
            if h < min_hours:
                 errors.append({'类型':'逻辑错误', '来源':'Source A', '_sys_id':'-', '行号':'-', '信息':f'人员[{n}]总工时({h}) < 阈值'})
        
        df_a['key'] = df_a[map_a['人员']].astype(str) + "_" + df_a[map_a['SPM']].astype(str)
        df_b['key'] = df_b[map_b['人员']].astype(str) + "_" + df_b[map_b['SPM']].astype(str)
        
        orphans = df_b[~df_b['key'].isin(df_a['key'])]
        for key in orphans['key'].unique():
             errors.append({'类型':'逻辑错误', '来源':'Source B', '_sys_id':'-', '行号':'-', '信息':f'孤立费用: {key}'})

        return errors, df_a, df_b

    @staticmethod
    def calculate(df_a, df_b, config_df, price_per_day, subsidy_tag):
        c = lambda t, s: DataEngine.get_col(config_df, t, s)
        
        col_a_user = c('人员', 'Source A')
        col_a_spm = c('SPM', 'Source A')
        col_a_hrs = c('耗时(小时)', 'Source A')
        dims = {
            'project': c('所属项目', 'Source A'), 'range': c('人事范围', 'Source A'), 'contract': c('合同主体', 'Source A'),
            'sales': c('销售人员', 'Source A'), 'dept': c('销售部门', 'Source A')
        }
        col_b_user = c('人员 (B)', 'Source B')
        col_b_spm = c('SPM (B)', 'Source B')
        col_b_amt = c('金额', 'Source B')
        col_b_type = c('费用类型', 'Source B')

        df_a[col_a_hrs] = DataEngine.clean_num(df_a, col_a_hrs)
        df_b[col_b_amt] = DataEngine.clean_num(df_b, col_b_amt)

        agg_rules = {col_a_hrs: 'sum'}
        for _, col in dims.items():
            if col: agg_rules[col] = 'first'
        df_a_gp = df_a.groupby([col_a_user, col_a_spm], as_index=False).agg(agg_rules)

        is_sub = df_b[col_b_type].astype(str).str.contains(subsidy_tag, na=False)
        grp_b = [col_b_user, col_b_spm]
        df_sub = df_b[is_sub].groupby(grp_b)[col_b_amt].sum().reset_index(name='差旅补助')
        df_fee = df_b[~is_sub].groupby(grp_b)[col_b_amt].sum().reset_index(name='差旅费控平台')

        df_a_gp[col_a_spm] = df_a_gp[col_a_spm].astype(str)
        df_sub[col_b_spm] = df_sub[col_b_spm].astype(str)
        df_fee[col_b_spm] = df_fee[col_b_spm].astype(str)
        
        res = pd.merge(df_a_gp, df_sub, left_on=[col_a_user, col_a_spm], right_on=[col_b_user, col_b_spm], how='left')
        res = pd.merge(res, df_fee, left_on=[col_a_user, col_a_spm], right_on=[col_b_user, col_b_spm], how='left')
        res = res.fillna(0)

        res['支持时间(人天)'] = res[col_a_hrs] / 8
        res['人力费用'] = res['支持时间(人天)'] * price_per_day
        res['结算费用合计'] = res['人力费用'] + res['差旅补助'] + res['差旅费控平台']

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

        dims_t2 = ['人事范围', '合同主体', '销售部门']
        if all(c in t3.columns for c in dims_t2):
            t2 = t3.groupby(dims_t2).agg({'结算费用合计': 'sum', '支持时间(人天)': 'sum'}).reset_index()
            t2.columns = ['销售公司', '采购公司', '采购部门', '金额(含税,单位:元)', '工作量(人天)']
            t2.insert(0, '序号', range(1, len(t2)+1))
        else:
            t2 = pd.DataFrame({'提示': ['缺少维度字段']})

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
            if st.button("🐱 字段映射 & 逻辑", help="查看映射逻辑"):
                st.session_state.page = 'mapping'
                st.rerun()
            return p, h, s

    @staticmethod
    def render_file_slot(key, title, data_store):
        data = data_store[key]
        has_file = data['df'] is not None
        with st.container(border=True):
            if not has_file:
                st.markdown(f"<div style='text-align:center; color:#888; margin-top:30px; margin-bottom:10px;'>{title}</div>", unsafe_allow_html=True)
                return st.file_uploader(title, type=['xlsx', 'csv'], key=f"uploader_{key}", label_visibility="collapsed")
            else:
                c1, c2 = st.columns([9, 1])
                c1.markdown(f"<div class='file-card-styled'><div><div style='font-size:0.8rem; color:#8b949e;'>{title.split(':')[0]}</div><div style='font-weight:bold; color:#fff;'>📄 {data['name']}</div><div style='font-size:0.8rem; color:#238636;'>✓ {len(data['df'])} 行</div></div></div>", unsafe_allow_html=True)
                if c2.button("Del", key=f"del_{key}"): return "DELETE"
        return None

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

    @staticmethod
    def render_tab_content(description, subset, is_edit, cols_a, cols_b):
        """渲染映射表内容 (Strict Aligned Table)"""
        # 4. 用途说明
        st.markdown(f'<div class="info-bar">ℹ️ {description}</div>', unsafe_allow_html=True)
        
        # 表格容器
        st.markdown('<div class="table-container">', unsafe_allow_html=True)
        
        # 表头 (2.5 - 3.5 - 1.5 - 2.5)
        c1, c2, c3, c4 = st.columns([2.5, 3.5, 1.5, 2.5])
        with c1: st.markdown('<div class="header-cell">目标字段</div>', unsafe_allow_html=True)
        with c2: st.markdown('<div class="header-cell">匹配字段</div>', unsafe_allow_html=True)
        with c3: st.markdown('<div class="header-cell">源表</div>', unsafe_allow_html=True)
        with c4: st.markdown('<div class="header-cell no-border-right">逻辑说明</div>', unsafe_allow_html=True)
        
        # 表内容
        for idx, row in subset.iterrows():
            is_result_source = row['源表'] == '结果表3'
            
            # 使用 container 模拟行容器，保持列宽一致
            c1, c2, c3, c4 = st.columns([2.5, 3.5, 1.5, 2.5])
            
            with c1:
                st.markdown(f'<div class="row-cell" style="font-weight:bold;">{row["目标字段"]}</div>', unsafe_allow_html=True)
            
            with c2:
                # 核心交互逻辑
                if st.session_state.is_editing_mapping and not is_result_source:
                    opts = cols_a if row['源表']=='Source A' else cols_b
                    cur = row['匹配字段']
                    if cur not in opts: opts = [cur] + opts
                    # 此时不包 row-cell div，让 selectbox 填满，CSS 已处理高度
                    new_val = st.selectbox("s", opts, index=opts.index(cur), key=f"s_{idx}", label_visibility="collapsed")
                    st.session_state.mapping_config.at[idx, '匹配字段'] = new_val
                else:
                    color = "#a5d6ff" if not is_result_source else "#8b949e"
                    st.markdown(f'<div class="row-cell" style="color:{color}; font-family:monospace;">{row["匹配字段"]}</div>', unsafe_allow_html=True)
            
            with c3:
                tag_cls = 'tag-result' if is_result_source else 'tag-source'
                st.markdown(f'<div class="row-cell"><span class="source-tag {tag_cls}">{row["源表"]}</span></div>', unsafe_allow_html=True)
            
            with c4:
                st.markdown(f'<div class="row-cell no-border-right" style="color:#888;">{row["计算逻辑"]}</div>', unsafe_allow_html=True)
                
        st.markdown('</div>', unsafe_allow_html=True) # End table-container

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
    # 1. 顶部导航 (Title + Return)
    c1, c2 = st.columns([9, 1])
    c1.markdown("<div class='nav-header'>🐱 字段映射 & 逻辑配置</div>", unsafe_allow_html=True)
    if c2.button("⬅️ 返回", use_container_width=True): 
        st.session_state.page = 'main'
        st.rerun()
    
    st.markdown("<hr style='margin-top:0; border-color:#30363d;'>", unsafe_allow_html=True)

    # 2. 模块标题 + 配置按钮
    c_title, c_spacer, c_action = st.columns([7, 1, 2])
    c_title.markdown("#### 🧬 数据血缘与逻辑配置")
    
    has_files = st.session_state.data_store['A']['df'] is not None and st.session_state.data_store['B']['df'] is not None
    
    with c_action:
        if not st.session_state.is_editing_mapping:
            if st.button("✏️ 编辑配置", type="primary", use_container_width=True):
                if not has_files: st.toast("请先在主页上传 A/B 表", icon="🚫")
                else:
                    st.session_state.is_editing_mapping = True
                    st.rerun()
        else:
            if st.button("💾 保存生效", type="primary", use_container_width=True):
                st.session_state.is_editing_mapping = False
                st.session_state.is_calculated = False
                st.session_state.block_auto_run = False
                st.session_state.error_report = None
                st.rerun()

    # 3. 结果表切换
    df_c = st.session_state.mapping_config
    t1, t2, t3 = st.tabs(["结果表3 (底表)", "结果表2 (结算)", "结果表1 (工时)"])
    
    cols_a = list(st.session_state.data_store['A']['df'].columns) if has_files else []
    cols_b = list(st.session_state.data_store['B']['df'].columns) if has_files else []

    # 4. 内容渲染
    with t1:
        UIComponents.render_tab_content(
            "全量明细底表：基于 Source A/B 进行清洗、聚合、关联计算后的宽表。",
            df_c[df_c['所属表']=='结果表3'], st.session_state.is_editing_mapping, cols_a, cols_b
        )
    with t2:
        UIComponents.render_tab_content(
            "结算汇总表：基于【结果表3】按公司/部门维度二次聚合的金额数据。",
            df_c[df_c['所属表']=='结果表2'], st.session_state.is_editing_mapping, cols_a, cols_b
        )
    with t3:
        UIComponents.render_tab_content(
            "工时统计表：基于【结果表3】按人员维度二次聚合的工时数据。",
            df_c[df_c['所属表']=='结果表1'], st.session_state.is_editing_mapping, cols_a, cols_b
        )
