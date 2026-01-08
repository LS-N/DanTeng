import streamlit as st
import pandas as pd
import io
import time
import zipfile

# ==============================================================================
# Zone 0: 全局配置 & 样式注入
# ==============================================================================
st.set_page_config(page_title="淡藤财务报表 Pro", page_icon="😈", layout="wide", initial_sidebar_state="expanded")

def inject_css():
    st.markdown("""
    <style>
        /* 基础配色变量 */
        :root { --bg-color: #0d1117; --card-bg: #161b22; --accent: #238636; --text: #c9d1d9; --border-color: #555c65; }
        .stApp { background-color: var(--bg-color); color: var(--text); }
        
        /* 自定义组件样式 */
        .nav-header { font-size: 1.2rem; font-weight: bold; display:flex; align-items:center; height: 100%; }
        
        .info-bar {
            background-color: rgba(56, 139, 253, 0.1);
            border-left: 4px solid #58a6ff;
            color: #c9d1d9;
            padding: 8px 15px;
            margin-bottom: 20px;
            font-size: 0.9rem;
            border-radius: 4px;
        }

        .error-box {
            border: 1px solid #ff7b72;
            background-color: rgba(255, 123, 114, 0.1);
            padding: 15px;
            border-radius: 6px;
            margin-bottom: 15px;
        }
    </style>
    """, unsafe_allow_html=True)

# ==============================================================================
# Zone A: 纯逻辑层 (DataEngine) - V20.0 内核
# ==============================================================================
class DataEngine:
    @staticmethod
    def get_default_config():
        return pd.DataFrame([
            # === 结果表 3 (14个核心字段 + 辅助) ===
            {"所属表": "结果表3", "序号": 1, "目标字段": "序号", "源表": "🔒 系统生成", "匹配字段": "-", "逻辑说明": "自增序列"},
            {"所属表": "结果表3", "序号": 2, "目标字段": "人员", "源表": "Source A", "匹配字段": "人员", "逻辑说明": "主键 (Join Key)"},
            {"所属表": "结果表3", "序号": 3, "目标字段": "所属项目", "源表": "Source A", "匹配字段": "项目", "逻辑说明": "维度 (First)"},
            {"所属表": "结果表3", "序号": 4, "目标字段": "人事范围", "源表": "Source A", "匹配字段": "人事范围", "逻辑说明": "维度 (->销售公司)"},
            {"所属表": "结果表3", "序号": 5, "目标字段": "SPM", "源表": "Source A", "匹配字段": "SPM", "逻辑说明": "主键 (Join Key)"},
            {"所属表": "结果表3", "序号": 6, "目标字段": "合同主体", "源表": "Source A", "匹配字段": "合同主体", "逻辑说明": "维度 (->采购公司)"},
            {"所属表": "结果表3", "序号": 7, "目标字段": "销售人员", "源表": "Source A", "匹配字段": "销售", "逻辑说明": "维度 (First)"},
            {"所属表": "结果表3", "序号": 8, "目标字段": "销售部门", "源表": "Source A", "匹配字段": "销售部门", "逻辑说明": "维度 (->采购部门)"},
            {"所属表": "结果表3", "序号": 9, "目标字段": "差旅补助", "源表": "Source B", "匹配字段": "金额", "逻辑说明": "筛选：产品类型='差旅补助'"},
            {"所属表": "结果表3", "序号": 10, "目标字段": "差旅费控平台", "源表": "Source B", "匹配字段": "金额", "逻辑说明": "筛选：产品类型!='差旅补助'"},
            {"所属表": "结果表3", "序号": 11, "目标字段": "耗时（小时）", "源表": "Source A", "匹配字段": "交付工时", "逻辑说明": "SUM聚合 (清洗)"},
            {"所属表": "结果表3", "序号": 12, "目标字段": "支持时间（人天）", "源表": "🔒 公式计算", "匹配字段": "-", "逻辑说明": "耗时 / 8"},
            {"所属表": "结果表3", "序号": 13, "目标字段": "人力费用", "源表": "🔒 公式计算", "匹配字段": "-", "逻辑说明": "人天 * 单价"},
            {"所属表": "结果表3", "序号": 14, "目标字段": "结算费用合计", "源表": "🔒 公式计算", "匹配字段": "-", "逻辑说明": "人力 + 差旅 + 费控"},
            
            # 辅助配置
            {"所属表": "结果表3", "序号": 15, "目标字段": "[配置] B表关联人", "源表": "Source B", "匹配字段": "出差人", "逻辑说明": "辅助：用于匹配A表人员"},
            {"所属表": "结果表3", "序号": 16, "目标字段": "[配置] B表关联SPM", "源表": "Source B", "匹配字段": "SPM", "逻辑说明": "辅助：用于匹配A表SPM"},
            {"所属表": "结果表3", "序号": 17, "目标字段": "[配置] B表类型列", "源表": "Source B", "匹配字段": "产品类型", "逻辑说明": "辅助：用于区分补助/费控"},

            # === 结果表 2 ===
            {"所属表": "结果表2", "序号": 1, "目标字段": "序号", "源表": "🔒 系统生成", "匹配字段": "-", "逻辑说明": "自增序列"},
            {"所属表": "结果表2", "序号": 2, "目标字段": "销售公司", "源表": "🔒 结果表3", "匹配字段": "人事范围", "逻辑说明": "维度分组"},
            {"所属表": "结果表2", "序号": 3, "目标字段": "采购公司", "源表": "🔒 结果表3", "匹配字段": "合同主体", "逻辑说明": "维度分组"},
            {"所属表": "结果表2", "序号": 4, "目标字段": "采购部门", "源表": "🔒 结果表3", "匹配字段": "销售部门", "逻辑说明": "维度分组"},
            {"所属表": "结果表2", "序号": 5, "目标字段": "金额（含税）", "源表": "🔒 结果表3", "匹配字段": "结算费用合计", "逻辑说明": "SUM聚合"},
            {"所属表": "结果表2", "序号": 6, "目标字段": "工作量（人天）", "源表": "🔒 结果表3", "匹配字段": "支持时间（人天）", "逻辑说明": "SUM聚合"},

            # === 结果表 1 ===
            {"所属表": "结果表1", "序号": 1, "目标字段": "序号", "源表": "🔒 系统生成", "匹配字段": "-", "逻辑说明": "自增序列"},
            {"所属表": "结果表1", "序号": 2, "目标字段": "人员", "源表": "🔒 结果表3", "匹配字段": "人员", "逻辑说明": "维度分组"},
            {"所属表": "结果表1", "序号": 3, "目标字段": "项目工时", "源表": "🔒 结果表3", "匹配字段": "耗时（小时）", "逻辑说明": "SUM聚合"},
        ])

    @staticmethod
    def get_col(config_df, target):
        row = config_df[config_df['目标字段'] == target]
        if row.empty: return None
        return str(row.iloc[0]['匹配字段']).strip()

    @staticmethod
    def clean_num(df, col):
        if col not in df.columns: return 0
        return pd.to_numeric(df[col].astype(str).str.replace(',', ''), errors='coerce').fillna(0)

    @staticmethod
    def validate(df_a, df_b, config_df, min_hours):
        errors = []
        c = lambda t: DataEngine.get_col(config_df, t)
        
        col_a_user = c('人员')
        col_a_spm = c('SPM')
        col_a_hrs = c('耗时（小时）')
        col_b_user = c('[配置] B表关联人')
        col_b_spm = c('[配置] B表关联SPM')
        col_b_amt = c('差旅补助')
        
        def check(df, col, src, tag):
            if col and col not in df.columns:
                errors.append({'类型':'逻辑错误', '来源':src, '_sys_id':'-', '行号':'-', '信息':f'缺列: {col} (用途:{tag})'})
                return False
            return True

        valid_a = check(df_a, col_a_user, 'Source A', '人员') and check(df_a, col_a_spm, 'Source A', 'SPM') and check(df_a, col_a_hrs, 'Source A', '工时')
        valid_b = check(df_b, col_b_user, 'Source B', '出差人') and check(df_b, col_b_spm, 'Source B', 'SPM') and check(df_b, col_b_amt, 'Source B', '金额')
        
        if not (valid_a and valid_b): return errors, df_a, df_b

        # 1. 基础清洗与负数校验
        df_a_clean = df_a.copy()
        df_a_clean[col_a_hrs] = DataEngine.clean_num(df_a_clean, col_a_hrs)
        for i, r in df_a_clean[df_a_clean[col_a_hrs] < 0].iterrows():
            errors.append({'类型':'数据错误', '来源':'Source A', '_sys_id':r.get('_sys_id','-'), '行号':r.get('_sys_id','-'), '信息':'工时为负'})
        
        # 2. 按人聚合校验工时阈值
        if col_a_user and col_a_hrs and min_hours > 0:
            user_sums = df_a_clean.groupby(col_a_user)[col_a_hrs].sum()
            invalid_users = user_sums[user_sums < min_hours]
            
            for user, total_hrs in invalid_users.items():
                sample_rows = df_a_clean[df_a_clean[col_a_user] == user]
                if not sample_rows.empty:
                    r = sample_rows.iloc[0]
                    errors.append({
                        '类型': '业务规则校验', 
                        '来源': 'Source A',
                        '_sys_id': r.get('_sys_id', '-'),
                        '行号': r.get('_sys_id', '-'), 
                        '信息': f'人员【{user}】总工时({total_hrs}h) 低于阈值({min_hours}h)'
                    })

        return errors, df_a, df_b

    @staticmethod
    def calculate(df_a, df_b, config_df, price_per_day, subsidy_tag):
        c = lambda t: DataEngine.get_col(config_df, t)
        
        col_a_user = c('人员')
        col_a_spm = c('SPM')
        col_a_hrs = c('耗时（小时）')
        dims_a = {
            'project': c('所属项目'), 'range': c('人事范围'), 'contract': c('合同主体'),
            'sales': c('销售人员'), 'dept': c('销售部门')
        }
        
        col_b_user = c('[配置] B表关联人')
        col_b_spm = c('[配置] B表关联SPM')
        col_b_amt = c('差旅补助')
        col_b_type = c('[配置] B表类型列')

        # 1. 数值清洗
        df_a[col_a_hrs] = DataEngine.clean_num(df_a, col_a_hrs)
        df_b[col_b_amt] = DataEngine.clean_num(df_b, col_b_amt)
        
        # 【核心修正】Source B 金额单位转换：万元 -> 元
        # 逻辑：先乘以 10000 变成元，再保留2位小数以防万一
        df_b[col_b_amt] = (df_b[col_b_amt] * 10000).round(2)

        # 2. B表人员名称清洗 (关联前处理)
        if col_b_user and col_b_user in df_b.columns:
            df_b[col_b_user] = df_b[col_b_user].astype(str).str.replace('_云计算', '', regex=False).str.strip()

        # 3. A表聚合
        agg_rules = {col_a_hrs: 'sum'}
        for _, col in dims_a.items():
            if col: agg_rules[col] = 'first'
        df_a_gp = df_a.groupby([col_a_user, col_a_spm], as_index=False).agg(agg_rules)

        # 4. B表分流与聚合
        is_sub = df_b[col_b_type].astype(str).str.contains(subsidy_tag, na=False)
        grp_b = [col_b_user, col_b_spm]
        df_sub = df_b[is_sub].groupby(grp_b)[col_b_amt].sum().reset_index(name='差旅补助')
        df_fee = df_b[~is_sub].groupby(grp_b)[col_b_amt].sum().reset_index(name='差旅费控平台')

        # 5. 统一数据类型 (防止 Merge 失败)
        for d in [df_a_gp, df_sub, df_fee]:
            k = col_a_spm if col_a_spm in d.columns else col_b_spm
            d[k] = d[k].astype(str)

        # 6. 关联计算 (Left Join)
        res = pd.merge(df_a_gp, df_sub, left_on=[col_a_user, col_a_spm], right_on=[col_b_user, col_b_spm], how='left')
        res = pd.merge(res, df_fee, left_on=[col_a_user, col_a_spm], right_on=[col_b_user, col_b_spm], how='left')
        res = res.fillna(0)

        # 7. 最终公式计算 & 【精度控制】
        # 保留2位小数
        res['支持时间（人天）'] = res[col_a_hrs] / 8
        res['人力费用'] = (res['支持时间（人天）'] * price_per_day).round(2)
        res['差旅补助'] = res['差旅补助'].round(2)
        res['差旅费控平台'] = res['差旅费控平台'].round(2)
        res['结算费用合计'] = (res['人力费用'] + res['差旅补助'] + res['差旅费控平台']).round(2)

        rename_map = {
            col_a_user: '人员', dims_a['project']: '所属项目', dims_a['range']: '人事范围',
            col_a_spm: 'SPM', dims_a['contract']: '合同主体', dims_a['sales']: '销售人员',
            dims_a['dept']: '销售部门', col_a_hrs: '耗时（小时）'
        }
        t3 = res.rename(columns=rename_map)
        
        final_cols = ['序号','人员','所属项目','人事范围','SPM','合同主体','销售人员','销售部门',
                      '差旅补助','差旅费控平台','耗时（小时）','支持时间（人天）','人力费用','结算费用合计']
        t3.insert(0, '序号', range(1, len(t3)+1))
        t3 = t3[[c for c in final_cols if c in t3.columns]]

        t2_cols = ['人事范围', '合同主体', '销售部门']
        if all(c in t3.columns for c in t2_cols):
            # 聚合表2也需要控制精度
            t2 = t3.groupby(t2_cols).agg({'结算费用合计':'sum', '支持时间（人天）':'sum'}).reset_index()
            t2['结算费用合计'] = t2['结算费用合计'].round(2) 
            t2.columns = ['销售公司', '采购公司', '采购部门', '金额（含税，单位：元）', '工作量（人天）']
            t2.insert(0, '序号', range(1, len(t2)+1))
        else: t2 = pd.DataFrame()

        t1 = t3.groupby('人员')['耗时（小时）'].sum().reset_index()
        t1.rename(columns={'耗时（小时）':'项目工时'}, inplace=True)
        t1.insert(0, '序号', range(1, len(t1)+1))

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
                st.markdown(f"**{title}**")
                return st.file_uploader(title, type=['xlsx', 'csv'], key=f"uploader_{key}", label_visibility="collapsed")
            else:
                c1, c2 = st.columns([9, 1])
                c1.markdown(f"✅ **{data['name']}**")
                if c2.button("Del", key=f"del_{key}"): return "DELETE"
        return None

    @staticmethod
    def render_error_report(err_df, on_fix):
        """
        渲染错误报告，并返回是否需要【重算】
        """
        fixable = err_df[err_df['类型']=='数据错误']
        logic = err_df[err_df['类型']=='逻辑错误']
        rule = err_df[err_df['类型']=='业务规则校验']
        
        st.markdown(f"<div class='error-box'><h3 style='margin:0'>🚨 校验失败</h3><p>发现 <b>{len(fixable)}</b> 个数据错误，<b>{len(rule)}</b> 个规则异常，<b>{len(logic)}</b> 个映射错误。</p></div>", unsafe_allow_html=True)
        
        st.dataframe(err_df[['类型','来源','行号','信息']], use_container_width=True, hide_index=True)
        
        c1, c2, c3 = st.columns([1, 1, 1])
        c1.download_button("📥 下载清单", err_df.to_csv(index=False).encode('utf-8-sig'), "err.csv", "text/csv", use_container_width=True)
        
        should_rerun = c2.button("🔄 参数已改，重新校验", type="secondary", use_container_width=True)
        
        if not fixable.empty:
            if c3.button("🛠️ 在线修复", type="primary", use_container_width=True): on_fix()
            
        return should_rerun

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
    def render_native_editor(desc, subset, is_edit, cols_a, cols_b):
        st.markdown(f'<div class="info-bar">ℹ️ {desc}</div>', unsafe_allow_html=True)
        
        # 1. 准备数据
        df_display = subset[['序号', '目标字段', '源表', '匹配字段', '逻辑说明']].copy().reset_index(drop=True)
        df_display['序号'] = df_display['序号'].astype(str)
        
        # 2. 列配置
        column_config = {
            "序号": st.column_config.TextColumn("序号", width="small", disabled=True),
            "目标字段": st.column_config.TextColumn("目标字段", disabled=True, width="medium"),
            "逻辑说明": st.column_config.TextColumn("逻辑说明", disabled=True, width="large"),
        }

        # 3. 动态配置可编辑列
        if is_edit:
            column_config["源表"] = st.column_config.SelectboxColumn(
                "源表", options=["Source A", "Source B"], width="small", required=True
            )
            column_config["匹配字段"] = st.column_config.SelectboxColumn(
                "匹配字段", options=cols_a + cols_b, width="medium", required=True
            )
        else:
            column_config["源表"] = st.column_config.TextColumn("源表", disabled=True)
            column_config["匹配字段"] = st.column_config.TextColumn("匹配字段", disabled=True)

        # 4. 动态计算表格高度
        calc_height = (len(df_display) + 1) * 35 + 10
        final_height = max(400, min(1000, calc_height))

        # 5. 渲染表格
        edited = st.data_editor(
            df_display,
            column_config=column_config,
            use_container_width=True,
            hide_index=True,          
            disabled=not is_edit,
            height=final_height,
            key=f"editor_{subset.iloc[0]['所属表']}"
        )

        # 6. 保存逻辑
        if is_edit:
            for i, row in edited.iterrows():
                orig_idx = subset.index[i]
                orig_row = st.session_state.mapping_config.loc[orig_idx]
                
                if 'Source' not in orig_row['源表']:
                    continue 
                
                if row['源表'] != orig_row['源表']:
                    st.session_state.mapping_config.at[orig_idx, '源表'] = row['源表']
                    target_opts = cols_a if row['源表'] == 'Source A' else cols_b
                    new_val = row['目标字段'] if row['目标字段'] in target_opts else (target_opts[0] if target_opts else None)
                    st.session_state.mapping_config.at[orig_idx, '匹配字段'] = new_val
                elif row['匹配字段'] != orig_row['匹配字段']:
                    valid_opts = cols_a if row['源表'] == 'Source A' else cols_b
                    if row['匹配字段'] in valid_opts:
                        st.session_state.mapping_config.at[orig_idx, '匹配字段'] = row['匹配字段']

# ==============================================================================
# Zone C: 控制层
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
                            na = st.data_editor(da, height=300, hide_index=True, column_config={"_sys_id": None}, key="fix_a")
                        else: st.info("无数据错误")
                    with t2:
                        if not db.empty:
                            nb = st.data_editor(db, height=300, hide_index=True, column_config={"_sys_id": None}, key="fix_b")
                        else: st.info("无数据错误")
                    
                    if st.button("💾 保存并重算", type="primary"):
                        if na is not None:
                            od = st.session_state.data_store['A']['df'].set_index('_sys_id')
                            od.update(na.set_index('_sys_id'))
                            st.session_state.data_store['A']['df'] = od.reset_index()
                        if nb is not None:
                            od = st.session_state.data_store['B']['df'].set_index('_sys_id')
                            od.update(nb.set_index('_sys_id'))
                            st.session_state.data_store['B']['df'] = od.reset_index()
                        st.session_state.error_report = None
                        st.session_state.block_auto_run = False
                        st.rerun()
                show_fix()
            
            should_rerun = UIComponents.render_error_report(st.session_state.error_report, fix_action)
            if should_rerun:
                st.session_state.error_report = None
                trigger = True

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
    c1, c2 = st.columns([9, 1])
    c1.markdown("<div class='nav-header'>🐱 字段映射 & 逻辑配置</div>", unsafe_allow_html=True)
    if c2.button("⬅️", use_container_width=True, help="返回主页"): 
        st.session_state.page = 'main'
        st.rerun()
    
    st.markdown("<hr style='margin-top:0; border-color:#30363d;'>", unsafe_allow_html=True)

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

    df_c = st.session_state.mapping_config
    t1, t2, t3 = st.tabs(["结果表3 (底表)", "结果表2 (结算)", "结果表1 (工时)"])
    
    cols_a = list(st.session_state.data_store['A']['df'].columns) if has_files else []
    cols_b = list(st.session_state.data_store['B']['df'].columns) if has_files else []

    with t1:
        UIComponents.render_native_editor(
            "全量明细底表：基于 Source A/B 进行清洗、聚合、关联计算后的宽表。",
            df_c[df_c['所属表']=='结果表3'], st.session_state.is_editing_mapping, cols_a, cols_b
        )
    with t2:
        UIComponents.render_native_editor(
            "结算汇总表：基于【结果表3】按公司/部门维度二次聚合的金额数据。",
            df_c[df_c['所属表']=='结果表2'], st.session_state.is_editing_mapping, cols_a, cols_b
        )
    with t3:
        UIComponents.render_native_editor(
            "工时统计表：基于【结果表3】按人员维度二次聚合的工时数据。",
            df_c[df_c['所属表']=='结果表1'], st.session_state.is_editing_mapping, cols_a, cols_b
        )
