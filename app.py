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
        /* [UI] 按钮 */
        .ghost-btn button { border: 1px dashed #444 !important; color: #888 !important; background: transparent !important; }
        .cat-btn button { border: 1px solid #30363d !important; background: #161b22 !important; color: #c9d1d9 !important; width: 100%; margin-top: 10px; }
        .cat-btn button:hover { border-color: #a371f7 !important; color: #a371f7 !important; }
        
        /* [UI] 映射表 */
        .mapping-table { border: 1px solid #30363d; border-radius: 6px; overflow: hidden; background-color: #0d1117; margin-bottom: 20px; }
        .map-header-row { background-color: #161b22; border-bottom: 1px solid #30363d; padding: 12px 10px; font-weight: 600; color: #8b949e; display: flex; align-items: center; }
        .map-data-row { padding: 10px 10px; border-bottom: 1px solid #21262d; display: flex; align-items: center; background-color: #0d1117; }
        .source-tag { background: rgba(56, 139, 253, 0.1); border: 1px solid rgba(56, 139, 253, 0.4); border-radius: 10px; padding: 2px 8px; font-size: 0.7rem; color: #58a6ff; }
        
        div[data-testid="stFileUploader"] section > div:first-child { display: none; }
        div[data-testid="stFileUploader"] { padding-top: 15px; }
    </style>
    """, unsafe_allow_html=True)

# ==============================================================================
# Zone A: 纯逻辑层 (DataEngine) - 核心清洗与计算
# ==============================================================================
class DataEngine:
    @staticmethod
    def get_default_config():
        """定义字段映射默认配置 - 基于用户提供的真实文件列名优化"""
        return pd.DataFrame([
            # --- 表A (工时) ---
            {"所属表": "表3", "目标字段": "人员", "来源": "Source A", "匹配字段": "人员", "计算逻辑": "主键 (Join Key)"},
            {"所属表": "表3", "目标字段": "SPM", "来源": "Source A", "匹配字段": "SPM", "计算逻辑": "主键 (Join Key)"},
            {"所属表": "表3", "目标字段": "工时", "来源": "Source A", "匹配字段": "交付工时", "计算逻辑": "SUM聚合 (清洗: 去逗号转数字)"},
            # 维度
            {"所属表": "表3", "目标字段": "所属项目", "来源": "Source A", "匹配字段": "所属项目", "计算逻辑": "维度 (First)"},
            {"所属表": "表3", "目标字段": "人事范围", "来源": "Source A", "匹配字段": "人事范围", "计算逻辑": "维度 (对应销售公司)"},
            {"所属表": "表3", "目标字段": "合同主体", "来源": "Source A", "匹配字段": "合同主体", "计算逻辑": "维度 (对应采购公司)"},
            {"所属表": "表3", "目标字段": "销售人员", "来源": "Source A", "匹配字段": "销售", "计算逻辑": "维度 (First)"},
            {"所属表": "表3", "目标字段": "销售部门", "来源": "Source A", "匹配字段": "销售部门", "计算逻辑": "维度 (对应采购部门)"},
            
            # --- 表B (费用) ---
            # 注意：根据用户注释，表2的人员字段叫“出差人”
            {"所属表": "表3", "目标字段": "人员 (B)", "来源": "Source B", "匹配字段": "出差人", "计算逻辑": "外键 (Join Key)"},
            {"所属表": "表3", "目标字段": "SPM (B)", "来源": "Source B", "匹配字段": "SPM", "计算逻辑": "外键 (Join Key)"},
            {"所属表": "表3", "目标字段": "金额", "来源": "Source B", "匹配字段": "金额", "计算逻辑": "SUM聚合 (清洗: 去逗号转数字)"},
            {"所属表": "表3", "目标字段": "费用类型", "来源": "Source B", "匹配字段": "产品类型", "计算逻辑": "分类依据 (含关键词则为补助)"},
        ])

    @staticmethod
    def get_col(config_df, target):
        """根据目标字段名获取用户配置的源列名"""
        row = config_df[config_df['目标字段'] == target]
        if row.empty: return None
        return str(row.iloc[0]['匹配字段']).strip()

    @staticmethod
    def clean_num(df, col):
        """清洗数字列: 转字符串 -> 去逗号 -> 转数字 -> 填0"""
        if col not in df.columns: return 0
        return pd.to_numeric(df[col].astype(str).str.replace(',', ''), errors='coerce').fillna(0)

    @staticmethod
    def validate(df_a, df_b, config_df, min_hours):
        """
        校验逻辑:
        1. 检查所有配置的列是否存在
        2. 检查工时/金额是否为负
        3. 检查SPM是否为空
        4. 检查工时阈值
        5. 检查孤立费用 (B表有费用但A表没工时)
        """
        errors = []
        c = lambda t: DataEngine.get_col(config_df, t)
        
        # 1. 映射字典
        map_a = {k: c(k) for k in ['人员','SPM','工时','所属项目','人事范围','合同主体','销售人员','销售部门']}
        map_b = {k: c(k) for k in ['人员 (B)','SPM (B)','金额','费用类型']}
        
        # 2. 列存在性检查
        for target, col in map_a.items():
            if col and col not in df_a.columns:
                errors.append({'类型':'逻辑错误', '来源':'Source A', '_sys_id':'-', '行号':'-', '信息':f'缺列: {col} (目标:{target})'})
        for target, col in map_b.items():
            if col and col not in df_b.columns:
                errors.append({'类型':'逻辑错误', '来源':'Source B', '_sys_id':'-', '行号':'-', '信息':f'缺列: {col} (目标:{target})'})
        
        if errors: return errors, df_a, df_b # 缺列直接阻断

        # 3. 数据清洗 (临时清洗用于校验)
        df_a_clean = df_a.copy()
        df_b_clean = df_b.copy()
        df_a_clean[map_a['工时']] = DataEngine.clean_num(df_a_clean, map_a['工时'])
        df_b_clean[map_b['金额']] = DataEngine.clean_num(df_b_clean, map_b['金额'])

        # 4. 数据错误校验
        # 负数检查
        for i, r in df_a_clean[df_a_clean[map_a['工时']] < 0].iterrows():
            errors.append({'类型':'数据错误', '来源':'Source A', '_sys_id':r['_sys_id'], '行号':r['_sys_id'], '信息':'工时为负'})
        for i, r in df_b_clean[df_b_clean[map_b['金额']] < 0].iterrows():
            errors.append({'类型':'数据错误', '来源':'Source B', '_sys_id':r['_sys_id'], '行号':r['_sys_id'], '信息':'金额为负'})
        # 空值检查
        for i, r in df_a[df_a[map_a['SPM']].isnull() | (df_a[map_a['SPM']] == '')].iterrows():
             errors.append({'类型':'数据错误', '来源':'Source A', '_sys_id':r['_sys_id'], '行号':r['_sys_id'], '信息':'SPM为空'})

        # 5. 业务逻辑校验
        # 阈值
        agg = df_a_clean.groupby(map_a['人员'])[map_a['工时']].sum()
        for n, h in agg.items():
            if h < min_hours:
                 errors.append({'类型':'逻辑错误', '来源':'Source A', '_sys_id':'-', '行号':'-', '信息':f'人员[{n}]总工时({h}) < 阈值'})
        
        # 孤立费用 (Match Key: 人员+SPM)
        df_a['key'] = df_a[map_a['人员']].astype(str) + "_" + df_a[map_a['SPM']].astype(str)
        df_b['key'] = df_b[map_b['人员 (B)']].astype(str) + "_" + df_b[map_b['SPM (B)']].astype(str)
        
        orphans = df_b[~df_b['key'].isin(df_a['key'])]
        # 去重报错，防止同一笔孤立费用报多次
        for key in orphans['key'].unique():
             errors.append({'类型':'逻辑错误', '来源':'Source B', '_sys_id':'-', '行号':'-', '信息':f'发现孤立费用，无法匹配到交付人员: {key}'})

        return errors, df_a, df_b

    @staticmethod
    def calculate(df_a, df_b, config_df, price_per_day, subsidy_tag):
        """
        核心计算逻辑:
        1. 清洗数据
        2. A表聚合 (工时求和)
        3. B表拆分 (补助/费控) 并聚合
        4. Left Join (以A表为主)
        5. 计算最终金额 (工时费 + 差旅费)
        """
        c = lambda t: DataEngine.get_col(config_df, t)
        
        # 获取映射列名
        col_a_user = c('人员')
        col_a_spm = c('SPM')
        col_a_hours = c('工时')
        # 维度
        dims_a = {
            'project': c('所属项目'), 'range': c('人事范围'), 'contract': c('合同主体'),
            'sales': c('销售人员'), 'dept': c('销售部门')
        }
        
        col_b_user = c('人员 (B)')
        col_b_spm = c('SPM (B)')
        col_b_amt = c('金额')
        col_b_type = c('费用类型')

        # 1. 清洗数值
        df_a[col_a_hours] = DataEngine.clean_num(df_a, col_a_hours)
        df_b[col_b_amt] = DataEngine.clean_num(df_b, col_b_amt)

        # 2. 聚合 A表 (工时)
        # 规则: 按[人员, SPM]分组, 工时求和, 其他维度取第一条
        agg_rules = {col_a_hours: 'sum'}
        for _, col in dims_a.items():
            if col: agg_rules[col] = 'first'
        
        df_a_gp = df_a.groupby([col_a_user, col_a_spm], as_index=False).agg(agg_rules)

        # 3. 拆分与聚合 B表 (费用)
        # 规则: 按[人员, SPM]分组, 分别计算补助和其他
        is_sub = df_b[col_b_type].astype(str).str.contains(subsidy_tag, na=False)
        grp_b = [col_b_user, col_b_spm]
        
        df_sub = df_b[is_sub].groupby(grp_b)[col_b_amt].sum().reset_index(name='差旅补助')
        df_fee = df_b[~is_sub].groupby(grp_b)[col_b_amt].sum().reset_index(name='差旅费控平台')

        # 4. 合并 (Left Join: A left join B)
        # 统一关联键类型
        df_a_gp[col_a_spm] = df_a_gp[col_a_spm].astype(str)
        df_sub[col_b_spm] = df_sub[col_b_spm].astype(str)
        df_fee[col_b_spm] = df_fee[col_b_spm].astype(str)
        
        res = pd.merge(df_a_gp, df_sub, left_on=[col_a_user, col_a_spm], right_on=[col_b_user, col_b_spm], how='left')
        res = pd.merge(res, df_fee, left_on=[col_a_user, col_a_spm], right_on=[col_b_user, col_b_spm], how='left')
        res = res.fillna(0) # 没匹配到的费用填0

        # 5. 计算金额
        # 核心公式: (交付工时 / 8) * 单价 + 补助 + 费控
        res['支持时间(人天)'] = res[col_a_hours] / 8
        res['人力费用'] = res['支持时间(人天)'] * price_per_day
        res['结算费用合计'] = res['人力费用'] + res['差旅补助'] + res['差旅费控平台']

        # 6. 生成表3 (详细明细 - 底表)
        rename_map = {
            col_a_user: '人员', dims_a['project']: '所属项目', dims_a['range']: '人事范围',
            col_a_spm: 'SPM', dims_a['contract']: '合同主体', dims_a['sales']: '销售人员',
            dims_a['dept']: '销售部门', col_a_hours: '耗时(小时)', 
            '支持时间(人天)': '支持时间(人天)', '结算费用合计': '结算费用合计'
        }
        # 仅重命名存在的列
        rename_map = {k:v for k,v in rename_map.items() if k in res.columns}
        t3 = res.rename(columns=rename_map)
        
        final_cols = ['序号','人员','所属项目','人事范围','SPM','合同主体','销售人员','销售部门',
                      '差旅补助','差旅费控平台','耗时(小时)','支持时间(人天)','人力费用','结算费用合计']
        t3.insert(0, '序号', range(1, len(t3)+1))
        # 按顺序提取存在的列
        t3 = t3[[c for c in final_cols if c in t3.columns]]

        # 7. 生成表2 (结算汇总)
        # 维度: 销售公司(人事范围) + 采购公司(合同主体) + 采购部门(销售部门)
        dims_t2 = ['人事范围', '合同主体', '销售部门']
        if all(c in t3.columns for c in dims_t2):
            t2 = t3.groupby(dims_t2).agg({'结算费用合计': 'sum', '支持时间(人天)': 'sum'}).reset_index()
            t2.columns = ['销售公司', '采购公司', '采购部门', '金额(含税,单位:元)', '工作量(人天)']
            t2.insert(0, '序号', range(1, len(t2)+1))
        else:
            t2 = pd.DataFrame({'提示': ['缺少维度字段，无法生成']})

        # 8. 生成表1 (工时统计)
        # 维度: 人员
        if '人员' in t3.columns and '耗时(小时)' in t3.columns:
            t1 = t3.groupby('人员')['耗时(小时)'].sum().reset_index()
            t1.rename(columns={'耗时(小时)': '项目工时'}, inplace=True)
            t1.insert(0, '序号', range(1, len(t1)+1))
        else:
            t1 = pd.DataFrame({'提示': ['缺少人员或工时字段']})

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
            if st.button("🐱 字段映射配置", help="自定义匹配规则"):
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
    def render_mapping_table(subset, is_edit, cols_a, cols_b):
        st.markdown("""<div class="mapping-table"><div class="map-header-row"><div style="width:25%">目标字段</div><div style="width:35%">匹配列</div><div style="width:15%">来源表</div><div style="width:25%">逻辑说明</div></div>""", unsafe_allow_html=True)
        for idx, row in subset.iterrows():
            st.markdown('<div class="map-data-row">', unsafe_allow_html=True)
            c1, c2, c3, c4 = st.columns([2.5, 3.5, 1.5, 2.5])
            c1.markdown(f"**{row['目标字段']}**")
            with c2:
                if is_edit:
                    opts = cols_a if row['来源']=='Source A' else cols_b
                    cur = row['匹配字段']
                    if cur not in opts: opts = [cur] + opts
                    new_val = st.selectbox("s", opts, index=opts.index(cur), key=f"s_{idx}", label_visibility="collapsed")
                    st.session_state.mapping_config.at[idx, '匹配字段'] = new_val
                else: st.markdown(f"<span style='color:#a5d6ff; font-family:monospace;'>{row['匹配字段']}</span>", unsafe_allow_html=True)
            c3.markdown(f"<span class='source-tag'>{row['来源']}</span>", unsafe_allow_html=True)
            c4.markdown(f"<span style='color:#666; font-size:0.8rem;'>{row['计算逻辑']}</span>", unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)

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
            with st.container(border=True):
                st.success("✅ 生成完毕")
                st.download_button("📦 批量下载 (ZIP)", st.session_state.result_zip, "report.zip", type="primary", use_container_width=True)
                st.markdown("---")
                c1, c2, c3 = st.columns(3)
                fs = st.session_state.result_files
                if 't1' in fs: c1.download_button("📥 表1", fs['t1'], "t1.xlsx", use_container_width=True)
                if 't2' in fs: c2.download_button("📥 表2", fs['t2'], "t2.xlsx", use_container_width=True)
                if 't3' in fs: c3.download_button("📥 表3", fs['t3'], "t3.xlsx", use_container_width=True)
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
    st.markdown("### 🐱 字段映射 & 逻辑配置")
    c1, c2 = st.columns([1, 4])
    if c1.button("⬅️ 返回主页", use_container_width=True): 
        st.session_state.page = 'main'
        st.rerun()
    
    with c2:
        c_status, c_edit = st.columns([3, 1])
        with c_edit:
            has_files = st.session_state.data_store['A']['df'] is not None
            if not st.session_state.is_editing_mapping:
                if st.button("✏️ 编辑配置", type="primary", use_container_width=True):
                    if not has_files: st.toast("请先上传文件", icon="🚫")
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
    
    st.divider()
    cols_a = list(st.session_state.data_store['A']['df'].columns) if st.session_state.data_store['A']['df'] is not None else []
    cols_b = list(st.session_state.data_store['B']['df'].columns) if st.session_state.data_store['B']['df'] is not None else []
    
    df_c = st.session_state.mapping_config
    t1, t2, t3 = st.tabs(["A表核心", "B表核心", "维度配置"])
    
    with t1: UIComponents.render_mapping_table(df_c[df_c['来源']=='Source A'][0:3], st.session_state.is_editing_mapping, cols_a, cols_b)
    with t2: UIComponents.render_mapping_table(df_c[df_c['来源']=='Source B'], st.session_state.is_editing_mapping, cols_a, cols_b)
    with t3: UIComponents.render_mapping_table(df_c[df_c['来源']=='Source A'][3:], st.session_state.is_editing_mapping, cols_a, cols_b)
