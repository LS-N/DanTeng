import streamlit as st
import pandas as pd
import io
import time
import zipfile
from st_aggrid import AgGrid, GridOptionsBuilder

# ==============================================================================
# Zone 0: 全局配置 & 样式注入 (Global Config)
# ==============================================================================
st.set_page_config(page_title="淡藤财务报表 Pro", page_icon="😈", layout="wide", initial_sidebar_state="expanded")

def inject_css():
    st.markdown("""
    <style>
        :root { --bg-color: #0d1117; --card-bg: #161b22; --accent: #238636; --text: #c9d1d9; --red: #da3633; --border: #30363d; }
        .stApp { background-color: var(--bg-color); color: var(--text); }
        
        /* [UI组件] 文件卡片 */
        .file-card-styled { 
            background: #21262d; border-left: 4px solid #238636; border-radius: 4px; padding: 15px; 
            width: 100%; height: 100%; display: flex; align-items: center; justify-content: space-between; 
        }
        .close-btn { cursor: pointer; color: #666; font-size: 1.2rem; line-height: 1; padding: 5px; }
        .close-btn:hover { color: #da3633; }

        /* [UI组件] 错误舱 */
        .error-box { border: 1px solid var(--red); background: rgba(218, 54, 51, 0.1); border-radius: 8px; padding: 1.5rem; margin-top: 1rem; }
        
        /* [UI组件] 幽灵按钮 */
        .ghost-btn button { border: 1px dashed #444 !important; color: #888 !important; background: transparent !important; padding: 0.2rem 0.8rem !important; height: auto !important; font-size: 0.8rem !important; }
        .ghost-btn button:hover { border-color: var(--red) !important; color: var(--red) !important; }

        /* [UI组件] 侧边栏猫猫按钮 */
        .cat-btn button {
            border: 1px solid #30363d !important; background: #161b22 !important; color: #c9d1d9 !important;
            border-radius: 6px !important; padding: 0.5rem 1rem !important; width: 100%; text-align: center; margin-top: 10px;
        }
        .cat-btn button:hover { border-color: #a371f7 !important; color: #a371f7 !important; }

        /* [UI组件] 映射表样式 */
        .map-header {
            background-color: #21262d; color: #8b949e; font-weight: bold; font-size: 0.9rem;
            padding: 10px 5px; border-bottom: 2px solid #30363d; margin-bottom: 5px;
        }
        .map-row {
            padding: 8px 5px; border-bottom: 1px solid #21262d; display: flex; align-items: center;
        }
        .source-tag {
            background: #10151b; border: 1px solid #30363d; border-radius: 4px; padding: 2px 6px; font-size: 0.75rem; color: #8b949e;
        }
        
        /* Streamlit 补丁 */
        div[data-testid="stDialog"] > div[role="dialog"] { width: 80vw !important; max-width: 1200px !important; margin: auto !important; }
        div[data-testid="stFileUploader"] section > div:first-child { display: none; }
        div[data-testid="stFileUploader"] { padding-top: 15px; }
        div[data-testid="stSelectbox"] > div > div { min-height: 38px; }
    </style>
    """, unsafe_allow_html=True)

# ==============================================================================
# Zone A: 纯逻辑层 (Logic Layer) - 严禁包含 UI 代码
# ==============================================================================
class DataEngine:
    @staticmethod
    def get_default_config():
        """返回默认字段映射配置"""
        return pd.DataFrame([
            # 表1: 工时
            {"所属表": "表1", "目标字段": "人员", "来源": "Source A", "匹配字段": "姓名", "计算逻辑": "主键 (分组依据)"},
            {"所属表": "表1", "目标字段": "项目工时", "来源": "Source A", "匹配字段": "交付工时", "计算逻辑": "SUM聚合"},
            # 表2: 结算
            {"所属表": "表2", "目标字段": "销售公司", "来源": "Source A", "匹配字段": "人事范围", "计算逻辑": "维度展示"},
            {"所属表": "表2", "目标字段": "采购公司", "来源": "Source A", "匹配字段": "合同主体", "计算逻辑": "维度展示"},
            {"所属表": "表2", "目标字段": "采购部门", "来源": "Source A", "匹配字段": "销售部门", "计算逻辑": "维度展示"},
            # 表3: 明细 (A)
            {"所属表": "表3", "目标字段": "SPM", "来源": "Source A", "匹配字段": "SPM", "计算逻辑": "主键 (匹配用)"},
            {"所属表": "表3", "目标字段": "工时", "来源": "Source A", "匹配字段": "交付工时", "计算逻辑": "核心计算"},
            {"所属表": "表3", "目标字段": "姓名", "来源": "Source A", "匹配字段": "姓名", "计算逻辑": "主键 (匹配用)"},
            {"所属表": "表3", "目标字段": "项目名称", "来源": "Source A", "匹配字段": "所属项目", "计算逻辑": "维度"},
            # 表3: 明细 (B)
            {"所属表": "表3", "目标字段": "SPM (B)", "来源": "Source B", "匹配字段": "费用归属项目", "计算逻辑": "外键"},
            {"所属表": "表3", "目标字段": "金额", "来源": "Source B", "匹配字段": "报销金额", "计算逻辑": "SUM"},
            {"所属表": "表3", "目标字段": "姓名 (B)", "来源": "Source B", "匹配字段": "报销人", "计算逻辑": "外键"},
            {"所属表": "表3", "目标字段": "费用类型", "来源": "Source B", "匹配字段": "费用类型", "计算逻辑": "分类"},
        ])

    @staticmethod
    def get_config_col(df_conf, source, target):
        """从配置表中查找匹配列名"""
        row = df_conf[(df_conf['来源'] == source) & (df_conf['目标字段'] == target)]
        if row.empty: return None
        return str(row.iloc[0]['匹配字段']).strip()

    @staticmethod
    def validate_and_clean(df_a, df_b, config_df, min_hours):
        """执行所有校验逻辑，返回错误列表和清洗后的数据"""
        errors = []
        
        # 1. 动态获取列名
        def gc(s, t): return DataEngine.get_config_col(config_df, s, t)
        
        ca_spm, ca_hrs, ca_name = gc('Source A','SPM'), gc('Source A','工时'), gc('Source A','姓名')
        cb_spm, cb_amt, cb_name = gc('Source B','SPM (B)'), gc('Source B','金额'), gc('Source B','姓名 (B)')
        
        # 2. 检查列是否存在 (逻辑错误)
        def check(df, col, src, target):
            if col not in df.columns:
                errors.append({'类型': '逻辑错误', '来源': src, '_sys_id': '-', '行号': '-', '信息': f'未找到列[{col}] (目标:{target})'})
                return False
            return True

        has_cols_a = check(df_a, ca_spm, 'Source A', 'SPM') and check(df_a, ca_hrs, 'Source A', '工时') and check(df_a, ca_name, 'Source A', '姓名')
        has_cols_b = check(df_b, cb_spm, 'Source B', 'SPM (B)') and check(df_b, cb_amt, 'Source B', '金额') and check(df_b, cb_name, 'Source B', '姓名 (B)')

        if not errors:
            # 3. 数据清洗
            df_a[ca_hrs] = pd.to_numeric(df_a[ca_hrs], errors='coerce').fillna(0)
            if df_b[cb_amt].dtype == object: df_b[cb_amt] = df_b[cb_amt].astype(str).str.replace(',', '')
            df_b[cb_amt] = pd.to_numeric(df_b[cb_amt], errors='coerce').fillna(0)

            # 4. 数据错误校验 (负数、空值)
            for i, r in df_a[df_a[ca_hrs] < 0].iterrows():
                errors.append({'类型': '数据错误', '来源': 'Source A', '_sys_id': r['_sys_id'], '行号': r['_sys_id'], '信息': '工时为负'})
            for i, r in df_b[df_b[cb_amt] < 0].iterrows():
                errors.append({'类型': '数据错误', '来源': 'Source B', '_sys_id': r['_sys_id'], '行号': r['_sys_id'], '信息': '金额为负'})
            for i, r in df_a[df_a[ca_spm].isnull() | (df_a[ca_spm] == '')].iterrows():
                errors.append({'类型': '数据错误', '来源': 'Source A', '_sys_id': r['_sys_id'], '行号': r['_sys_id'], '信息': 'SPM为空'})

            # 5. 业务逻辑校验 (阈值、孤立匹配)
            agg = df_a.groupby(ca_name)[ca_hrs].sum()
            for n, h in agg.items():
                if h < min_hours:
                    errors.append({'类型': '逻辑错误', '来源': 'Source A', '_sys_id': '-', '行号': '-', '信息': f'人员[{n}]总工时({h}) < 阈值'})
            
            df_a['key'] = df_a[ca_name].astype(str) + "_" + df_a[ca_spm].astype(str)
            df_b['key'] = df_b[cb_name].astype(str) + "_" + df_b[cb_spm].astype(str)
            orphans = df_b[~df_b['key'].isin(df_a['key'])]
            for i, r in orphans.iterrows():
                errors.append({'类型': '逻辑错误', '来源': 'Source B', '_sys_id': '-', '行号': '-', '信息': f'无法匹配到交付人员: {r["key"]}'})

        return errors, df_a, df_b

    @staticmethod
    def calculate_results(df_a, df_b, config_df, price_per_day, subsidy_tag):
        """执行合并与计算，返回三个结果表"""
        def gc(s, t): return DataEngine.get_config_col(config_df, s, t)
        
        # 字段别名准备
        ca_spm, ca_hrs, ca_name = gc('Source A','SPM'), gc('Source A','工时'), gc('Source A','姓名')
        cb_spm, cb_amt, cb_name = gc('Source B','SPM (B)'), gc('Source B','金额'), gc('Source B','姓名 (B)')
        
        # 辅助维度
        ca_proj = gc('Source A','项目名称') or df_a.columns[0]
        ca_range = gc('Source A','人事范围') or df_a.columns[0]
        ca_contract = gc('Source A','合同主体') or df_a.columns[0]
        ca_sales = gc('Source A','销售部门') or df_a.columns[0]
        cb_type = gc('Source B','费用类型') or df_b.columns[0]

        # 1. 聚合 A
        agg_rules = {ca_hrs: 'sum'}
        for c in [ca_proj, ca_range, ca_contract, ca_sales]: agg_rules[c] = 'first'
        df_a_gp = df_a.groupby([ca_name, ca_spm], as_index=False).agg(agg_rules)

        # 2. 拆分 B
        is_sub = df_b[cb_type].astype(str).str.contains(subsidy_tag, na=False)
        grp_b = [cb_name, cb_spm]
        df_sub = df_b[is_sub].groupby(grp_b)[cb_amt].sum().reset_index(name='差旅补助')
        df_fee = df_b[~is_sub].groupby(grp_b)[cb_amt].sum().reset_index(name='差旅费控平台')

        # 3. 合并
        for d in [df_a_gp, df_sub, df_fee]:
            k = ca_spm if ca_spm in d.columns else cb_spm
            d[k] = d[k].astype(str)

        res = pd.merge(df_a_gp, df_sub, left_on=[ca_name, ca_spm], right_on=[cb_name, cb_spm], how='left')
        res = pd.merge(res, df_fee, left_on=[ca_name, ca_spm], right_on=[cb_name, cb_spm], how='left')
        res = res.fillna(0)

        # 4. 金额计算
        res['支持时间(人天)'] = res[ca_hrs] / 8
        res['人力费用'] = res['支持时间(人天)'] * price_per_day
        res['结算费用合计'] = res['人力费用'] + res['差旅补助'] + res['差旅费控平台']

        # 5. 生成 T3
        t3 = res.rename(columns={ca_name:'人员', ca_proj:'所属项目', ca_range:'人事范围', ca_spm:'SPM',
                                 ca_contract:'合同主体', ca_sales:'销售部门', ca_hrs:'耗时(小时)'})
        final_cols = ['序号','人员','所属项目','人事范围','SPM','合同主体','销售部门',
                      '差旅补助','差旅费控平台','耗时(小时)','支持时间(人天)','人力费用','结算费用合计']
        t3.insert(0, '序号', range(1, len(t3)+1))
        t3 = t3[[c for c in final_cols if c in t3.columns]]

        # 6. 生成 T2
        dims = [c for c in ['人事范围','合同主体','销售部门'] if c in t3.columns]
        if dims:
            t2 = t3.groupby(dims).agg({'结算费用合计':'sum', '支持时间(人天)':'sum'}).reset_index()
            t2.columns = ['销售公司','采购公司','采购部门','金额(含税,单位:元)','工作量(人天)']
            t2.insert(0, '序号', range(1, len(t2)+1))
        else: t2 = pd.DataFrame()

        # 7. 生成 T1
        t1 = t3.groupby('人员')['耗时(小时)'].sum().reset_index()
        t1.rename(columns={'耗时(小时)':'项目工时'}, inplace=True)
        t1.insert(0, '序号', range(1, len(t1)+1))

        return {'t1': t1, 't2': t2, 't3': t3}

    @staticmethod
    def to_bytes(df):
        b = io.BytesIO()
        # 导出时剔除系统列
        out = df.drop(columns=['_sys_id'], errors='ignore')
        out.to_excel(b, index=False)
        return b.getvalue()

# ==============================================================================
# Zone B: 纯 UI 组件层 (View Layer) - 只负责渲染
# ==============================================================================
class UIComponents:
    @staticmethod
    def render_sidebar():
        """侧边栏渲染"""
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
        """文件上传/展示槽位"""
        data = data_store[key]
        has_file = data['df'] is not None
        
        # 容器
        with st.container(height=180, border=True):
            if not has_file:
                st.markdown(f"<div style='text-align:center; color:#888; margin-top:30px; margin-bottom:10px;'>{title}</div>", unsafe_allow_html=True)
                return st.file_uploader(title, type=['xlsx', 'csv'], key=f"uploader_{key}", label_visibility="collapsed")
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
                    if st.button("Del", key=f"del_{key}"): return "DELETE"
        return None

    @staticmethod
    def render_error_report(err_df, on_fix_click):
        """错误报告舱"""
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
        
        if not fixable.empty:
            if c2.button("🛠️ 打开在线修复", type="primary", use_container_width=True): on_fix_click()
        else:
            c2.warning("⚠️ 存在逻辑错误，无法在线修复，请检查源文件或配置。")

    @staticmethod
    def render_mapping_table_row(row, is_editing, cols_a, cols_b, idx):
        """渲染映射表的一行"""
        r1, r2, r3, r4 = st.columns([2, 3, 1.5, 3])
        with r1: st.markdown(f"<div style='padding-top:10px; font-weight:bold;'>{row['目标字段']}</div>", unsafe_allow_html=True)
        with r2:
            if is_editing:
                options = cols_a if row['来源'] == 'Source A' else cols_b
                curr = row['匹配字段']
                if curr not in options: options = [curr] + options # 容错
                idx_val = options.index(curr) if curr in options else 0
                new_val = st.selectbox("s", options, index=idx_val, key=f"s_{idx}", label_visibility="collapsed")
                return new_val
            else:
                st.markdown(f"<div style='padding-top:10px; color:#a5d6ff; font-family:monospace;'>{row['匹配字段']}</div>", unsafe_allow_html=True)
                return None
        with r3: st.markdown(f"<div style='padding-top:10px;'><span class='source-tag'>{row['来源']}</span></div>", unsafe_allow_html=True)
        with r4: st.markdown(f"<div style='padding-top:10px; font-size:0.85rem; color:#888;'>{row['计算逻辑']}</div>", unsafe_allow_html=True)
        st.markdown("<div style='border-bottom: 1px solid #21262d; margin-bottom: 5px;'></div>", unsafe_allow_html=True)
        return None

    @staticmethod
    def render_download_zone(result_files, result_zip):
        """下载区"""
        with st.container(border=True):
            st.success("✅ 生成完毕")
            st.download_button("📦 批量下载 (ZIP)", result_zip, "report.zip", type="primary", use_container_width=True)
            st.markdown("---")
            c1, c2, c3 = st.columns(3)
            if 't1' in result_files: c1.download_button("📥 表1", result_files['t1'], "t1.xlsx", use_container_width=True)
            if 't2' in result_files: c2.download_button("📥 表2", result_files['t2'], "t2.xlsx", use_container_width=True)
            if 't3' in result_files: c3.download_button("📥 表3", result_files['t3'], "t3.xlsx", use_container_width=True)

# ==============================================================================
# Zone C: 控制层 (Controller) - 负责调度
# ==============================================================================

# 1. 状态初始化
if 'page' not in st.session_state: st.session_state.page = 'main'
if 'data_store' not in st.session_state: st.session_state.data_store = {'A': {'df': None, 'name': None}, 'B': {'df': None, 'name': None}}
if 'mapping_config' not in st.session_state: st.session_state.mapping_config = DataEngine.get_default_config()
if 'is_calculated' not in st.session_state: st.session_state.is_calculated = False
if 'error_report' not in st.session_state: st.session_state.error_report = None
if 'block_auto_run' not in st.session_state: st.session_state.block_auto_run = False
if 'is_editing_mapping' not in st.session_state: st.session_state.is_editing_mapping = False

# 2. 注入样式
inject_css()

# 3. 路由分发
if st.session_state.page == 'main':
    # --- 主页面流程 ---
    
    # 3.1 渲染侧边栏
    price, hours_limit, sub_tag = UIComponents.render_sidebar()
    
    st.title("😈 淡藤财务报表 Pro")
    
    # 3.2 渲染数据控制台
    with st.container(border=True):
        c_h1, c_h2 = st.columns([8, 1])
        c_h1.markdown("### 📂 数据源控制台")
        with c_h2:
            st.markdown('<div class="ghost-btn">', unsafe_allow_html=True)
            if st.button("🗑️ 重置"): 
                st.session_state.clear()
                st.rerun()
            st.markdown('</div>', unsafe_allow_html=True)

        c_u1, c_u2 = st.columns(2)
        
        # 3.3 处理文件上传/删除事件
        def handle_file(key, title):
            res = UIComponents.render_file_slot(key, title, st.session_state.data_store)
            if res == "DELETE":
                st.session_state.data_store[key] = {'df': None, 'name': None}
                st.session_state.is_calculated = False
                st.session_state.error_report = None
                st.session_state.block_auto_run = False
                st.rerun()
            elif res is not None:
                try:
                    if res.name.endswith('.csv'): df = pd.read_csv(res)
                    else: df = pd.read_excel(res)
                    df.columns = [str(c).strip() for c in df.columns]
                    df['_sys_id'] = range(1, len(df)+1)
                    st.session_state.data_store[key] = {'df': df, 'name': res.name}
                    if st.session_state.block_auto_run: st.session_state.error_report = None
                    st.rerun()
                except Exception as e: st.error(f"Err: {e}")

        handle_file('A', "Source A: 投入明细")
        handle_file('B', "Source B: 差旅明细")

    st.divider()
    
    # 3.4 核心逻辑判断
    has_files = st.session_state.data_store['A']['df'] is not None and st.session_state.data_store['B']['df'] is not None
    trigger_run = False
    
    if has_files:
        if st.session_state.is_calculated:
            # 已完成态
            UIComponents.render_download_zone(st.session_state.result_files, st.session_state.result_zip)
            
        elif st.session_state.error_report is not None:
            # 报错态
            def open_fix_dialog():
                @st.dialog("🛠️ 在线修复", width="large")
                def dialog_body():
                    fixable = st.session_state.error_report[st.session_state.error_report['类型']=='数据错误']
                    
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
                dialog_body()

            UIComponents.render_error_report(st.session_state.error_report, open_fix_dialog)
            
        elif st.session_state.block_auto_run:
            # 阻断等待态
            st.info("ℹ️ 源文件已更新，等待确认...")
            if st.button("▶️ 重新校验并计算", type="primary", use_container_width=True): trigger_run = True
        else:
            # 自动运行态
            trigger_run = True
            
    # 3.5 执行运算
    if trigger_run:
        progress = st.progress(0, "启动引擎...")
        errs, df_a, df_b = DataEngine.validate_and_clean(
            st.session_state.data_store['A']['df'].copy(),
            st.session_state.data_store['B']['df'].copy(),
            st.session_state.mapping_config,
            hours_limit
        )
        time.sleep(0.2)
        
        if errs:
            progress.empty()
            err_df = pd.DataFrame(errs)
            # 补全列名防止Bug
            for c in ['类型','来源','行号','信息','_sys_id']: 
                if c not in err_df.columns: err_df[c] = '-'
            st.session_state.error_report = err_df
            st.session_state.block_auto_run = True
            st.rerun()
        else:
            progress.progress(50, "计算中...")
            results = DataEngine.calculate_results(df_a, df_b, st.session_state.mapping_config, price, sub_tag)
            
            # 打包
            st.session_state.result_files = {k: DataEngine.to_bytes(v) for k, v in results.items()}
            buf = io.BytesIO()
            with zipfile.ZipFile(buf, 'w') as z:
                z.writestr("表1.xlsx", st.session_state.result_files['t1'])
                z.writestr("表2.xlsx", st.session_state.result_files['t2'])
                z.writestr("表3.xlsx", st.session_state.result_files['t3'])
            st.session_state.result_zip = buf.getvalue()
            
            st.session_state.is_calculated = True
            progress.progress(100)
            st.rerun()

elif st.session_state.page == 'mapping':
    # --- 映射配置页面 ---
    st.markdown("### 🐱 字段映射 & 逻辑配置")
    
    col1, col2 = st.columns([1, 4])
    with col1:
        if st.button("⬅️ 返回主页", use_container_width=True): 
            st.session_state.page = 'main'
            st.rerun()
            
    with col2:
        c_status, c_edit = st.columns([3, 1])
        with c_edit:
            has_files = st.session_state.data_store['A']['df'] is not None and st.session_state.data_store['B']['df'] is not None
            if not st.session_state.is_editing_mapping:
                if st.button("✏️ 编辑配置", type="primary", use_container_width=True):
                    if not has_files: st.toast("⚠️ 请先在主页上传文件以获取列名", icon="🚫")
                    else:
                        st.session_state.is_editing_mapping = True
                        st.rerun()
            else:
                if st.button("💾 保存生效", type="primary", use_container_width=True):
                    st.session_state.is_editing_mapping = False
                    st.session_state.is_calculated = False
                    st.session_state.block_auto_run = False
                    st.session_state.error_report = None
                    st.success("配置已更新")
                    time.sleep(0.5)
                    st.rerun()
                    
    st.divider()
    
    # 准备下拉选项
    cols_a = list(st.session_state.data_store['A']['df'].columns) if st.session_state.data_store['A']['df'] is not None else []
    cols_b = list(st.session_state.data_store['B']['df'].columns) if st.session_state.data_store['B']['df'] is not None else []
    
    tab1, tab2, tab3 = st.tabs(["表1: 工时统计", "表2: 结算汇总", "表3: 详细明细"])
    
    def render_tab(t_name):
        df_c = st.session_state.mapping_config
        subset = df_c[df_c['所属表'] == t_name]
        
        # 表头
        c1, c2, c3, c4 = st.columns([2, 3, 1.5, 3])
        c1.markdown("<div class='map-header'>目标字段</div>", unsafe_allow_html=True)
        c2.markdown("<div class='map-header'>匹配列 (可编辑)</div>", unsafe_allow_html=True)
        c3.markdown("<div class='map-header'>来源表</div>", unsafe_allow_html=True)
        c4.markdown("<div class='map-header'>逻辑说明</div>", unsafe_allow_html=True)
        
        for idx, row in subset.iterrows():
            new_val = UIComponents.render_mapping_table_row(
                row, 
                st.session_state.is_editing_mapping, 
                cols_a, cols_b, idx
            )
            if new_val:
                st.session_state.mapping_config.at[idx, '匹配字段'] = new_val

    with tab1: render_tab("表1")
    with tab2: render_tab("表2")
    with tab3: render_tab("表3")
