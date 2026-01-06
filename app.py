import streamlit as st
import pandas as pd
import io
import time
import zipfile
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode

# ==========================================
# 0. 基础配置
# ==========================================
st.set_page_config(page_title="淡藤财务报表 Pro", page_icon="😈", layout="wide", initial_sidebar_state="collapsed")

# 强制深色极简 CSS (优化了卡片间距)
st.markdown("""
<style>
    :root { --bg-color: #000000; --accent-color: #4ade80; --error-color: #ef4444; }
    .stApp { background-color: var(--bg-color); color: #ffffff; }
    
    /* 1. 头部样式 */
    h1 { font-size: 2.2rem !important; font-weight: 800 !important; color: #fff; }
    .stCaption { font-size: 0.9rem !important; opacity: 0.7; }
    
    /* 2. 按钮样式 */
    .stButton > button { 
        border: 1px solid #ffffff; color: #000000; background: #ffffff; font-weight: bold; 
        transition: all 0.3s ease;
    }
    .stButton > button:hover { 
        box-shadow: 0 0 15px rgba(74, 222, 128, 0.8); transform: scale(1.02); background: var(--accent-color); border-color: var(--accent-color);
    }
    
    /* 3. 进度条自定义 */
    .stProgress > div > div > div > div { background-color: var(--accent-color) !important; }
    
    /* 4. 分区卡片优化 */
    div[data-testid="stVerticalBlock"] > div[data-testid="stVerticalBlock"] {
        gap: 1.5rem; /* 增加组件垂直间距 */
    }
</style>
""", unsafe_allow_html=True)

# ==========================================
# 1. 头部与侧边栏
# ==========================================
st.title("😈 淡藤财务报表 Pro")
st.caption("Minimalist Financial Settlement System")

with st.sidebar:
    st.header("⚙️ 参数设置")
    PRICE_PER_DAY = st.number_input("人力单价 (元/天)", 1500)
    SUBSIDY_TAG = st.text_input("补助关键词", "差旅补助")

# ==========================================
# 2. 数据输入区 (卡片化)
# ==========================================
# 使用 container(border=True) 制造视觉分区
with st.container(border=True):
    st.markdown("### 📂 数据源上传")
    
    # 状态初始化
    if 'df_a_state' not in st.session_state: st.session_state.df_a_state = None
    if 'df_b_state' not in st.session_state: st.session_state.df_b_state = None
    if 'file_a_id' not in st.session_state: st.session_state.file_a_id = None
    if 'file_b_id' not in st.session_state: st.session_state.file_b_id = None

    col_up1, col_up2 = st.columns(2)
    with col_up1:
        f_a = st.file_uploader("Source A: 交付明细", type=['xlsx', 'csv'], key='f_a')
    with col_up2:
        f_b = st.file_uploader("Source B: 差旅明细", type=['xlsx', 'csv'], key='f_b')

    # 数据读取
    def init_df(file, key_id, state_key):
        if file and file.file_id != st.session_state[key_id]:
            try:
                if file.name.endswith('.csv'): df = pd.read_csv(file)
                else: df = pd.read_excel(file)
                df.columns = [str(c).strip() for c in df.columns]
                st.session_state[state_key] = df
                st.session_state[key_id] = file.file_id
            except Exception as e:
                st.error(f"读取失败: {e}")

    init_df(f_a, 'file_a_id', 'df_a_state')
    init_df(f_b, 'file_b_id', 'df_b_state')

# ==========================================
# 3. 逻辑处理核心区
# ==========================================
if st.session_state.df_a_state is not None and st.session_state.df_b_state is not None:
    
    # --- 统一进度条 ---
    # 放在输入区下方，作为连接输入和输出的桥梁
    st.divider()
    progress_bar = st.progress(0, text="等待开始...")
    
    # 辅助函数：列名查找
    def find_col(df, candidates):
        for c in candidates:
            if c in df.columns: return c
        return None
        
    def auto_map_cols(df, type_key):
        mapping = {}
        cols = df.columns
        if type_key == 'A':
            mapping['spm'] = find_col(df, ['SPM', '标识符', '项目编号'])
            mapping['hours'] = find_col(df, ['交付工时', '工时', '投入工时'])
            mapping['contract'] = find_col(df, ['合同主体'])
            mapping['user'] = find_col(df, ['人员', '姓名', '员工姓名'])
            mapping['project'] = find_col(df, ['项目', '所属项目']) or cols[0]
            mapping['range'] = find_col(df, ['人事范围']) or cols[0]
            mapping['sales'] = find_col(df, ['销售', '销售人员']) or cols[0]
            mapping['dept'] = find_col(df, ['销售部门']) or cols[0]
        else:
            mapping['spm'] = find_col(df, ['SPM', '项目编号', '费用归属项目'])
            mapping['amount'] = find_col(df, ['金额', '总金额', '报销金额'])
            mapping['user'] = find_col(df, ['出差人', '姓名', '报销人'])
            mapping['type'] = find_col(df, ['产品类型', '费用类型'])
        return mapping

    def validate_and_trace(df, source_name, required_cols_map):
        if df is None: return False, None, None
        errors = []
        final_map = {}
        for key, col_name in required_cols_map.items():
            if col_name is None:
                return False, pd.DataFrame([{'错误类型': '列缺失', '详情': f"无法在{source_name}中找到对应 {key} 的列"}]), None
            final_map[key] = col_name

        check_keys = ['spm', 'hours', 'contract', 'amount']
        for k in check_keys:
            if k in final_map:
                col = final_map[k]
                null_rows = df[df[col].isnull() | (df[col].astype(str).str.strip() == '')]
                if not null_rows.empty:
                    for idx, row in null_rows.iterrows():
                        errors.append({
                            '来源表': source_name, 'Excel行号': idx + 2,
                            '错误列': col, '错误原因': '数值为空',
                            '快照': " | ".join(row.astype(str).values[:4])
                        })
        if errors: return False, pd.DataFrame(errors), final_map
        return True, None, final_map

    # --- 阶段 1: 校验 (进度 0% -> 30%) ---
    progress_bar.progress(30, text="🔄 正在校验数据完整性...")
    
    map_a = auto_map_cols(st.session_state.df_a_state, 'A')
    map_b = auto_map_cols(st.session_state.df_b_state, 'B')
    
    valid_a, err_a, final_map_a = validate_and_trace(st.session_state.df_a_state, "表A", map_a)
    valid_b, err_b, final_map_b = validate_and_trace(st.session_state.df_b_state, "表B", map_b)
    all_errors = pd.concat([err_a, err_b], ignore_index=True) if (err_a is not None or err_b is not None) else None

    # [校验失败分支]
    if all_errors is not None and not all_errors.empty:
        progress_bar.progress(0, text="🚨 校验失败，请修复下方数据")
        
        st.error(f"发现 {len(all_errors)} 处数据异常，请修正后点击保存：")
        
        # 错误修正区
        col_e1, col_e2 = st.columns(2)
        new_df_a, new_df_b = None, None
        
        def show_editor(df, key):
            gb = GridOptionsBuilder.from_dataframe(df)
            gb.configure_default_column(editable=True, resizable=True)
            gb.configure_selection('single')
            return AgGrid(df, gridOptions=gb.build(), update_mode=GridUpdateMode.MANUAL, height=250, theme='balham', key=key)

        with col_e1:
            if not valid_a:
                st.caption("🔴 表A 修正区")
                res_a = show_editor(st.session_state.df_a_state, 'ga')
                new_df_a = res_a['data']
        with col_e2:
            if not valid_b:
                st.caption("🔴 表B 修正区")
                res_b = show_editor(st.session_state.df_b_state, 'gb')
                new_df_b = res_b['data']

        if st.button("✅ 保存修正并重新计算", type="primary", use_container_width=True):
            if new_df_a is not None: st.session_state.df_a_state = pd.DataFrame(new_df_a)
            if new_df_b is not None: st.session_state.df_b_state = pd.DataFrame(new_df_b)
            st.rerun()
            
        st.stop() # 阻断后续运行

    # --- 阶段 2: 计算 (进度 30% -> 80%) ---
    progress_bar.progress(60, text="⚡ 校验通过 | 正在进行核心计算...")
    time.sleep(0.5) # 稍微停顿让用户看清状态变化（可选）
    
    try:
        # 数据拷贝与清洗
        df_a = st.session_state.df_a_state.copy()
        df_b = st.session_state.df_b_state.copy()
        
        def clean_num(df, col): return pd.to_numeric(df[col].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
        df_a[final_map_a['hours']] = clean_num(df_a, final_map_a['hours'])
        df_b[final_map_b['amount']] = clean_num(df_b, final_map_b['amount'])
        
        # 聚合逻辑
        agg = {final_map_a['hours']: 'sum'}
        for k in ['project', 'range', 'contract', 'sales', 'dept']: 
            if final_map_a[k]: agg[final_map_a[k]] = 'first'
        df_a_gp = df_a.groupby([final_map_a['user'], final_map_a['spm']], as_index=False).agg(agg)
        
        # 费用计算
        is_sub = df_b[final_map_b['type']].astype(str).str.contains(SUBSIDY_TAG, na=False)
        grp_b = [final_map_b['user'], final_map_b['spm']]
        df_sub = df_b[is_sub].groupby(grp_b)[final_map_b['amount']].sum().reset_index(name='差旅补助')
        df_fee = df_b[~is_sub].groupby(grp_b)[final_map_b['amount']].sum().reset_index(name='差旅费控平台')
        
        # Merge前统一类型
        key_cols = [final_map_a['user'], final_map_a['spm']]
        df_a_gp[final_map_a['spm']] = df_a_gp[final_map_a['spm']].astype(str)
        df_sub[final_map_b['spm']] = df_sub[final_map_b['spm']].astype(str)
        df_fee[final_map_b['spm']] = df_fee[final_map_b['spm']].astype(str)
        
        res = pd.merge(df_a_gp, df_sub, left_on=key_cols, right_on=[final_map_b['user'], final_map_b['spm']], how='left')
        res = pd.merge(res, df_fee, left_on=key_cols, right_on=[final_map_b['user'], final_map_b['spm']], how='left')
        res = res.fillna(0)
        
        # 最终计算
        res['支持时间'] = res[final_map_a['hours']] / 8
        res['人力费用'] = res['支持时间'] * PRICE_PER_DAY
        res['结算费用合计'] = res['人力费用'] + res['差旅补助'] + res['差旅费控平台']
        
        # 格式化输出 (表3)
        rename_map = {
            key_cols[0]: '人员', final_map_a['project']: '所属项目', final_map_a['range']: '人事范围',
            key_cols[1]: 'SPM', final_map_a['contract']: '合同主体', final_map_a['sales']: '销售人员',
            final_map_a['dept']: '销售部门', final_map_a['hours']: '耗时（小时）'
        }
        t3 = res.rename(columns=rename_map)
        cols_order = ['人员', '所属项目', '人事范围', 'SPM', '合同主体', '销售人员', '销售部门',
                      '差旅补助', '差旅费控平台', '耗时（小时）', '支持时间', '人力费用', '结算费用合计']
        t3 = t3[[c for c in cols_order if c in t3.columns]]
        t3.rename(columns={'支持时间': '支持时间（人天）'}, inplace=True)
        t3.insert(0, '序号', range(1, len(t3)+1))
        
        # (表2) 结算表
        t2_cols = ['人事范围', '合同主体', '销售部门']
        if all(c in t3.columns for c in t2_cols):
            t2 = t3.groupby(t2_cols).agg({'结算费用合计': 'sum', '支持时间（人天）': 'sum'}).reset_index()
            t2.columns = ['销售公司', '采购公司', '采购部门', '金额（含税，单位：元）', '工作量（人天）']
            t2.insert(0, '序号', range(1, len(t2)+1))
        else: t2 = pd.DataFrame({'提示': ['缺少维度字段']})
        
        # (表1) 工时表
        t1 = t3.groupby('人员')['耗时（小时）'].sum().reset_index()
        t1.rename(columns={'耗时（小时）': '项目工时'}, inplace=True)
        t1.insert(0, '序号', range(1, len(t1)+1))
        
        progress_bar.progress(90, text="📦 正在打包下载文件...")
        
        # --- 阶段 3: 输出结果 (进度 100%) ---
        
        # 生成二进制流
        def to_bytes(df):
            out = io.BytesIO()
            with pd.ExcelWriter(out, engine='xlsxwriter') as w: df.to_excel(w, index=False)
            return out.getvalue()
        
        b1, b2, b3 = to_bytes(t1), to_bytes(t2), to_bytes(t3)
        z_out = io.BytesIO()
        with zipfile.ZipFile(z_out, 'w', zipfile.ZIP_DEFLATED) as z:
            z.writestr("表1_工时统计.xlsx", b1)
            z.writestr("表2_结算汇总.xlsx", b2)
            z.writestr("表3_详细明细.xlsx", b3)

        progress_bar.progress(100, text="✅ 报表生成完毕！")
        
        # ==========================================
        # 4. 结果展示区 (卡片化)
        # ==========================================
        # 只有在 100% 后才渲染此区域，解决"未生成先显示"的问题
        with st.container(border=True):
            st.markdown("### 📥 报表下载")
            
            st.download_button(
                label="📦 批量下载所有报表 (ZIP)",
                data=z_out.getvalue(),
                file_name="淡藤财务报表汇总.zip",
                mime="application/zip",
                type="primary",
                use_container_width=True
            )
            
            st.markdown("---") # 分割线
            
            cols_d = st.columns(3)
            # 修复了文件名重复的问题
            cols_d[0].download_button("📥 表1 (工时)", b1, "表1_工时统计.xlsx", use_container_width=True)
            cols_d[1].download_button("📥 表2 (结算)", b2, "表2_结算汇总.xlsx", use_container_width=True)
            cols_d[2].download_button("📥 表3 (明细)", b3, "表3_详细明细.xlsx", use_container_width=True)

    except Exception as e:
        st.error(f"计算过程发生未知错误: {e}")
