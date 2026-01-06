import streamlit as st
import pandas as pd
import io
import time
import zipfile
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode, DataReturnMode

# ==========================================
# 1. 系统配置
# ==========================================
st.set_page_config(page_title="淡藤财务报表 Pro", page_icon="😈", layout="wide", initial_sidebar_state="collapsed")

# --- 样式：保留深色极简风 ---
st.markdown("""
<style>
    :root { --bg-color: #000000; --accent-color: #4ade80; --error-color: #ef4444; }
    .stApp { background-color: var(--bg-color); color: #ffffff; }
    .stButton > button { border: 1px solid #ffffff; color: #000000; background: #ffffff; font-weight: bold; }
    .stButton > button:hover { box-shadow: 0 0 8px rgba(74, 222, 128, 0.6); }
    .error-box { border: 1px solid var(--error-color); background: #1a0505; padding: 1rem; border-radius: 6px; margin-bottom: 1rem; }
    /* 进度条颜色 */
    .stProgress > div > div > div > div { background-color: var(--accent-color) !important; }
</style>
""", unsafe_allow_html=True)

st.title("😈 淡藤财务报表 Pro")
st.caption("可视化错误追踪 (Error Traceability) & 在线修正 (In-App Correction)")
st.divider()

# ==========================================
# 2. 状态管理 (Session State)
# ==========================================
if 'df_a_state' not in st.session_state: st.session_state.df_a_state = None
if 'df_b_state' not in st.session_state: st.session_state.df_b_state = None
if 'file_a_id' not in st.session_state: st.session_state.file_a_id = None
if 'file_b_id' not in st.session_state: st.session_state.file_b_id = None

# ==========================================
# 3. 数据上传与初始化
# ==========================================
with st.sidebar:
    st.header("⚙️ 参数")
    PRICE_PER_DAY = st.number_input("人力单价", 1500)
    SUBSIDY_TAG = st.text_input("补助关键词", "差旅补助")

col_up1, col_up2 = st.columns(2)
with col_up1:
    f_a = st.file_uploader("Source A: 交付明细", type=['xlsx', 'csv'], key='f_a')
with col_up2:
    f_b = st.file_uploader("Source B: 差旅明细", type=['xlsx', 'csv'], key='f_b')

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
# 4. 字段映射与校验逻辑
# ==========================================
MAP_CFG = {
    'A': {'spm': 'SPM', 'hours': '交付工时', 'contract': '合同主体', 'user': '姓名', 'project': '项目', 'range': '人事范围', 'sales': '销售', 'dept': '销售部门'},
    'B': {'spm': 'SPM', 'amount': '金额', 'user': '出差人', 'type': '费用类型'}
}

# 智能查找列名
def find_col(df, candidates):
    for c in candidates:
        if c in df.columns: return c
    return None

# 尝试自动修正列名映射 (简单版)
def auto_map_cols(df, type_key):
    mapping = {}
    cols = df.columns
    if type_key == 'A':
        mapping['spm'] = find_col(df, ['SPM', '标识符', '项目编号'])
        mapping['hours'] = find_col(df, ['交付工时', '工时', '投入工时'])
        mapping['contract'] = find_col(df, ['合同主体'])
        mapping['user'] = find_col(df, ['人员', '姓名', '员工姓名'])
        # 非必填字段
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
    
    # 检查列是否存在
    final_map = {}
    for key, col_name in required_cols_map.items():
        if col_name is None:
            return False, pd.DataFrame([{'错误类型': '列缺失', '详情': f"无法在{source_name}中找到对应 {key} 的列"}]), None
        final_map[key] = col_name

    # 检查空值 (只检查关键字段)
    check_keys = ['spm', 'hours', 'contract', 'amount'] # 根据表不同来定
    for k in check_keys:
        if k in final_map:
            col = final_map[k]
            null_rows = df[df[col].isnull() | (df[col].astype(str).str.strip() == '')]
            if not null_rows.empty:
                for idx, row in null_rows.iterrows():
                    snapshot = " | ".join(row.astype(str).values[:5])
                    errors.append({
                        '来源表': source_name,
                        'Excel行号': idx + 2,
                        '错误列': col,
                        '错误原因': '数值为空',
                        '行数据快照': snapshot
                    })
    
    if errors: return False, pd.DataFrame(errors), final_map
    return True, None, final_map

# ==========================================
# 5. 主流程
# ==========================================
if st.session_state.df_a_state is not None and st.session_state.df_b_state is not None:
    
    # 自动映射
    map_a = auto_map_cols(st.session_state.df_a_state, 'A')
    map_b = auto_map_cols(st.session_state.df_b_state, 'B')

    # 校验
    valid_a, err_a, final_map_a = validate_and_trace(st.session_state.df_a_state, "表A", map_a)
    valid_b, err_b, final_map_b = validate_and_trace(st.session_state.df_b_state, "表B", map_b)

    all_errors = pd.concat([err_a, err_b], ignore_index=True) if (err_a is not None or err_b is not None) else None

    # --- 分支 A: 有错误，显示修复界面 ---
    if all_errors is not None and not all_errors.empty:
        st.markdown(f"""<div class="error-box"><h3>🚨 发现 {len(all_errors)} 处数据异常</h3><p>请在下方直接修改红色高亮区域，然后点击保存。</p></div>""", unsafe_allow_html=True)
        
        # 下载错误清单
        csv_buffer = io.BytesIO()
        all_errors.to_excel(csv_buffer, index=False)
        st.download_button("📥 下载错误清单", csv_buffer, "错误清单.xlsx")

        def show_editor(df, key, title):
            gb = GridOptionsBuilder.from_dataframe(df)
            gb.configure_default_column(editable=True, resizable=True)
            gb.configure_selection('single')
            return AgGrid(df, gridOptions=gb.build(), update_mode=GridUpdateMode.MANUAL, height=300, theme='balham', key=key)

        new_df_a, new_df_b = None, None
        
        c1, c2 = st.columns(2)
        with c1:
            if not valid_a:
                st.markdown("**表A 编辑区**")
                res_a = show_editor(st.session_state.df_a_state, 'ga', 'A')
                new_df_a = res_a['data']
        with c2:
            if not valid_b:
                st.markdown("**表B 编辑区**")
                res_b = show_editor(st.session_state.df_b_state, 'gb', 'B')
                new_df_b = res_b['data']

        if st.button("✅ 保存修正并重新校验", type="primary", use_container_width=True):
            if new_df_a is not None: st.session_state.df_a_state = pd.DataFrame(new_df_a)
            if new_df_b is not None: st.session_state.df_b_state = pd.DataFrame(new_df_b)
            st.rerun()
        st.stop()

    # --- 分支 B: 校验通过，执行计算 (Happy Path) ---
    st.success("✅ 数据校验通过！正在生成报表...")
    progress_bar = st.progress(0)
    
    try:
        # 1. 准备数据
        df_a = st.session_state.df_a_state.copy()
        df_b = st.session_state.df_b_state.copy()
        
        # 2. 清洗数值 (关键步骤)
        def clean_num(df, col):
            return pd.to_numeric(df[col].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
        
        df_a[final_map_a['hours']] = clean_num(df_a, final_map_a['hours'])
        df_b[final_map_b['amount']] = clean_num(df_b, final_map_b['amount'])
        progress_bar.progress(30)

        # 3. 聚合逻辑
        # A表聚合
        agg = {final_map_a['hours']: 'sum'}
        for k in ['project', 'range', 'contract', 'sales', 'dept']: 
            if final_map_a[k]: agg[final_map_a[k]] = 'first'
        
        df_a_gp = df_a.groupby([final_map_a['user'], final_map_a['spm']], as_index=False).agg(agg)
        
        # B表拆分
        is_sub = df_b[final_map_b['type']].astype(str).str.contains(SUBSIDY_TAG, na=False)
        grp_b = [final_map_b['user'], final_map_b['spm']]
        df_sub = df_b[is_sub].groupby(grp_b)[final_map_b['amount']].sum().reset_index(name='差旅补助')
        df_fee = df_b[~is_sub].groupby(grp_b)[final_map_b['amount']].sum().reset_index(name='差旅费控平台')
        progress_bar.progress(60)

        # 4. 合并逻辑
        # 统一关联键类型
        key_a_user, key_a_spm = final_map_a['user'], final_map_a['spm']
        key_b_user, key_b_spm = final_map_b['user'], final_map_b['spm']
        
        df_a_gp[key_a_spm] = df_a_gp[key_a_spm].astype(str)
        df_sub[key_b_spm] = df_sub[key_b_spm].astype(str)
        df_fee[key_b_spm] = df_fee[key_b_spm].astype(str)

        res = pd.merge(df_a_gp, df_sub, left_on=[key_a_user, key_a_spm], right_on=[key_b_user, key_b_spm], how='left')
        res = pd.merge(res, df_fee, left_on=[key_a_user, key_a_spm], right_on=[key_b_user, key_b_spm], how='left')
        
        res = res.fillna(0)
        res['支持时间'] = res[final_map_a['hours']] / 8
        res['人力费用'] = res['支持时间'] * PRICE_PER_DAY
        res['结算费用合计'] = res['人力费用'] + res['差旅补助'] + res['差旅费控平台']
        progress_bar.progress(80)

        # 5. 生成结果表
        # 表3 明细
        rename_dict = {
            key_a_user: '人员', final_map_a['project']: '所属项目', final_map_a['range']: '人事范围',
            key_a_spm: 'SPM', final_map_a['contract']: '合同主体', final_map_a['sales']: '销售人员',
            final_map_a['dept']: '销售部门', final_map_a['hours']: '耗时（小时）'
        }
        t3 = res.rename(columns=rename_dict)
        wanted_cols = ['人员', '所属项目', '人事范围', 'SPM', '合同主体', '销售人员', '销售部门',
                       '差旅补助', '差旅费控平台', '耗时（小时）', '支持时间', '人力费用', '结算费用合计']
        t3 = t3[[c for c in wanted_cols if c in t3.columns]] # 只保留存在的列
        t3.rename(columns={'支持时间': '支持时间（人天）'}, inplace=True)
        t3.insert(0, '序号', range(1, len(t3)+1))

        # 表2 结算单
        t2_cols = ['人事范围', '合同主体', '销售部门']
        if all(c in t3.columns for c in t2_cols):
            t2 = t3.groupby(t2_cols).agg({'结算费用合计': 'sum', '支持时间（人天）': 'sum'}).reset_index()
            t2.columns = ['销售公司', '采购公司', '采购部门', '金额（含税，单位：元）', '工作量（人天）']
            t2.insert(0, '序号', range(1, len(t2)+1))
        else:
            t2 = pd.DataFrame({'提示': ['缺少必要维度，无法生成结算单']})

        # 表1 工时
        t1 = t3.groupby('人员')['耗时（小时）'].sum().reset_index()
        t1.rename(columns={'耗时（小时）': '项目工时'}, inplace=True)
        t1.insert(0, '序号', range(1, len(t1)+1))
        
        progress_bar.progress(100)
        time.sleep(0.5)
        progress_bar.empty()

        # 6. 下载区域
        st.markdown("### 🎉 报表已生成")
        
        def to_bytes(df):
            out = io.BytesIO()
            with pd.ExcelWriter(out, engine='xlsxwriter') as w: df.to_excel(w, index=False)
            return out.getvalue()

        b1, b2, b3 = to_bytes(t1), to_bytes(t2), to_bytes(t3)
        
        # ZIP
        z_out = io.BytesIO()
        with zipfile.ZipFile(z_out, 'w', zipfile.ZIP_DEFLATED) as z:
            z.writestr("表1_工时.xlsx", b1)
            z.writestr("表2_结算.xlsx", b2)
            z.writestr("表3_明细.xlsx", b3)
        
        st.download_button("📦 批量下载所有报表 (ZIP)", z_out.getvalue(), "财务报表汇总.zip", "application/zip", type="primary", use_container_width=True)
        
        d1, d2, d3 = st.columns(3)
        d1.download_button("📥 表1 (工时)", b1, "表1.xlsx", use_container_width=True)
        d2.download_button("📥 表2 (结算)", b2, "表2.xlsx", use_container_width=True)
        d3.download_button("📥 表3 (明细)", b3, "表3.xlsx", use_container_width=True)

    except Exception as e:
        st.error(f"计算过程出错: {e}")
