
import streamlit as st
import pandas as pd
import io
import time

# ==========================================
# 1. 系统配置与强制深色极简皮肤
# ==========================================
st.set_page_config(
    page_title="淡藤财务报表",
    page_icon="😈", 
    layout="wide",
    initial_sidebar_state="collapsed"
)

# --- 强制深色极简 CSS ---
st.markdown("""
<style>
    /* 强制定义深色变量 */
    :root {
        --bg-color: #000000;
        --text-color: #ffffff;
        --card-bg: #121212;
        --card-border: #333333;
        --btn-bg: #ffffff;
        --btn-text: #000000;
        --accent-color: #4ade80; /* 荧光绿 */
        --error-color: #ef4444;
    }

    /* 全局背景与文字 */
    .stApp {
        background-color: var(--bg-color);
        color: var(--text-color);
    }
    
    h1, h2, h3, p, div, span, label, li {
        color: var(--text-color) !important;
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
    }

    /* 极简卡片容器 */
    div[data-testid="stVerticalBlock"] > div[style*="background-color"] {
        background-color: var(--card-bg);
    }

    /* 按钮：黑底白字反转风格 */
    .stButton > button {
        background-color: var(--btn-bg);
        color: var(--btn-text) !important;
        border: 1px solid var(--btn-bg);
        border-radius: 4px;
        padding: 0.5rem 2rem;
        font-weight: 600;
        transition: all 0.2s;
    }
    .stButton > button:hover {
        opacity: 0.85;
        box-shadow: 0 0 8px rgba(255, 255, 255, 0.3);
    }
    
    /* 进度条修复 */
    .stProgress > div > div > div > div {
        background-color: var(--accent-color) !important;
    }

    /* 上传组件边框 */
    div[data-testid="stFileUploader"] {
        border: 1px dashed #555;
        border-radius: 6px;
        background-color: #1a1a1a;
    }

    /* 侧边栏 */
    section[data-testid="stSidebar"] {
        background-color: #0a0a0a;
        border-right: 1px solid #333;
    }
    
    /* 错误提示框美化 */
    .stAlert {
        background-color: #1a1a1a;
        border: 1px solid var(--error-color);
        color: var(--error-color);
    }
</style>
""", unsafe_allow_html=True)

# ==========================================
# 2. 头部 (Header)
# ==========================================
st.markdown("# 😈 淡藤财务报表")
st.caption("Minimalist Financial Settlement System | Dark Mode Only")
st.markdown("---")

# ==========================================
# 3. 参数配置 (侧边栏)
# ==========================================
with st.sidebar:
    st.header("⚙️ 参数设置")
    PRICE_PER_DAY = st.number_input("人力单价 (元/天)", value=1500, step=100)
    SUBSIDY_TAG = st.text_input("补助关键词", value="差旅补助")

# ==========================================
# 4. 系统介绍 & 文件上传
# ==========================================
st.markdown("### 1. 核心流程")
st.info("步骤：上传文件 -> 自动映射 -> 强制校验 -> 结果留存下载")

st.markdown("### 2. 数据源上传")

with st.container(border=True):
    col_a, col_b = st.columns(2)
    with col_a:
        st.markdown("**Source A: 交付明细**")
        file_a = st.file_uploader("拖拽或点击上传", type=['xlsx', 'csv'], key='a', label_visibility="collapsed")
    with col_b:
        st.markdown("**Source B: 差旅明细**")
        file_b = st.file_uploader("拖拽或点击上传", type=['xlsx', 'csv'], key='b', label_visibility="collapsed")

# 初始化 Session State
if 'results' not in st.session_state:
    st.session_state.results = None

@st.cache_data(ttl=600)
def load_data(file):
    if not file: return None
    try:
        if file.name.endswith('.csv'):
            try: return pd.read_csv(file)
            except: return pd.read_csv(file, encoding='gbk')
        else: return pd.read_excel(file)
    except: return None

df_a = load_data(file_a)
df_b = load_data(file_b)

# ==========================================
# 5. 报表校验与处理
# ==========================================
if df_a is not None and df_b is not None:
    
    # 隐式映射区域
    df_a.columns = [str(c).strip() for c in df_a.columns]
    df_b.columns = [str(c).strip() for c in df_b.columns]
    
    with st.expander("🛠️ 字段映射设置 (默认自动匹配)", expanded=False):
        mc1, mc2 = st.columns(2)
        cols_a = list(df_a.columns)
        cols_b = list(df_b.columns)
        
        def smart_idx(opts, keys):
            for k in keys: 
                if k in opts: return opts.index(k)
            return 0

        with mc1:
            st.caption("表A 映射")
            map_a = {}
            cfg_a = {
                'user': ['人员', '姓名'], 'spm': ['SPM', '标识符'], 'hours': ['交付工时', '工时'],
                'project': ['项目', '所属项目'], 'range': ['人事范围'], 'contract': ['合同主体'],
                'sales': ['销售', '销售人员'], 'dept': ['销售部门']
            }
            for k, v in cfg_a.items():
                map_a[k] = st.selectbox(f"{k}", cols_a, index=smart_idx(cols_a, v), key=f"a_{k}")
        
        with mc2:
            st.caption("表B 映射")
            map_b = {}
            cfg_b = {
                'user': ['出差人', '姓名'], 'spm': ['SPM', '项目编号'],
                'amount': ['金额', '总金额'], 'type': ['产品类型', '费用类型']
            }
            for k, v in cfg_b.items():
                map_b[k] = st.selectbox(f"{k}", cols_b, index=smart_idx(cols_b, v), key=f"b_{k}")

    st.markdown("### 3. 执行生成")

    # 执行按钮
    if st.button("开始校验并生成报表", use_container_width=True):
        
        progress_bar = st.progress(0)
        status_text = st.empty()
        error_msgs = [] # 用于收集所有错误信息

        try:
            # --- 阶段 1: 强制数据完整性校验 ---
            status_text.text("正在进行三重数据完整性校验...")
            time.sleep(0.2)
            progress_bar.progress(10)

            # 1. 校验 表B SPM (原有)
            missing_spm_b = df_b[map_b['spm']].isnull().sum()
            if missing_spm_b > 0:
                error_msgs.append(f"❌ 表B 错误：发现 {missing_spm_b} 条缺少 [SPM编号] 的数据")

            # 2. 校验 表A 合同主体 (新增)
            missing_contract_a = df_a[map_a['contract']].isnull().sum()
            if missing_contract_a > 0:
                error_msgs.append(f"❌ 表A 错误：发现 {missing_contract_a} 条缺少 [合同主体] 的数据")

            # 3. 校验 表A 交付工时 (新增)
            missing_hours_a = df_a[map_a['hours']].isnull().sum()
            if missing_hours_a > 0:
                error_msgs.append(f"❌ 表A 错误：发现 {missing_hours_a} 条缺少 [交付工时] 的数据")

            # 如果有任何错误，统一显示并阻断
            if error_msgs:
                progress_bar.empty()
                status_text.empty()
                for msg in error_msgs:
                    st.error(msg)
                st.stop() # 强制停止后续逻辑
            
            # --- 阶段 2: 清洗 ---
            status_text.text("校验全部通过！正在清洗数据...")
            progress_bar.progress(40)
            
            # A表聚合
            agg = {map_a['hours']: 'sum'}
            for k in ['project', 'range', 'contract', 'sales', 'dept']: agg[map_a[k]] = 'first'
            df_a_cl = df_a.dropna(subset=[map_a['spm']]).copy()
            df_a_gp = df_a_cl.groupby([map_a['user'], map_a['spm']], as_index=False).agg(agg)

            # B表拆分
            status_text.text("正在拆分差旅费用...")
            progress_bar.progress(60)
            df_b_cl = df_b.dropna(subset=[map_b['spm']]).copy()
            is_sub = df_b_cl[map_b['type']] == SUBSIDY_TAG
            df_sub = df_b_cl[is_sub].groupby([map_b['user'], map_b['spm']])[map_b['amount']].sum().reset_index(name='差旅补助')
            df_fee = df_b_cl[~is_sub].groupby([map_b['user'], map_b['spm']])[map_b['amount']].sum().reset_index(name='差旅费控平台')

            # 合并
            status_text.text("正在合并生成报表...")
            progress_bar.progress(80)
            key_a = [map_a['user'], map_a['spm']]
            key_b = [map_b['user'], map_b['spm']]
            res = pd.merge(df_a_gp, df_sub, left_on=key_a, right_on=key_b, how='left')
            res = pd.merge(res, df_fee, left_on=key_a, right_on=key_b, how='left')
            res.fillna(0, inplace=True)
            
            res['支持时间'] = res[map_a['hours']] / 8
            res['人力费用'] = res['支持时间'] * PRICE_PER_DAY
            res['结算费用合计'] = res['人力费用'] + res['差旅补助'] + res['差旅费控平台']

            # --- 阶段 3: 格式化 ---
            # 表3
            rename = {
                map_a['user']: '人员', map_a['project']: '所属项目', map_a['range']: '人事范围',
                map_a['spm']: 'SPM', map_a['contract']: '合同主体', map_a['sales']: '销售人员',
                map_a['dept']: '销售部门', map_a['hours']: '耗时（小时）'
            }
            t3 = res.rename(columns=rename)
            cols = ['人员', '所属项目', '人事范围', 'SPM', '合同主体', '销售人员', '销售部门',
                    '差旅补助', '差旅费控平台', '耗时（小时）', '支持时间', '人力费用', '结算费用合计']
            t3 = t3[[c for c in cols if c in t3.columns]]
            t3.rename(columns={'支持时间': '支持时间（人天）'}, inplace=True)
            t3.insert(0, '序号', range(1, len(t3)+1))

            # 表2
            grp = ['人事范围', '合同主体', '销售部门']
            if all(c in t3.columns for c in grp):
                t2 = t3.groupby(grp).agg({'结算费用合计': 'sum', '支持时间（人天）': 'sum'}).reset_index()
                t2.columns = ['销售公司', '采购公司', '采购部门', '金额（含税，单位：元）', '工作量（人天）']
                t2.insert(0, '序号', range(1, len(t2)+1))
            else: t2 = pd.DataFrame()

            # 表1
            t1 = t3.groupby('人员')['耗时（小时）'].sum().reset_index()
            t1.rename(columns={'耗时（小时）': '项目工时'}, inplace=True)
            t1.insert(0, '序号', range(1, len(t1)+1))

            # 存入 Session State
            st.session_state.results = {'t1': t1, 't2': t2, 't3': t3}

            progress_bar.progress(100)
            status_text.text("✅ 生成完成！")
            time.sleep(0.5)
            progress_bar.empty()
            status_text.empty()

        except Exception as e:
            st.error(f"处理错误: {e}")

# ==========================================
# 6. 文件下载区域 (持久化显示)
# ==========================================
if st.session_state.results is not None:
    st.divider()
    st.markdown("### 4. 报表下载")
    st.success("报表已就绪。点击下方按钮下载（无需重复生成）。")

    with st.container(border=True):
        d1, d2, d3 = st.columns(3)
        results = st.session_state.results
        
        def to_excel(df):
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False)
            return output.getvalue()

        with d1:
            st.download_button("📥 结果表1 (工时)", to_excel(results['t1']), "结果表1.xlsx", use_container_width=True)
        with d2:
            st.download_button("📥 结果表2 (结算)", to_excel(results['t2']), "结果表2.xlsx", use_container_width=True)
        with d3:
            st.download_button("📥 结果表3 (明细)", to_excel(results['t3']), "结果表3.xlsx", use_container_width=True)
