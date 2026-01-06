import streamlit as st
import pandas as pd
import io
import time

# ==========================================
# 1. 系统配置与黑白极简皮肤
# ==========================================
st.set_page_config(
    page_title="淡藤财务报表",
    page_icon="😈", 
    layout="wide",
    initial_sidebar_state="collapsed"
)

# 状态管理
if 'theme' not in st.session_state:
    st.session_state.theme = 'light'

def toggle_theme():
    st.session_state.theme = 'dark' if st.session_state.theme == 'light' else 'light'

# --- 极简黑白 CSS 变量 ---
theme_light = """
    --bg-color: #ffffff;
    --text-color: #000000;
    --card-bg: #ffffff;
    --card-border: #e5e5e5;
    --btn-bg: #000000;
    --btn-text: #ffffff;
    --success-bg: #f4f4f5;
"""

theme_dark = """
    --bg-color: #000000;
    --text-color: #ffffff;
    --card-bg: #121212;
    --card-border: #333333;
    --btn-bg: #ffffff;
    --btn-text: #000000;
    --success-bg: #1a1a1a;
"""

current_theme = theme_light if st.session_state.theme == 'light' else theme_dark

st.markdown(f"""
<style>
    :root {{ {current_theme} }}

    /* 全局背景与文字 */
    .stApp {{
        background-color: var(--bg-color);
        color: var(--text-color);
    }}
    
    h1, h2, h3, p, div, span, label {{
        color: var(--text-color) !important;
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
    }}

    /* 极简卡片容器 */
    .minimal-card {{
        border: 1px solid var(--card-border);
        background-color: var(--card-bg);
        border-radius: 8px;
        padding: 24px;
        margin-bottom: 24px;
    }}

    /* 按钮：黑白反转 */
    .stButton > button {{
        background-color: var(--btn-bg);
        color: var(--btn-text) !important;
        border: 1px solid var(--btn-bg);
        border-radius: 4px;
        padding: 0.5rem 2rem;
        font-weight: 500;
        transition: all 0.2s;
    }}
    .stButton > button:hover {{
        opacity: 0.8;
        border-color: var(--text-color);
    }}
    
    /* 上传组件边框 */
    div[data-testid="stFileUploader"] {{
        border: 1px dashed var(--card-border);
        border-radius: 6px;
    }}

    /* 顶部导航对齐 */
    .header-row {{
        display: flex;
        align-items: center;
        justify-content: space-between;
        padding-bottom: 20px;
        border-bottom: 1px solid var(--card-border);
        margin-bottom: 30px;
    }}
    
    /* 侧边栏 */
    section[data-testid="stSidebar"] {{
        background-color: var(--card-bg);
        border-right: 1px solid var(--card-border);
    }}
</style>
""", unsafe_allow_html=True)

# ==========================================
# 2. 头部 (Header) & 系统介绍
# ==========================================
# 使用 columns 布局头部
c_head, c_toggle = st.columns([9, 1])
with c_head:
    st.markdown("# 😈 淡藤财务报表")
    st.caption("Minimalist Financial Settlement System")

with c_toggle:
    # 切换按钮
    icon = "🌙" if st.session_state.theme == 'light' else "🌞"
    if st.button(icon, help="切换深色/浅色模式"):
        toggle_theme()
        st.rerun()

# 1. 系统介绍
st.markdown("### 1. 系统介绍")
st.markdown("""
本系统用于自动化合并 **交付明细** 与 **差旅明细**。请按照下方步骤上传文件，
系统将自动执行字段映射、数据校验，并生成标准的三张结算报表。
""")
st.divider()

# ==========================================
# 3. 参数配置 (侧边栏)
# ==========================================
with st.sidebar:
    st.header("⚙️ 参数设置")
    PRICE_PER_DAY = st.number_input("人力单价 (元/天)", value=1500, step=100)
    SUBSIDY_TAG = st.text_input("补助关键词", value="差旅补助")

# ==========================================
# 4. 文件上传区域
# ==========================================
st.markdown("### 2. 文件上传")

# 使用原生容器保持整洁
with st.container(border=True):
    col_a, col_b = st.columns(2)
    with col_a:
        st.markdown("**Source A: 交付明细**")
        file_a = st.file_uploader("拖拽或点击上传", type=['xlsx', 'csv'], key='a', label_visibility="collapsed")
    with col_b:
        st.markdown("**Source B: 差旅明细**")
        file_b = st.file_uploader("拖拽或点击上传", type=['xlsx', 'csv'], key='b', label_visibility="collapsed")

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
# 5. 报表校验与处理 (核心逻辑)
# ==========================================
if df_a is not None and df_b is not None:
    
    # --- 隐式映射区域 (保持极简，但功能必须有) ---
    # 清洗列名
    df_a.columns = [str(c).strip() for c in df_a.columns]
    df_b.columns = [str(c).strip() for c in df_b.columns]
    
    # 简单的两列布局显示映射，不抢眼
    with st.expander("🛠️ 字段映射设置 (默认已自动匹配，点击展开修改)", expanded=False):
        mc1, mc2 = st.columns(2)
        cols_a = list(df_a.columns)
        cols_b = list(df_b.columns)
        
        def smart_idx(opts, keys):
            for k in keys: 
                if k in opts: return opts.index(k)
            return 0

        with mc1:
            st.caption("表A 映射关系")
            map_a = {}
            cfg_a = {
                'user': ['人员', '姓名'], 'spm': ['SPM', '标识符'], 'hours': ['交付工时', '工时'],
                'project': ['项目', '所属项目'], 'range': ['人事范围'], 'contract': ['合同主体'],
                'sales': ['销售', '销售人员'], 'dept': ['销售部门']
            }
            for k, v in cfg_a.items():
                map_a[k] = st.selectbox(f"{k}", cols_a, index=smart_idx(cols_a, v), key=f"a_{k}")
        
        with mc2:
            st.caption("表B 映射关系")
            map_b = {}
            cfg_b = {
                'user': ['出差人', '姓名'], 'spm': ['SPM', '项目编号'],
                'amount': ['金额', '总金额'], 'type': ['产品类型', '费用类型']
            }
            for k, v in cfg_b.items():
                map_b[k] = st.selectbox(f"{k}", cols_b, index=smart_idx(cols_b, v), key=f"b_{k}")

    st.divider()
    st.markdown("### 3. 报表校验与生成")

    # 执行按钮
    if st.button("开始校验并生成报表", use_container_width=True):
        
        # 1. 进度条容器
        progress_bar = st.progress(0)
        status_text = st.empty()

        try:
            # --- 阶段 1: 校验 ---
            status_text.text("正在进行数据完整性校验...")
            time.sleep(0.3) # 模拟体验
            progress_bar.progress(20)

            missing_spm = df_b[map_b['spm']].isnull().sum()
            if missing_spm > 0:
                progress_bar.empty()
                status_text.empty()
                st.error(f"❌ 校验未通过：表 B 中发现 {missing_spm} 条数据缺少 SPM 编号。请修改源文件后重新上传。")
                st.stop()
            
            # --- 阶段 2: 清洗与计算 ---
            status_text.text("校验通过，正在清洗数据...")
            progress_bar.progress(50)
            
            # A表聚合
            agg = {map_a['hours']: 'sum'}
            for k in ['project', 'range', 'contract', 'sales', 'dept']: agg[map_a[k]] = 'first'
            df_a_cl = df_a.dropna(subset=[map_a['spm']]).copy()
            df_a_gp = df_a_cl.groupby([map_a['user'], map_a['spm']], as_index=False).agg(agg)

            # B表拆分
            status_text.text("正在拆分差旅费用 (补助/费控)...")
            progress_bar.progress(70)
            df_b_cl = df_b.dropna(subset=[map_b['spm']]).copy()
            is_sub = df_b_cl[map_b['type']] == SUBSIDY_TAG
            df_sub = df_b_cl[is_sub].groupby([map_b['user'], map_b['spm']])[map_b['amount']].sum().reset_index(name='差旅补助')
            df_fee = df_b_cl[~is_sub].groupby([map_b['user'], map_b['spm']])[map_b['amount']].sum().reset_index(name='差旅费控平台')

            # 合并计算
            key_a = [map_a['user'], map_a['spm']]
            key_b = [map_b['user'], map_b['spm']]
            res = pd.merge(df_a_gp, df_sub, left_on=key_a, right_on=key_b, how='left')
            res = pd.merge(res, df_fee, left_on=key_a, right_on=key_b, how='left')
            res.fillna(0, inplace=True)
            
            res['支持时间'] = res[map_a['hours']] / 8
            res['人力费用'] = res['支持时间'] * PRICE_PER_DAY
            res['结算费用合计'] = res['人力费用'] + res['差旅补助'] + res['差旅费控平台']

            # --- 阶段 3: 格式化 ---
            progress_bar.progress(90)
            
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

            # --- 完成 ---
            progress_bar.progress(100)
            status_text.text("✅ 处理完成")
            time.sleep(0.5)
            progress_bar.empty()
            status_text.empty()

            # ==========================================
            # 6. 文件下载区域 (仅在通过后显示)
            # ==========================================
            st.divider()
            st.markdown("### 4. 报表下载")
            st.success("校验通过！报表已生成，请在下方下载。")

            with st.container(border=True):
                d1, d2, d3 = st.columns(3)
                
                def to_excel(df):
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                        df.to_excel(writer, index=False)
                    return output.getvalue()

                with d1:
                    st.download_button("📥 结果表1 (工时)", to_excel(t1), "结果表1.xlsx", use_container_width=True)
                with d2:
                    st.download_button("📥 结果表2 (结算)", to_excel(t2), "结果表2.xlsx", use_container_width=True)
                with d3:
                    st.download_button("📥 结果表3 (明细)", to_excel(t3), "结果表3.xlsx", use_container_width=True)

        except Exception as e:
            st.error(f"处理过程中发生系统错误: {e}")
