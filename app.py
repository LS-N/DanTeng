import streamlit as st
import pandas as pd
import io
import time

# ==========================================
# 1. 核心配置与 高级 UI 注入
# ==========================================
st.set_page_config(
    page_title="AutoFinance Pro",
    page_icon="💳",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# 注入 CSS：这是一套现代商务风格的皮肤
st.markdown("""
<style>
    /* --- 1. 全局背景：高级灰 --- */
    .stApp {
        background-color: #f4f6f9; /* 柔和的灰背景 */
        color: #2c3e50;
    }

    /* --- 2. 侧边栏：深邃夜空 --- */
    section[data-testid="stSidebar"] {
        background-color: #1a202c; /* 近似黑色的深蓝 */
    }
    section[data-testid="stSidebar"] h2, section[data-testid="stSidebar"] span {
        color: #e2e8f0 !important;
    }

    /* --- 3. 卡片化容器 (关键美化) --- */
    .stMarkdown, .stDataFrame, .stAlert, div[data-testid="stVerticalBlock"] > div {
        /* 这里不直接给所有元素加框，而是通过特定类控制 */
    }
    
    /* 自定义卡片类 */
    .card-box {
        background-color: white;
        padding: 25px;
        border-radius: 12px;
        box-shadow: 0 4px 20px rgba(0,0,0,0.05); /* 极柔和的阴影 */
        margin-bottom: 20px;
        border: 1px solid #edf2f7;
    }

    /* --- 4. 标题优化 --- */
    h1 {
        font-family: 'Helvetica Neue', sans-serif;
        font-weight: 700;
        color: #1e293b;
        letter-spacing: -0.5px;
    }
    h3 {
        font-weight: 600;
        color: #334155;
        margin-top: 0 !important;
    }

    /* --- 5. 按钮升级：渐变商务蓝 --- */
    .stButton>button {
        background: linear-gradient(135deg, #3b82f6 0%, #2563eb 100%);
        color: white;
        border: none;
        padding: 0.6rem 1.2rem;
        border-radius: 8px;
        font-weight: 500;
        box-shadow: 0 4px 6px rgba(37, 99, 235, 0.2);
        transition: all 0.3s ease;
        width: 100%;
    }
    .stButton>button:hover {
        transform: translateY(-2px);
        box-shadow: 0 6px 15px rgba(37, 99, 235, 0.3);
    }
    
    /* --- 6. 上传框美化 --- */
    div[data-testid="stFileUploader"] {
        border: 1px dashed #cbd5e1;
        border-radius: 10px;
        padding: 10px;
        background-color: #f8fafc;
    }

    /* --- 7. 映射区的连接线效果 --- */
    .mapping-row {
        display: flex;
        align-items: center;
        padding: 8px 0;
        border-bottom: 1px solid #f1f5f9;
    }
    .mapping-label {
        font-size: 0.9rem;
        font-weight: 600;
        color: #64748b;
        width: 30%;
    }
    
    /* 进度条颜色 */
    .stProgress > div > div > div > div {
        background-color: #3b82f6;
    }

</style>
""", unsafe_allow_html=True)

# ==========================================
# 2. 侧边栏：隐藏式参数
# ==========================================
with st.sidebar:
    st.markdown("### ⚙️ 全局设置")
    st.info("调整下方参数可即时影响计算结果")
    
    PRICE_PER_DAY = st.number_input(
        "人力单价 (CNY/Day)", 
        value=1500, step=100
    )
    SUBSIDY_TAG = st.text_input(
        "差旅补助关键词", 
        value="差旅补助"
    )
    st.markdown("---")
    st.caption("AutoFinance System v3.0")

# ==========================================
# 3. 头部与文件上传 (卡片式)
# ==========================================
st.title("AutoFinance 财务结算中心")
st.markdown("Automated Settlement Pipeline")

# 使用 HTML 容器包裹，创造卡片视觉
st.markdown('<div class="card-box">', unsafe_allow_html=True)
c1, c2 = st.columns(2)
with c1:
    st.markdown("#### 📂 交付明细表 (表A)")
    file_a = st.file_uploader("Upload Project Data", type=['xlsx', 'csv'], key='a', label_visibility="collapsed")
with c2:
    st.markdown("#### ✈️ 差旅明细表 (表B)")
    file_b = st.file_uploader("Upload Travel Data", type=['xlsx', 'csv'], key='b', label_visibility="collapsed")
st.markdown('</div>', unsafe_allow_html=True)

# 缓存读取
@st.cache_data(ttl=600)
def load_data(file):
    if not file: return None
    try:
        if file.name.endswith('.csv'):
            try: return pd.read_csv(file)
            except: return pd.read_csv(file, encoding='gbk')
        else: return pd.read_excel(file)
    except Exception as e:
        st.error(f"Error: {e}")
        return None

df_a_raw = load_data(file_a)
df_b_raw = load_data(file_b)

# ==========================================
# 4. 字段映射 (极简科技感)
# ==========================================
if df_a_raw is not None and df_b_raw is not None:
    
    # 清洗列名
    df_a_raw.columns = [str(c).strip() for c in df_a_raw.columns]
    df_b_raw.columns = [str(c).strip() for c in df_b_raw.columns]
    cols_a = list(df_a_raw.columns)
    cols_b = list(df_b_raw.columns)

    def smart_idx(options, keywords):
        for k in keywords:
            if k in options: return options.index(k)
        return 0

    # 映射卡片
    st.markdown('<div class="card-box">', unsafe_allow_html=True)
    st.markdown("#### 🔗 字段映射配置 (Field Mapping)")
    st.markdown("<p style='font-size:14px; color:#94a3b8; margin-bottom:20px'>请确认左侧系统标准字段与右侧 Excel 列名的对应关系</p>", unsafe_allow_html=True)
    
    mc1, mc2 = st.columns([1, 1], gap="large")

    with mc1:
        st.caption("SOURCE A: PROJECT DATA")
        map_a = {}
        cfg_a = {
            'user': ['👤 人员姓名', ['人员', '姓名']],
            'spm': ['🆔 SPM 编号', ['SPM', '标识符']],
            'hours': ['⏱️ 交付工时', ['交付工时（h）', '工时']],
            'project': ['📁 项目名称', ['项目', '所属项目']],
            'range': ['🏢 人事范围', ['人事范围']],
            'contract': ['📜 合同主体', ['合同主体']],
            'sales': ['💼 销售人员', ['销售', '销售人员']],
            'dept': ['📊 销售部门', ['销售部门']]
        }
        for k, v in cfg_a.items():
            # 自定义 HTML 布局来实现对齐
            col_label, col_select = st.columns([4, 6])
            col_label.markdown(f"<div style='margin-top: 10px; font-weight:500; font-size:14px'>{v[0]}</div>", unsafe_allow_html=True)
            map_a[k] = col_select.selectbox(f"match {k}", cols_a, index=smart_idx(cols_a, v[1]), label_visibility="collapsed")

    with mc2:
        st.caption("SOURCE B: TRAVEL DATA")
        map_b = {}
        cfg_b = {
            'user': ['👤 出差人', ['出差人', '姓名', '人员']],
            'spm': ['🆔 SPM 编号', ['SPM', '项目编号']],
            'amount': ['💰 报销金额', ['金额', '总金额']],
            'type': ['🏷️ 费用类型', ['产品类型', '费用类型']]
        }
        for k, v in cfg_b.items():
            col_label, col_select = st.columns([4, 6])
            col_label.markdown(f"<div style='margin-top: 10px; font-weight:500; font-size:14px'>{v[0]}</div>", unsafe_allow_html=True)
            map_b[k] = col_select.selectbox(f"match {k}", cols_b, index=smart_idx(cols_b, v[1]), label_visibility="collapsed")
    
    st.markdown('</div>', unsafe_allow_html=True)

    # ==========================================
    # 5. 执行与控制台
    # ==========================================
    if st.button("开始计算 (Run Calculation)"):
        
        # 控制台 UI
        st.markdown('<div class="card-box" style="background-color:#1e293b; color:#10b981; font-family:monospace; padding:15px;">', unsafe_allow_html=True)
        console = st.empty()
        
        def log(msg, type="info"):
            color = "#10b981" if type=="success" else "#f59e0b" if type=="warn" else "#ef4444"
            console.markdown(f"<span style='color:{color}'> > {msg}</span>", unsafe_allow_html=True)
            time.sleep(0.2)

        try:
            log("Initializing pipeline...")
            
            # --- 校验 ---
            log("Validating data integrity...")
            missing_spm = df_b_raw[map_b['spm']].isnull().sum()
            if missing_spm > 0:
                log(f"CRITICAL ERROR: Found {missing_spm} missing SPM records in Table B!", "error")
                st.toast(f"❌ 校验失败：表B 存在 {missing_spm} 条缺失 SPM 的数据", icon="🚨")
                st.stop()
            
            # --- 清洗 ---
            log("Cleaning and transforming data...")
            # A表聚合
            agg_rules = {map_a['hours']: 'sum'}
            for k in ['project', 'range', 'contract', 'sales', 'dept']:
                agg_rules[map_a[k]] = 'first'
            
            df_a_clean = df_a_raw.dropna(subset=[map_a['spm']]).copy()
            df_a_gp = df_a_clean.groupby([map_a['user'], map_a['spm']], as_index=False).agg(agg_rules)

            # B表拆分
            df_b_clean = df_b_raw.dropna(subset=[map_b['spm']]).copy()
            is_sub = df_b_clean[map_b['type']] == SUBSIDY_TAG
            
            df_sub = df_b_clean[is_sub].groupby([map_b['user'], map_b['spm']])[map_b['amount']].sum().reset_index(name='差旅补助')
            df_fee = df_b_clean[~is_sub].groupby([map_b['user'], map_b['spm']])[map_b['amount']].sum().reset_index(name='差旅费控平台')
            
            log("Mapping expenses to projects...")

            # --- 计算 ---
            key_a = [map_a['user'], map_a['spm']]
            key_b = [map_b['user'], map_b['spm']]
            
            df_main = pd.merge(df_a_gp, df_sub, left_on=key_a, right_on=key_b, how='left')
            df_main = pd.merge(df_main, df_fee, left_on=key_a, right_on=key_b, how='left')
            df_main.fillna(0, inplace=True)
            
            df_main['支持时间'] = df_main[map_a['hours']] / 8
            df_main['人力费用'] = df_main['支持时间'] * PRICE_PER_DAY
            df_main['结算费用合计'] = df_main['人力费用'] + df_main['差旅补助'] + df_main['差旅费控平台']
            
            log("Calculating financial metrics...", "success")

            # --- 生成 ---
            # 表3
            rename_dict = {
                map_a['user']: '人员', map_a['project']: '所属项目', map_a['range']: '人事范围',
                map_a['spm']: 'SPM', map_a['contract']: '合同主体', map_a['sales']: '销售人员',
                map_a['dept']: '销售部门', map_a['hours']: '耗时（小时）'
            }
            t3 = df_main.rename(columns=rename_dict)
            cols_final = ['人员', '所属项目', '人事范围', 'SPM', '合同主体', '销售人员', '销售部门',
                          '差旅补助', '差旅费控平台', '耗时（小时）', '支持时间', '人力费用', '结算费用合计']
            t3 = t3[[c for c in cols_final if c in t3.columns]] # 容错
            t3.rename(columns={'支持时间': '支持时间（人天）'}, inplace=True)
            t3.insert(0, '序号', range(1, len(t3)+1))

            # 表2
            grp_cols = ['人事范围', '合同主体', '销售部门']
            if all(c in t3.columns for c in grp_cols):
                t2 = t3.groupby(grp_cols).agg({'结算费用合计': 'sum', '支持时间（人天）': 'sum'}).reset_index()
                t2.columns = ['销售公司', '采购公司', '采购部门', '金额（含税，单位：元）', '工作量（人天）']
                t2['备注'] = ''
                t2.insert(0, '序号', range(1, len(t2)+1))
            else:
                t2 = pd.DataFrame()

            # 表1
            t1 = t3.groupby('人员')['耗时（小时）'].sum().reset_index()
            t1.rename(columns={'耗时（小时）': '项目工时'}, inplace=True)
            t1.insert(0, '序号', range(1, len(t1)+1))
            
            log("PIPELINE COMPLETED SUCCESSFULLY.", "success")
            st.markdown('</div>', unsafe_allow_html=True) # 结束控制台div
            
            st.toast("任务完成！正在渲染结果...", icon="✅")

            # --- 结果下载 (卡片式) ---
            st.markdown('<div class="card-box">', unsafe_allow_html=True)
            st.subheader("📊 结算报表下载")
            
            tab1, tab2, tab3 = st.tabs(["结果表3 (明细)", "结果表2 (结算)", "结果表1 (工时)"])
            
            def to_excel(df):
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df.to_excel(writer, index=False)
                return output.getvalue()

            with tab1:
                st.dataframe(t3, height=300)
                st.download_button("📥 下载 结果表3.xlsx", to_excel(t3), "结果表3.xlsx", type="primary")
            with tab2:
                st.dataframe(t2, height=300)
                st.download_button("📥 下载 结果表2.xlsx", to_excel(t2), "结果表2.xlsx", type="primary")
            with tab3:
                st.dataframe(t1, height=300)
                st.download_button("📥 下载 结果表1.xlsx", to_excel(t1), "结果表1.xlsx", type="primary")
            
            st.markdown('</div>', unsafe_allow_html=True)

        except Exception as e:
            st.error(f"System Error: {str(e)}")
            st.markdown('</div>', unsafe_allow_html=True)

else:
    # 引导提示
    st.info("👋 Welcome! Please unfold the sidebar settings and upload data files to begin.")
