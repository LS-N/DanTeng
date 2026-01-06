import streamlit as st
import pandas as pd
import io
import time

# ==========================================
# 1. 核心配置与 CSS 美化 (UI/UX)
# ==========================================
st.set_page_config(
    page_title="AutoFinance Pro | 财务自动化结算系统",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="collapsed" # 默认收起侧边栏，保持界面干净
)

# 注入自定义 CSS 以实现“科技感”和“干净”风格
st.markdown("""
<style>
    /* 全局字体与背景 */
    .stApp {
        background-color: #f8f9fa;
        font-family: 'Segoe UI', Roboto, Helvetica, Arial, sans-serif;
    }
    
    /* 标题样式 */
    h1 { color: #1e3a8a; font-weight: 700; letter-spacing: -1px; }
    h2, h3 { color: #334155; }
    
    /* 卡片式容器 */
    .css-1r6slb0, .stMarkdown, .stDataFrame {
        border-radius: 8px;
    }
    
    /* 映射区域的工程感 */
    .mapping-box {
        background-color: #ffffff;
        padding: 20px;
        border-radius: 10px;
        border-left: 5px solid #3b82f6; /* 蓝色科技条 */
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
        margin-bottom: 20px;
    }
    
    /* 按钮美化 */
    .stButton>button {
        background-color: #2563eb;
        color: white;
        border-radius: 6px;
        font-weight: 600;
        transition: all 0.3s ease;
    }
    .stButton>button:hover {
        background-color: #1d4ed8;
        transform: translateY(-1px);
        box-shadow: 0 4px 12px rgba(37, 99, 235, 0.3);
    }
    
    /* 侧边栏美化 */
    section[data-testid="stSidebar"] {
        background-color: #1e293b; /* 深色侧边栏 */
        color: white;
    }
</style>
""", unsafe_allow_html=True)

# ==========================================
# 2. 侧边栏：全局参数 (Global Params)
# ==========================================
with st.sidebar:
    st.header("⚙️ 全局参数配置")
    st.markdown("---")
    
    # 使用 number_input 的 format 参数增加专业感
    PRICE_PER_DAY = st.number_input(
        "人力单价 (CNY/Day)", 
        value=1500, 
        step=100, 
        help="用于计算支持时间对应的标准人力费用"
    )
    
    SUBSIDY_TAG = st.text_input(
        "差旅补助关键词", 
        value="差旅补助", 
        help="在表B的产品类型中，匹配此关键词的金额将计入补助，其余计入费控。"
    )
    
    st.markdown("---")
    st.caption("v2.0.1 Build 2026")
    st.caption("Designed for Finance Ops")

# ==========================================
# 3. 主界面：头部与上传
# ==========================================
st.title("⚡ 财务自动化结算系统 Pro")
st.markdown("##### Enterprise Grade ETL Pipeline for Financial Settlement")
st.divider()

# 文件上传区
c1, c2 = st.columns(2)
with c1:
    st.markdown("### 📂 Source A")
    file_a = st.file_uploader("上传 交付明细 (Project Hours)", type=['xlsx', 'csv'], key='a')
with c2:
    st.markdown("### 📂 Source B")
    file_b = st.file_uploader("上传 实施差旅 (Travel Expenses)", type=['xlsx', 'csv'], key='b')

# 核心读取函数
@st.cache_data(ttl=600) # 增加缓存，提升体验
def load_data(file):
    if not file: return None
    try:
        if file.name.endswith('.csv'):
            try: return pd.read_csv(file)
            except: return pd.read_csv(file, encoding='gbk')
        else: return pd.read_excel(file)
    except Exception as e:
        st.error(f"无法读取文件: {e}")
        return None

df_a_raw = load_data(file_a)
df_b_raw = load_data(file_b)

# ==========================================
# 4. 字段映射模块 (工程感 UI)
# ==========================================
if df_a_raw is not None and df_b_raw is not None:
    
    st.markdown("<br>", unsafe_allow_html=True)
    st.subheader("🔗 字段拓扑映射 (Field Mapping)")
    
    # 预处理列名
    df_a_raw.columns = [str(c).strip() for c in df_a_raw.columns]
    df_b_raw.columns = [str(c).strip() for c in df_b_raw.columns]
    cols_a = list(df_a_raw.columns)
    cols_b = list(df_b_raw.columns)

    # 智能匹配辅助函数
    def smart_idx(options, keywords):
        for k in keywords:
            if k in options: return options.index(k)
        return 0

    # 容器化布局 - 更有层次感
    with st.container():
        st.markdown('<div class="mapping-box">', unsafe_allow_html=True)
        
        col_map_1, col_map_2 = st.columns([1, 1], gap="large")
        
        # 左侧：表A 映射
        with col_map_1:
            st.markdown("#### 🛠️ 表A (交付明细) 配置")
            map_a = {}
            # 格式: 键: [显示名, 默认匹配词]
            cfg_a = {
                'user': ['人员姓名', ['人员', '姓名']],
                'spm': ['SPM 编号', ['SPM', '标识符']],
                'hours': ['交付工时', ['交付工时（h）', '工时']],
                'project': ['项目名称', ['项目', '所属项目']],
                'range': ['人事范围', ['人事范围']],
                'contract': ['合同主体', ['合同主体']],
                'sales': ['销售人员', ['销售', '销售人员']],
                'dept': ['销售部门', ['销售部门']]
            }
            for k, v in cfg_a.items():
                # 使用列布局模拟 "Link" 效果
                sub_c1, sub_c2 = st.columns([3, 7])
                sub_c1.markdown(f"**{v[0]}**")
                sub_c1.caption("System Field")
                map_a[k] = sub_c2.selectbox(f"Select match for {v[0]}", cols_a, index=smart_idx(cols_a, v[1]), label_visibility="collapsed")

        # 右侧：表B 映射
        with col_map_2:
            st.markdown("#### ✈️ 表B (实施差旅) 配置")
            map_b = {}
            cfg_b = {
                'user': ['出差人', ['出差人', '姓名', '人员']],
                'spm': ['SPM 编号', ['SPM', '项目编号']],
                'amount': ['报销金额', ['金额', '总金额']],
                'type': ['费用类型', ['产品类型', '费用类型']]
            }
            for k, v in cfg_b.items():
                sub_c1, sub_c2 = st.columns([3, 7])
                sub_c1.markdown(f"**{v[0]}**")
                sub_c1.caption("System Field")
                map_b[k] = sub_c2.selectbox(f"Select match for {v[0]}", cols_b, index=smart_idx(cols_b, v[1]), label_visibility="collapsed")
        
        st.markdown('</div>', unsafe_allow_html=True)

    # ==========================================
    # 5. 执行引擎 (带控制台日志)
    # ==========================================
    run_btn = st.button("🚀 启动自动化处理流程 (Start Pipeline)", type="primary", use_container_width=True)

    if run_btn:
        # 1. 占位符：用于显示控制台日志
        console = st.empty()
        
        # 日志函数
        def log(msg, status="INFO"):
            icon = "✅" if status=="SUCCESS" else "⏳" if status=="INFO" else "❌"
            console.code(f"[{time.strftime('%H:%M:%S')}] {icon} {msg}", language="bash")
            time.sleep(0.3) # 模拟处理延迟，增加真实感

        try:
            log("正在初始化计算引擎...", "INFO")
            
            # --- 阶段 1: 校验 (Validation) ---
            log("正在扫描数据完整性...", "INFO")
            
            # 校验 SPM
            spm_col_b = map_b['spm']
            missing_spm = df_b_raw[spm_col_b].isnull().sum()
            if missing_spm > 0:
                log(f"校验失败: 表B {spm_col_b} 列存在 {missing_spm} 个空值", "ERROR")
                st.toast(f"❌ 校验失败：表B存在 {missing_spm} 条没有SPM号的数据！", icon="🚫")
                st.stop()
            
            log("数据校验通过。", "SUCCESS")

            # --- 阶段 2: 数据清洗 (Transform) ---
            log("正在执行聚合与清洗逻辑...", "INFO")
            
            # 清洗 表A
            agg_rules = {map_a['hours']: 'sum'}
            for k in ['project', 'range', 'contract', 'sales', 'dept']:
                agg_rules[map_a[k]] = 'first'
            
            df_a_clean = df_a_raw.dropna(subset=[map_a['spm']]).copy()
            df_a_gp = df_a_clean.groupby([map_a['user'], map_a['spm']], as_index=False).agg(agg_rules)
            
            # 清洗 表B
            df_b_clean = df_b_raw.dropna(subset=[map_b['spm']]).copy()
            
            # 拆分逻辑
            is_sub = df_b_clean[map_b['type']] == SUBSIDY_TAG
            
            # 补助表
            df_sub = df_b_clean[is_sub].groupby([map_b['user'], map_b['spm']])[map_b['amount']].sum().reset_index(name='差旅补助')
            
            # 费控表
            df_fee = df_b_clean[~is_sub].groupby([map_b['user'], map_b['spm']])[map_b['amount']].sum().reset_index(name='差旅费控平台')
            
            log("费用类型拆分完成 (补助/费控)。", "SUCCESS")

            # --- 阶段 3: 关联与计算 (Calculation) ---
            log("正在执行多表关联 (Left Join)...", "INFO")
            
            # 关联 Key
            key_a = [map_a['user'], map_a['spm']]
            key_b = [map_b['user'], map_b['spm']]
            
            # Merge
            df_main = pd.merge(df_a_gp, df_sub, left_on=key_a, right_on=key_b, how='left')
            df_main = pd.merge(df_main, df_fee, left_on=key_a, right_on=key_b, how='left')
            
            # 填补 NaN
            df_main[['差旅补助', '差旅费控平台']] = df_main[['差旅补助', '差旅费控平台']].fillna(0)
            
            # 算钱
            df_main['支持时间（人天）'] = df_main[map_a['hours']] / 8
            df_main['人力费用'] = df_main['支持时间（人天）'] * PRICE_PER_DAY
            df_main['结算费用合计'] = df_main['人力费用'] + df_main['差旅补助'] + df_main['差旅费控平台']
            
            log("核心财务指标计算完成。", "SUCCESS")

            # --- 阶段 4: 格式化输出 (Load) ---
            log("正在生成最终报表...", "INFO")
            
            # 生成表 3
            rename_dict = {
                map_a['user']: '人员',
                map_a['project']: '所属项目',
                map_a['range']: '人事范围',
                map_a['spm']: 'SPM',
                map_a['contract']: '合同主体',
                map_a['sales']: '销售人员',
                map_a['dept']: '销售部门',
                map_a['hours']: '耗时（小时）'
            }
            t3 = df_main.rename(columns=rename_dict)
            cols_final = ['人员', '所属项目', '人事范围', 'SPM', '合同主体', '销售人员', '销售部门',
                          '差旅补助', '差旅费控平台', '耗时（小时）', '支持时间（人天）', '人力费用', '结算费用合计']
            # 容错提取
            t3 = t3[[c for c in cols_final if c in t3.columns]]
            t3.insert(0, '序号', range(1, len(t3)+1))
            
            # 生成表 2
            grp_cols = ['人事范围', '合同主体', '销售部门']
            if all(c in t3.columns for c in grp_cols):
                t2 = t3.groupby(grp_cols).agg({
                    '结算费用合计': 'sum',
                    '支持时间（人天）': 'sum'
                }).reset_index()
                t2.columns = ['销售公司', '采购公司', '采购部门', '金额（含税，单位：元）', '工作量（人天）']
                t2['备注'] = ''
                t2.insert(0, '序号', range(1, len(t2)+1))
            else:
                t2 = pd.DataFrame()
                log("无法生成表2：缺少必要聚合列", "ERROR")

            # 生成表 1
            t1 = t3.groupby('人员')['耗时（小时）'].sum().reset_index()
            t1.rename(columns={'耗时（小时）': '项目工时'}, inplace=True)
            t1.insert(0, '序号', range(1, len(t1)+1))
            
            log("所有任务执行完毕。", "SUCCESS")
            
            # 弹窗提示成功
            st.toast("🎉 处理完成！结果表已准备就绪。", icon="✅")

            # --- 结果展示区 ---
            st.divider()
            st.subheader("📊 结果看板 (Dashboard)")
            
            tab1, tab2, tab3 = st.tabs(["📋 结果表3 (明细)", "💰 结果表2 (结算)", "⏱️ 结果表1 (工时)"])
            
            def to_excel(df):
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df.to_excel(writer, index=False)
                return output.getvalue()

            with tab1:
                st.dataframe(t3, use_container_width=True)
                st.download_button("📥 下载 结果表3.xlsx", to_excel(t3), "结果表3.xlsx")
            with tab2:
                st.dataframe(t2, use_container_width=True)
                st.download_button("📥 下载 结果表2.xlsx", to_excel(t2), "结果表2.xlsx")
            with tab3:
                st.dataframe(t1, use_container_width=True)
                st.download_button("📥 下载 结果表1.xlsx", to_excel(t1), "结果表1.xlsx")

        except Exception as e:
            st.error(f"系统运行崩溃，请截图联系开发者: {str(e)}")
            import traceback
            st.exception(e)

else:
    # 空状态美化
    st.info("👈 请点击左上角箭头展开参数配置，并在上方上传数据文件以开始。")
