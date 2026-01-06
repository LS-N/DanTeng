import streamlit as st
import pandas as pd
import io
import time

# ==========================================
# 1. 状态初始化与主题引擎
# ==========================================
st.set_page_config(
    page_title="淡藤财务财务报表",
    page_icon="😈",  # <--- 已修改为小恶魔
    layout="wide",
    initial_sidebar_state="collapsed"
)

# 初始化主题状态 (默认为浅色 'light')
if 'theme' not in st.session_state:
    st.session_state.theme = 'light'

def toggle_theme():
    if st.session_state.theme == 'light':
        st.session_state.theme = 'dark'
    else:
        st.session_state.theme = 'light'

# --- 定义两套皮肤的 CSS 变量 ---
# 浅色模式 (Light)
theme_light = """
    --bg-color: #f8f9fc;
    --text-color: #2c3e50;
    --card-bg: #ffffff;
    --card-border: #e2e8f0;
    --primary-color: #6366f1;
    --accent-color: #818cf8;
    --mapping-line: #cbd5e1;
    --console-bg: #1e293b;
    --console-text: #10b981;
"""

# 深色模式 (Dark)
theme_dark = """
    --bg-color: #0f172a;
    --text-color: #e2e8f0;
    --card-bg: #1e293b;
    --card-border: #334155;
    --primary-color: #818cf8;
    --accent-color: #6366f1;
    --mapping-line: #475569;
    --console-bg: #000000;
    --console-text: #34d399;
"""

current_theme = theme_light if st.session_state.theme == 'light' else theme_dark

# 注入动态 CSS
st.markdown(f"""
<style>
    :root {{
        {current_theme}
    }}

    /* 全局应用变量 */
    .stApp {{
        background-color: var(--bg-color);
        color: var(--text-color);
        transition: background-color 0.3s ease;
    }}
    
    h1, h2, h3, h4, p, span, div {{
        color: var(--text-color) !important;
    }}

    /* 映射连接器 (科技感) */
    .connector-row {{
        display: flex;
        align-items: center;
        margin-bottom: 12px;
        position: relative;
    }}
    .connector-line {{
        flex-grow: 1;
        height: 2px;
        background: repeating-linear-gradient(
            90deg,
            var(--mapping-line),
            var(--mapping-line) 4px,
            transparent 4px,
            transparent 8px
        );
        margin: 0 15px;
        opacity: 0.6;
    }}
    .field-label {{
        background-color: var(--bg-color);
        padding: 4px 10px;
        border-radius: 4px;
        border: 1px solid var(--primary-color);
        color: var(--primary-color) !important;
        font-family: 'Courier New', monospace;
        font-weight: bold;
        font-size: 0.85rem;
        white-space: nowrap;
    }}

    /* 按钮样式 */
    .stButton>button {{
        background-color: var(--primary-color);
        color: white !important;
        border: none;
    }}
    .stButton>button:hover {{
        opacity: 0.9;
    }}

    /* 侧边栏 */
    section[data-testid="stSidebar"] {{
        background-color: var(--card-bg);
        border-right: 1px solid var(--card-border);
    }}
    
</style>
""", unsafe_allow_html=True)

# ==========================================
# 2. 顶部导航栏 (Title + Toggle)
# ==========================================
col_header, col_toggle = st.columns([9, 1])

with col_header:
    st.markdown("<h1>🍇 淡藤财务财务报表</h1>", unsafe_allow_html=True)
    st.markdown("<p style='opacity:0.7; margin-top:-10px;'>Automated Financial Settlement System | Enterprise Edition</p>", unsafe_allow_html=True)

with col_toggle:
    # 切换按钮：小恶魔图标
    btn_icon = "😈"  # <--- 已修改为小恶魔
    if st.button(btn_icon, help="切换深色/浅色模式"):
        toggle_theme()
        st.rerun()

st.divider()

# ==========================================
# 3. 侧边栏配置
# ==========================================
with st.sidebar:
    st.header("🛠️ 参数控制台")
    
    PRICE_PER_DAY = st.number_input(
        "人力单价 (CNY)", 
        value=1500, step=100
    )
    SUBSIDY_TAG = st.text_input(
        "补助关键词", 
        value="差旅补助"
    )
    st.markdown("---")
    st.caption("DanTeng Finance System v4.1")

# ==========================================
# 4. 文件上传 (使用原生边框容器修复空条BUG)
# ==========================================
# 这里改用了 with st.container(border=True) 替代了之前的 html hack
with st.container(border=True):
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("#### 📂 交付明细 (Source A)")
        file_a = st.file_uploader("Upload Project Data", type=['xlsx', 'csv'], key='a', label_visibility="collapsed")
    with c2:
        st.markdown("#### ✈️ 差旅明细 (Source B)")
        file_b = st.file_uploader("Upload Travel Data", type=['xlsx', 'csv'], key='b', label_visibility="collapsed")

# 读取逻辑
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
# 5. 字段映射 (工程感 UI)
# ==========================================
if df_a_raw is not None and df_b_raw is not None:
    
    # 列名清洗
    df_a_raw.columns = [str(c).strip() for c in df_a_raw.columns]
    df_b_raw.columns = [str(c).strip() for c in df_b_raw.columns]
    cols_a = list(df_a_raw.columns)
    cols_b = list(df_b_raw.columns)

    def smart_idx(options, keywords):
        for k in keywords:
            if k in options: return options.index(k)
        return 0

    with st.container(border=True):
        st.subheader("🔗 字段映射配置 (Mapping Topology)")
        
        mc1, mc2 = st.columns([1, 1], gap="large")

        # 渲染连接器样式的函数
        def render_connector(label, key, options, default_keys, prefix):
            # 使用 HTML/CSS 绘制左侧 Label --虚线--> 右侧 Selectbox
            c_left, c_right = st.columns([4, 6])
            with c_left:
                st.markdown(f"""
                <div class="connector-row">
                    <span class="field-label">{label}</span>
                    <div class="connector-line"></div>
                </div>
                """, unsafe_allow_html=True)
            with c_right:
                return st.selectbox(
                    f"Map {label}", options, 
                    index=smart_idx(options, default_keys), 
                    key=f"{prefix}_{key}", 
                    label_visibility="collapsed"
                )

        with mc1:
            st.caption("SOURCE A: 交付明细表")
            map_a = {}
            cfg_a = {
                'user': ['人员姓名', ['人员', '姓名']],
                'spm': ['SPM_编号', ['SPM', '标识符']],
                'hours': ['交付工时', ['交付工时（h）', '工时']],
                'project': ['项目名称', ['项目', '所属项目']],
                'range': ['人事范围', ['人事范围']],
                'contract': ['合同主体', ['合同主体']],
                'sales': ['销售人员', ['销售', '销售人员']],
                'dept': ['销售部门', ['销售部门']]
            }
            for k, v in cfg_a.items():
                map_a[k] = render_connector(v[0], k, cols_a, v[1], "a")

        with mc2:
            st.caption("SOURCE B: 实施差旅表")
            map_b = {}
            cfg_b = {
                'user': ['出差人员', ['出差人', '姓名', '人员']],
                'spm': ['SPM_编号', ['SPM', '项目编号']],
                'amount': ['报销金额', ['金额', '总金额']],
                'type': ['费用类型', ['产品类型', '费用类型']]
            }
            for k, v in cfg_b.items():
                map_b[k] = render_connector(v[0], k, cols_b, v[1], "b")

    # ==========================================
    # 6. 执行引擎
    # ==========================================
    if st.button("🚀 初始化计算引擎 (Execute Pipeline)", type="primary", use_container_width=True):
        
        # 控制台样式区域
        st.markdown(f"""
        <div style="background-color: var(--console-bg); border:1px solid #333; padding: 15px; border-radius: 8px; margin-bottom: 20px;">
            <div style="color: var(--console-text); font-family: 'Consolas', monospace; font-size: 0.9rem;">
                <span style="opacity:0.5">root@danteng-finance:~$</span> ./run_settlement.sh<br>
        """, unsafe_allow_html=True)
        
        console = st.empty()
        
        def log(msg, type="info"):
            icon = "INFO" if type=="info" else "PASS" if type=="success" else "FAIL"
            color = "var(--console-text)" if type!="error" else "#ef4444"
            console.markdown(f"""
            <div style="color: {color}; font-family: 'Consolas', monospace; margin-left: 20px;">
                [{time.strftime('%H:%M:%S')}] [{icon}] {msg}
            </div>
            """, unsafe_allow_html=True)
            time.sleep(0.1)

        try:
            log("System Check Initiated...")
            
            # 1. 校验
            missing_spm = df_b_raw[map_b['spm']].isnull().sum()
            if missing_spm > 0:
                st.markdown("</div>", unsafe_allow_html=True) # 关闭控制台
                # 弹窗报错
                st.toast(f"❌ 校验被阻断：表B 发现 {missing_spm} 条空 SPM 数据！", icon="🚨")
                st.error(f"Critical Error: Table B contains {missing_spm} missing SPM records.")
                st.stop()
            
            log(f"Integrity Check Passed. Processing {len(df_a_raw)} records from Source A...", "success")

            # 2. 清洗
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
            
            log("Expense Classification Complete (Subsidy vs Fee Control).", "success")

            # 3. 计算
            key_a = [map_a['user'], map_a['spm']]
            key_b = [map_b['user'], map_b['spm']]
            
            df_main = pd.merge(df_a_gp, df_sub, left_on=key_a, right_on=key_b, how='left')
            df_main = pd.merge(df_main, df_fee, left_on=key_a, right_on=key_b, how='left')
            df_main.fillna(0, inplace=True)
            
            df_main['支持时间'] = df_main[map_a['hours']] / 8
            df_main['人力费用'] = df_main['支持时间'] * PRICE_PER_DAY
            df_main['结算费用合计'] = df_main['人力费用'] + df_main['差旅补助'] + df_main['差旅费控平台']
            
            log("Financial Calculation & Ledger Generation Complete.", "success")
            st.markdown("</div></div>", unsafe_allow_html=True) # 关闭控制台

            # 成功弹窗
            st.toast("✅ 计算成功！报表已生成。", icon="🎉")

            # 4. 结果下载 (使用原生容器)
            with st.container(border=True):
                st.subheader("📊 报表下载中心 (Download Center)")
                
                # 生成表 3
                rename_dict = {
                    map_a['user']: '人员', map_a['project']: '所属项目', map_a['range']: '人事范围',
                    map_a['spm']: 'SPM', map_a['contract']: '合同主体', map_a['sales']: '销售人员',
                    map_a['dept']: '销售部门', map_a['hours']: '耗时（小时）'
                }
                t3 = df_main.rename(columns=rename_dict)
                cols_final = ['人员', '所属项目', '人事范围', 'SPM', '合同主体', '销售人员', '销售部门',
                              '差旅补助', '差旅费控平台', '耗时（小时）', '支持时间', '人力费用', '结算费用合计']
                t3 = t3[[c for c in cols_final if c in t3.columns]] 
                t3.rename(columns={'支持时间': '支持时间（人天）'}, inplace=True)
                t3.insert(0, '序号', range(1, len(t3)+1))

                # 生成表 2
                grp_cols = ['人事范围', '合同主体', '销售部门']
                if all(c in t3.columns for c in grp_cols):
                    t2 = t3.groupby(grp_cols).agg({'结算费用合计': 'sum', '支持时间（人天）': 'sum'}).reset_index()
                    t2.columns = ['销售公司', '采购公司', '采购部门', '金额（含税，单位：元）', '工作量（人天）']
                    t2['备注'] = ''
                    t2.insert(0, '序号', range(1, len(t2)+1))
                else:
                    t2 = pd.DataFrame()

                # 生成表 1
                t1 = t3.groupby('人员')['耗时（小时）'].sum().reset_index()
                t1.rename(columns={'耗时（小时）': '项目工时'}, inplace=True)
                t1.insert(0, '序号', range(1, len(t1)+1))

                def to_excel(df):
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                        df.to_excel(writer, index=False)
                    return output.getvalue()

                tab1, tab2, tab3 = st.tabs(["结果表3 (明细)", "结果表2 (结算)", "结果表1 (工时)"])
                with tab1:
                    st.download_button("📥 下载 结果表3.xlsx", to_excel(t3), "结果表3.xlsx", use_container_width=True)
                    st.dataframe(t3, height=250)
                with tab2:
                    st.download_button("📥 下载 结果表2.xlsx", to_excel(t2), "结果表2.xlsx", use_container_width=True)
                    st.dataframe(t2, height=250)
                with tab3:
                    st.download_button("📥 下载 结果表1.xlsx", to_excel(t1), "结果表1.xlsx", use_container_width=True)
                    st.dataframe(t1, height=250)

        except Exception as e:
            st.markdown("</div>", unsafe_allow_html=True)
            st.toast(f"系统运行错误: {str(e)}", icon="🔥")
            st.error(f"Error Log: {str(e)}")

else:
    # 空状态美化
    st.info("👈 请在上方上传数据文件以激活映射配置面板。")
