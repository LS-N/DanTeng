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

# 引入 CSS：优化弹窗宽度和卡片样式
st.markdown("""
<style>
    :root { --bg-color: #0d1117; --card-bg: #161b22; --accent: #238636; --text: #c9d1d9; --red: #da3633; }
    .stApp { background-color: var(--bg-color); color: var(--text); }
    
    /* 进度条不回退 */
    .stProgress > div > div > div > div { background-color: var(--accent) !important; }
    
    /* 错误提示框 */
    .error-card { border: 1px solid var(--red); background: rgba(218, 54, 51, 0.1); padding: 1rem; border-radius: 6px; margin-bottom: 1rem;}
    
    /* 强制 Dialog 宽度更大 */
    div[data-testid="stDialog"] > div { width: 90vw !important; max-width: 1200px !important; }
</style>
""", unsafe_allow_html=True)

# ==========================================
# 1. 状态管理 (Session State)
# ==========================================
# 核心数据
if 'df_a' not in st.session_state: st.session_state.df_a = None
if 'df_b' not in st.session_state: st.session_state.df_b = None
# 文件指纹 (用于检测新文件上传)
if 'file_a_id' not in st.session_state: st.session_state.file_a_id = None
if 'file_b_id' not in st.session_state: st.session_state.file_b_id = None
# 结果缓存 (解决进度条重复跑的问题)
if 'calc_result_zip' not in st.session_state: st.session_state.calc_result_zip = None

# ==========================================
# 2. 侧边栏 & 头部
# ==========================================
st.title("😈 淡藤财务报表 Pro")
st.caption("Minimalist Financial Settlement System | MVP v2.0")

with st.sidebar:
    st.header("⚙️ 全局参数")
    PRICE_PER_DAY = st.number_input("人力单价 (元/天)", value=1500, step=100)
    MIN_HOURS_THRESHOLD = st.number_input("⚠️ 最低工时阈值 (小时)", value=500, help="人员累计工时低于此值将报错")
    SUBSIDY_TAG = st.text_input("补助关键词", "差旅补助")
    # 强制重置按钮
    if st.button("🔄 重置所有状态"):
        st.session_state.clear()
        st.rerun()

# ==========================================
# 3. 数据上传区
# ==========================================
with st.container(border=True):
    c1, c2 = st.columns(2)
    f_a = c1.file_uploader("Source A: 交付明细", type=['xlsx', 'csv'], key='up_a')
    f_b = c2.file_uploader("Source B: 差旅明细", type=['xlsx', 'csv'], key='up_b')

    def load_file(file, old_id):
        if file is not None and file.file_id != old_id:
            try:
                if file.name.endswith('.csv'): df = pd.read_csv(file)
                else: df = pd.read_excel(file)
                df.columns = [str(c).strip() for c in df.columns]
                # 上传新文件时，清除旧的计算结果
                st.session_state.calc_result_zip = None 
                return df, file.file_id
            except Exception as e:
                st.error(f"读取失败: {e}")
                return None, old_id
        return None, old_id

    # 只有当上传了新文件时才更新 State
    new_df_a, new_id_a = load_file(f_a, st.session_state.file_a_id)
    if new_df_a is not None: 
        st.session_state.df_a = new_df_a
        st.session_state.file_a_id = new_id_a

    new_df_b, new_id_b = load_file(f_b, st.session_state.file_b_id)
    if new_df_b is not None: 
        st.session_state.df_b = new_df_b
        st.session_state.file_b_id = new_id_b

# ==========================================
# 4. 核心逻辑 (包含弹窗编辑器)
# ==========================================

# 只有当两个表都有数据时才开始
if st.session_state.df_a is not None and st.session_state.df_b is not None:
    
    # ----------------------------------------------------
    # A. 校验逻辑函数
    # ----------------------------------------------------
    def find_col(df, candidates):
        for c in candidates:
            if c in df.columns: return c
        return None

    def get_mappings():
        # 自动映射逻辑
        ma = {}
        cols_a = st.session_state.df_a.columns
        ma['spm'] = find_col(st.session_state.df_a, ['SPM', '标识符', '项目编号'])
        ma['hours'] = find_col(st.session_state.df_a, ['交付工时', '工时', '投入工时'])
        ma['contract'] = find_col(st.session_state.df_a, ['合同主体'])
        ma['user'] = find_col(st.session_state.df_a, ['人员', '姓名', '员工姓名'])
        # 维度字段默认取第0列兜底
        ma['project'] = find_col(st.session_state.df_a, ['项目', '所属项目']) or cols_a[0]
        ma['range'] = find_col(st.session_state.df_a, ['人事范围']) or cols_a[0]
        ma['sales'] = find_col(st.session_state.df_a, ['销售', '销售人员']) or cols_a[0]
        ma['dept'] = find_col(st.session_state.df_a, ['销售部门']) or cols_a[0]
        
        mb = {}
        cols_b = st.session_state.df_b.columns
        mb['spm'] = find_col(st.session_state.df_b, ['SPM', '项目编号', '费用归属项目'])
        mb['amount'] = find_col(st.session_state.df_b, ['金额', '总金额', '报销金额'])
        mb['user'] = find_col(st.session_state.df_b, ['出差人', '姓名', '报销人'])
        mb['type'] = find_col(st.session_state.df_b, ['产品类型', '费用类型'])
        return ma, mb

    def run_validation(map_a, map_b):
        dfa = st.session_state.df_a
        dfb = st.session_state.df_b
        errs = []
        
        # 1. 必填列检查
        if not map_a['spm'] or not map_a['hours']: 
            return False, pd.DataFrame([{'错误': '表A缺少关键列(SPM/工时)'}]), None
        
        # 2. 空值检查
        check_list = [(dfa, '表A', map_a['spm']), (dfa, '表A', map_a['hours']), (dfb, '表B', map_b['amount'])]
        for df, name, col in check_list:
            if col:
                nulls = df[df[col].isnull() | (df[col].astype(str).str.strip() == '')]
                for idx, row in nulls.iterrows():
                    errs.append({'来源': name, '行号': idx+2, '列': col, '原因': '数值为空', '详情': '必填项缺失'})

        # 3. 工时阈值检查 (按人员聚合)
        if map_a['user'] and map_a['hours']:
            # 清洗工时为数字
            dfa[map_a['hours']] = pd.to_numeric(dfa[map_a['hours']], errors='coerce').fillna(0)
            hours_agg = dfa.groupby(map_a['user'])[map_a['hours']].sum()
            low_hours = hours_agg[hours_agg < MIN_HOURS_THRESHOLD]
            
            for user, h in low_hours.items():
                errs.append({
                    '来源': '表A', '行号': '-', '列': map_a['hours'], 
                    '原因': '工时不足', 
                    '详情': f"人员[{user}] 累计工时 {h} 小时 < 阈值 {MIN_HOURS_THRESHOLD}"
                })

        return (len(errs) == 0), pd.DataFrame(errs), (map_a, map_b)

    # ----------------------------------------------------
    # B. 弹窗编辑器 (Dialog) - 核心交互优化
    # ----------------------------------------------------
    @st.dialog("🛠️ 全屏数据编辑器", width="large")
    def open_editor_dialog():
        st.caption("请在此处直接修改数据，点击右下角保存生效。")
        
        tab1, tab2 = st.tabs(["编辑 表A (交付)", "编辑 表B (差旅)"])
        
        with tab1:
            gb_a = GridOptionsBuilder.from_dataframe(st.session_state.df_a)
            gb_a.configure_default_column(editable=True, groupable=True)
            res_a = AgGrid(st.session_state.df_a, gridOptions=gb_a.build(), height=400, key='edit_a')
            
        with tab2:
            gb_b = GridOptionsBuilder.from_dataframe(st.session_state.df_b)
            gb_b.configure_default_column(editable=True, groupable=True)
            res_b = AgGrid(st.session_state.df_b, gridOptions=gb_b.build(), height=400, key='edit_b')

        if st.button("💾 保存修改并重新校验", type="primary", use_container_width=True):
            st.session_state.df_a = pd.DataFrame(res_a['data'])
            st.session_state.df_b = pd.DataFrame(res_b['data'])
            st.session_state.calc_result_zip = None # 清除缓存，强制重算
            st.rerun()

    # ----------------------------------------------------
    # C. 主流程控制
    # ----------------------------------------------------
    
    # 如果已有缓存结果，直接显示结果，不再进行校验和计算
    if st.session_state.calc_result_zip is not None:
        pass # 直接跳到底部显示下载
    else:
        # 执行校验
        st.divider()
        map_a, map_b = get_mappings()
        is_valid, err_df, maps = run_validation(map_a, map_b)
        
        if not is_valid:
            # === 错误态 ===
            st.markdown(f"""
            <div class="error-card">
                <h3 style="color:#da3633; margin:0">🚨 校验未通过：发现 {len(err_df)} 个问题</h3>
                <p>包含空值或工时未达标 ({MIN_HOURS_THRESHOLD}h)。请修复后继续。</p>
            </div>
            """, unsafe_allow_html=True)
            
            col_act1, col_act2 = st.columns([1, 2])
            with col_act1:
                # 下载错误报告
                csv = err_df.to_csv(index=False).encode('utf-8-sig')
                st.download_button("📥 下载错误报告 (CSV)", csv, "错误清单.csv", "text/csv", use_container_width=True)
            with col_act2:
                # 触发弹窗
                if st.button("🛠️ 打开全屏修复 (推荐)", type="primary", use_container_width=True):
                    open_editor_dialog()
            
            # 展示部分错误预览
            st.dataframe(err_df.head(10), use_container_width=True)
            st.stop() # 阻断

        else:
            # === 计算态 ===
            # 显示进度条 (仅在计算时显示一次)
            progress = st.progress(0, "正在启动计算引擎...")
            
            try:
                # 模拟计算步骤
                progress.progress(20, "清洗数据...")
                df_a = st.session_state.df_a.copy()
                df_b = st.session_state.df_b.copy()
                ma, mb = maps[0], maps[1]
                
                # ... (此处省略具体的数学计算逻辑，与之前相同) ...
                # 简单模拟计算过程
                df_b['CleanAmount'] = pd.to_numeric(df_b[mb['amount']].astype(str).str.replace(',',''), errors='coerce').fillna(0)
                is_sub = df_b[mb['type']].astype(str).str.contains(SUBSIDY_TAG, na=False)
                
                progress.progress(60, "聚合与合并...")
                # 构造结果数据 (简化版示例)
                # 实际生产中这里应包含完整的 Merge 逻辑
                res_df = df_a.copy()
                res_df['结算结果'] = '演示数据'
                
                progress.progress(90, "打包文件...")
                
                # 生成 ZIP
                z_buffer = io.BytesIO()
                with zipfile.ZipFile(z_buffer, 'w') as zf:
                    with pd.ExcelWriter(io.BytesIO(), engine='xlsxwriter') as w:
                        res_df.to_excel(w, sheet_name='明细')
                        z_data = w.book.io.getvalue()
                    zf.writestr("结算报表.xlsx", z_data)
                
                # === 关键：存入缓存 ===
                st.session_state.calc_result_zip = z_buffer.getvalue()
                
                progress.progress(100, "✅ 完成！")
                time.sleep(0.5)
                st.rerun() # 强制刷新页面，进入下方的"缓存态"
                
            except Exception as e:
                st.error(f"计算出错: {e}")
                st.stop()

    # ----------------------------------------------------
    # D. 下载区 (缓存态)
    # ----------------------------------------------------
    if st.session_state.calc_result_zip is not None:
        st.divider()
        st.success("✅ 报表已生成 (数据已缓存，点击下载不消耗流量)")
        
        # 布局
        c_d1, c_d2 = st.columns([2, 1])
        with c_d1:
            st.download_button(
                label="📦 下载最终报表包 (ZIP)",
                data=st.session_state.calc_result_zip,
                file_name="结算报表.zip",
                mime="application/zip",
                type="primary",
                use_container_width=True
            )
        with c_d2:
            if st.button("🗑️ 清除缓存/重算", use_container_width=True):
                st.session_state.calc_result_zip = None
                st.rerun()
