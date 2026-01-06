import streamlit as st
import pandas as pd
import io

# ==========================================
# 1. 页面初始化
# ==========================================
st.set_page_config(page_title="财务自动化结算系统 Pro", layout="wide", page_icon="📊")
st.title("📊 财务自动化结算系统 Pro")
st.markdown("### 自定义字段映射 + 智能校验 + 自动化计算")
st.divider()

# ==========================================
# 2. 侧边栏：全局参数
# ==========================================
st.sidebar.header("⚙️ 全局参数")
PRICE_PER_DAY = st.sidebar.number_input("人力单价 (元/天)", value=1500, step=100)
SUBSIDY_TAG = st.sidebar.text_input("差旅补助标识", value="差旅补助")

# ==========================================
# 3. 文件上传模块
# ==========================================
col1, col2 = st.columns(2)
with col1:
    st.info("步骤 1: 上传 交付明细 (表A)")
    file_a = st.file_uploader("支持 Excel/CSV", type=["xlsx", "xls", "csv"], key="a")
with col2:
    st.info("步骤 2: 上传 实施差旅 (表B)")
    file_b = st.file_uploader("需包含 SPM 列", type=["xlsx", "xls", "csv"], key="b")

# 定义读取函数
def load_file(file):
    if not file: return None
    try:
        if file.name.endswith('.csv'):
            try: return pd.read_csv(file)
            except: return pd.read_csv(file, encoding='gbk')
        else: return pd.read_excel(file)
    except Exception as e:
        st.error(f"读取文件失败: {e}")
        return None

df_a_raw = load_file(file_a)
df_b_raw = load_file(file_b)

# ==========================================
# 4. 字段映射模块 (核心新增)
# ==========================================
if df_a_raw is not None and df_b_raw is not None:
    st.divider()
    st.subheader("🛠️ 字段映射配置")
    st.info("系统已自动识别表头。如果识别不准确，请手动调整下拉菜单。")

    # 去除列名空格
    df_a_raw.columns = [str(c).strip() for c in df_a_raw.columns]
    df_b_raw.columns = [str(c).strip() for c in df_b_raw.columns]

    # --- 定义标准字段需求 ---
    # 格式: '标准键': {'label': '显示名称', 'default': ['可能的列名1', '可能的列名2']}
    
    # 表A 映射配置
    cols_a = list(df_a_raw.columns)
    map_a_config = {
        'user': {'label': '人员姓名', 'default': ['人员', '姓名', 'User']},
        'spm': {'label': 'SPM编号', 'default': ['SPM', '项目编号', '标识符']},
        'hours': {'label': '交付工时', 'default': ['交付工时', '交付工时（h）', '工时', '时长']},
        'project': {'label': '项目名称', 'default': ['项目', '所属项目', '项目名']},
        'range': {'label': '人事范围', 'default': ['人事范围', '公司范围']},
        'contract': {'label': '合同主体', 'default': ['合同主体', '签约主体']},
        'sales_dept': {'label': '销售部门', 'default': ['销售部门', '部门']},
        'sales_user': {'label': '销售人员', 'default': ['销售', '销售人员']}
    }

    # 表B 映射配置
    cols_b = list(df_b_raw.columns)
    map_b_config = {
        'user': {'label': '出差人姓名', 'default': ['出差人', '姓名', '人员']},
        'spm': {'label': 'SPM编号', 'default': ['SPM', '项目编号']},
        'amount': {'label': '金额', 'default': ['金额', '总金额']},
        'type': {'label': '产品类型', 'default': ['产品类型', '费用类型']}
    }

    # 渲染映射界面
    c1, c2 = st.columns(2)
    
    # 辅助函数: 自动匹配默认值
    def get_default_index(options, defaults):
        for d in defaults:
            if d in options: return options.index(d)
        return 0

    mapping_a = {}
    with c1:
        st.markdown("#### 表A (交付明细) 映射")
        for key, cfg in map_a_config.items():
            idx = get_default_index(cols_a, cfg['default'])
            mapping_a[key] = st.selectbox(
                f"[{cfg['label']}] 对应列:", cols_a, index=idx, key=f"a_{key}"
            )

    mapping_b = {}
    with c2:
        st.markdown("#### 表B (实施差旅) 映射")
        for key, cfg in map_b_config.items():
            idx = get_default_index(cols_b, cfg['default'])
            mapping_b[key] = st.selectbox(
                f"[{cfg['label']}] 对应列:", cols_b, index=idx, key=f"b_{key}"
            )

    # ==========================================
    # 5. 校验与计算模块
    # ==========================================
    st.divider()
    if st.button("🚀 开始校验并生成报表", type="primary", use_container_width=True):
        
        # --- 步骤 1: 数据校验 (Validate) ---
        errors = []
        
        # 1.1 检查表B的 SPM 是否有空值
        col_spm_b = mapping_b['spm']
        if df_b_raw[col_spm_b].isnull().any():
            missing_count = df_b_raw[col_spm_b].isnull().sum()
            errors.append(f"❌ 校验失败: 表B的 [{col_spm_b}] 列存在 {missing_count} 个空值。请补全后重试。")
            
        # 1.2 检查表A的 人员/SPM/工时 是否有空值 (关键字段)
        for k in ['user', 'spm', 'hours']:
            col_name = mapping_a[k]
            if df_a_raw[col_name].isnull().any():
                 errors.append(f"⚠️ 警告: 表A的 [{col_name}] 列存在空值，这些行将在计算时被忽略。")

        if any("❌" in e for e in errors):
            for e in errors: st.error(e)
            st.stop()
        elif errors:
            for e in errors: st.warning(e)

        # --- 步骤 2: 数据标准化 (Transform) ---
        with st.spinner("正在执行清洗与计算逻辑..."):
            try:
                # 2.1 提取表A数据并重命名为标准键
                # 聚合规则: 工时求和，其他取第一条
                agg_rules = {mapping_a['hours']: 'sum'}
                meta_cols = ['project', 'range', 'contract', 'sales_dept', 'sales_user']
                for k in meta_cols:
                    agg_rules[mapping_a[k]] = 'first'
                
                # 过滤无SPM数据
                df_a_clean = df_a_raw.dropna(subset=[mapping_a['spm']]).copy()
                
                # 按 人员+SPM 分组
                df_a_gp = df_a_clean.groupby([mapping_a['user'], mapping_a['spm']], as_index=False).agg(agg_rules)
                
                # 2.2 提取表B数据
                df_b_clean = df_b_raw.dropna(subset=[mapping_b['spm']]).copy()
                
                # 拆分补助与费控
                # 统一列名方便 merge
                b_user = mapping_b['user']
                b_spm = mapping_b['spm']
                b_amt = mapping_b['amount']
                b_type = mapping_b['type']
                
                # 补助
                mask_sub = df_b_clean[b_type] == SUBSIDY_TAG
                df_sub = df_b_clean[mask_sub].groupby([b_user, b_spm])[b_amt].sum().reset_index(name='差旅补助')
                
                # 费控
                df_fee = df_b_clean[~mask_sub].groupby([b_user, b_spm])[b_amt].sum().reset_index(name='差旅费控平台')

                # --- 步骤 3: 关联与计算 ---
                # Left Join: A为主
                key_a = [mapping_a['user'], mapping_a['spm']]
                key_b = [b_user, b_spm]
                
                merged = pd.merge(df_a_gp, df_sub, left_on=key_a, right_on=key_b, how='left')
                merged = pd.merge(merged, df_fee, left_on=key_a, right_on=key_b, how='left')
                
                # 填充0
                merged[['差旅补助', '差旅费控平台']] = merged[['差旅补助', '差旅费控平台']].fillna(0)
                
                # 算钱
                col_hours = mapping_a['hours']
                merged['支持时间（人天）'] = merged[col_hours] / 8
                merged['人力费用'] = merged['支持时间（人天）'] * PRICE_PER_DAY
                merged['结算费用合计'] = merged['人力费用'] + merged['差旅补助'] + merged['差旅费控平台']

                # --- 步骤 4: 生成结果表 (Load) ---
                
                # >>> 生成 表3: 结算工时总表 <<<
                # 重命名为模板要求的名字
                rename_map = {
                    mapping_a['user']: '人员',
                    mapping_a['project']: '所属项目',
                    mapping_a['range']: '人事范围',
                    mapping_a['spm']: 'SPM',
                    mapping_a['contract']: '合同主体',
                    mapping_a['sales_user']: '销售人员',
                    mapping_a['sales_dept']: '销售部门',
                    mapping_a['hours']: '耗时（小时）'
                }
                table3 = merged.rename(columns=rename_map)
                
                # 整理列顺序
                cols_order = [
                    '人员', '所属项目', '人事范围', 'SPM', '合同主体', '销售人员', '销售部门',
                    '差旅补助', '差旅费控平台', '耗时（小时）', '支持时间（人天）', '人力费用', '结算费用合计'
                ]
                # 容错：只保留存在的列
                final_cols = [c for c in cols_order if c in table3.columns]
                table3 = table3[final_cols]
                table3.insert(0, '序号', range(1, len(table3)+1))

                # >>> 生成 表2: 采销结算单 <<<
                # 分组: 人事范围(销售公司) + 合同主体(采购公司) + 销售部门(采购部门)
                grp_cols = ['人事范围', '合同主体', '销售部门']
                # 确保这些列存在
                valid_grp = [c for c in grp_cols if c in table3.columns]
                
                if valid_grp:
                    table2 = table3.groupby(valid_grp).agg({
                        '结算费用合计': 'sum',
                        '支持时间（人天）': 'sum'
                    }).reset_index()
                    table2.columns = ['销售公司', '采购公司', '采购部门', '金额（含税，单位：元）', '工作量（人天）']
                    table2['备注'] = ''
                    table2.insert(0, '序号', range(1, len(table2)+1))
                else:
                    table2 = pd.DataFrame({'错误': ['缺少必要的聚合列，无法生成表2']})

                # >>> 生成 表1: 工时统计 <<<
                table1 = table3.groupby('人员')['耗时（小时）'].sum().reset_index()
                table1.rename(columns={'耗时（小时）': '项目工时'}, inplace=True)
                table1.insert(0, '序号', range(1, len(table1)+1))

                # --- 步骤 5: 展示与下载 ---
                st.success("✅ 计算成功！")
                
                tab1, tab2, tab3 = st.tabs(["结果表3 (明细)", "结果表2 (结算)", "结果表1 (工时)"])
                
                def to_excel(df):
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                        df.to_excel(writer, index=False)
                    return output.getvalue()

                with tab1:
                    st.dataframe(table3)
                    st.download_button("📥 下载 结果表3.xlsx", to_excel(table3), "结果表3-结算工时总表.xlsx")
                
                with tab2:
                    st.dataframe(table2)
                    st.download_button("📥 下载 结果表2.xlsx", to_excel(table2), "结果表2-采销结算单.xlsx")
                    
                with tab3:
                    st.dataframe(table1)
                    st.download_button("📥 下载 结果表1.xlsx", to_excel(table1), "结果表1-工时统计.xlsx")

            except Exception as e:
                import traceback
                st.error(f"处理过程中发生错误: {str(e)}")
                st.code(traceback.format_exc())

else:
    st.info("👋 请先上传两个文件以开始配置映射。")
