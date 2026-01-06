import streamlit as st
import pandas as pd
import io
import time

# ==========================================
# 1. 系统配置与界面设计
# ==========================================
st.set_page_config(page_title="自动化财务结算系统", layout="wide")

st.title("📊 自动化财务工时结算系统")
st.markdown("### 自动关联交付明细(表A)与差旅报销(表B)，生成结算报表")

# 侧边栏：全局参数配置
st.sidebar.header("⚙️ 参数配置 (工作流)")
with st.sidebar:
    price_per_day = st.number_input("人力单价 (元/天)", value=1500, step=100)
    spm_col_a = st.text_input("表A 关联键 (SPM)", value="SPM")
    spm_col_b = st.text_input("表B 关联键 (SPM)", value="SPM")
    subsidy_tag = st.text_input("差旅补助标识", value="差旅补助")

# ==========================================
# 2. 文件上传区
# ==========================================
col1, col2 = st.columns(2)

with col1:
    st.info("步骤 1: 上传交付明细 (表A)")
    file_a = st.file_uploader("请上传表A (支持 Excel/CSV)", type=["xlsx", "xls", "csv"], key="a")

with col2:
    st.info("步骤 2: 上传实施差旅 (表B)")
    file_b = st.file_uploader("请上传表B (需包含 SPM 列)", type=["xlsx", "xls", "csv"], key="b")

# ==========================================
# 3. 核心处理逻辑 (后端引擎)
# ==========================================
def load_file(file):
    if file.name.endswith('.csv'):
        try:
            return pd.read_csv(file)
        except:
            return pd.read_csv(file, encoding='gbk')
    else:
        return pd.read_excel(file)

def process_data(df_a, df_b):
    # --- 1. 清洗表 A ---
    # 确保没有空的 SPM
    df_a_clean = df_a.dropna(subset=[spm_col_a]).copy()
    
    # 定义聚合规则
    # 注意：这里的字段名必须与你真实的 Excel 表头一致，如果不一致会报错
    # 假设表头标准为：人员, SPM, 交付工时（h）, 人事范围, 合同主体, 销售部门
    agg_rules = {
        '交付工时（h）': 'sum',
        '人事范围': 'first',
        '合同主体': 'first',
        '销售部门': 'first',
        '销售': 'first',
        '项目': 'first'
    }
    
    # 容错处理：如果列不存在，就不聚合该列
    valid_agg = {k: v for k, v in agg_rules.items() if k in df_a_clean.columns}
    
    # 核心聚合：按 人员 + SPM
    df_a_grouped = df_a_clean.groupby(['人员', spm_col_a], as_index=False).agg(valid_agg)
    
    # --- 2. 清洗表 B ---
    # 必须确保表 B 有 SPM 列
    if spm_col_b not in df_b.columns:
        return None, f"错误：表B中找不到 '{spm_col_b}' 列，请先手动添加该列再上传。"
        
    # 拆分补助与费控
    # 假设 '出差人' 对应表 A 的 '人员'
    group_keys = ['出差人', spm_col_b]
    
    # 补助
    df_b_sub = df_b[df_b['产品类型'] == subsidy_tag].groupby(group_keys)['金额'].sum().reset_index()
    df_b_sub.rename(columns={'金额': '差旅补助'}, inplace=True)
    
    # 费控
    df_b_fee = df_b[df_b['产品类型'] != subsidy_tag].groupby(group_keys)['金额'].sum().reset_index()
    df_b_fee.rename(columns={'金额': '差旅费控平台'}, inplace=True)

    # --- 3. 关联 (Left Join) ---
    merged = pd.merge(df_a_grouped, df_b_sub, 
                      left_on=['人员', spm_col_a], right_on=['出差人', spm_col_b], how='left')
    merged = pd.merge(merged, df_b_fee, 
                      left_on=['人员', spm_col_a], right_on=['出差人', spm_col_b], how='left')

    # 填充 0
    merged['差旅补助'] = merged['差旅补助'].fillna(0)
    merged['差旅费控平台'] = merged['差旅费控平台'].fillna(0)

    # --- 4. 算钱 ---
    merged['支持时间（人天）'] = merged['交付工时（h）'] / 8
    merged['人力费用'] = merged['支持时间（人天）'] * price_per_day
    merged['结算费用合计'] = merged['人力费用'] + merged['差旅补助'] + merged['差旅费控平台']

    return merged, "SUCCESS"

# ==========================================
# 4. 执行与展示
# ==========================================
if file_a and file_b:
    if st.button("🚀 开始自动化计算", type="primary"):
        with st.spinner("正在清洗数据、关联表单、计算费用..."):
            try:
                df_a_raw = load_file(file_a)
                df_b_raw = load_file(file_b)
                
                result_df, msg = process_data(df_a_raw, df_b_raw)
                
                if msg != "SUCCESS":
                    st.error(msg)
                else:
                    st.success("✅ 计算完成！数据已生成。")
                    
                    # --- 生成三张表 ---
                    
                    # 表 3: 明细表
                    st.subheader("📝 结果表 3: 结算明细预览")
                    # 整理字段显示
                    cols_show = ['人员', '项目', 'SPM', '人力费用', '差旅补助', '结算费用合计']
                    st.dataframe(result_df[cols_show].head())
                    
                    # 表 1: 人员汇总
                    df_r1 = result_df.groupby('人员')['交付工时（h）'].sum().reset_index()
                    df_r1.rename(columns={'交付工时（h）': '项目工时'}, inplace=True)
                    
                    # 表 2: 结算汇总
                    # 容错：确保有人事范围等字段
                    group_cols = [c for c in ['人事范围', '合同主体', '销售部门'] if c in result_df.columns]
                    if group_cols:
                        df_r2 = result_df.groupby(group_cols)[['结算费用合计', '支持时间（人天）']].sum().reset_index()
                        df_r2.columns = ['销售公司', '采购公司', '采购部门', '金额（含税）', '工作量（人天）']
                    else:
                        df_r2 = pd.DataFrame() # 空表防报错

                    # --- 下载区 ---
                    st.write("---")
                    c1, c2, c3 = st.columns(3)
                    
                    # 导出函数
                    def to_excel(df):
                        output = io.BytesIO()S
                        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                            df.to_excel(writer, index=False)
                        return output.getvalue()

                    with c1:
                        st.download_button("📥 下载 结果表1 (工时统计)", data=to_excel(df_r1), file_name="结果表1.xlsx")
                    with c2:
                        st.download_button("📥 下载 结果表2 (结算汇总)", data=to_excel(df_r2), file_name="结果表2.xlsx")
                    with c3:
                        st.download_button("📥 下载 结果表3 (完整明细)", data=to_excel(result_df), file_name="结果表3.xlsx")

            except Exception as e:
                st.error(f"运行出错: {str(e)}")
                st.warning("请检查上传的 Excel 表头名称是否与代码中的字段匹配。")

else:
    st.info("👋 请在上方上传两个文件以开始。")