import streamlit as st
import pandas as pd
import io

# ==========================================
# 1. 基础配置
# ==========================================
st.set_page_config(page_title="财务自动化系统", layout="wide")
st.title("🛠️ 财务自动化结算系统 (调试版)")

# ==========================================
# 2. 文件上传
# ==========================================
st.info("请上传文件，系统将自动读取列名供你确认。")
col1, col2 = st.columns(2)
with col1:
    file_a = st.file_uploader("步骤1：上传表A (交付明细)", type=["xlsx", "xls", "csv"], key="a")
with col2:
    file_b = st.file_uploader("步骤2：上传表B (实施差旅)", type=["xlsx", "xls", "csv"], key="b")

# ==========================================
# 3. 核心逻辑
# ==========================================
def load_data(file):
    try:
        if file.name.endswith('.csv'):
            return pd.read_csv(file)
        else:
            return pd.read_excel(file)
    except Exception as e:
        # 如果UTF-8失败，尝试GBK
        try:
            return pd.read_csv(file, encoding='gbk')
        except:
            st.error(f"无法读取文件 {file.name}，请检查格式。")
            return None

if file_a and file_b:
    df_a = load_data(file_a)
    df_b = load_data(file_b)

    if df_a is not None and df_b is not None:
        st.divider()
        st.subheader("⚙️ 关键字段映射 (请确认)")
        
        # --- 这里的下拉框解决了字段名对不上的问题 ---
        c1, c2, c3, c4 = st.columns(4)
        
        # 自动猜测默认值
        default_spm_a = next((c for c in df_a.columns if 'SPM' in str(c).upper()), df_a.columns[0])
        default_work_a = next((c for c in df_a.columns if '工时' in str(c)), df_a.columns[0])
        default_user_a = next((c for c in df_a.columns if '人员' in str(c) or '姓名' in str(c)), df_a.columns[0])
        
        default_spm_b = next((c for c in df_b.columns if 'SPM' in str(c).upper()), df_b.columns[0])
        
        with c1:
            col_user_a = st.selectbox("表A-人员列", df_a.columns, index=list(df_a.columns).index(default_user_a))
        with c2:
            col_spm_a = st.selectbox("表A-SPM列", df_a.columns, index=list(df_a.columns).index(default_spm_a))
        with c3:
            col_work_a = st.selectbox("表A-工时列", df_a.columns, index=list(df_a.columns).index(default_work_a))
        with c4:
            col_spm_b = st.selectbox("表B-SPM列", df_b.columns, index=list(df_b.columns).index(default_spm_b))

        price = st.number_input("人力单价 (元/天)", value=1500)

        # --- 开始计算 ---
        if st.button("🚀 开始运行逻辑", type="primary"):
            try:
                # 1. 清洗 A 表
                # 必须有的列
                req_cols_a = [col_user_a, col_spm_a, col_work_a]
                # 可选列(用于展示)
                opt_cols_a = ['人事范围', '合同主体', '销售部门', '销售', '项目']
                # 实际存在的列
                real_opt_cols = [c for c in opt_cols_a if c in df_a.columns]
                
                # 聚合规则
                agg_dict = {col_work_a: 'sum'}
                for c in real_opt_cols:
                    agg_dict[c] = 'first'
                
                # 去除空SPM
                df_a_clean = df_a.dropna(subset=[col_spm_a]).copy()
                # 聚合
                df_a_gp = df_a_clean.groupby([col_user_a, col_spm_a], as_index=False).agg(agg_dict)
                
                # 2. 清洗 B 表
                # 确保有产品类型和金额
                if '产品类型' not in df_b.columns or '金额' not in df_b.columns:
                    st.error("表B 缺少 '产品类型' 或 '金额' 列，请检查。")
                    st.stop()
                    
                df_b_clean = df_b.dropna(subset=[col_spm_b]).copy()
                # 假设表B的人员列叫'出差人'
                col_user_b = '出差人' if '出差人' in df_b.columns else df_b.columns[0]
                
                # 拆分补助
                mask_sub = df_b_clean['产品类型'] == '差旅补助'
                df_b_sub = df_b_clean[mask_sub].groupby([col_user_b, col_spm_b])['金额'].sum().reset_index()
                df_b_sub.rename(columns={'金额': '差旅补助'}, inplace=True)
                
                # 拆分费控
                df_b_fee = df_b_clean[~mask_sub].groupby([col_user_b, col_spm_b])['金额'].sum().reset_index()
                df_b_fee.rename(columns={'金额': '差旅费控平台'}, inplace=True)
                
                # 3. 关联
                # A join B_sub
                merged = pd.merge(df_a_gp, df_b_sub, 
                                  left_on=[col_user_a, col_spm_a], 
                                  right_on=[col_user_b, col_spm_b], how='left')
                # join B_fee
                merged = pd.merge(merged, df_b_fee, 
                                  left_on=[col_user_a, col_spm_a], 
                                  right_on=[col_user_b, col_spm_b], how='left')
                
                # 4. 计算
                merged['差旅补助'] = merged['差旅补助'].fillna(0)
                merged['差旅费控平台'] = merged['差旅费控平台'].fillna(0)
                merged['支持时间'] = merged[col_work_a] / 8
                merged['人力费用'] = merged['支持时间'] * price
                merged['结算费用合计'] = merged['人力费用'] + merged['差旅补助'] + merged['差旅费控平台']
                
                st.success("✅ 计算成功！")
                st.dataframe(merged.head())
                
                # 5. 下载
                def to_excel(df):
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                        df.to_excel(writer, index=False)
                    return output.getvalue()

                st.download_button("📥 下载结果表3 (明细)", to_excel(merged), "结果表3.xlsx")

            except Exception as e:
                st.error(f"❌ 运行出错: {str(e)}")
                # 打印详细错误方便你发给我看
                import traceback
                st.text(traceback.format_exc())
