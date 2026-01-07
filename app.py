import streamlit as st
import pandas as pd

# -----------------------------------------------------------------------------
# 1. 核心工具函数：HTML 居中渲染器 (解决"结果表居中且不可改"的问题)
# -----------------------------------------------------------------------------
def render_centered_table(df: pd.DataFrame):
    """
    将 DataFrame 转换为带有自定义 CSS 的 HTML，实现：
    1. 页面居中显示
    2. 样式美化 (边框、表头背景)
    3. 纯静态展示 (Read-only)
    """
    # 将数据转换为 HTML，不带索引
    table_html = df.to_html(index=False, classes="styled-table")

    # 定义 CSS 样式
    custom_css = """
    <style>
        /* 外层容器：使用 Flexbox 实现水平居中 */
        .table-container {
            display: flex;
            justify-content: center;
            margin-top: 20px;
            margin-bottom: 40px;
        }
        
        /* 表格样式 */
        .styled-table {
            border-collapse: collapse;
            margin: 25px 0;
            font-size: 0.9em;
            font-family: sans-serif;
            min-width: 600px; /* 设定一个最小宽度，保证气势 */
            box-shadow: 0 0 20px rgba(0, 0, 0, 0.15);
            border-radius: 8px; /* 圆角表格 */
            overflow: hidden;
        }
        
        /* 表头样式 */
        .styled-table thead tr {
            background-color: #009879; /* 经典的 Excel 绿或深色 */
            color: #ffffff;
            text-align: center;
        }
        
        /* 单元格样式 */
        .styled-table th,
        .styled-table td {
            padding: 12px 15px;
            text-align: center; /* 内容居中 */
            border-bottom: 1px solid #dddddd;
        }
        
        /* 偶数行变色 (斑马纹) */
        .styled-table tbody tr:nth-of-type(even) {
            background-color: #f3f3f3;
        }
        
        /* 最后一行边框 */
        .styled-table tbody tr:last-of-type {
            border-bottom: 2px solid #009879;
        }
        
        /* 浅色模式适配文字颜色 */
        body {
            color: #333;
        }
    </style>
    """

    # 组合 HTML：CSS + 容器 div + 表格
    final_html = f"{custom_css}<div class='table-container'>{table_html}</div>"
    
    # 渲染
    st.markdown(final_html, unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# 2. 页面主逻辑
# -----------------------------------------------------------------------------
def main():
    st.set_page_config(page_title="淡藤财务报表 Pro", layout="wide", page_icon="📊")

    st.title("📊 淡藤财务报表 Pro - 配置与生成")
    st.caption("架构模式：源表配置可编辑 (Data Editor) | 结果报表只读居中 (HTML/CSS)")

    st.divider()

    # --- A. 数据准备 (模拟 Session State) ---
    if "config_data" not in st.session_state:
        # 初始化配置数据
        st.session_state.config_data = pd.DataFrame([
            {"目标列名": "营业收入", "源表类型": "利润表", "匹配字段": "revenue_total", "计算逻辑": "直接取值"},
            {"目标列名": "营业成本", "源表类型": "利润表", "匹配字段": "cost_operating", "计算逻辑": "直接取值"},
            {"目标列名": "销售费用", "源表类型": "费用明细表", "匹配字段": "expense_sales", "计算逻辑": "聚合求和"},
            {"目标列名": "净利润",   "源表类型": "利润表", "匹配字段": "net_profit", "计算逻辑": "公式计算"},
        ])

    # --- B. 配置区域 (Editable) ---
    st.subheader("🛠️ 1. 报表映射配置 (可修改)")
    st.info("💡 说明：请在下方表格中修改 **源表类型** 和 **匹配字段**。修改后点击“生成报表”即可生效。")

    # 使用 st.data_editor 允许用户修改
    # column_config 用于增强体验，比如把“源表类型”变成下拉框
    edited_df = st.data_editor(
        st.session_state.config_data,
        use_container_width=True,
        num_rows="dynamic",
        column_config={
            "源表类型": st.column_config.SelectboxColumn(
                "源表类型 (点击选择)",
                help="选择该字段数据来源的表",
                width="medium",
                options=[
                    "利润表",
                    "资产负债表",
                    "现金流量表",
                    "费用明细表",
                ],
                required=True,
            ),
            "匹配字段": st.column_config.TextColumn(
                "匹配字段 (可编辑)",
                help="输入数据库中的原始字段名",
                width="medium",
                validate="^[a-z_]+$", # 简单的正则验证，只允许小写字母和下划线
            ),
            "目标列名": st.column_config.TextColumn(
                "目标列名 (只读)",
                disabled=True # 锁定目标列名，不让用户改
            )
        },
        hide_index=True,
    )

    # 将修改后的数据写回 session_state (以便在其他地方使用)
    st.session_state.config_data = edited_df

    st.divider()

    # --- C. 动作区域 ---
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        generate_btn = st.button("🚀 生成 / 刷新 预览报表", use_container_width=True, type="primary")

    # --- D. 结果展示区域 (Read-Only & Centered) ---
    if generate_btn:
        st.subheader("📈 2. 最终报表预览 (居中·只读)")
        
        # 模拟生成逻辑：根据上面的配置，生成一个假的“结果数据”
        # 这里仅做演示，实际逻辑应调用你的后端计算函数
        result_data = {
            "项目 (Item)": edited_df["目标列名"].tolist(),
            "本期金额 (Current Period)": ["¥ 1,200,000.00", "¥ 800,000.00", "¥ 50,000.00", "¥ 350,000.00"],
            "上期金额 (Last Period)": ["¥ 1,000,000.00", "¥ 750,000.00", "¥ 45,000.00", "¥ 205,000.00"],
            "同比 (YoY)": ["+20%", "+6.6%", "+11.1%", "+70.7%"]
        }
        result_df = pd.DataFrame(result_data)

        # 核心调用：使用我们定义的 HTML 渲染函数
        render_centered_table(result_df)
        
        st.success("✅ 报表已根据最新配置重新生成！")

if __name__ == "__main__":
    main()
