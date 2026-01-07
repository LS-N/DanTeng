import streamlit as st
import pandas as pd
import time

# ==========================================
# 1. 全局配置与常量定义
# ==========================================

# 设置页面基础信息
st.set_page_config(
    page_title="淡藤财务报表 Pro",
    page_icon="😈",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ⚡️ 核心架构约束：结果表契约 (不可变域 Immutable Scope)
# 这部分定义了下游系统要求的“死格式”，无论前端如何操作，这部分数据结构不可动摇
LOCKED_RESULT_SCHEMA = {
    "table_name": "finance_final_settlement_v1",
    "description": "财务月度结算标准输出表",
    "required_fields": [
        {"name": "report_id", "type": "String", "desc": "报表唯一ID"},
        {"name": "account_code", "type": "Varchar(50)", "desc": "会计科目代码"},
        {"name": "total_amount", "type": "Decimal(18,2)", "desc": "结算金额"},
        {"name": "biz_date", "type": "Date", "desc": "业务日期"},
        {"name": "cost_center", "type": "String", "desc": "成本中心"}
    ]
}

# ==========================================
# 2. UI 组件与交互逻辑
# ==========================================

def main():
    st.title("😈 淡藤财务报表 Pro - 数据映射控制台")
    st.markdown("### 任务配置：源数据 -> 结果表")
    
    # 使用两列布局
    col_src, col_arrow, col_res = st.columns([4, 1, 4])

    # -------------------------------------------------
    # 左侧：源表配置 (Mutable - 允许修改)
    # -------------------------------------------------
    with col_src:
        st.info("📂 源数据配置 (输入端)")
        
        # 1. 源表类型选择 (可变)
        source_type = st.selectbox(
            "选择数据源类型",
            ["MySQL - 核心交易库", "Excel - 线下导入", "API - 支付网关"],
            help="你可以随意修改数据的来源渠道"
        )
        
        # 2. 模拟根据源类型加载的字段 (可变)
        st.write("配置源字段:")
        # 这里用 data_editor 模拟源表字段的定义，允许用户增删改
        default_src_data = pd.DataFrame([
            {"field": "id", "type": "int"},
            {"field": "money", "type": "float"},
            {"field": "create_time", "type": "datetime"}
        ])
        edited_source_fields = st.data_editor(
            default_src_data, 
            num_rows="dynamic",
            use_container_width=True,
            key="source_editor"
        )

    # -------------------------------------------------
    # 中间：映射指示器
    # -------------------------------------------------
    with col_arrow:
        st.markdown("<br><br><br><br>", unsafe_allow_html=True) # 简单的占位
        st.markdown("<h1 style='text-align: center;'>➡️</h1>", unsafe_allow_html=True)

    # -------------------------------------------------
    # 右侧：结果表配置 (Immutable - 严格锁定)
    # -------------------------------------------------
    with col_res:
        st.warning("🔒 结果表配置 (输出端 - 已锁定)")
        
        # 1. 表名 (UI 禁用)
        st.text_input(
            "目标表名 (不可改)",
            value=LOCKED_RESULT_SCHEMA["table_name"],
            disabled=True, # ⚡️ 关键：前端禁止交互
            help="系统契约锁定，无法修改"
        )
        
        # 2. 描述 (UI 禁用)
        st.text_area(
            "业务描述 (不可改)",
            value=LOCKED_RESULT_SCHEMA["description"],
            disabled=True
        )

        # 3. 字段结构展示 (只读表格)
        st.write("标准输出结构:")
        df_schema = pd.DataFrame(LOCKED_RESULT_SCHEMA["required_fields"])
        st.dataframe(
            df_schema, 
            hide_index=True, 
            use_container_width=True,
            # Streamlit 的 dataframe 默认就是只读展示，除非用 data_editor
        )

    st.markdown("---")

    # ==========================================
    # 3. 提交与保存逻辑 (后端清洗)
    # ==========================================
    
    # 模拟保存按钮
    if st.button("🚀 保存并生成任务 (去吧皮卡丘)", type="primary"):
        with st.spinner("正在校验架构契约..."):
            time.sleep(1) # 模拟处理耗时
            
            # --- ⚡️ 核心逻辑清洗层 ---
            # 无论前端传来了什么临时状态，我们构造最终配置时
            # 强制使用 LOCKED_RESULT_SCHEMA
            
            final_config = {
                "task_id": "TASK_20260107_001",
                # 源端：取自用户输入
                "source_config": {
                    "type": source_type,
                    "fields_defined": edited_source_fields.to_dict(orient="records")
                },
                # 目标端：强制覆盖，忽略任何可能的篡改
                "target_config": LOCKED_RESULT_SCHEMA
            }
            
            # 成功反馈
            st.success("✅ 配置保存成功！结果表结构已通过契约校验。")
            
            # 打印最终生成的 JSON，证明逻辑生效
            with st.expander("查看生成的最终配置 JSON"):
                st.json(final_config)

# 运行主程序
if __name__ == "__main__":
    main()
