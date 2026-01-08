import streamlit as st
import pandas as pd
import io
import zipfile
from docx import Document

# --- 1. 模拟生成 Word 的函数 (接收不同参数) ---
def generate_word_for_person(name, amount):
    doc = Document()
    doc.add_heading(f'{name} 的结算单', 0)
    doc.add_paragraph(f"尊敬的 {name}：")
    doc.add_paragraph(f"您本月的结算金额为：{amount} 元。")
    doc.add_paragraph("请确认无误。")
    
    output = io.BytesIO()
    doc.save(output)
    return output.getvalue()

# --- 2. 模拟 Excel 生成 (保持不变) ---
def to_excel_bytes(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()

# --- 3. ZIP 打包函数 (保持不变) ---
def create_zip_bytes(files_dict):
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for file_name, file_data in files_dict.items():
            zf.writestr(file_name, file_data)
    return zip_buffer.getvalue()

# ================= 业务逻辑开始 =================

st.title("动态报表生成器")

# 1. 准备 Excel 数据 (固定的 3 个表)
excel_files = {
    "表1_工时统计.xlsx": to_excel_bytes(pd.DataFrame({"A": [1]})),
    "表2_结算汇总.xlsx": to_excel_bytes(pd.DataFrame({"B": [2]})),
    "表3_详细明细.xlsx": to_excel_bytes(pd.DataFrame({"C": [3]}))
}

# 2. 动态生成 Word 数据 (比如根据员工名单，生成 N 个)
# 假设这是你的业务数据来源
employee_list = [
    {"name": "张三", "money": 5000},
    {"name": "李四", "money": 6200},
    {"name": "王五", "money": 4800},
    {"name": "赵六", "money": 7000},
    # ... 哪怕这里有 100 个人也没问题
]

word_files = {}
for emp in employee_list:
    # 动态生成文件名和内容
    file_name = f"结算单_{emp['name']}.docx"
    file_content = generate_word_for_person(emp['name'], emp['money'])
    word_files[file_name] = file_content

# 3. 合并所有文件 (用于 ZIP 打包)
# 语法说明：**字典1, **字典2 是 Python 合并字典的快捷写法
all_download_files = {**excel_files, **word_files}

# 4. 生成 ZIP
zip_data = create_zip_bytes(all_download_files)

st.success(f"✅ 生成完毕！共生成 {len(all_download_files)} 个文件 (3个Excel + {len(word_files)}个Word)")

# ================= UI 展示部分 =================

# --- 区域1: 批量下载 (核心功能) ---
st.download_button(
    label=f"📦 批量下载所有文件 (共{len(all_download_files)}个)",
    data=zip_data,
    file_name="所有结算资料打包.zip",
    mime="application/zip",
    type="primary",
    use_container_width=True
)

st.divider()

# --- 区域2: Excel 单独下载 (固定布局) ---
st.subheader("📊 基础数据表 (Excel)")
cols = st.columns(3)
keys = list(excel_files.keys())
# 遍历展示 3 个 Excel 按钮
for i, col in enumerate(cols):
    with col:
        st.download_button(
            label=f"📥 {keys[i]}",
            data=excel_files[keys[i]],
            file_name=keys[i],
            use_container_width=True
        )

# --- 区域3: Word 动态单独下载 (处理 N 个文件) ---
st.subheader(f"📝 个人结算单 (Word - 共{len(word_files)}个)")

# 方式 A：使用折叠面板 (Expander) - 推荐，省空间
with st.expander("点击展开查看所有结算单列表", expanded=True):
    # 这里我们遍历刚刚生成的 word_files 字典
    for name, data in word_files.items():
        # 使用两列布局：左边显示文件名，右边显示下载按钮
        c1, c2 = st.columns([3, 1]) 
        with c1:
            st.text(f"📄 {name}") # 仅显示文件名文本
        with c2:
            st.download_button(
                label="下载",
                data=data,
                file_name=name,
                key=f"btn_{name}", # 必须设置唯一的 key
                use_container_width=True
            )

# 方式 B (备选)：如果文件实在太多(比如100个)，建议用下拉框选择下载
# selected_file = st.selectbox("选择要下载的结算单", list(word_files.keys()))
# st.download_button("下载选中的文件", data=word_files[selected_file], file_name=selected_file)
