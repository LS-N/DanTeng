class WordGenerator:
    @staticmethod
    def set_cell_style(cell, text, font_size=10, bold=False, align="center"):
        """
        核心辅助函数：强制设置单元格内的字体（中英文混排）、大小、对齐
        解决默认字体不是宋体的问题
        """
        # 清除旧内容
        cell.text = ""
        paragraph = cell.paragraphs[0]
        
        # 设置对齐
        if align == "center":
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        elif align == "left":
            paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
        elif align == "right":
            paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT

        run = paragraph.add_run(str(text))
        run.font.bold = bold
        run.font.size = Pt(font_size)
        run.font.name = 'Times New Roman'
        # 强制设置中文字体为宋体
        run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
        
        # 垂直居中
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    @staticmethod
    def create_hardcoded_template(purchase_comp, sales_comp, dept_name, period_text):
        doc = docx.Document()
        
        # 1. 页面设置 (可选：调整页边距以防表格太宽)
        section = doc.sections[0]
        section.top_margin = Cm(2.0)
        section.bottom_margin = Cm(2.0)
        section.left_margin = Cm(2.0)
        section.right_margin = Cm(2.0)

        # 2. 大标题
        # 修正：移除"与"前后的空格，匹配模板 
        # 格式：公司A与公司B,2025年第三季度...
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(f"{purchase_comp}与{sales_comp},{period_text}项目交付与运维费用结算账单")
        run.font.name = 'Times New Roman'
        run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
        run.font.size = Pt(16)
        run.font.bold = True
        
        doc.add_paragraph() # 空一行

        # 3. 绘制【汇总表】
        # 修正：改为 6 行 10 列 (原代码是7行，导致多余空行)
        # 行0: 表头1
        # 行1: 表头2
        # 行2: 数据行
        # 行3: (空行 - 间隔)
        # 行4: 部门行
        # 行5: 签字行
        table0 = doc.add_table(rows=6, cols=10)
        table0.style = 'Table Grid'
        table0.autofit = False 
        # 可以根据需要手动锁定列宽，这里暂且自适应

        # --- 设置表头第一层 (行0) ---
        r0 = table0.rows[0]
        # 合并0,1,2 -> 工作量
        c0 = r0.cells[0].merge(r0.cells[2])
        WordGenerator.set_cell_style(c0, "工作量 （单位：人/天）", bold=True)
        
        # 合并3,4,5 -> 人力费用
        c1 = r0.cells[3].merge(r0.cells[5])
        WordGenerator.set_cell_style(c1, "人力费用 （单位：元）", bold=True)
        
        # 合并6,7 -> 差旅费用
        c2 = r0.cells[6].merge(r0.cells[7])
        WordGenerator.set_cell_style(c2, "差旅费用 （单位：元）", bold=True)
        
        # 合并8,9 -> 合计
        c3 = r0.cells[8].merge(r0.cells[9])
        WordGenerator.set_cell_style(c3, "合计", bold=True)

        # --- 设置表头第二层 (行1) ---
        headers = [
            "项目标准交付", "项目数据治理", "项目运维服务", 
            "项目标准交付", "项目数据治理", "项目运维服务", 
            "差旅补助", "商旅平台费用", "工作量", "合计费用（单位：元）"
        ]
        # 注意：模板最后一列是“合计费用（单位：元）” 
        for i, h in enumerate(headers):
            WordGenerator.set_cell_style(table0.rows[1].cells[i], h, font_size=9)

        # --- 数据行 (行2) 预留，稍后填充 ---
        
        # --- 间隔行 (行3) --- 保持为空
        
        # --- 部门行 (行4) ---
        r4 = table0.rows[4]
        WordGenerator.set_cell_style(r4.cells[0], "项目所属区域", bold=True)
        # 合并剩余列
        merged_dept = r4.cells[1].merge(r4.cells[9])
        # 格式：中原区-永城 (即 dept_name)
        WordGenerator.set_cell_style(merged_dept, str(dept_name), align="left")

        # --- 签字行 (行5) ---
        r5 = table0.rows[5]
        WordGenerator.set_cell_style(r5.cells[0], "项目所属区域销售确认", bold=True)
        merged_sign = r5.cells[1].merge(r5.cells[9])
        # 增加空格以匹配 visually
        sign_text = "确认意见：\t\t\t\t签字（签章）：\t\t\t\t日期：    年    月    日"
        WordGenerator.set_cell_style(merged_sign, sign_text, align="left")

        doc.add_paragraph("\n费用详单：")

        # 4. 绘制【明细表】表头
        # 修正：表头文字完全匹配模板 
        table1 = doc.add_table(rows=1, cols=11)
        table1.style = 'Table Grid'
        
        # 注意：模板中"人事范围"可能因为列宽不够显示为两行，但在代码里我们设为"人事范围"即可
        cols_text = [
            '人员', '人事范围', '项目名称', '项目合同主体', 
            '销售人员', '销售所在大区', '支持人天', '人力费用', 
            '差旅补助', '差旅平台费用', '总费用（元）'
        ]
        
        for i, c in enumerate(cols_text):
            WordGenerator.set_cell_style(table1.rows[0].cells[i], c, bold=True, font_size=9)

        return doc

    @staticmethod
    def generate(df_result, period_text):
        if not HAS_DOCX: return {}, "缺少 python-docx 库"
        
        req_cols = ['合同主体', '人事范围', '销售部门']
        if not all(c in df_result.columns for c in req_cols):
            return {}, "数据中缺少必要列（合同主体/人事范围/销售部门），无法拆分结算单"
        
        # 确保数据不为空
        if df_result.empty: return {}, "结果数据为空"

        pairs = df_result[req_cols].dropna().drop_duplicates().values
        output_files = {}

        for purchase_comp, sales_comp, dept_name in pairs:
            # 筛选数据
            df_curr = df_result[
                (df_result['合同主体'] == purchase_comp) & 
                (df_result['人事范围'] == sales_comp) &
                (df_result['销售部门'] == dept_name)
            ].copy()
            
            if df_curr.empty: continue

            # 1. 动态“画”出模板
            doc = WordGenerator.create_hardcoded_template(purchase_comp, sales_comp, dept_name, period_text)

            # 2. 填充数据 - 汇总表 (Table 0)
            table0 = doc.tables[0]
            
            total_days = df_curr['支持时间（人天）'].sum()
            total_labor = df_curr['人力费用'].sum()
            total_sub = df_curr['差旅补助'].sum()
            total_fee = df_curr['差旅费控平台'].sum()
            grand_total = df_curr['结算费用合计'].sum()
            
            fmt = lambda x: "{:,.2f}".format(x)
            fmt_d = lambda x: "{:,.1f}".format(x)

            # 定位数据行 (现在是 Index 2)
            cells = table0.rows[2].cells
            
            # 填充并应用样式
            WordGenerator.set_cell_style(cells[0], fmt_d(total_days)) # 标准交付-工时
            WordGenerator.set_cell_style(cells[1], "0.0")             # 数据治理
            WordGenerator.set_cell_style(cells[2], "0.0")             # 运维
            WordGenerator.set_cell_style(cells[3], fmt(total_labor))  # 标准交付-费用
            WordGenerator.set_cell_style(cells[4], "0.00")
            WordGenerator.set_cell_style(cells[5], "0.00")
            WordGenerator.set_cell_style(cells[6], fmt(total_sub))
            WordGenerator.set_cell_style(cells[7], fmt(total_fee))
            WordGenerator.set_cell_style(cells[8], fmt_d(total_days)) # 合计工时
            WordGenerator.set_cell_style(cells[9], fmt(grand_total))  # 合计费用

            # 3. 填充数据 - 明细表 (Table 1)
            table1 = doc.tables[1]
            # 映射关系：DataFrame列名 -> Word表格列序
            # 注意：DataFrame 中列名是 "所属项目", Word 中是 "项目名称"
            # DataFrame 中列名是 "销售部门", Word 中是 "销售所在大区"
            cols_map_df = [
                '人员', '人事范围', '所属项目', '合同主体', 
                '销售人员', '销售部门', '支持时间（人天）', 
                '人力费用', '差旅补助', '差旅费控平台', '结算费用合计'
            ]

            for _, row in df_curr.iterrows():
                new_row = table1.add_row()
                for i, col_name in enumerate(cols_map_df):
                    val = row.get(col_name, '')
                    text_val = ""
                    if isinstance(val, (int, float)):
                        if '人天' in col_name:
                            text_val = "{:,.1f}".format(val)
                        else:
                            text_val = "{:,.2f}".format(val)
                    else:
                        text_val = str(val)
                    
                    # 应用样式 (字体小一点)
                    WordGenerator.set_cell_style(new_row.cells[i], text_val, font_size=9)
            
            # 增加明细表合计行
            sum_row = table1.add_row()
            WordGenerator.set_cell_style(sum_row.cells[0], "合计", bold=True, font_size=9)
            # 合并前几列? 模板里没有合并，直接填空
            
            WordGenerator.set_cell_style(sum_row.cells[6], fmt_d(total_days), bold=True, font_size=9)
            WordGenerator.set_cell_style(sum_row.cells[7], fmt(total_labor), bold=True, font_size=9)
            WordGenerator.set_cell_style(sum_row.cells[8], fmt(total_sub), bold=True, font_size=9)
            WordGenerator.set_cell_style(sum_row.cells[9], fmt(total_fee), bold=True, font_size=9)
            WordGenerator.set_cell_style(sum_row.cells[10], fmt(grand_total), bold=True, font_size=9)

            # 保存
            out = io.BytesIO()
            doc.save(out)
            safe_dept = str(dept_name).replace('/', '_').replace('\\', '_')
            fname = f"结算单_{purchase_comp}_{sales_comp}_{safe_dept}.docx"
            output_files[fname] = out.getvalue()

        return output_files, None
