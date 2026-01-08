import streamlit as st
import pandas as pd
import io
import time
import zipfile

# 尝试导入 python-docx
try:
    import docx
    from docx.shared import Pt, Cm
    from docx.oxml.ns import qn
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_ALIGN_VERTICAL
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False

# ==============================================================================
# Zone 0: 全局配置 & 样式注入
# ==============================================================================
st.set_page_config(page_title="淡藤财务报表 Pro", page_icon="😈", layout="wide", initial_sidebar_state="expanded")

def inject_css():
    st.markdown("""
    <style>
        :root { --bg-color: #0d1117; --card-bg: #161b22; --accent: #238636; --text: #c9d1d9; --border-color: #555c65; }
        .stApp { background-color: var(--bg-color); color: var(--text); }
        .nav-header { font-size: 1.2rem; font-weight: bold; display:flex; align-items:center; height: 100%; }
        .info-bar { background-color: rgba(56, 139, 253, 0.1); border-left: 4px solid #58a6ff; color: #c9d1d9; padding: 8px 15px; margin-bottom: 20px; font-size: 0.9rem; border-radius: 4px; }
        .error-box { border: 1px solid #ff7b72; background-color: rgba(255, 123, 114, 0.1); padding: 15px; border-radius: 6px; margin-bottom: 15px; }
    </style>
    """, unsafe_allow_html=True)

# ==============================================================================
# Zone A: 纯逻辑层 (DataEngine & WordGenerator)
# ==============================================================================
class DataEngine:
    @staticmethod
    def get_default_config():
        return pd.DataFrame([
            {"所属表": "结果表3", "序号": 1, "目标字段": "序号", "源表": "🔒 系统生成", "匹配字段": "-", "逻辑说明": "自增序列"},
            {"所属表": "结果表3", "序号": 2, "目标字段": "人员", "源表": "Source A", "匹配字段": "人员", "逻辑说明": "主键 (Join Key)"},
            {"所属表": "结果表3", "序号": 3, "目标字段": "所属项目", "源表": "Source A", "匹配字段": "项目", "逻辑说明": "维度 (First)"},
            {"所属表": "结果表3", "序号": 4, "目标字段": "人事范围", "源表": "Source A", "匹配字段": "人事范围", "逻辑说明": "维度 (->销售公司)"},
            {"所属表": "结果表3", "序号": 5, "目标字段": "SPM", "源表": "Source A", "匹配字段": "SPM", "逻辑说明": "主键 (Join Key)"},
            {"所属表": "结果表3", "序号": 6, "目标字段": "合同主体", "源表": "Source A", "匹配字段": "合同主体", "逻辑说明": "维度 (->采购公司)"},
            {"所属表": "结果表3", "序号": 7, "目标字段": "销售人员", "源表": "Source A", "匹配字段": "销售", "逻辑说明": "维度 (First)"},
            {"所属表": "结果表3", "序号": 8, "目标字段": "销售部门", "源表": "Source A", "匹配字段": "销售部门", "逻辑说明": "维度 (->采购部门)"},
            {"所属表": "结果表3", "序号": 9, "目标字段": "差旅补助", "源表": "Source B", "匹配字段": "金额", "逻辑说明": "筛选：类型='差旅补助' | 清洗：x10000转元"},
            {"所属表": "结果表3", "序号": 10, "目标字段": "差旅费控平台", "源表": "Source B", "匹配字段": "金额", "逻辑说明": "筛选：类型!='差旅补助' | 清洗：x10000转元"},
            {"所属表": "结果表3", "序号": 11, "目标字段": "耗时（小时）", "源表": "Source A", "匹配字段": "交付工时", "逻辑说明": "SUM聚合 (清洗去逗号)"},
            {"所属表": "结果表3", "序号": 12, "目标字段": "支持时间（人天）", "源表": "🔒 公式计算", "匹配字段": "-", "逻辑说明": "耗时 / 8"},
            {"所属表": "结果表3", "序号": 13, "目标字段": "人力费用", "源表": "🔒 公式计算", "匹配字段": "-", "逻辑说明": "人天 * 单价 (保留2位小数)"},
            {"所属表": "结果表3", "序号": 14, "目标字段": "结算费用合计", "源表": "🔒 公式计算", "匹配字段": "-", "逻辑说明": "人力 + 差旅 + 费控 (保留2位小数)"},
            {"所属表": "结果表3", "序号": 15, "目标字段": "[配置] B表关联人", "源表": "Source B", "匹配字段": "出差人", "逻辑说明": "辅助：用于匹配A表人员 (自动去'_云计算'后缀)"},
            {"所属表": "结果表3", "序号": 16, "目标字段": "[配置] B表关联SPM", "源表": "Source B", "匹配字段": "SPM", "逻辑说明": "辅助：用于匹配A表SPM"},
            {"所属表": "结果表3", "序号": 17, "目标字段": "[配置] B表类型列", "源表": "Source B", "匹配字段": "产品类型", "逻辑说明": "辅助：用于区分补助/费控"},
            {"所属表": "结果表2", "序号": 1, "目标字段": "序号", "源表": "🔒 系统生成", "匹配字段": "-", "逻辑说明": "自增序列"},
            {"所属表": "结果表2", "序号": 2, "目标字段": "销售公司", "源表": "🔒 结果表3", "匹配字段": "人事范围", "逻辑说明": "维度分组"},
            {"所属表": "结果表2", "序号": 3, "目标字段": "采购公司", "源表": "🔒 结果表3", "匹配字段": "合同主体", "逻辑说明": "维度分组"},
            {"所属表": "结果表2", "序号": 4, "目标字段": "采购部门", "源表": "🔒 结果表3", "匹配字段": "销售部门", "逻辑说明": "维度分组"},
            {"所属表": "结果表2", "序号": 5, "目标字段": "金额（含税）", "源表": "🔒 结果表3", "匹配字段": "结算费用合计", "逻辑说明": "SUM聚合 (保留2位小数)"},
            {"所属表": "结果表2", "序号": 6, "目标字段": "工作量（人天）", "源表": "🔒 结果表3", "匹配字段": "支持时间（人天）", "逻辑说明": "SUM聚合"},
            {"所属表": "结果表1", "序号": 1, "目标字段": "序号", "源表": "🔒 系统生成", "匹配字段": "-", "逻辑说明": "自增序列"},
            {"所属表": "结果表1", "序号": 2, "目标字段": "人员", "源表": "🔒 结果表3", "匹配字段": "人员", "逻辑说明": "维度分组"},
            {"所属表": "结果表1", "序号": 3, "目标字段": "项目工时", "源表": "🔒 结果表3", "匹配字段": "耗时（小时）", "逻辑说明": "SUM聚合"},
        ])

    @staticmethod
    def get_col(config_df, target):
        row = config_df[config_df['目标字段'] == target]
        if row.empty: return None
        return str(row.iloc[0]['匹配字段']).strip()

    @staticmethod
    def clean_num(df, col):
        if col not in df.columns: return 0
        return pd.to_numeric(df[col].astype(str).str.replace(',', ''), errors='coerce').fillna(0)

    @staticmethod
    def validate(df_a, df_b, config_df, min_hours):
        errors = []
        c = lambda t: DataEngine.get_col(config_df, t)
        col_a_user = c('人员')
        col_a_spm = c('SPM')
        col_a_hrs = c('耗时（小时）')
        col_b_user = c('[配置] B表关联人')
        col_b_spm = c('[配置] B表关联SPM')
        col_b_amt = c('差旅补助')
        
        def check(df, col, src, tag):
            if col and col not in df.columns:
                errors.append({'类型':'逻辑错误', '来源':src, '_sys_id':'-', '行号':'-', '信息':f'缺列: {col} (用途:{tag})'})
                return False
            return True

        valid_a = check(df_a, col_a_user, 'Source A', '人员') and check(df_a, col_a_spm, 'Source A', 'SPM') and check(df_a, col_a_hrs, 'Source A', '工时')
        valid_b = check(df_b, col_b_user, 'Source B', '出差人') and check(df_b, col_b_spm, 'Source B', 'SPM') and check(df_b, col_b_amt, 'Source B', '金额')
        if not (valid_a and valid_b): return errors, df_a, df_b

        df_a_clean = df_a.copy()
        df_a_clean[col_a_hrs] = DataEngine.clean_num(df_a_clean, col_a_hrs)
        for i, r in df_a_clean[df_a_clean[col_a_hrs] < 0].iterrows():
            errors.append({'类型':'数据错误', '来源':'Source A', '_sys_id':r.get('_sys_id','-'), '行号':r.get('_sys_id','-'), '信息':'工时为负'})
        
        if col_a_user and col_a_hrs and min_hours > 0:
            user_sums = df_a_clean.groupby(col_a_user)[col_a_hrs].sum()
            invalid_users = user_sums[user_sums < min_hours]
            for user, total_hrs in invalid_users.items():
                sample_rows = df_a_clean[df_a_clean[col_a_user] == user]
                if not sample_rows.empty:
                    r = sample_rows.iloc[0]
                    errors.append({'类型': '业务规则校验', '来源': 'Source A', '_sys_id': r.get('_sys_id', '-'), '行号': r.get('_sys_id', '-'), '信息': f'人员【{user}】总工时({total_hrs}h) 低于阈值({min_hours}h)'})
        return errors, df_a, df_b

    @staticmethod
    def calculate(df_a, df_b, config_df, price_per_day, subsidy_tag):
        c = lambda t: DataEngine.get_col(config_df, t)
        col_a_user = c('人员')
        col_a_spm = c('SPM')
        col_a_hrs = c('耗时（小时）')
        dims_a = {'project': c('所属项目'), 'range': c('人事范围'), 'contract': c('合同主体'), 'sales': c('销售人员'), 'dept': c('销售部门')}
        col_b_user = c('[配置] B表关联人')
        col_b_spm = c('[配置] B表关联SPM')
        col_b_amt = c('差旅补助')
        col_b_type = c('[配置] B表类型列')

        df_a[col_a_hrs] = DataEngine.clean_num(df_a, col_a_hrs)
        df_b[col_b_amt] = DataEngine.clean_num(df_b, col_b_amt)
        # 核心逻辑：万元转元
        df_b[col_b_amt] = (df_b[col_b_amt] * 10000).round(2)

        if col_b_user and col_b_user in df_b.columns:
            df_b[col_b_user] = df_b[col_b_user].astype(str).str.replace('_云计算', '', regex=False).str.strip()

        agg_rules = {col_a_hrs: 'sum'}
        for _, col in dims_a.items():
            if col: agg_rules[col] = 'first'
        df_a_gp = df_a.groupby([col_a_user, col_a_spm], as_index=False).agg(agg_rules)

        is_sub = df_b[col_b_type].astype(str).str.contains(subsidy_tag, na=False)
        grp_b = [col_b_user, col_b_spm]
        df_sub = df_b[is_sub].groupby(grp_b)[col_b_amt].sum().reset_index(name='差旅补助')
        df_fee = df_b[~is_sub].groupby(grp_b)[col_b_amt].sum().reset_index(name='差旅费控平台')

        for d in [df_a_gp, df_sub, df_fee]:
            k = col_a_spm if col_a_spm in d.columns else col_b_spm
            d[k] = d[k].astype(str)

        res = pd.merge(df_a_gp, df_sub, left_on=[col_a_user, col_a_spm], right_on=[col_b_user, col_b_spm], how='left')
        res = pd.merge(res, df_fee, left_on=[col_a_user, col_a_spm], right_on=[col_b_user, col_b_spm], how='left')
        res = res.fillna(0)

        res['支持时间（人天）'] = res[col_a_hrs] / 8
        res['人力费用'] = (res['支持时间（人天）'] * price_per_day).round(2)
        res['差旅补助'] = res['差旅补助'].round(2)
        res['差旅费控平台'] = res['差旅费控平台'].round(2)
        res['结算费用合计'] = (res['人力费用'] + res['差旅补助'] + res['差旅费控平台']).round(2)

        rename_map = {col_a_user: '人员', dims_a['project']: '所属项目', dims_a['range']: '人事范围', col_a_spm: 'SPM', dims_a['contract']: '合同主体', dims_a['sales']: '销售人员', dims_a['dept']: '销售部门', col_a_hrs: '耗时（小时）'}
        t3 = res.rename(columns=rename_map)
        
        final_cols = ['序号','人员','所属项目','人事范围','SPM','合同主体','销售人员','销售部门','差旅补助','差旅费控平台','耗时（小时）','支持时间（人天）','人力费用','结算费用合计']
        t3.insert(0, '序号', range(1, len(t3)+1))
        t3 = t3[[c for c in final_cols if c in t3.columns]]

        # 表2生成逻辑：三重分组 (人事范围=销售, 合同主体=采购, 销售部门=采购部门)
        t2_cols = ['人事范围', '合同主体', '销售部门']
        if all(c in t3.columns for c in t2_cols):
            t2 = t3.groupby(t2_cols).agg({'结算费用合计':'sum', '支持时间（人天）':'sum'}).reset_index()
            t2['结算费用合计'] = t2['结算费用合计'].round(2)
            t2.columns = ['销售公司', '采购公司', '采购部门', '金额（含税，单位：元）', '工作量（人天）']
            t2.insert(0, '序号', range(1, len(t2)+1))
        else: t2 = pd.DataFrame()

        t1 = t3.groupby('人员')['耗时（小时）'].sum().reset_index()
        t1.rename(columns={'耗时（小时）':'项目工时'}, inplace=True)
        t1.insert(0, '序号', range(1, len(t1)+1))

        return {'t1': t1, 't2': t2, 't3': t3}

    @staticmethod
    def to_bytes(df):
        b = io.BytesIO()
        out = df.drop(columns=['_sys_id'], errors='ignore')
        # 【关键修正】使用 openpyxl 引擎防止 xlsxwriter 报错
        with pd.ExcelWriter(b, engine='openpyxl') as writer:
            out.to_excel(writer, index=False)
        return b.getvalue()

class WordGenerator:
    @staticmethod
    def create_base_template(purchase_comp, sales_comp, dept_name, period_text):
        """
        【核心升级】内置模板生成引擎
        """
        doc = docx.Document()
        
        # 1. 设置默认字体为宋体
        doc.styles['Normal'].font.name = 'Times New Roman'
        doc.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
        
        # 2. 标题
        p_title = doc.add_paragraph()
        p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p_title.add_run(f"{purchase_comp} 与 {sales_comp} {period_text} 项目交付与运维费用结算账单")
        run.font.size = Pt(14)
        run.font.bold = True
        
        # 3. 汇总表 (Table 0) - 7行10列 (含合并逻辑)
        table0 = doc.add_table(rows=7, cols=10)
        table0.style = 'Table Grid'
        
        # --- 绘制复杂表头 ---
        r0 = table0.rows[0]
        r0.cells[0].merge(r0.cells[2]).text = "工作量 （单位：人/天）"
        r0.cells[3].merge(r0.cells[5]).text = "人力费用 （单位：元）"
        r0.cells[6].merge(r0.cells[7]).text = "差旅费用 （单位：元）"
        r0.cells[8].merge(r0.cells[9]).text = "合计"
        
        # Row 1: Sub-headers
        r1 = table0.rows[1]
        headers_r1 = ["项目标准交付", "项目数据治理", "项目运维服务", "项目标准交付", "项目数据治理", "项目运维服务", "差旅补助", "商旅平台费用", "工作量", "合计费用"]
        for i, text in enumerate(headers_r1):
            r1.cells[i].text = text
            
        # Row 5: 区域行
        r5 = table0.rows[5]
        r5.cells[0].text = "项目所属区域"
        r5.cells[1].merge(r5.cells[9]).text = str(dept_name)
        
        # Row 6: 签字行
        r6 = table0.rows[6]
        r6.cells[0].text = "项目所属区域销售确认"
        r6.cells[1].merge(r6.cells[9]).text = "确认意见：            签字（签章）：            日期：    年    月    日"

        # 设置表格样式细节 (居中)
        for row in table0.rows:
            for cell in row.cells:
                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                for paragraph in cell.paragraphs:
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

        doc.add_paragraph("\n费用详单：")
        
        # 4. 明细表 (Table 1) - 表头
        table1 = doc.add_table(rows=1, cols=11)
        table1.style = 'Table Grid'
        cols_map = ['人员', '人事范围', '项目名称', '项目合同主体', '销售人员', '销售所在大区', '支持人天', '人力费用', '差旅补助', '差旅平台费用', '总费用（元）']
        for i, col_name in enumerate(cols_map):
            cell = table1.rows[0].cells[i]
            cell.text = col_name
            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            
        return doc

    @staticmethod
    def generate(df_result, period_text):
        """
        基于内置模板引擎生成 Word，返回字典 {文件名: 二进制流}
        """
        if not HAS_DOCX: return {}, "缺少 python-docx 库"
        
        req_cols = ['合同主体', '人事范围', '销售部门']
        if not all(c in df_result.columns for c in req_cols):
            return {}, "数据中缺少必要列（合同主体/人事范围/销售部门），无法拆分结算单"
        
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

            # === 调用内置引擎绘制“空卷子” ===
            doc = WordGenerator.create_base_template(purchase_comp, sales_comp, dept_name, period_text)

            # === 填充汇总数据 (Table 0) ===
            table0 = doc.tables[0]
            
            # 计算汇总 (归一法)
            total_days = df_curr['支持时间（人天）'].sum()
            total_labor = df_curr['人力费用'].sum()
            total_sub = df_curr['差旅补助'].sum()
            total_fee = df_curr['差旅费控平台'].sum()
            grand_total = df_curr['结算费用合计'].sum()
            
            # 填入 Row 2
            cells = table0.rows[2].cells # 第三行填数据
            
            fmt = lambda x: "{:,.2f}".format(x)
            fmt_d = lambda x: "{:,.1f}".format(x)
            
            cells[0].text = fmt_d(total_days) # 标准-工时
            cells[1].text = "0.0"
            cells[2].text = "0.0"
            cells[3].text = fmt(total_labor)  # 标准-费用
            cells[4].text = "0.00"
            cells[5].text = "0.00"
            cells[6].text = fmt(total_sub)    # 补助
            cells[7].text = fmt(total_fee)    # 平台
            cells[8].text = fmt_d(total_days) # 总工时
            cells[9].text = fmt(grand_total)  # 总费用
            
            # 居中对齐数值
            for cell in cells:
                for p in cell.paragraphs: p.alignment = WD_ALIGN_PARAGRAPH.CENTER

            # === 填充明细数据 (Table 1) ===
            table1 = doc.tables[1]
            cols_map = ['人员', '人事范围', '所属项目', '合同主体', '销售人员', '销售部门', '支持时间（人天）', '人力费用', '差旅补助', '差旅费控平台', '结算费用合计']
            
            for _, row in df_curr.iterrows():
                new_row = table1.add_row()
                for i, col_name in enumerate(cols_map):
                    val = row.get(col_name, '')
                    cell = new_row.cells[i]
                    if isinstance(val, (int, float)):
                        if '人天' in col_name:
                            cell.text = "{:,.1f}".format(val)
                        else:
                            cell.text = "{:,.2f}".format(val)
                    else:
                        cell.text = str(val)
                    cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

            # === 增加合计行到明细表底部 ===
            sum_row = table1.add_row()
            sum_row.cells[0].text = "合计"
            sum_row.cells[6].text = fmt_d(total_days)
            sum_row.cells[7].text = fmt(total_labor)
            sum_row.cells[8].text = fmt(total_sub)
            sum_row.cells[9].text = fmt(total_fee)
            sum_row.cells[10].text = fmt(grand_total)
            for c in sum_row.cells: c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

            out = io.BytesIO()
            doc.save(out)
            safe_dept = str(dept_name).replace('/', '_').replace('\\', '_')
            fname = f"结算单_{purchase_comp}_{sales_comp}_{safe_dept}.docx"
            output_files[fname] = out.getvalue()

        return output_files, None

# ==============================================================================
# Zone B: UI 组件层 (View)
# ==============================================================================
class UIComponents:
    @staticmethod
    def render_sidebar():
        with st.sidebar:
            st.header("⚙️ 参数配置")
            p = st.number_input("人力单价 (元/天)", value=1500, step=100)
            h = st.number_input("工时阈值 (小时)", value=100)
            s = st.text_input("补助关键词", "差旅补助")
            period = st.text_input("结算周期文案", "2025年第三季度")
            
            st.markdown("---")
            if st.button("🐱 字段映射 & 逻辑", help="查看映射逻辑"):
                st.session_state.page = 'mapping'
                st.rerun()
            return p, h, s, period

    @staticmethod
    def render_file_slot(key, title, data_store):
        data = data_store[key]
        has_file = data['df'] is not None
        with st.container(border=True):
            if not has_file:
                st.markdown(f"**{title}**")
                return st.file_uploader(title, type=['xlsx', 'csv'], key=f"uploader_{key}", label_visibility="collapsed")
            else:
                c1, c2 = st.columns([9, 1])
                c1.markdown(f"✅ **{data['name']}**")
                if c2.button("Del", key=f"del_{key}"): return "DELETE"
        return None

    @staticmethod
    def render_error_report(err_df, on_fix):
        fixable = err_df[err_df['类型']=='数据错误']
        logic = err_df[err_df['类型']=='逻辑错误']
        rule = err_df[err_df['类型']=='业务规则校验']
        st.markdown(f"<div class='error-box'><h3 style='margin:0'>🚨 校验失败</h3><p>发现 <b>{len(fixable)}</b> 个数据错误，<b>{len(rule)}</b> 个规则异常，<b>{len(logic)}</b> 个映射错误。</p></div>", unsafe_allow_html=True)
        st.dataframe(err_df[['类型','来源','行号','信息']], use_container_width=True, hide_index=True)
        c1, c2, c3 = st.columns([1, 1, 1])
        c1.download_button("📥 下载清单", err_df.to_csv(index=False).encode('utf-8-sig'), "err.csv", "text/csv", use_container_width=True)
        should_rerun = c2.button("🔄 参数已改，重新校验", type="secondary", use_container_width=True)
        if not fixable.empty:
            if c3.button("🛠️ 在线修复", type="primary", use_container_width=True): on_fix()
        return should_rerun

    @staticmethod
    def render_download_zone(result_files, all_in_one_zip, word_files_dict):
        with st.container(border=True):
            st.success("✅ 所有报表生成完毕")
            
            # 1. 顶部：超级批量下载按钮
            st.subheader("📦 批量下载")
            st.download_button(
                label="🚀 一键下载所有文件 (Excel + Word 打包)",
                data=all_in_one_zip,
                file_name="项目结算资料全集.zip",
                mime="application/zip",
                type="primary",
                use_container_width=True
            )

            st.divider()

            # 2. 中部：Excel 单独下载
            st.subheader("📊 基础数据表 (Excel)")
            c1, c2, c3 = st.columns(3)
            if 't1' in result_files: c1.download_button("📥 表1: 工时统计", result_files['t1'], "t1.xlsx", use_container_width=True)
            if 't2' in result_files: c2.download_button("📥 表2: 结算汇总", result_files['t2'], "t2.xlsx", use_container_width=True)
            if 't3' in result_files: c3.download_button("📥 表3: 详细明细", result_files['t3'], "t3.xlsx", use_container_width=True)

            # 3. 底部：Word 列表 (动态生成)
            st.subheader(f"📝 结算单 (Word - 共{len(word_files_dict)}个)")
            if not word_files_dict:
                st.info("没有生成 Word 文件（可能是缺少 docx 库或无数据）。")
            else:
                with st.expander(f"点击展开查看 {len(word_files_dict)} 份结算单", expanded=False):
                    for fname, fbytes in word_files_dict.items():
                        col_text, col_btn = st.columns([4, 1])
                        with col_text:
                            st.text(f"📄 {fname}")
                        with col_btn:
                            st.download_button(
                                "下载", 
                                fbytes, 
                                fname, 
                                key=f"btn_{fname}",
                                use_container_width=True
                            )

    @staticmethod
    def render_native_editor(desc, subset, is_edit, cols_a, cols_b):
        st.markdown(f'<div class="info-bar">ℹ️ {desc}</div>', unsafe_allow_html=True)
        df_display = subset[['序号', '目标字段', '源表', '匹配字段', '逻辑说明']].copy().reset_index(drop=True)
        df_display['序号'] = df_display['序号'].astype(str)
        column_config = {
            "序号": st.column_config.TextColumn("序号", width="small", disabled=True),
            "目标字段": st.column_config.TextColumn("目标字段", disabled=True, width="medium"),
            "逻辑说明": st.column_config.TextColumn("逻辑说明", disabled=True, width="large"),
        }
        if is_edit:
            column_config["源表"] = st.column_config.SelectboxColumn("源表", options=["Source A", "Source B"], width="small", required=True)
            column_config["匹配字段"] = st.column_config.SelectboxColumn("匹配字段", options=cols_a + cols_b, width="medium", required=True)
        else:
            column_config["源表"] = st.column_config.TextColumn("源表", disabled=True)
            column_config["匹配字段"] = st.column_config.TextColumn("匹配字段", disabled=True)

        calc_height = (len(df_display) + 1) * 35 + 10
        final_height = max(400, min(1000, calc_height))

        edited = st.data_editor(df_display, column_config=column_config, use_container_width=True, hide_index=True, disabled=not is_edit, height=final_height, key=f"editor_{subset.iloc[0]['所属表']}")

        if is_edit:
            for i, row in edited.iterrows():
                orig_idx = subset.index[i]
                orig_row = st.session_state.mapping_config.loc[orig_idx]
                if 'Source' not in orig_row['源表']: continue 
                if row['源表'] != orig_row['源表']:
                    st.session_state.mapping_config.at[orig_idx, '源表'] = row['源表']
                    target_opts = cols_a if row['源表'] == 'Source A' else cols_b
                    new_val = row['目标字段'] if row['目标字段'] in target_opts else (target_opts[0] if target_opts else None)
                    st.session_state.mapping_config.at[orig_idx, '匹配字段'] = new_val
                elif row['匹配字段'] != orig_row['匹配字段']:
                    valid_opts = cols_a if row['源表'] == 'Source A' else cols_b
                    if row['匹配字段'] in valid_opts:
                        st.session_state.mapping_config.at[orig_idx, '匹配字段'] = row['匹配字段']

# ==============================================================================
# Zone C: 控制层
# ==============================================================================
if 'page' not in st.session_state: st.session_state.page = 'main'
if 'data_store' not in st.session_state: st.session_state.data_store = {'A': {'df': None, 'name': None}, 'B': {'df': None, 'name': None}}
if 'mapping_config' not in st.session_state: st.session_state.mapping_config = DataEngine.get_default_config()
if 'is_calculated' not in st.session_state: st.session_state.is_calculated = False
if 'error_report' not in st.session_state: st.session_state.error_report = None
if 'block_auto_run' not in st.session_state: st.session_state.block_auto_run = False
if 'is_editing_mapping' not in st.session_state: st.session_state.is_editing_mapping = False
if 'all_zip' not in st.session_state: st.session_state.all_zip = None
if 'word_files' not in st.session_state: st.session_state.word_files = {}

inject_css()

if st.session_state.page == 'main':
    price, hours_limit, sub_tag, period_text = UIComponents.render_sidebar()
    st.title("😈 淡藤财务报表 Pro")

    with st.container(border=True):
        c_h1, c_h2 = st.columns([8, 1])
        c_h1.markdown("### 📂 数据源控制台")
        if c_h2.button("🗑️ 重置"): 
            st.session_state.clear()
            st.rerun()
        
        c1, c2 = st.columns(2)
        def handle(key, title):
            res = UIComponents.render_file_slot(key, title, st.session_state.data_store)
            if res == "DELETE":
                st.session_state.data_store[key] = {'df': None, 'name': None}
                st.session_state.is_calculated = False
                st.session_state.error_report = None
                st.session_state.block_auto_run = False
                st.session_state.all_zip = None
                st.rerun()
            elif res:
                try:
                    df = pd.read_csv(res) if res.name.endswith('.csv') else pd.read_excel(res)
                    df.columns = [str(c).strip() for c in df.columns]
                    df['_sys_id'] = range(1, len(df)+1)
                    st.session_state.data_store[key] = {'df': df, 'name': res.name}
                    if st.session_state.block_auto_run: st.session_state.error_report = None
                    st.rerun()
                except Exception as e: st.error(f"Error: {e}")
        handle('A', "Source A: 投入明细")
        handle('B', "Source B: 差旅明细")

    st.divider()
    has_files = st.session_state.data_store['A']['df'] is not None and st.session_state.data_store['B']['df'] is not None
    trigger = False

    if has_files:
        if st.session_state.is_calculated:
            # 【修复】调用新的大一统下载渲染函数
            UIComponents.render_download_zone(
                st.session_state.result_files, 
                st.session_state.all_zip,
                st.session_state.word_files
            )
        
        elif st.session_state.error_report is not None:
            def fix_action():
                @st.dialog("🛠️ 在线修复", width="large")
                def show_fix():
                    fixable = st.session_state.error_report[st.session_state.error_report['类型']=='数据错误']
                    def get_df(src):
                        ids = fixable[fixable['来源']==src]['_sys_id'].unique()
                        if len(ids)==0: return pd.DataFrame()
                        full = st.session_state.data_store[src.split()[-1]]['df']
                        return full[full['_sys_id'].isin(ids)].copy()
                    
                    da, db = get_df('Source A'), get_df('Source B')
                    t1, t2 = st.tabs([f"A ({len(da)})", f"B ({len(db)})"])
                    na, nb = None, None
                    
                    with t1:
                        if not da.empty:
                            na = st.data_editor(da, height=300, hide_index=True, column_config={"_sys_id": None}, key="fix_a")
                        else: st.info("无数据错误")
                    with t2:
                        if not db.empty:
                            nb = st.data_editor(db, height=300, hide_index=True, column_config={"_sys_id": None}, key="fix_b")
                        else: st.info("无数据错误")
                    
                    if st.button("💾 保存并重算", type="primary"):
                        if na is not None:
                            od = st.session_state.data_store['A']['df'].set_index('_sys_id')
                            od.update(na.set_index('_sys_id'))
                            st.session_state.data_store['A']['df'] = od.reset_index()
                        if nb is not None:
                            od = st.session_state.data_store['B']['df'].set_index('_sys_id')
                            od.update(nb.set_index('_sys_id'))
                            st.session_state.data_store['B']['df'] = od.reset_index()
                        st.session_state.error_report = None
                        st.session_state.block_auto_run = False
                        st.rerun()
                show_fix()
            
            should_rerun = UIComponents.render_error_report(st.session_state.error_report, fix_action)
            if should_rerun:
                st.session_state.error_report = None
                trigger = True

        elif st.session_state.block_auto_run:
            st.info("ℹ️ 源文件已更新，等待确认...")
            if st.button("▶️ 重新校验并计算", type="primary", use_container_width=True): trigger = True
        else:
            trigger = True

    if trigger:
        with st.spinner("🚀 正在计算数据并生成全套报表..."):
            # 1. 校验逻辑
            errs, df_a, df_b = DataEngine.validate(
                st.session_state.data_store['A']['df'].copy(),
                st.session_state.data_store['B']['df'].copy(),
                st.session_state.mapping_config,
                hours_limit
            )

            if errs:
                st.session_state.error_report = pd.DataFrame(errs)
                st.session_state.block_auto_run = True
                st.rerun()
            else:
                # 2. 计算核心逻辑
                res = DataEngine.calculate(df_a, df_b, st.session_state.mapping_config, price, sub_tag)
                
                # 3. 生成 Excel 二进制流 (openpyxl)
                excel_files_dict = {
                    "表1_工时统计.xlsx": DataEngine.to_bytes(res['t1']),
                    "表2_结算汇总.xlsx": DataEngine.to_bytes(res['t2']),
                    "表3_详细明细.xlsx": DataEngine.to_bytes(res['t3'])
                }
                st.session_state.result_files = excel_files_dict
                
                # 4. 生成 Word 二进制流 (动态N个)
                word_files_dict, err_msg = WordGenerator.generate(res['t3'], period_text)
                if err_msg:
                    st.warning(f"Word生成部分受限: {err_msg}")
                    word_files_dict = {}
                st.session_state.word_files = word_files_dict

                # 5. 【关键】合并所有文件 -> 打包 ZIP
                all_files_to_zip = {**excel_files_dict, **word_files_dict}
                
                buf_zip = io.BytesIO()
                with zipfile.ZipFile(buf_zip, 'w', zipfile.ZIP_DEFLATED) as z:
                    for fname, fcontent in all_files_to_zip.items():
                        z.writestr(fname, fcontent)
                
                st.session_state.all_zip = buf_zip.getvalue()
                st.session_state.is_calculated = True
                st.rerun()

elif st.session_state.page == 'mapping':
    c1, c2 = st.columns([9, 1])
    c1.markdown("<div class='nav-header'>🐱 字段映射 & 逻辑配置</div>", unsafe_allow_html=True)
    if c2.button("⬅️", use_container_width=True, help="返回主页"): 
        st.session_state.page = 'main'
        st.rerun()
    
    st.markdown("<hr style='margin-top:0; border-color:#30363d;'>", unsafe_allow_html=True)

    c_title, c_spacer, c_action = st.columns([7, 1, 2])
    c_title.markdown("#### 🧬 数据血缘与逻辑配置")
    
    has_files = st.session_state.data_store['A']['df'] is not None and st.session_state.data_store['B']['df'] is not None
    
    with c_action:
        if not st.session_state.is_editing_mapping:
            if st.button("✏️ 编辑配置", type="primary", use_container_width=True):
                if not has_files: st.toast("请先在主页上传 A/B 表", icon="🚫")
                else:
                    st.session_state.is_editing_mapping = True
                    st.rerun()
        else:
            if st.button("💾 保存生效", type="primary", use_container_width=True):
                st.session_state.is_editing_mapping = False
                st.session_state.is_calculated = False
                st.session_state.block_auto_run = False
                st.session_state.error_report = None
                st.rerun()

    df_c = st.session_state.mapping_config
    t1, t2, t3 = st.tabs(["结果表3 (底表)", "结果表2 (结算)", "结果表1 (工时)"])
    
    cols_a = list(st.session_state.data_store['A']['df'].columns) if has_files else []
    cols_b = list(st.session_state.data_store['B']['df'].columns) if has_files else []

    with t1:
        UIComponents.render_native_editor(
            "全量明细底表：基于 Source A/B 进行清洗、聚合、关联计算后的宽表。",
            df_c[df_c['所属表']=='结果表3'], st.session_state.is_editing_mapping, cols_a, cols_b
        )
    with t2:
        UIComponents.render_native_editor(
            "结算汇总表：基于【结果表3】按公司/部门维度二次聚合的金额数据。",
            df_c[df_c['所属表']=='结果表2'], st.session_state.is_editing_mapping, cols_a, cols_b
        )
    with t3:
        UIComponents.render_native_editor(
            "工时统计表：基于【结果表3】按人员维度二次聚合的工时数据。",
            df_c[df_c['所属表']=='结果表1'], st.session_state.is_editing_mapping, cols_a, cols_b
        )
