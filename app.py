import streamlit as st
import pandas as pd
import io
import time
import zipfile
import re

# ==============================================================================
# 依赖库检查与导入 (python-docx)
# ==============================================================================
try:
    import docx
    from docx.shared import Pt, Cm, RGBColor
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
    from docx.enum.table import WD_ALIGN_VERTICAL, WD_ROW_HEIGHT_RULE
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False

# Excel样式依赖
from openpyxl.styles import Border, Side, Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# ==============================================================================
# Zone 0: 全局配置 & 样式注入 (CSS 深度修复版)
# ==============================================================================
st.set_page_config(page_title="淡藤财务报表 Pro", page_icon="😈", layout="wide", initial_sidebar_state="expanded")

def inject_css():
    st.markdown("""
    <style>
        :root { --bg-color: #0d1117; --card-bg: #161b22; --text: #c9d1d9; --border-color: #30363d; }
        .stApp { background-color: var(--bg-color); color: var(--text); }
        
        /* 强制按钮文字不换行 */
        button p { white-space: nowrap !important; }

        /* === 文件卡片样式 === */
        .file-name { font-weight: 600; font-size: 14px; color: #e6edf3; display: block; line-height: 1.2; }
        .file-stats { font-size: 12px; color: #8b949e; display: block; margin-top: 2px; }
        .file-icon { font-size: 24px; display: flex; align-items: center; justify-content: center; height: 100%; }
        
        /* === 修复 X 按钮不显示的核心 CSS === */
        /* 1. 确保列内容不被隐藏 */
        div[data-testid="column"] { overflow: visible !important; }

        /* 2. 针对删除按钮 (secondary) 的强制样式 */
        div[data-testid="column"] button[kind="secondary"] {
            border: 1px solid rgba(255,255,255,0.1) !important; /* 微弱边框，确保可见 */
            background-color: rgba(255,255,255,0.05) !important; /* 微弱背景 */
            color: #c9d1d9 !important; /* 亮灰色文字 */
            padding: 0px !important;
            margin: 0px !important;
            height: 42px !important; /* 强制高度 */
            width: 100% !important;  /* 强制撑满 */
            min-width: 40px !important;
            display: flex !important;
            align-items: center !important;
            justify-content: center !important;
            float: none !important;
            border-radius: 6px !important;
            transition: all 0.2s;
        }

        /* 3. 悬停效果 */
        div[data-testid="column"] button[kind="secondary"]:hover {
            color: #ff7b72 !important; /* 红色文字 */
            border-color: #ff7b72 !important; /* 红色边框 */
            background-color: rgba(255, 123, 114, 0.1) !important;
        }

        /* 4. 兼容性修复 */
        div[data-testid="column"] button[kind="secondary"]:active,
        div[data-testid="column"] button[kind="secondary"]:focus {
            box-shadow: none !important;
            outline: none !important;
        }
        
        /* 顶部信息栏样式 */
        .info-bar { background-color: rgba(56, 139, 253, 0.1); border-left: 4px solid #58a6ff; color: #c9d1d9; padding: 8px 15px; margin-bottom: 20px; font-size: 0.9rem; border-radius: 4px; }
        .error-box { border: 1px solid #ff7b72; background-color: rgba(255, 123, 114, 0.1); padding: 15px; border-radius: 6px; margin-bottom: 15px; }
        .balance-box-ok { border: 1px solid #238636; background-color: rgba(35, 134, 54, 0.1); padding: 10px; border-radius: 6px; margin-bottom: 15px; color: #3fb950; }
        .balance-box-err { border: 1px solid #da3633; background-color: rgba(218, 54, 51, 0.1); padding: 10px; border-radius: 6px; margin-bottom: 15px; color: #f85149; font-weight: bold;}

    </style>
    """, unsafe_allow_html=True)

# ==============================================================================
# Zone A: 纯逻辑层 (包含 T4 校验与空白页修复)
# ==============================================================================
class DataEngine:
    @staticmethod
    def get_quarter_str(text):
        if not text: return "2025Qx"
        s = str(text)
        year_match = re.search(r'(\d{4})', s)
        year = year_match.group(1) if year_match else "2025"
        quarter = "Q1" 
        if '一' in s: quarter = "Q1"
        elif '二' in s: quarter = "Q2"
        elif '三' in s: quarter = "Q3"
        elif '四' in s: quarter = "Q4"
        else:
            rem = s.replace(year, '')
            if '1' in rem: quarter = "Q1"
            elif '2' in rem: quarter = "Q2"
            elif '3' in rem: quarter = "Q3"
            elif '4' in rem: quarter = "Q4"
        return f"{year}{quarter}"

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
        """核心校验逻辑"""
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

        if col_b_user in df_b.columns and col_a_user in df_a.columns:
            users_in_a = set(df_a[col_a_user].astype(str).str.strip())
            b_clean_series = df_b[col_b_user].astype(str).str.replace('_云计算', '', regex=False).str.strip()
            users_in_b = set(b_clean_series)
            ghost_users = users_in_b - users_in_a
            if ghost_users:
                for u in ghost_users:
                    example_rows = df_b[b_clean_series == u]
                    for idx, row in example_rows.iterrows():
                        sys_id = row.get('_sys_id', idx+1)
                        errors.append({
                            '类型': '业务规则校验', 
                            '来源': 'Source B', 
                            '_sys_id': sys_id, 
                            '行号': sys_id, 
                            '信息': f'异常差旅：人员【{u}】产生差旅费用，但在 Source A 中无对应交付记录'
                        })

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
        df_b[col_b_amt] = df_b[col_b_amt].round(2)

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
            d[k] = d[k].astype(str).str.strip()

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

        t2_cols = ['人事范围', '合同主体', '销售部门']
        if all(c in t3.columns for c in t2_cols):
            t2 = t3.groupby(t2_cols).agg({'结算费用合计':'sum', '支持时间（人天）':'sum'}).reset_index()
            t2['结算费用合计'] = t2['结算费用合计'].round(2)
            t2.columns = ['销售公司', '采购公司', '采购部门', '金额（含税，单位：元）', '工作量（人天）']
            t2.insert(0, '序号', range(1, len(t2)+1))
        else: t2 = pd.DataFrame()

        t1 = t3.groupby('人员')['耗时（小时）'].sum().reset_index()
        t1.rename(columns={'耗时（小时）':'项目工时'}, inplace=True)
        t1['人员类型'] = t1['人员'].apply(lambda x: '实施交付部' if str(x).strip() == '黄毅兵' else '实施交付部云交付小组')
        t1['备注'] = ''
        t1.insert(0, '序号', range(1, len(t1)+1))
        t1 = t1[['序号', '人员', '人员类型', '项目工时', '备注']]

        return {'t1': t1, 't2': t2, 't3': t3}

    @staticmethod
    def verify_balance(df_a, df_b, results_dict, config_df):
        """
        执行全链路数据平衡校验 & 内部勾稽关系校验 (包含 T4 vs T1/T2/T3)
        """
        messages = []
        is_balanced = True
        
        c = lambda t: DataEngine.get_col(config_df, t)
        col_a_hrs = c('耗时（小时）')
        col_b_amt = c('差旅补助')
        
        df_t1 = results_dict['t1']
        df_t2 = results_dict['t2']
        df_t3 = results_dict['t3']
        
        # --- 1. 源数据(Input) vs 结果数据(Output) 总额校验 ---
        clean_a_hrs = DataEngine.clean_num(df_a, col_a_hrs).sum()
        res_hrs = df_t3['耗时（小时）'].sum()
        
        if abs(clean_a_hrs - res_hrs) > 0.1:
            is_balanced = False
            messages.append(f"❌ [输入输出] 工时丢失：源表({clean_a_hrs:,.1f}) != 明细表({res_hrs:,.1f})")
            
        clean_b_amt = DataEngine.clean_num(df_b, col_b_amt).sum()
        res_amt = df_t3['差旅补助'].sum() + df_t3['差旅费控平台'].sum()
        
        if abs(clean_b_amt - res_amt) > 0.1:
            is_balanced = False
            messages.append(f"❌ [输入输出] 金额丢失：源表({clean_b_amt:,.2f}) != 明细表({res_amt:,.2f}) (可能原因：B表有SPM未匹配到A表)")

        # --- 2. 内部勾稽关系校验 (Simulate Result 4) ---
        req_cols = ['合同主体', '人事范围', '销售部门']
        if all(c in df_t3.columns for c in req_cols):
            # T4: 按照分单维度聚合
            df_t4 = df_t3.groupby(req_cols)[['结算费用合计', '支持时间（人天）']].sum().reset_index()
            
            # (A) 校验 T4 vs T2 (部门汇总表)
            t2_sum_amt = df_t2['金额（含税，单位：元）'].sum()
            t4_sum_amt = df_t4['结算费用合计'].sum()
            if abs(t2_sum_amt - t4_sum_amt) > 0.05:
                is_balanced = False
                messages.append(f"❌ [内部勾稽] 结算汇总表(T2)与分单合集(T4)金额不平: {t2_sum_amt:,.2f} vs {t4_sum_amt:,.2f}")
            
            t2_sum_days = df_t2['工作量（人天）'].sum()
            t4_sum_days = df_t4['支持时间（人天）'].sum()
            if abs(t2_sum_days - t4_sum_days) > 0.05:
                is_balanced = False
                messages.append(f"❌ [内部勾稽] 结算汇总表(T2)与分单合集(T4)人天不平")

            # (B) 校验 T4 vs T1 (人员工时表)
            t1_sum_hrs = df_t1['项目工时'].sum()
            t4_calc_hrs = df_t4['支持时间（人天）'].sum() * 8
            if abs(t1_sum_hrs - t4_calc_hrs) > 0.1:
                is_balanced = False
                messages.append(f"❌ [内部勾稽] 工时统计表(T1)与分单合集(T4)工时转换不平: {t1_sum_hrs:,.1f} vs {t4_calc_hrs:,.1f}")

            # (C) 校验 T4 vs T3 (明细底表)
            t3_sum_amt = df_t3['结算费用合计'].sum()
            if abs(t3_sum_amt - t4_sum_amt) > 0.05:
                 is_balanced = False
                 messages.append(f"❌ [内部勾稽] 明细底表(T3)与分单合集(T4)金额聚合不平")

        if is_balanced:
            return True, "✅ 全链路校验通过：输入输出平衡，且 Result 1/2/3/4 内部勾稽完全一致。"
        else:
            return False, " | ".join(messages)

    @staticmethod
    def to_bytes(df):
        b = io.BytesIO()
        out = df.drop(columns=['_sys_id'], errors='ignore')
        with pd.ExcelWriter(b, engine='openpyxl') as writer:
            out.to_excel(writer, index=False, sheet_name='Sheet1')
            worksheet = writer.sheets['Sheet1']
            thin = Side(border_style="thin", color="000000")
            border = Border(top=thin, left=thin, right=thin, bottom=thin)
            align_center = Alignment(horizontal='center', vertical='center', wrap_text=False)
            header_font = Font(bold=True)
            for row in worksheet.iter_rows(min_row=1, max_row=len(out)+1, min_col=1, max_col=len(out.columns)):
                for cell in row:
                    cell.border = border
                    cell.alignment = align_center
                    if cell.row == 1: cell.font = header_font
            for i, col in enumerate(out.columns):
                max_len = 0
                try: max_len = len(str(col).encode('gbk'))
                except: max_len = len(str(col))
                for val in out[col]:
                    v_len = 0
                    try: v_len = len(str(val).encode('gbk'))
                    except: v_len = len(str(val))
                    if v_len > max_len: max_len = v_len
                adjusted_width = min((max_len + 2) * 1.1, 60) 
                worksheet.column_dimensions[get_column_letter(i + 1)].width = adjusted_width
        return b.getvalue()

class WordGenerator:
    @staticmethod
    def set_cell_style(cell, text, font_size=10, bold=False, align="center"):
        cell.text = ""
        paragraph = cell.paragraphs[0]
        if align == "center": paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        elif align == "left": paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
        elif align == "right": paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        run = paragraph.add_run(str(text))
        run.font.bold = bold
        run.font.size = Pt(font_size)
        try:
            run.font.name = 'SimSun' 
            run._element.rPr.rFonts.set(qn('w:eastAsia'), 'SimSun')
        except: pass
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    @staticmethod
    def set_row_height(row, height_cm):
        tr = row._tr
        trPr = tr.get_or_add_trPr()
        trHeight = OxmlElement('w:trHeight')
        trHeight.set(qn('w:val'), str(int(height_cm * 567))) 
        trHeight.set(qn('w:hRule'), "atLeast") 
        trPr.append(trHeight)

    @staticmethod
    def create_hardcoded_template(purchase_comp, sales_comp, dept_name, period_text):
        doc = docx.Document()
        section = doc.sections[0]
        section.top_margin = Cm(2.0); section.bottom_margin = Cm(2.0)
        section.left_margin = Cm(2.0); section.right_margin = Cm(2.0)

        title_line_1 = f"{purchase_comp}与云软件事业部-实施交付部"
        title_line_2 = f"{period_text}项目交付与运维费用结算账单"
        
        p1 = doc.add_paragraph()
        p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run1 = p1.add_run(title_line_1)
        run1.font.bold = True; run1.font.size = Pt(14)
        try: run1.font.name = 'SimSun'; run1._element.rPr.rFonts.set(qn('w:eastAsia'), 'SimSun')
        except: pass
        
        p2 = doc.add_paragraph()
        p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run2 = p2.add_run(title_line_2)
        run2.font.size = Pt(14)
        try: run2.font.name = 'SimSun'; run2._element.rPr.rFonts.set(qn('w:eastAsia'), 'SimSun')
        except: pass
        doc.add_paragraph() 

        table0 = doc.add_table(rows=6, cols=10)
        table0.style = 'Table Grid'
        col_widths = [1.3, 1.3, 1.3, 1.6, 1.6, 1.6, 1.6, 1.8, 1.3, 2.0]
        for row in table0.rows:
            for idx, width in enumerate(col_widths): row.cells[idx].width = Cm(width)

        WordGenerator.set_row_height(table0.rows[0], 1.63)
        c0 = table0.rows[0].cells[0].merge(table0.rows[0].cells[9])
        WordGenerator.set_cell_style(c0, title_line_1 + "\n" + title_line_2, font_size=10, bold=True)

        WordGenerator.set_row_height(table0.rows[1], 0.92)
        WordGenerator.set_cell_style(table0.rows[1].cells[0].merge(table0.rows[1].cells[2]), "工作量\n(单位: 人/天)")
        WordGenerator.set_cell_style(table0.rows[1].cells[3].merge(table0.rows[1].cells[5]), "人力费用\n(单位: 元)")
        WordGenerator.set_cell_style(table0.rows[1].cells[6].merge(table0.rows[1].cells[7]), "差旅费用\n(单位: 元)")
        WordGenerator.set_cell_style(table0.rows[1].cells[8].merge(table0.rows[1].cells[9]), "合计")

        WordGenerator.set_row_height(table0.rows[2], 0.92)
        headers = [(0,"项目标\n准交付"), (1,"项目数\n据治理"), (2,"项目运\n维服务"), (3,"项目标\n准交付"), (4,"项目数\n据治理"), (5,"项目运\n维服务"), (6,"差旅\n补助"), (7,"商旅平\n台费用"), (8,"工作量"), (9,"合计费用\n（单位:元）")]
        for idx, txt in headers: WordGenerator.set_cell_style(table0.rows[2].cells[idx], txt)

        WordGenerator.set_row_height(table0.rows[3], 0.92) # Data row

        WordGenerator.set_row_height(table0.rows[4], 1.13)
        WordGenerator.set_cell_style(table0.rows[4].cells[0].merge(table0.rows[4].cells[2]), "项目所属区域")
        WordGenerator.set_cell_style(table0.rows[4].cells[3].merge(table0.rows[4].cells[9]), str(dept_name), align="left")

        WordGenerator.set_row_height(table0.rows[5], 4.17)
        WordGenerator.set_cell_style(table0.rows[5].cells[0].merge(table0.rows[5].cells[2]), "项目所属\n区域销售\n确认")
        
        c_sign = table0.rows[5].cells[3].merge(table0.rows[5].cells[9])
        c_sign.text = ""
        p = c_sign.add_paragraph("确认意见：\n\n\n\n签字（签章）：\n\n\n")
        p.runs[0].font.size = Pt(10)
        try: p.runs[0].font.name = 'SimSun'; p.runs[0]._element.rPr.rFonts.set(qn('w:eastAsia'), 'SimSun')
        except: pass
        p_date = c_sign.add_paragraph("日期：    年    月    日        ")
        p_date.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        try: p_date.runs[0].font.name = 'SimSun'; p_date.runs[0]._element.rPr.rFonts.set(qn('w:eastAsia'), 'SimSun')
        except: pass

        doc.add_paragraph("\n")
        
        table1 = doc.add_table(rows=1, cols=11)
        table1.style = 'Table Grid'
        t1_widths = [0.9, 1.2, 1.5, 1.5, 1.0, 1.0, 1.0, 2.2, 2.3, 2.3, 2.6]
        for row in table1.rows:
            for idx, width in enumerate(t1_widths): row.cells[idx].width = Cm(width)
        
        headers_1 = ['人员', '人事\n范围', '项目\n名称', '合同\n名称', '销售\n人员', '销售所\n在大区', '支持\n人天', '人力\n费用', '差旅\n补助', '差旅平\n台费用', '总费用\n（元）']
        for i, text in enumerate(headers_1):
            WordGenerator.set_cell_style(table1.rows[0].cells[i], text)
            
        return doc

    @staticmethod
    def generate(df_result, period_text):
        if not HAS_DOCX: return {}, "缺少 python-docx 库"
        if df_result.empty: return {}, "结果数据为空"
        
        req_cols = ['合同主体', '人事范围', '销售部门']
        if not all(c in df_result.columns for c in req_cols): return {}, "缺少必要列"

        pairs = df_result[req_cols].dropna().drop_duplicates().values
        output_files = {}

        for purchase_comp, sales_comp, dept_name in pairs:
            df_curr = df_result[
                (df_result['合同主体'] == purchase_comp) & 
                (df_result['人事范围'] == sales_comp) &
                (df_result['销售部门'] == dept_name)
            ].copy()
            if df_curr.empty: continue

            doc = WordGenerator.create_hardcoded_template(purchase_comp, sales_comp, dept_name, period_text)
            table0 = doc.tables[0]
            
            total_days = df_curr['支持时间（人天）'].sum()
            total_labor = df_curr['人力费用'].sum()
            total_sub = df_curr['差旅补助'].sum()
            total_fee = df_curr['差旅费控平台'].sum()
            grand_total = df_curr['结算费用合计'].sum()
            
            cells = table0.rows[3].cells
            vals = [
                "{:,.1f}".format(total_days), "0.0", "0.0",
                "{:,.2f}".format(total_labor), "0.00", "0.00",
                "{:,.2f}".format(total_sub), "{:,.2f}".format(total_fee),
                "{:,.1f}".format(total_days), "{:,.2f}".format(grand_total)
            ]
            for i, v in enumerate(vals): WordGenerator.set_cell_style(cells[i], v)

            table1 = doc.tables[1]
            cols_map = ['人员', '人事范围', '所属项目', '合同主体', '销售人员', '销售部门', '支持时间（人天）', '人力费用', '差旅补助', '差旅费控平台', '结算费用合计']
            t1_widths = [0.9, 1.2, 1.5, 1.5, 1.0, 1.0, 1.0, 2.2, 2.3, 2.3, 2.6]
            
            for _, row in df_curr.iterrows():
                new_row = table1.add_row()
                WordGenerator.set_row_height(new_row, 1.0) 
                for idx, width in enumerate(t1_widths): new_row.cells[idx].width = Cm(width)
                for i, col_name in enumerate(cols_map):
                    val = row.get(col_name, '')
                    if isinstance(val, (int, float)):
                        text_val = "{:,.1f}".format(val) if '人天' in col_name else "{:,.2f}".format(val)
                    else: text_val = str(val)
                    WordGenerator.set_cell_style(new_row.cells[i], text_val)
            
            # ================================================================
            # 🛠️ 空白页修复：添加零高度段落
            # ================================================================
            last_p = doc.add_paragraph()
            p_fmt = last_p.paragraph_format
            p_fmt.space_before = Pt(0)
            p_fmt.space_after = Pt(0)
            p_fmt.line_spacing = Pt(0)
            p_fmt.line_spacing_rule = WD_LINE_SPACING.EXACTLY
            last_p.add_run().font.size = Pt(0)
            # ================================================================

            out = io.BytesIO()
            doc.save(out)
            safe_dept = str(dept_name).replace('/', '_').replace('\\', '_')
            fname = f"结算单_{purchase_comp}_{sales_comp}_{safe_dept}.docx"
            output_files[fname] = out.getvalue()

        return output_files, None

# ==============================================================================
# Zone B: UI 组件层 (修正列宽与垂直对齐)
# ==============================================================================
class UIComponents:
    @staticmethod
    def render_sidebar(has_error):
        with st.sidebar:
            st.header("⚙️ 参数配置")
            
            if 'params' not in st.session_state:
                st.session_state.params = {
                    'price': 1500, 'hours_limit': 100, 'sub_tag': "差旅补助", 
                    'period': "2025年第三季度"
                }

            st.session_state.params['price'] = st.number_input("人力单价 (元/天)", value=st.session_state.params['price'], step=100)
            if has_error: st.error("🚨 请调整工时", icon=None)
            st.session_state.params['hours_limit'] = st.number_input("工时阈值 (小时)", value=st.session_state.params['hours_limit'])
            st.session_state.params['sub_tag'] = st.text_input("补助关键词", value=st.session_state.params['sub_tag'])
            st.session_state.params['period'] = st.text_input("结算周期文案", value=st.session_state.params['period'])
            
            current_params = st.session_state.params.copy()
            last_run_params = st.session_state.get('last_run_params', None)
            has_files = st.session_state.data_store['A']['df'] is not None and st.session_state.data_store['B']['df'] is not None
            param_changed = last_run_params is not None and current_params != last_run_params
            
            trigger_recalc = False
            if has_files and param_changed:
                st.write("") 
                if st.button("重新运算", type="primary", use_container_width=True): trigger_recalc = True
            
            st.markdown("---")
            if st.button("🐱 字段映射 & 逻辑"):
                st.session_state.page = 'mapping'
                st.rerun()
                
            return current_params, trigger_recalc

    @staticmethod
    @st.cache_data(ttl=3600)
    def load_data_cached(file_content, file_name):
        try:
            if file_name.endswith('.csv'): df = pd.read_csv(file_content)
            else: df = pd.read_excel(file_content, engine='openpyxl')
            df.columns = [str(c).strip() for c in df.columns]
            df['_sys_id'] = range(1, len(df)+1)
            return df
        except Exception: return None

    @staticmethod
    def render_file_slot(key, title, data_store):
        data = data_store[key]
        has_file = data['df'] is not None
        
        with st.container(border=True):
            if not has_file:
                st.markdown(f"**{title}**")
                file = st.file_uploader(title, type=['xlsx', 'csv'], key=f"uploader_{key}", label_visibility="collapsed")
                
                if file:
                    df = UIComponents.load_data_cached(file, file.name)
                    if df is not None:
                        st.session_state.data_store[key] = {'df': df, 'name': file.name}
                        st.session_state.is_calculated = False
                        st.session_state.error_report = None
                        st.session_state.balance_check = (True, "")
                        st.session_state.last_run_params = None
                        st.rerun()
            else:
                # 调整列宽：0.15 (15%) 给按钮，防止被隐藏
                c_icon, c_info, c_close = st.columns([0.15, 0.70, 0.15], vertical_alignment="center")
                
                with c_icon:
                    st.markdown('<div class="file-icon">📄</div>', unsafe_allow_html=True)
                
                with c_info:
                    row_count = len(data['df'])
                    row_str = "{:,}".format(row_count)
                    st.markdown(f"""
                    <div style="display: flex; flex-direction: column; justify-content: center; height: 100%;">
                        <span class="file-name">{data['name']}</span>
                        <span class="file-stats">📊 已加载 {row_str} 条数据</span>
                    </div>
                    """, unsafe_allow_html=True)
                
                with c_close:
                    if st.button("✕", key=f"del_{key}", help="移除此文件", type="secondary"): 
                        st.session_state.data_store[key] = {'df': None, 'name': None}
                        st.session_state.is_calculated = False
                        st.session_state.error_report = None
                        st.session_state.balance_check = (True, "")
                        st.session_state.last_run_params = None
                        st.rerun()

    @staticmethod
    def render_error_report(err_df):
        fixable = err_df[err_df['类型']=='数据错误']
        rule = err_df[err_df['类型']=='业务规则校验']
        st.markdown(f"<div class='error-box'><h3 style='margin:0'>🚨 校验未通过</h3><p>发现 <b>{len(fixable)}</b> 个数据错误，<b>{len(rule)}</b> 个业务规则警告。</p></div>", unsafe_allow_html=True)
        st.dataframe(err_df[['类型','来源','行号','信息']], use_container_width=True, hide_index=True)
        st.download_button("📥 下载错误清单", err_df.to_csv(index=False).encode('utf-8-sig'), "err.csv", "text/csv", use_container_width=True)
        return any("阈值" in str(x) for x in err_df['信息'])

    @staticmethod
    def render_download_zone(result_files, all_in_one_zip, word_files_dict, period_str, balance_check):
        is_bal, bal_msg = balance_check
        
        css_class = "balance-box-ok" if is_bal else "balance-box-err"
        st.markdown(f"<div class='{css_class}'>{bal_msg}</div>", unsafe_allow_html=True)
        
        if not is_bal:
            st.warning("⚠️ 严重警告：总额或内部勾稽不平，请务必检查上方错误信息！")

        with st.container(border=True):
            st.success("✅ 计算完成 | 报表已生成")
            st.subheader("📦 批量下载")
            st.download_button("🚀 一键下载 (Excel + Word)", all_in_one_zip, "项目结算资料全集.zip", "application/zip", type="primary", use_container_width=True)
            st.divider()
            
            q_str = DataEngine.get_quarter_str(period_str)
            name_t1 = f"实施交付部项目情况汇总_部门工时统计-{q_str}.xlsx"
            name_t3 = f"实施交付部项目情况汇总_结算工时总表-{q_str}.xlsx"
            name_t2 = f"{q_str}实施交付部项目投入考核调整总表.xlsx"

            c1, c2, c3 = st.columns(3)
            if result_files and 't1' in result_files: c1.download_button(f"📥 {name_t1}", result_files['t1'], name_t1, use_container_width=True)
            if result_files and 't2' in result_files: c2.download_button(f"📥 {name_t2}", result_files['t2'], name_t2, use_container_width=True)
            if result_files and 't3' in result_files: c3.download_button(f"📥 {name_t3}", result_files['t3'], name_t3, use_container_width=True)
            
            if word_files_dict:
                with st.expander(f"📄 结算单 ({len(word_files_dict)})", expanded=False):
                    for fname, fbytes in word_files_dict.items():
                        c_t, c_b = st.columns([4, 1])
                        c_t.text(f"📄 {fname}")
                        c_b.download_button("下载", fbytes, fname, key=f"btn_{fname}")

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
if 'is_editing_mapping' not in st.session_state: st.session_state.is_editing_mapping = False
if 'all_zip' not in st.session_state: st.session_state.all_zip = None
if 'word_files' not in st.session_state: st.session_state.word_files = {}
if 'result_files' not in st.session_state: st.session_state.result_files = {}
if 'last_run_params' not in st.session_state: st.session_state.last_run_params = None
if 'threshold_error_flag' not in st.session_state: st.session_state.threshold_error_flag = False
if 'balance_check' not in st.session_state: st.session_state.balance_check = (True, "")

inject_css()

if st.session_state.page == 'main':
    current_params, manual_recalc = UIComponents.render_sidebar(st.session_state.threshold_error_flag)
    st.title("😈 淡藤财务报表 Pro")

    with st.container(border=True):
        c_h1, c_h2 = st.columns([0.88, 0.12], vertical_alignment="bottom")
        c_h1.markdown("### 📂 数据源控制台")
        
        if c_h2.button("🗑️ 重置", use_container_width=True): 
            st.session_state.data_store = {'A': {'df': None, 'name': None}, 'B': {'df': None, 'name': None}}
            st.session_state.is_calculated = False
            st.session_state.error_report = None
            st.session_state.all_zip = None
            st.session_state.last_run_params = None
            st.session_state.threshold_error_flag = False
            st.rerun()
        
        c1, c2 = st.columns(2)
        with c1: UIComponents.render_file_slot('A', "Source A: 投入明细 (工时)", st.session_state.data_store)
        with c2: UIComponents.render_file_slot('B', "Source B: 差旅明细 (费用)", st.session_state.data_store)

    st.divider()
    
    has_files = st.session_state.data_store['A']['df'] is not None and st.session_state.data_store['B']['df'] is not None
    should_calculate = False

    if has_files:
        if st.session_state.last_run_params is None: should_calculate = True
        elif manual_recalc: should_calculate = True
        
        if should_calculate:
            with st.spinner("🚀 正在校验并计算数据..."):
                st.session_state.threshold_error_flag = False
                errs, df_a, df_b = DataEngine.validate(
                    st.session_state.data_store['A']['df'].copy(),
                    st.session_state.data_store['B']['df'].copy(),
                    st.session_state.mapping_config,
                    current_params['hours_limit']
                )

                if errs:
                    st.session_state.error_report = pd.DataFrame(errs)
                    st.session_state.is_calculated = False
                    st.session_state.last_run_params = current_params.copy()
                    st.rerun()
                else:
                    # 1. 计算
                    res = DataEngine.calculate(df_a, df_b, st.session_state.mapping_config, current_params['price'], current_params['sub_tag'])
                    
                    # 2. 校验
                    st.session_state.balance_check = DataEngine.verify_balance(
                        df_a, df_b, res, st.session_state.mapping_config
                    )
                    
                    # 3. 结果生成
                    excel_files_dict = {
                        "t1": DataEngine.to_bytes(res['t1']),
                        "t2": DataEngine.to_bytes(res['t2']),
                        "t3": DataEngine.to_bytes(res['t3'])
                    }
                    st.session_state.result_files = excel_files_dict
                    word_files_dict, err_msg = WordGenerator.generate(res['t3'], current_params['period'])
                    if err_msg: st.warning(f"Word生成受限: {err_msg}")
                    st.session_state.word_files = word_files_dict
                    
                    all_files_to_zip = {}
                    q_str = DataEngine.get_quarter_str(current_params['period'])
                    all_files_to_zip[f"实施交付部项目情况汇总_部门工时统计-{q_str}.xlsx"] = excel_files_dict['t1']
                    all_files_to_zip[f"{q_str}实施交付部项目投入考核调整总表.xlsx"] = excel_files_dict['t2']
                    all_files_to_zip[f"实施交付部项目情况汇总_结算工时总表-{q_str}.xlsx"] = excel_files_dict['t3']
                    all_files_to_zip.update(word_files_dict)
                    
                    buf_zip = io.BytesIO()
                    with zipfile.ZipFile(buf_zip, 'w', zipfile.ZIP_DEFLATED) as z:
                        for fname, fcontent in all_files_to_zip.items(): z.writestr(fname, fcontent)
                    st.session_state.all_zip = buf_zip.getvalue()
                    
                    st.session_state.is_calculated = True
                    st.session_state.error_report = None
                    st.session_state.last_run_params = current_params.copy()
                    st.rerun()

    if st.session_state.error_report is not None:
        is_threshold_err = UIComponents.render_error_report(st.session_state.error_report)
        if is_threshold_err != st.session_state.threshold_error_flag:
            st.session_state.threshold_error_flag = is_threshold_err
            st.rerun()
            
    elif st.session_state.is_calculated:
        UIComponents.render_download_zone(
            st.session_state.result_files, 
            st.session_state.all_zip, 
            st.session_state.word_files, 
            current_params['period'],
            st.session_state.balance_check
        )

elif st.session_state.page == 'mapping':
    c1, c2 = st.columns([9, 1])
    c1.markdown("<div class='nav-header'>🐱 字段映射 & 逻辑配置</div>", unsafe_allow_html=True)
    if c2.button("⬅️", use_container_width=True): 
        st.session_state.page = 'main'
        st.rerun()
    st.markdown("<hr style='margin-top:0; border-color:#30363d;'>", unsafe_allow_html=True)
    c_action = st.columns([7, 1, 2])[2]
    
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
                st.session_state.error_report = None
                st.session_state.last_run_params = None
                st.rerun()

    df_c = st.session_state.mapping_config
    t1, t2, t3 = st.tabs(["结果表3 (底表)", "结果表2 (结算)", "结果表1 (工时)"])
    cols_a = list(st.session_state.data_store['A']['df'].columns) if has_files else []
    cols_b = list(st.session_state.data_store['B']['df'].columns) if has_files else []
    with t1: UIComponents.render_native_editor("全量明细底表", df_c[df_c['所属表']=='结果表3'], st.session_state.is_editing_mapping, cols_a, cols_b)
    with t2: UIComponents.render_native_editor("结算汇总表", df_c[df_c['所属表']=='结果表2'], st.session_state.is_editing_mapping, cols_a, cols_b)
    with t3: UIComponents.render_native_editor("工时统计表", df_c[df_c['所属表']=='结果表1'], st.session_state.is_editing_mapping, cols_a, cols_b)
