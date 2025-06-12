import xlsxwriter
import pandas as pd
from config import Config
from validation_rules.validation_core import get_validation_issues
from validation_rules.name_extraction import extract_name_from_report

def get_column_letter(df, column_name):
    col_idx = df.columns.get_loc(column_name) + 1
    return xlsxwriter.utility.xl_col_to_name(col_idx - 1)

def apply_format(worksheet, row_idx, col_letter, value, condition_met, format_obj):
    if condition_met:
        worksheet.write(f'{col_letter}{row_idx + 2}', value if pd.notna(value) else '', format_obj)
    else:
        worksheet.write(f'{col_letter}{row_idx + 2}', value if pd.notna(value) else '')

def format_excel(df, mismatch_indices, output_path, issues_list):
    with pd.ExcelWriter(output_path, engine='xlsxwriter', engine_kwargs={'options': {'nan_inf_to_errors': True}}) as writer:
        df_str = df.fillna('').astype(str)
        df_str.to_excel(writer, sheet_name='Sheet1', index=False)
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']
        for col in range(df_str.shape[1]):
            worksheet.set_column(col, col, None, workbook.add_format({'num_format': '@'}))
        red_format = workbook.add_format({'bg_color': Config.FORMATS["red"]})
        yellow_format = workbook.add_format({'bg_color': Config.FORMATS["yellow"]})

        # 独立处理每行样式
        for idx in range(len(df)):
            row = df.iloc[idx]
            # 受理时间
            if Config.COLUMN_MAPPINGS["acceptance_time"] in df.columns:
                value = row[Config.COLUMN_MAPPINGS["acceptance_time"]]
                apply_format(worksheet, idx, 'AF', value, any(issue == Config.VALIDATION_RULES["confirm_acceptance_time"] for i, issue in issues_list if i == idx), yellow_format)
            # 处置情况报告为空
            report_text = row["处置情况报告"] if "处置情况报告" in df.columns else ''
            if pd.isna(report_text):
                apply_format(worksheet, idx, 'AB', report_text, any(issue == Config.VALIDATION_RULES["empty_report"] for i, issue in issues_list if i == idx), yellow_format)

            # 检查“收缴金额（万元）”字段
            if "收缴金额（万元）" in df.columns:
                col_letter = get_column_letter(df, "收缴金额（万元）")
                value = row["收缴金额（万元）"]
                if any(issue == Config.VALIDATION_RULES["highlight_collection_amount"] for i, issue in issues_list if i == idx):
                    apply_format(worksheet, idx, col_letter, value, True, yellow_format)

            # 检查“没收金额”字段
            if "没收金额" in df.columns:
                col_letter = get_column_letter(df, "没收金额")
                value = row["没收金额"]
                if any(issue == Config.VALIDATION_RULES["highlight_confiscation_amount"] for i, issue in issues_list if i == idx):
                    apply_format(worksheet, idx, col_letter, value, True, yellow_format)

            # 检查“责令退赔金额”字段
            if "责令退赔金额" in df.columns:
                col_letter = get_column_letter(df, "责令退赔金额")
                value = row["责令退赔金额"]
                if any(issue == Config.VALIDATION_RULES["highlight_compensation_amount"] for i, issue in issues_list if i == idx):
                    apply_format(worksheet, idx, col_letter, value, True, yellow_format)

            # 检查“登记上交金额”字段
            if "登记上交金额" in df.columns:
                col_letter = get_column_letter(df, "登记上交金额")
                value = row["登记上交金额"]
                if any(issue == Config.VALIDATION_RULES["highlight_registration_amount"] for i, issue in issues_list if i == idx):
                    apply_format(worksheet, idx, col_letter, value, True, yellow_format)

        # 基于 mismatch_indices 和 issues_list 标红
        for idx in range(len(df)):
            row = df.iloc[idx]
            # agency
            if idx in mismatch_indices and any(issue == Config.VALIDATION_RULES["inconsistent_agency"] for _, issue in issues_list if _ == idx):
                apply_format(worksheet, idx, 'C', row["填报单位名称"], True, red_format)
                apply_format(worksheet, idx, 'H', row["办理机关"], True, red_format)
            # 被反映人
            if "被反映人" in df.columns:
                reported_person = str(row["被反映人"]).strip() if pd.notna(row["被反映人"]) else ''
                report_name = extract_name_from_report(report_text)
                if any(issue == Config.VALIDATION_RULES["inconsistent_name"] for _, issue in issues_list if _ == idx):
                    apply_format(worksheet, idx, 'E', row["被反映人"], True, red_format)
            # 组织措施
            if Config.COLUMN_MAPPINGS["organization_measure"] in df.columns:
                col_letter = get_column_letter(df, Config.COLUMN_MAPPINGS["organization_measure"])
                organization_measure = str(row[Config.COLUMN_MAPPINGS["organization_measure"]].strip()) if pd.notna(row[Config.COLUMN_MAPPINGS["organization_measure"]]) else ''
                if any(issue == Config.VALIDATION_RULES["inconsistent_organization_measure"] for _, issue in issues_list if _ == idx):
                    apply_format(worksheet, idx, col_letter, organization_measure, True, red_format)
            # 入党时间
            if Config.COLUMN_MAPPINGS["joining_party_time"] in df.columns:
                col_letter = get_column_letter(df, Config.COLUMN_MAPPINGS["joining_party_time"])
                joining_party_time = str(row[Config.COLUMN_MAPPINGS["joining_party_time"]].strip()) if pd.notna(row[Config.COLUMN_MAPPINGS["joining_party_time"]]) else ''
                if any(issue == Config.VALIDATION_RULES["inconsistent_joining_party_time"] for _, issue in issues_list if _ == idx):
                    apply_format(worksheet, idx, col_letter, joining_party_time, True, red_format)