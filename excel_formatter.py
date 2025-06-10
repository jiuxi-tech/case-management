import xlsxwriter
import pandas as pd
from config import Config
from validation_rules import extract_name_from_report

def get_column_letter(df, column_name):
    col_idx = df.columns.get_loc(column_name) + 1
    return xlsxwriter.utility.xl_col_to_name(col_idx - 1)

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

        # 处理所有行，基于 issues_list 进行独立标注（如受理时间）
        for idx in range(len(df)):
            row = df.iloc[idx]
            # 受理时间标注
            if Config.COLUMN_MAPPINGS["acceptance_time"] in df.columns and pd.notna(row[Config.COLUMN_MAPPINGS["acceptance_time"]]):
                if any(issue == Config.VALIDATION_RULES["confirm_acceptance_time"] for i, issue in issues_list if i == idx):
                    worksheet.write(f'AF{idx + 2}', str(row[Config.COLUMN_MAPPINGS["acceptance_time"]]) if pd.notna(row[Config.COLUMN_MAPPINGS["acceptance_time"]]) else '', yellow_format)
                else:
                    worksheet.write(f'AF{idx + 2}', str(row[Config.COLUMN_MAPPINGS["acceptance_time"]]) if pd.notna(row[Config.COLUMN_MAPPINGS["acceptance_time"]]) else '')
            # 处置情况报告为空时标注
            report_text = row["处置情况报告"] if "处置情况报告" in df.columns else ''
            if pd.isna(report_text):
                worksheet.write(f'AB{idx + 2}', str(report_text) if pd.notna(report_text) else '', yellow_format)

        # 处理 mismatch_indices 内的行，基于校验结果标红
        for idx in mismatch_indices:
            row = df.iloc[idx]
            # agency 相关标注
            if any(issue == Config.VALIDATION_RULES["inconsistent_agency"] for _, issue in issues_list if _ == idx):
                worksheet.write(f'C{idx + 2}', str(row["填报单位名称"]) if pd.notna(row["填报单位名称"]) else '', red_format)
                worksheet.write(f'H{idx + 2}', str(row["办理机关"]) if pd.notna(row["办理机关"]) else '', red_format)
            else:
                worksheet.write(f'C{idx + 2}', str(row["填报单位名称"]) if pd.notna(row["填报单位名称"]) else '')
                worksheet.write(f'H{idx + 2}', str(row["办理机关"]) if pd.notna(row["办理机关"]) else '')
            # 被反映人标注
            if "被反映人" in df.columns:
                reported_person = str(row["被反映人"]).strip() if pd.notna(row["被反映人"]) else ''
                report_name = extract_name_from_report(report_text)
                if any(issue == Config.VALIDATION_RULES["inconsistent_name"] for _, issue in issues_list if _ == idx):
                    worksheet.write(f'E{idx + 2}', str(row["被反映人"]) if pd.notna(row["被反映人"]) else '', red_format)
                else:
                    worksheet.write(f'E{idx + 2}', str(row["被反映人"]) if pd.notna(row["被反映人"]) else '')
            # 组织措施标注
            if Config.COLUMN_MAPPINGS["organization_measure"] in df.columns:
                col_letter = get_column_letter(df, Config.COLUMN_MAPPINGS["organization_measure"])
                organization_measure = str(row[Config.COLUMN_MAPPINGS["organization_measure"]]).strip() if pd.notna(row[Config.COLUMN_MAPPINGS["organization_measure"]]) else ''
                if any(issue == Config.VALIDATION_RULES["inconsistent_organization_measure"] for _, issue in issues_list if _ == idx):
                    worksheet.write(f'{col_letter}{idx + 2}', organization_measure, red_format)
                else:
                    worksheet.write(f'{col_letter}{idx + 2}', organization_measure)
            # 入党时间标注
            if Config.COLUMN_MAPPINGS["joining_party_time"] in df.columns:
                col_letter = get_column_letter(df, Config.COLUMN_MAPPINGS["joining_party_time"])
                joining_party_time = str(row[Config.COLUMN_MAPPINGS["joining_party_time"]]).strip() if pd.notna(row[Config.COLUMN_MAPPINGS["joining_party_time"]]) else ''
                if any(issue == Config.VALIDATION_RULES["inconsistent_joining_party_time"] for _, issue in issues_list if _ == idx):
                    worksheet.write(f'{col_letter}{idx + 2}', joining_party_time, red_format)
                else:
                    worksheet.write(f'{col_letter}{idx + 2}', joining_party_time)