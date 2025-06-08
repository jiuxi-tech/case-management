import xlsxwriter
import pandas as pd
from config import Config
from validation_rules import extract_name_from_report

def get_column_letter(df, column_name):
    col_idx = df.columns.get_loc(column_name) + 1
    return xlsxwriter.utility.xl_col_to_name(col_idx - 1)

def format_excel(df, mismatch_indices, output_path):
    with pd.ExcelWriter(output_path, engine='xlsxwriter', engine_kwargs={'options': {'nan_inf_to_errors': True}}) as writer:
        df.to_excel(writer, sheet_name='Sheet1', index=False)
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']
        red_format = workbook.add_format({'bg_color': Config.FORMATS["red"]})
        yellow_format = workbook.add_format({'bg_color': Config.FORMATS["yellow"]})

        for idx in mismatch_indices:
            row = df.iloc[idx]
            # 填报单位名称 vs 办理机关
            worksheet.write(f'C{idx + 2}', str(row["填报单位名称"]) if pd.notna(row["填报单位名称"]) else '', red_format)
            worksheet.write(f'H{idx + 2}', str(row["办理机关"]) if pd.notna(row["办理机关"]) else '', red_format)
            # 被反映人 vs 处置情况报告
            if "处置情况报告" in df.columns:
                report_text = row["处置情况报告"]
                if pd.isna(report_text):
                    worksheet.write(f'AB{idx + 2}', str(report_text) if pd.notna(report_text) else '', yellow_format)
                if "被反映人" in df.columns:
                    reported_person = str(row["被反映人"]).strip() if pd.notna(row["被反映人"]) else ''
                    report_name = extract_name_from_report(report_text)
                    if reported_person and report_name and reported_person != report_name:
                        worksheet.write(f'E{idx + 2}', str(row["被反映人"]) if pd.notna(row["被反映人"]) else '', red_format)
            # 受理时间
            if Config.COLUMN_MAPPINGS["acceptance_time"] in df.columns and pd.notna(row[Config.COLUMN_MAPPINGS["acceptance_time"]]):
                worksheet.write(f'AF{idx + 2}', str(row[Config.COLUMN_MAPPINGS["acceptance_time"]]) if pd.notna(row[Config.COLUMN_MAPPINGS["acceptance_time"]]) else '', yellow_format)
            # 组织措施
            if Config.COLUMN_MAPPINGS["organization_measure"] in df.columns:
                col_letter = get_column_letter(df, Config.COLUMN_MAPPINGS["organization_measure"])
                organization_measure = str(row[Config.COLUMN_MAPPINGS["organization_measure"]]).strip() if pd.notna(row[Config.COLUMN_MAPPINGS["organization_measure"]]) else ''
                report_text = str(row["处置情况报告"]).strip() if pd.notna(row["处置情况报告"]) else ''
                if not organization_measure or organization_measure not in Config.ORGANIZATION_MEASURES or (organization_measure and organization_measure not in report_text):
                    worksheet.write(f'{col_letter}{idx + 2}', str(row[Config.COLUMN_MAPPINGS["organization_measure"]]) if pd.notna(row[Config.COLUMN_MAPPINGS["organization_measure"]]) else '', red_format)