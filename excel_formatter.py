import xlsxwriter
import pandas as pd
from config import Config
from validation_rules.validation_core import get_validation_issues
from validation_rules.name_extraction import extract_name_from_report

def get_column_letter(df, column_name):
    # 获取列的索引，并转换为Excel列字母（1-based index）
    col_idx = df.columns.get_loc(column_name) + 1
    return xlsxwriter.utility.xl_col_to_name(col_idx - 1)

def apply_format(worksheet, row_idx, col_letter, value, condition_met, format_obj):
    """
    根据条件判断是否应用格式。
    row_idx 是 DataFrame 的索引 (0-based)，Excel 行号是 row_idx + 2 (因为有表头和Pandas的默认0行)
    """
    excel_row = row_idx + 2  # Excel行号从1开始，且有表头，所以加2
    if condition_met:
        # 如果条件满足，写入带格式的值
        worksheet.write(f'{col_letter}{excel_row}', value if pd.notna(value) else '', format_obj)
    else:
        # 如果条件不满足，写入不带格式的值
        worksheet.write(f'{col_letter}{excel_row}', value if pd.notna(value) else '')

def format_excel(df, mismatch_indices, output_path, issues_list, gender_mismatch_indices=set(), age_mismatch_indices=set(),
                 birth_date_mismatch_indices=set(), education_mismatch_indices=set(), ethnicity_mismatch_indices=set(),
                 party_member_mismatch_indices=set(), party_joining_date_mismatch_indices=set(),
                 brief_case_details_mismatch_indices=set(), filing_time_mismatch_indices=set(),
                 disciplinary_committee_filing_time_mismatch_indices=set(),
                 disciplinary_committee_filing_authority_mismatch_indices=set(),
                 supervisory_committee_filing_time_mismatch_indices=set(),
                 supervisory_committee_filing_authority_mismatch_indices=set(),
                 case_report_keyword_mismatch_indices=set(), disposal_spirit_mismatch_indices=set(),
                 voluntary_confession_highlight_indices=set(), closing_time_mismatch_indices=set()):
    """
    格式化Excel文件，根据验证问题对单元格进行着色。
    df: 原始DataFrame
    mismatch_indices: 包含不一致行的索引的集合
    output_path: 输出Excel文件的路径
    issues_list: 从 validation_core.py 或 case_validators.py 获取的问题列表，
                 每个元素是 (original_df_index, clue_code_value, issue_description)
    其他索引集合: 用于标红或标黄特定字段
    """
    with pd.ExcelWriter(output_path, engine='xlsxwriter', engine_kwargs={'options': {'nan_inf_to_errors': True}}) as writer:
        # 将DataFrame写入Excel，并处理NaN值为Dask DataFrame期望的空字符串
        df_str = df.fillna('').astype(str)
        df_str.to_excel(writer, sheet_name='Sheet1', index=False)
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']

        # 设置列格式，确保所有单元格都以文本格式写入
        for col in range(df_str.shape[1]):
            worksheet.set_column(col, col, None, workbook.add_format({'num_format': '@'}))

        # 定义红色和黄色格式
        red_format = workbook.add_format({'bg_color': Config.FORMATS["red"]})
        yellow_format = workbook.add_format({'bg_color': Config.FORMATS["yellow"]})

        # 遍历DataFrame的每一行，应用黄色和红色格式
        for idx in range(len(df)):
            row = df.iloc[idx]

            # --- 应用黄色格式的检查 ---
            # 受理时间
            if Config.COLUMN_MAPPINGS["acceptance_time"] in df.columns:
                value = row[Config.COLUMN_MAPPINGS["acceptance_time"]]
                condition = any(issue_desc == Config.VALIDATION_RULES["confirm_acceptance_time"] for i, _, issue_desc in issues_list if i == idx)
                apply_format(worksheet, idx, get_column_letter(df, Config.COLUMN_MAPPINGS["acceptance_time"]), value, condition, yellow_format)

            # 处置情况报告为空
            report_text = row.get("处置情况报告", '')
            if pd.isna(report_text) or report_text == '':
                condition = any(issue_desc == Config.VALIDATION_RULES["empty_report"] for i, _, issue_desc in issues_list if i == idx)
                apply_format(worksheet, idx, get_column_letter(df, "处置情况报告"), report_text, condition, yellow_format)

            # 检查金额字段
            for field, rule in [
                ("收缴金额（万元）", Config.VALIDATION_RULES["highlight_collection_amount"]),
                ("没收金额", Config.VALIDATION_RULES["highlight_confiscation_amount"]),
                ("责令退赔金额", Config.VALIDATION_RULES["highlight_compensation_amount"]),
                ("登记上交金额", Config.VALIDATION_RULES["highlight_registration_amount"]),
                ("追缴失职渎职滥用职权造成的损失金额", Config.VALIDATION_RULES["highlight_recovery_amount"])
            ]:
                if field in df.columns:
                    col_letter = get_column_letter(df, field)
                    value = row[field]
                    condition = any(issue_desc == rule for i, _, issue_desc in issues_list if i == idx)
                    apply_format(worksheet, idx, col_letter, value, condition, yellow_format)

            # --- 应用红色格式的检查 ---
            # agency
            if idx in mismatch_indices and any(issue_desc == Config.VALIDATION_RULES["inconsistent_agency"] for i, _, issue_desc in issues_list if i == idx):
                apply_format(worksheet, idx, get_column_letter(df, "填报单位名称"), row["填报单位名称"], True, red_format)
                apply_format(worksheet, idx, get_column_letter(df, "办理机关"), row["办理机关"], True, red_format)

            # 被反映人
            if "被反映人" in df.columns and any(issue_desc == Config.VALIDATION_RULES["inconsistent_name"] for i, _, issue_desc in issues_list if i == idx):
                apply_format(worksheet, idx, get_column_letter(df, "被反映人"), row["被反映人"], True, red_format)

            # 组织措施
            if Config.COLUMN_MAPPINGS["organization_measure"] in df.columns and \
               any(issue_desc == Config.VALIDATION_RULES["inconsistent_organization_measure"] for i, _, issue_desc in issues_list if i == idx):
                col_letter = get_column_letter(df, Config.COLUMN_MAPPINGS["organization_measure"])
                apply_format(worksheet, idx, col_letter, row[Config.COLUMN_MAPPINGS["organization_measure"]], True, red_format)

            # 入党时间
            if Config.COLUMN_MAPPINGS["joining_party_time"] in df.columns and \
               any(issue_desc == Config.VALIDATION_RULES["inconsistent_joining_party_time"] for i, _, issue_desc in issues_list if i == idx):
                col_letter = get_column_letter(df, Config.COLUMN_MAPPINGS["joining_party_time"])
                apply_format(worksheet, idx, col_letter, row[Config.COLUMN_MAPPINGS["joining_party_time"]], True, red_format)

            # 民族
            if Config.COLUMN_MAPPINGS["ethnicity"] in df.columns and \
               any(issue_desc == Config.VALIDATION_RULES["inconsistent_ethnicity"] for i, _, issue_desc in issues_list if i == idx):
                col_letter = get_column_letter(df, Config.COLUMN_MAPPINGS["ethnicity"])
                apply_format(worksheet, idx, col_letter, row[Config.COLUMN_MAPPINGS["ethnicity"]], True, red_format)

            # 出生年月
            if Config.COLUMN_MAPPINGS["birth_date"] in df.columns and \
               any(issue_desc == Config.VALIDATION_RULES["highlight_birth_date"] for i, _, issue_desc in issues_list if i == idx):
                col_letter = get_column_letter(df, Config.COLUMN_MAPPINGS["birth_date"])
                apply_format(worksheet, idx, col_letter, row[Config.COLUMN_MAPPINGS["birth_date"]], True, red_format)

            # 办结时间
            if Config.COLUMN_MAPPINGS["completion_time"] in df.columns and \
               any(issue_desc == Config.VALIDATION_RULES["highlight_completion_time"] for i, _, issue_desc in issues_list if i == idx):
                col_letter = get_column_letter(df, Config.COLUMN_MAPPINGS["completion_time"])
                apply_format(worksheet, idx, col_letter, row[Config.COLUMN_MAPPINGS["completion_time"]], True, red_format)

            # 处置方式1二级
            if Config.COLUMN_MAPPINGS["disposal_method_1"] in df.columns and \
               any(issue_desc == Config.VALIDATION_RULES["highlight_disposal_method_1"] for i, _, issue_desc in issues_list if i == idx):
                col_letter = get_column_letter(df, Config.COLUMN_MAPPINGS["disposal_method_1"])
                apply_format(worksheet, idx, col_letter, row[Config.COLUMN_MAPPINGS["disposal_method_1"]], True, yellow_format)

            # 新增：是否违反中央八项规定精神
            if "是否违反中央八项规定精神" in df.columns and idx in disposal_spirit_mismatch_indices:
                col_letter = get_column_letter(df, "是否违反中央八项规定精神")
                apply_format(worksheet, idx, col_letter, row["是否违反中央八项规定精神"], True, red_format)

            # 新增：是否主动交代问题
            if "是否主动交代问题" in df.columns and idx in voluntary_confession_highlight_indices:
                col_letter = get_column_letter(df, "是否主动交代问题")
                apply_format(worksheet, idx, col_letter, row["是否主动交代问题"], True, yellow_format)

            # 新增：结案时间
            if "结案时间" in df.columns and idx in closing_time_mismatch_indices:
                col_letter = get_column_letter(df, "结案时间")
                apply_format(worksheet, idx, col_letter, row["结案时间"], True, red_format)
