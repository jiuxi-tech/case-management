import xlsxwriter
import pandas as pd
from config import Config
import logging # Added import for logging

logger = logging.getLogger(__name__) # Added logger initialization

def get_column_letter(df, column_name):
    """
    Gets the column index and converts it to an Excel column letter (1-based index).
    """
    col_idx = df.columns.get_loc(column_name) + 1
    return xlsxwriter.utility.xl_col_to_name(col_idx - 1)

def apply_format(worksheet, row_idx, col_letter, value, condition_met, format_obj):
    """
    Applies formatting based on a condition.
    row_idx is the DataFrame's index (0-based), Excel row number is row_idx + 2 (due to header and Pandas' default 0-row).
    """
    excel_row = row_idx + 2       # Excel rows start from 1, and there's a header, so add 2
    if condition_met:
        # If the condition is met, write the value with formatting
        worksheet.write(f'{col_letter}{excel_row}', value if pd.notna(value) else '', format_obj)
    else:
        # If the condition is not met, write the value without formatting
        worksheet.write(f'{col_letter}{excel_row}', value if pd.notna(value) else '')

def format_excel(df, mismatch_indices, output_path, issues_list,
                 gender_mismatch_indices=set(), age_mismatch_indices=set(),
                 birth_date_mismatch_indices=set(), education_mismatch_indices=set(), ethnicity_mismatch_indices=set(),
                 party_member_mismatch_indices=set(), party_joining_date_mismatch_indices=set(),
                 brief_case_details_mismatch_indices=set(), filing_time_mismatch_indices=set(),
                 disciplinary_committee_filing_time_mismatch_indices=set(),
                 disciplinary_committee_filing_authority_mismatch_indices=set(),
                 supervisory_committee_filing_time_mismatch_indices=set(),
                 supervisory_committee_filing_authority_mismatch_indices=set(),
                 case_report_keyword_mismatch_indices=set(), disposal_spirit_mismatch_indices=set(),
                 voluntary_confession_highlight_indices=set(), closing_time_mismatch_indices=set(),
                 no_party_position_warning_mismatch_indices=set(), 
                 recovery_amount_highlight_indices=set(), 
                 trial_acceptance_time_mismatch_indices=set(), 
                 trial_closing_time_mismatch_indices=set(), 
                 trial_authority_agency_mismatch_indices=set(),
                 disposal_decision_keyword_mismatch_indices=set(), 
                 trial_report_non_representative_mismatch_indices=set(), 
                 trial_report_detention_mismatch_indices=set(),
                 confiscation_amount_indices=set(), 
                 confiscation_of_property_amount_indices=set(), 
                 compensation_amount_highlight_indices=set(),
                 registered_handover_amount_indices=set(),
                 # --- 【新增】这里添加 disciplinary_sanction_mismatch_indices 参数 ---
                 disciplinary_sanction_mismatch_indices=set() 
                 ):
    """
    Formats the Excel file, coloring cells based on validation issues.
    df: Original DataFrame
    mismatch_indices: Set of indices for inconsistent rows
    output_path: Path for the output Excel file
    issues_list: List of issues obtained from validation_core.py or case_validators.py,
                 each element might be (original_df_index, clue_code_value, issue_description) (3 values)
                 or (original_df_index, case_code_value, person_code_value, issue_description) (4 values)
    Other index sets: Used for red or yellow highlighting of specific fields
    confiscation_amount_indices: Set of row indices for "收缴金额（万元）" to be highlighted.
    confiscation_of_property_amount_indices: Set of row indices for "没收金额" to be highlighted.
    compensation_amount_highlight_indices: Set of row indices for "责令退赔金额" to be highlighted.
    registered_handover_amount_indices: Set of row indices for "登记上交金额" to be highlighted.
    disciplinary_sanction_mismatch_indices: Set of row indices for "党纪处分" to be highlighted. # <--- NEW: Parameter description
    """
    try:
        with pd.ExcelWriter(output_path, engine='xlsxwriter', engine_kwargs={'options': {'nan_inf_to_errors': True}}) as writer:
            df_str = df.fillna('').astype(str)
            df_str.to_excel(writer, sheet_name='Sheet1', index=False)
            workbook = writer.book
            worksheet = writer.sheets['Sheet1']

            for col in range(df_str.shape[1]):
                worksheet.set_column(col, col, None, workbook.add_format({'num_format': '@'}))

            red_format = workbook.add_format({'bg_color': Config.FORMATS["red"]})
            yellow_format = workbook.add_format({'bg_color': Config.FORMATS["yellow"]})
            
            # 【关键修改】判断 issues_list 中元素的结构
            is_case_table_issues = False
            # 这里的判断应该更健壮，考虑 issues_list 可能为空的情况
            if issues_list and isinstance(issues_list[0], dict): # 假设现在 issues_list 元素是字典
                # 检查字典中是否存在涉案人员编码，这通常是案件表的特征
                if "涉案人员编码" in issues_list[0]:
                    is_case_table_issues = True
            elif issues_list and isinstance(issues_list[0], tuple) and len(issues_list[0]) == 4: # 兼容旧的元组格式（4个元素）
                is_case_table_issues = True


            for idx in range(len(df)):
                row = df.iloc[idx]

                # --- Apply yellow formatting checks (mainly for clue table) ---
                # Acceptance Time (Clue Table)
                if Config.COLUMN_MAPPINGS.get("acceptance_time") in df.columns: 
                    value = row.get(Config.COLUMN_MAPPINGS["acceptance_time"])
                    condition = False
                    if issues_list and isinstance(issues_list[0], dict): # 优先处理字典格式
                         condition = any(issue_item.get('问题描述') == Config.VALIDATION_RULES["confirm_acceptance_time"] and issue_item.get('行号', 0) - 2 == idx for issue_item in issues_list)
                    elif issues_list and isinstance(issues_list[0], tuple): # 兼容元组格式
                        if is_case_table_issues: # 4个元素的元组
                            condition = any(issue_desc == Config.VALIDATION_RULES["confirm_acceptance_time"] for i, _, _, issue_desc in issues_list if i == idx)
                        else: # 3个元素的元组 (线索表)
                            condition = any(issue_desc == Config.VALIDATION_RULES["confirm_acceptance_time"] for i, _, issue_desc in issues_list if i == idx)
                    apply_format(worksheet, idx, get_column_letter(df, Config.COLUMN_MAPPINGS["acceptance_time"]), value, condition, yellow_format)

                # Empty Disposal Report (Clue Table)
                report_text = row.get("处置情况报告", '')
                if "处置情况报告" in df.columns and (pd.isna(report_text) or report_text == ''):
                    condition = False
                    if issues_list and isinstance(issues_list[0], dict): # 优先处理字典格式
                        condition = any(issue_item.get('问题描述') == Config.VALIDATION_RULES["empty_report"] and issue_item.get('行号', 0) - 2 == idx for issue_item in issues_list)
                    elif issues_list and isinstance(issues_list[0], tuple): # 兼容元组格式
                        if is_case_table_issues:
                            condition = any(issue_desc == Config.VALIDATION_RULES["empty_report"] for i, _, _, issue_desc in issues_list if i == idx)
                        else: # Clue table scenario
                            condition = any(issue_desc == Config.VALIDATION_RULES["empty_report"] for i, _, issue_desc in issues_list if i == idx)
                    apply_format(worksheet, idx, get_column_letter(df, "处置情况报告"), report_text, condition, yellow_format)

                # Check amount fields (Clue Table) - **【CORE CHANGE: REMOVED "没收金额" FROM THIS LOOP】**
                for field, rule in [
                    ("收缴金额（万元）", Config.VALIDATION_RULES["highlight_collection_amount"]),
                    ("责令退赔金额", Config.VALIDATION_RULES["highlight_compensation_amount"]),
                    ("登记上交金额", Config.VALIDATION_RULES["highlight_registration_amount"]) 
                ]:
                    if field in df.columns:
                        col_letter = get_column_letter(df, field)
                        value = row.get(field)
                        condition = False
                        if issues_list and isinstance(issues_list[0], dict):
                            condition = any(issue_item.get('问题描述') == rule and issue_item.get('行号', 0) - 2 == idx for issue_item in issues_list)
                        elif issues_list and isinstance(issues_list[0], tuple):
                            if is_case_table_issues:
                                condition = any(issue_desc == rule for i, _, _, issue_desc in issues_list if i == idx)
                            else: # Clue table scenario
                                condition = any(issue_desc == rule for i, _, issue_desc in issues_list if i == idx)
                        apply_format(worksheet, idx, col_letter, value, condition, yellow_format)

                # --- Apply red formatting checks (mainly for clue table) ---
                # Agency (Clue Table)
                if "填报单位名称" in df.columns and "办理机关" in df.columns:
                    condition = False
                    if issues_list and isinstance(issues_list[0], dict):
                        condition = any(issue_item.get('问题描述') == Config.VALIDATION_RULES["inconsistent_agency"] and issue_item.get('行号', 0) - 2 == idx for issue_item in issues_list)
                    elif issues_list and isinstance(issues_list[0], tuple):
                        if is_case_table_issues:
                            condition = any(issue_desc == Config.VALIDATION_RULES["inconsistent_agency"] for i, _, _, issue_desc in issues_list if i == idx)
                        else:
                            condition = any(issue_desc == Config.VALIDATION_RULES["inconsistent_agency"] for i, _, issue_desc in issues_list if i == idx)
                    if condition:
                        apply_format(worksheet, idx, get_column_letter(df, "填报单位名称"), row.get("填报单位名称"), True, red_format)
                        apply_format(worksheet, idx, get_column_letter(df, "办理机关"), row.get("办理机关"), True, red_format)

                # Reflected Person (Clue Table)
                if "被反映人" in df.columns:
                    condition = False
                    if issues_list and isinstance(issues_list[0], dict):
                        condition = any(issue_item.get('问题描述') == Config.VALIDATION_RULES["inconsistent_name"] and issue_item.get('行号', 0) - 2 == idx for issue_item in issues_list)
                    elif issues_list and isinstance(issues_list[0], tuple):
                        if is_case_table_issues:
                            condition = any(issue_desc == Config.VALIDATION_RULES["inconsistent_name"] for i, _, _, issue_desc in issues_list if i == idx)
                        else:
                            condition = any(issue_desc == Config.VALIDATION_RULES["inconsistent_name"] for i, _, issue_desc in issues_list if i == idx)
                    if condition:
                        apply_format(worksheet, idx, get_column_letter(df, "被反映人"), row.get("被反映人"), True, red_format)

                # Organizational Measures (Clue Table)
                if Config.COLUMN_MAPPINGS.get("organization_measure") in df.columns:
                    condition = False
                    if issues_list and isinstance(issues_list[0], dict):
                        condition = any(issue_item.get('问题描述') == Config.VALIDATION_RULES["inconsistent_organization_measure"] and issue_item.get('行号', 0) - 2 == idx for issue_item in issues_list)
                    elif issues_list and isinstance(issues_list[0], tuple):
                        if is_case_table_issues:
                            condition = any(issue_desc == Config.VALIDATION_RULES["inconsistent_organization_measure"] for i, _, _, issue_desc in issues_list if i == idx)
                        else:
                            condition = any(issue_desc == Config.VALIDATION_RULES["inconsistent_organization_measure"] for i, _, issue_desc in issues_list if i == idx)
                    if condition:
                        col_letter = get_column_letter(df, Config.COLUMN_MAPPINGS["organization_measure"])
                        apply_format(worksheet, idx, col_letter, row.get(Config.COLUMN_MAPPINGS["organization_measure"]), True, red_format)

                # Party Joining Time (Clue Table)
                if Config.COLUMN_MAPPINGS.get("joining_party_time") in df.columns:
                    condition = False
                    if issues_list and isinstance(issues_list[0], dict):
                        condition = any(issue_item.get('问题描述') == Config.VALIDATION_RULES["inconsistent_joining_party_time"] and issue_item.get('行号', 0) - 2 == idx for issue_item in issues_list)
                    elif issues_list and isinstance(issues_list[0], tuple):
                        if is_case_table_issues:
                            condition = any(issue_desc == Config.VALIDATION_RULES["inconsistent_joining_party_time"] for i, _, _, issue_desc in issues_list if i == idx)
                        else:
                            condition = any(issue_desc == Config.VALIDATION_RULES["inconsistent_joining_party_time"] for i, _, issue_desc in issues_list if i == idx)
                    if condition:
                        col_letter = get_column_letter(df, Config.COLUMN_MAPPINGS["joining_party_time"])
                        apply_format(worksheet, idx, col_letter, row.get(Config.COLUMN_MAPPINGS["joining_party_time"]), True, red_format)

                # Ethnicity (Clue Table)
                if Config.COLUMN_MAPPINGS.get("ethnicity") in df.columns:
                    condition = False
                    if issues_list and isinstance(issues_list[0], dict):
                        condition = any(issue_item.get('问题描述') == Config.VALIDATION_RULES["inconsistent_ethnicity"] and issue_item.get('行号', 0) - 2 == idx for issue_item in issues_list)
                    elif issues_list and isinstance(issues_list[0], tuple):
                        if is_case_table_issues:
                            condition = any(issue_desc == Config.VALIDATION_RULES["inconsistent_ethnicity"] for i, _, _, issue_desc in issues_list if i == idx)
                        else:
                            condition = any(issue_desc == Config.VALIDATION_RULES["inconsistent_ethnicity"] for i, _, issue_desc in issues_list if i == idx)
                    if condition:
                        col_letter = get_column_letter(df, Config.COLUMN_MAPPINGS["ethnicity"])
                        apply_format(worksheet, idx, col_letter, row.get(Config.COLUMN_MAPPINGS["ethnicity"]), True, red_format)

                # Birth Date (Clue Table)
                if Config.COLUMN_MAPPINGS.get("birth_date") in df.columns:
                    condition = False
                    if issues_list and isinstance(issues_list[0], dict):
                        condition = any(issue_item.get('问题描述') == Config.VALIDATION_RULES["highlight_birth_date"] and issue_item.get('行号', 0) - 2 == idx for issue_item in issues_list)
                    elif issues_list and isinstance(issues_list[0], tuple):
                        if is_case_table_issues:
                            condition = any(issue_desc == Config.VALIDATION_RULES["highlight_birth_date"] for i, _, _, issue_desc in issues_list if i == idx)
                        else:
                            condition = any(issue_desc == Config.VALIDATION_RULES["highlight_birth_date"] for i, _, issue_desc in issues_list if i == idx)
                    if condition:
                        col_letter = get_column_letter(df, Config.COLUMN_MAPPINGS["birth_date"])
                        apply_format(worksheet, idx, col_letter, row.get(Config.COLUMN_MAPPINGS["birth_date"]), True, red_format)

                # Completion Time (Clue Table)
                if Config.COLUMN_MAPPINGS.get("completion_time") in df.columns:
                    condition = False
                    if issues_list and isinstance(issues_list[0], dict):
                        condition = any(issue_item.get('问题描述') == Config.VALIDATION_RULES["highlight_completion_time"] and issue_item.get('行号', 0) - 2 == idx for issue_item in issues_list)
                    elif issues_list and isinstance(issues_list[0], tuple):
                        if is_case_table_issues:
                            condition = any(issue_desc == Config.VALIDATION_RULES["highlight_completion_time"] for i, _, _, issue_desc in issues_list if i == idx)
                        else:
                            condition = any(issue_desc == Config.VALIDATION_RULES["highlight_completion_time"] for i, _, issue_desc in issues_list if i == idx)
                    if condition:
                        col_letter = get_column_letter(df, Config.COLUMN_MAPPINGS["completion_time"])
                        apply_format(worksheet, idx, col_letter, row.get(Config.COLUMN_MAPPINGS["completion_time"]), True, red_format)

                # Disposal Method 1 Secondary (Clue Table)
                if Config.COLUMN_MAPPINGS.get("disposal_method_1") in df.columns:
                    condition = False
                    if issues_list and isinstance(issues_list[0], dict):
                        condition = any(issue_item.get('问题描述') == Config.VALIDATION_RULES["highlight_disposal_method_1"] and issue_item.get('行号', 0) - 2 == idx for issue_item in issues_list)
                    elif issues_list and isinstance(issues_list[0], tuple):
                        if is_case_table_issues:
                            condition = any(issue_desc == Config.VALIDATION_RULES["highlight_disposal_method_1"] for i, _, _, issue_desc in issues_list if i == idx)
                        else:
                            condition = any(issue_desc == Config.VALIDATION_RULES["highlight_disposal_method_1"] for i, _, issue_desc in issues_list if i == idx)
                    if condition:
                        col_letter = get_column_letter(df, Config.COLUMN_MAPPINGS["disposal_method_1"])
                        apply_format(worksheet, idx, col_letter, row.get(Config.COLUMN_MAPPINGS["disposal_method_1"]), True, yellow_format)


                # --- Formatting logic for Case Registration Table ---

                # Inconsistent Investigator Name (red)
                if "被调查人" in df.columns and idx in mismatch_indices:
                    apply_format(worksheet, idx, get_column_letter(df, "被调查人"), row.get("被调查人"), True, red_format)
                    for report_col_name, rule_key in [
                        ("立案报告", "inconsistent_case_name_report"),
                        ("处分决定", "inconsistent_case_name_decision"),
                        ("审查调查报告", "inconsistent_case_name_investigation"),
                        ("审理报告", "inconsistent_case_name_trial")
                    ]:
                        if report_col_name in df.columns:
                            condition = False
                            if issues_list and isinstance(issues_list[0], dict):
                                condition = any(issue_item.get('问题描述') == Config.VALIDATION_RULES[rule_key] and issue_item.get('行号', 0) - 2 == idx for issue_item in issues_list)
                            elif issues_list and isinstance(issues_list[0], tuple):
                                if is_case_table_issues:
                                    condition = any(issue_desc == Config.VALIDATION_RULES[rule_key] for i, _, _, issue_desc in issues_list if i == idx)
                                else:
                                    condition = any(issue_desc == Config.VALIDATION_RULES[rule_key] for i, _, issue_desc in issues_list if i == idx)
                            if condition:
                                apply_format(worksheet, idx, get_column_letter(df, report_col_name), row.get(report_col_name), True, red_format)

                # Inconsistent Gender (red)
                if "性别" in df.columns and idx in gender_mismatch_indices:
                    apply_format(worksheet, idx, get_column_letter(df, "性别"), row.get("性别"), True, red_format)

                # Inconsistent Age (red)
                if "年龄" in df.columns and idx in age_mismatch_indices:
                    apply_format(worksheet, idx, get_column_letter(df, "年龄"), row.get("年龄"), True, red_format)

                # Inconsistent Brief Case Details (red)
                if "简要案情" in df.columns and idx in brief_case_details_mismatch_indices:
                    apply_format(worksheet, idx, get_column_letter(df, "简要案情"), row.get("简要案情"), True, red_format)

                # Inconsistent Birth Date (red)
                if "出生年月" in df.columns and idx in birth_date_mismatch_indices:
                    apply_format(worksheet, idx, get_column_letter(df, "出生年月"), row.get("出生年月"), True, red_format)

                # Inconsistent Education (red)
                if "学历" in df.columns and idx in education_mismatch_indices:
                    apply_format(worksheet, idx, get_column_letter(df, "学历"), row.get("学历"), True, red_format)

                # Inconsistent Ethnicity (red)
                if "民族" in df.columns and idx in ethnicity_mismatch_indices:
                    apply_format(worksheet, idx, get_column_letter(df, "民族"), row.get("民族"), True, red_format)

                # Inconsistent Party Member Status (red)
                if "是否中共党员" in df.columns and idx in party_member_mismatch_indices:
                    apply_format(worksheet, idx, get_column_letter(df, "是否中共党员"), row.get("是否中共党员"), True, red_format)

                # Inconsistent Party Joining Time (red)
                if "入党时间" in df.columns and idx in party_joining_date_mismatch_indices:
                    apply_format(worksheet, idx, get_column_letter(df, "入党时间"), row.get("入党时间"), True, red_format)

                # Filing Time related (red)
                if "立案时间" in df.columns and idx in filing_time_mismatch_indices:
                    apply_format(worksheet, idx, get_column_letter(df, "立案时间"), row.get("立案时间"), True, red_format)

                # Disciplinary Committee Filing Time (red)
                if "纪委立案时间" in df.columns and idx in disciplinary_committee_filing_time_mismatch_indices:
                    apply_format(worksheet, idx, get_column_letter(df, "纪委立案时间"), row.get("纪委立案时间"), True, red_format)

                # Disciplinary Committee Filing Authority (red)
                if "纪委立案机关" in df.columns and idx in disciplinary_committee_filing_authority_mismatch_indices:
                    apply_format(worksheet, idx, get_column_letter(df, "纪委立案机关"), row.get("纪委立案机关"), True, red_format)

                # Supervisory Committee Filing Time (red)
                if "监委立案时间" in df.columns and idx in supervisory_committee_filing_time_mismatch_indices:
                    apply_format(worksheet, idx, get_column_letter(df, "监委立案时间"), row.get("监委立案时间"), True, red_format)

                # Supervisory Committee Filing Authority (red)
                if "监委立案机关" in df.columns and idx in supervisory_committee_filing_authority_mismatch_indices:
                    apply_format(worksheet, idx, get_column_letter(df, "监委立案机关"), row.get("监委立案机关"), True, red_format)

                # Case Report Keywords (red)
                if "立案报告" in df.columns and idx in case_report_keyword_mismatch_indices:
                    apply_format(worksheet, idx, get_column_letter(df, "立案报告"), row.get("立案报告"), True, red_format)

                # Violation of Central Eight Provisions Spirit (red) 
                if "是否违反中央八项规定精神" in df.columns and idx in disposal_spirit_mismatch_indices:
                    apply_format(worksheet, idx, get_column_letter(df, "是否违反中央八项规定精神"), row.get("是否违反中央八项规定精神"), True, red_format)

                # Voluntary Confession (yellow) 
                if "是否主动交代问题" in df.columns and idx in voluntary_confession_highlight_indices:
                    apply_format(worksheet, idx, get_column_letter(df, "是否主动交代问题"), row.get("是否主动交代问题"), True, yellow_format)

                # Closing Time (red) 
                if "结案时间" in df.columns and idx in closing_time_mismatch_indices:
                    apply_format(worksheet, idx, get_column_letter(df, "结案时间"), row.get("结案时间"), True, red_format)

                # Should have revoked party position... (red)
                if "是否属于本应撤销党内职务，但本人没有党内职务而给予严重警告处分" in df.columns and idx in no_party_position_warning_mismatch_indices:
                    apply_format(worksheet, idx, get_column_letter(df, "是否属于本应撤销党内职务，但本人没有党内职务而给予严重警告处分"), row.get("是否属于本应撤销党内职务，但本人没有党内职务而给予严重警告处分"), True, red_format)

                # Recovery amount for dereliction of duty (yellow) 
                if "追缴失职渎职滥用职权造成的损失金额" in df.columns and idx in recovery_amount_highlight_indices:
                    apply_format(worksheet, idx, get_column_letter(df, "追缴失职渎职滥用职权造成的损失金额"), row.get("追缴失职渎职滥用职权造成的损失金额"), True, yellow_format)

                # Trial Acceptance Time (red)
                if "审理受理时间" in df.columns and idx in trial_acceptance_time_mismatch_indices:
                    apply_format(worksheet, idx, get_column_letter(df, "审理受理时间"), row.get("审理受理时间"), True, red_format)

                # Trial Closing Time (red)
                if "审结时间" in df.columns and idx in trial_closing_time_mismatch_indices:
                    apply_format(worksheet, idx, get_column_letter(df, "审结时间"), row.get("审结时间"), True, red_format)

                # Trial Authority vs. Reporting Unit (red)
                if "审理机关" in df.columns and "填报单位名称" in df.columns and idx in trial_authority_agency_mismatch_indices:
                    apply_format(worksheet, idx, get_column_letter(df, "审理机关"), row.get("审理机关"), True, red_format)
                    apply_format(worksheet, idx, get_column_letter(df, "填报单位名称"), row.get("填报单位名称"), True, red_format)

                # Disposal Decision Keywords (red)
                if "处分决定" in df.columns and idx in disposal_decision_keyword_mismatch_indices:
                    apply_format(worksheet, idx, get_column_letter(df, "处分决定"), row.get("处分决定"), True, red_format)
                
                # --- START OF NEW TRIAL REPORT HIGHLIGHTING ---
                # Trial Report - Non-representative keywords (red)
                if "审理报告" in df.columns and idx in trial_report_non_representative_mismatch_indices:
                    apply_format(worksheet, idx, get_column_letter(df, "审理报告"), row.get("审理报告"), True, red_format)

                # Trial Report - "扣押" keyword (red)
                if "审理报告" in df.columns and idx in trial_report_detention_mismatch_indices:
                    apply_format(worksheet, idx, get_column_letter(df, "审理报告"), row.get("审理报告"), True, red_format)

                # Confiscation amount (yellow)
                if "收缴金额（万元）" in df.columns and idx in confiscation_amount_indices:
                    apply_format(worksheet, idx, get_column_letter(df, "收缴金额（万元）"), row.get("收缴金额（万元）"), True, yellow_format)
                
                # Confiscation of property amount (yellow)
                if "没收金额" in df.columns and idx in confiscation_of_property_amount_indices:
                    apply_format(worksheet, idx, get_column_letter(df, "没收金额"), row.get("没收金额"), True, yellow_format)

                # Compensation amount to highlight (yellow)
                if "责令退赔金额" in df.columns and idx in compensation_amount_highlight_indices:
                    apply_format(worksheet, idx, get_column_letter(df, "责令退赔金额"), row.get("责令退赔金额"), True, yellow_format)
                
                # Registered handover amount to highlight (yellow)
                if "登记上交金额" in df.columns and idx in registered_handover_amount_indices:
                    apply_format(worksheet, idx, get_column_letter(df, "登记上交金额"), row.get("登记上交金额"), True, yellow_format)

                # --- NEW: Party Disciplinary Sanction (red) ---
                if "党纪处分" in df.columns and idx in disciplinary_sanction_mismatch_indices:
                    apply_format(worksheet, idx, get_column_letter(df, "党纪处分"), row.get("党纪处分"), True, red_format)
                # --- END OF NEW PARAMETER HIGHLIGHTING ---

            # Create a new sheet for issues if issues_list is not empty
            if issues_list:
                # Determine columns for issues_df based on the structure of issues_list
                # Now that issues_list is expected to contain dictionaries from case_validators.py,
                # we can rely on dictionary keys.
                if issues_list and isinstance(issues_list[0], dict):
                    # Check for "涉案人员编码" to differentiate between clue and case issues
                    if "涉案人员编码" in issues_list[0]:
                        issues_df = pd.DataFrame([
                            {'序号': i + 1, '案件编码': item.get('案件编码', ''), '涉案人员编码': item.get('涉案人员编码', ''), '问题': item.get('问题描述', '')}
                            for i, item in enumerate(issues_list)
                        ])
                    else: # Assume it's for clue table with '受理线索编码'
                         issues_df = pd.DataFrame([
                            {'序号': i + 1, '受理线索编码': item.get('受理线索编码', ''), '问题': item.get('问题描述', '')}
                            for i, item in enumerate(issues_list)
                        ])
                elif issues_list and isinstance(issues_list[0], tuple): # Fallback for old tuple format, should be rare now
                    # This branch should ideally not be hit if case_validators.py is fixed to return dictionaries
                    if len(issues_list[0]) == 4: # Case table issues (index, case_code, person_code, description)
                        issues_df = pd.DataFrame([
                            {'序号': i + 1, '案件编码': item[1], '涉案人员编码': item[2], '问题': item[3]}
                            for i, item in enumerate(issues_list)
                        ])
                    else: # Clue table issues (index, clue_code, description)
                        issues_df = pd.DataFrame([
                            {'序号': i + 1, '受理线索编码': item[1], '问题': item[2]}
                            for i, item in enumerate(issues_list)
                        ])
                
                issues_df.to_excel(writer, sheet_name='问题列表', index=False)
                logger.info(f"Issues written to '问题列表' sheet.")
            else:
                # If no issues, create a sheet indicating no issues
                no_issues_df = pd.DataFrame({'提示': ['未发现任何问题。']})
                no_issues_df.to_excel(writer, sheet_name='问题列表', index=False)
                logger.info("No issues found. '问题列表' sheet created with a no-issues message.")

        logger.info(f"Excel file formatted and saved successfully: {output_path}")
        return True

    except Exception as e:
        logger.error(f"Error formatting Excel file: {e}", exc_info=True)
        return False