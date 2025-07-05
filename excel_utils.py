import pandas as pd
from config import Config
import logging

logger = logging.getLogger(__name__)

def get_column_letter(df, column_name):
    """
    Gets the Excel column letter for a given column name.
    """
    if column_name in df.columns:
        return pd.Index(df.columns).get_loc(column_name)
    return None

def apply_format(worksheet, row_idx, col_idx, value, condition, cell_format):
    """
    Applies a given format to a cell if the condition is met.
    """
    if condition:
        worksheet.write(row_idx + 1, col_idx, str(value), cell_format)

def _check_issue_condition(issues_list, idx, rule, is_case_table_issues):
    """Helper to check if a specific issue exists for a given row and rule."""
    if not issues_list:
        return False
    
    if isinstance(issues_list[0], dict):
        return any(issue_item.get('问题描述') == rule and issue_item.get('行号', 0) - 2 == idx for issue_item in issues_list)
    elif isinstance(issues_list[0], tuple):
        if is_case_table_issues:
            return any(issue_desc == rule for i, _, _, issue_desc in issues_list if i == idx)
        else:
            return any(issue_desc == rule for i, _, issue_desc in issues_list if i == idx)
    return False

def apply_clue_table_formats(worksheet, df, row, idx, issues_list, is_case_table_issues, yellow_format, red_format):
    """
    Applies yellow and red formatting checks specific to the clue table.
    """
    # Acceptance Time (Clue Table) - Yellow
    if Config.COLUMN_MAPPINGS.get("acceptance_time") in df.columns:
        value = row.get(Config.COLUMN_MAPPINGS["acceptance_time"])
        condition = _check_issue_condition(issues_list, idx, Config.VALIDATION_RULES["confirm_acceptance_time"], is_case_table_issues)
        apply_format(worksheet, idx, get_column_letter(df, Config.COLUMN_MAPPINGS["acceptance_time"]), value, condition, yellow_format)

    # Empty Disposal Report (Clue Table) - Yellow
    report_text = row.get("处置情况报告", '')
    if "处置情况报告" in df.columns and (pd.isna(report_text) or report_text == ''):
        condition = _check_issue_condition(issues_list, idx, Config.VALIDATION_RULES["empty_report"], is_case_table_issues)
        apply_format(worksheet, idx, get_column_letter(df, "处置情况报告"), report_text, condition, yellow_format)

    # Check amount fields (Clue Table) - Yellow
    for field, rule in [
        ("收缴金额（万元）", Config.VALIDATION_RULES["highlight_collection_amount"]),
        ("责令退赔金额", Config.VALIDATION_RULES["highlight_compensation_amount"]),
        ("登记上交金额", Config.VALIDATION_RULES["highlight_registration_amount"])
    ]:
        if field in df.columns:
            col_letter = get_column_letter(df, field)
            value = row.get(field)
            condition = _check_issue_condition(issues_list, idx, rule, is_case_table_issues)
            apply_format(worksheet, idx, col_letter, value, condition, yellow_format)

    # Agency (Clue Table) - Red
    # Agency (Clue Table) - Red
    if Config.COLUMN_MAPPINGS.get("reporting_agency") in df.columns and Config.COLUMN_MAPPINGS.get("authority") in df.columns:
        # 检查是否有针对 'inconsistent_agency_clue' 规则的问题
        condition_agency_inconsistent = _check_issue_condition(issues_list, idx, Config.VALIDATION_RULES["inconsistent_agency_clue"], is_case_table_issues)
        if condition_agency_inconsistent:
            # 标红“填报单位名称”和“办理机关”
            apply_format(worksheet, idx, get_column_letter(df, Config.COLUMN_MAPPINGS["reporting_agency"]), row.get(Config.COLUMN_MAPPINGS["reporting_agency"]), True, red_format)
            apply_format(worksheet, idx, get_column_letter(df, Config.COLUMN_MAPPINGS["authority"]), row.get(Config.COLUMN_MAPPINGS["authority"]), True, red_format)


    # Reflected Person (Clue Table) - Red
    if "被反映人" in df.columns:
        condition = _check_issue_condition(issues_list, idx, Config.VALIDATION_RULES["inconsistent_name"], is_case_table_issues)
        if condition:
            apply_format(worksheet, idx, get_column_letter(df, "被反映人"), row.get("被反映人"), True, red_format)

    # Organizational Measures (Clue Table) - Red
    if Config.COLUMN_MAPPINGS.get("organization_measure") in df.columns:
        condition = _check_issue_condition(issues_list, idx, Config.VALIDATION_RULES["inconsistent_organization_measure"], is_case_table_issues)
        if condition:
            col_letter = get_column_letter(df, Config.COLUMN_MAPPINGS["organization_measure"])
            apply_format(worksheet, idx, col_letter, row.get(Config.COLUMN_MAPPINGS["organization_measure"]), True, red_format)

    # Party Joining Time (Clue Table) - Red
    if Config.COLUMN_MAPPINGS.get("joining_party_time") in df.columns:
        condition = _check_issue_condition(issues_list, idx, Config.VALIDATION_RULES["inconsistent_joining_party_time"], is_case_table_issues)
        if condition:
            col_letter = get_column_letter(df, Config.COLUMN_MAPPINGS["joining_party_time"])
            apply_format(worksheet, idx, col_letter, row.get(Config.COLUMN_MAPPINGS["joining_party_time"]), True, red_format)

    # Ethnicity (Clue Table) - Red
    if Config.COLUMN_MAPPINGS.get("ethnicity") in df.columns:
        condition = _check_issue_condition(issues_list, idx, Config.VALIDATION_RULES["inconsistent_ethnicity"], is_case_table_issues)
        if condition:
            col_letter = get_column_letter(df, Config.COLUMN_MAPPINGS["ethnicity"])
            apply_format(worksheet, idx, col_letter, row.get(Config.COLUMN_MAPPINGS["ethnicity"]), True, red_format)

    # Birth Date (Clue Table) - Red
    if Config.COLUMN_MAPPINGS.get("birth_date") in df.columns:
        condition = _check_issue_condition(issues_list, idx, Config.VALIDATION_RULES["highlight_birth_date"], is_case_table_issues)
        if condition:
            col_letter = get_column_letter(df, Config.COLUMN_MAPPINGS["birth_date"])
            apply_format(worksheet, idx, col_letter, row.get(Config.COLUMN_MAPPINGS["birth_date"]), True, red_format)

    # Completion Time (Clue Table) - Red
    if Config.COLUMN_MAPPINGS.get("completion_time") in df.columns:
        condition = _check_issue_condition(issues_list, idx, Config.VALIDATION_RULES["highlight_completion_time"], is_case_table_issues)
        if condition:
            col_letter = get_column_letter(df, Config.COLUMN_MAPPINGS["completion_time"])
            apply_format(worksheet, idx, col_letter, row.get(Config.COLUMN_MAPPINGS["completion_time"]), True, red_format)

    # Disposal Method 1 Secondary (Clue Table) - Yellow
    if Config.COLUMN_MAPPINGS.get("disposal_method_1") in df.columns:
        condition = _check_issue_condition(issues_list, idx, Config.VALIDATION_RULES["highlight_disposal_method_1"], is_case_table_issues)
        if condition:
            col_letter = get_column_letter(df, Config.COLUMN_MAPPINGS["disposal_method_1"])
            apply_format(worksheet, idx, col_letter, row.get(Config.COLUMN_MAPPINGS["disposal_method_1"]), True, yellow_format)


def apply_case_table_formats(worksheet, df, row, idx, mismatch_indices, issues_list, is_case_table_issues,
                             gender_mismatch_indices, age_mismatch_indices, birth_date_mismatch_indices,
                             education_mismatch_indices, ethnicity_mismatch_indices, party_member_mismatch_indices,
                             party_joining_date_mismatch_indices, brief_case_details_mismatch_indices,
                             filing_time_mismatch_indices, disciplinary_committee_filing_time_mismatch_indices,
                             disciplinary_committee_filing_authority_mismatch_indices,
                             supervisory_committee_filing_time_mismatch_indices,
                             supervisory_committee_filing_authority_mismatch_indices,
                             case_report_keyword_mismatch_indices, disposal_spirit_mismatch_indices,
                             voluntary_confession_highlight_indices, closing_time_mismatch_indices,
                             no_party_position_warning_mismatch_indices, recovery_amount_highlight_indices,
                             trial_acceptance_time_mismatch_indices, trial_closing_time_mismatch_indices,
                             trial_authority_agency_mismatch_indices, disposal_decision_keyword_mismatch_indices,
                             trial_report_non_representative_mismatch_indices, trial_report_detention_mismatch_indices,
                             confiscation_amount_indices, confiscation_of_property_amount_indices,
                             compensation_amount_highlight_indices, registered_handover_amount_indices,
                             disciplinary_sanction_mismatch_indices,
                             administrative_sanction_mismatch_indices, # <-- 【新增】在这里添加这个参数
                             yellow_format, red_format):
    """
    Applies red and yellow formatting checks specific to the case registration table.
    """
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
                condition = _check_issue_condition(issues_list, idx, Config.VALIDATION_RULES[rule_key], is_case_table_issues)
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

    # Party Disciplinary Sanction (red)
    if "党纪处分" in df.columns and idx in disciplinary_sanction_mismatch_indices:
        apply_format(worksheet, idx, get_column_letter(df, "党纪处分"), row.get("党纪处分"), True, red_format)

    # Administrative Sanction (red) # <-- 【新增】在这里添加政务处分的高亮逻辑
    if "政务处分" in df.columns and idx in administrative_sanction_mismatch_indices:
        apply_format(worksheet, idx, get_column_letter(df, "政务处分"), row.get("政务处分"), True, red_format)


def create_clue_issues_sheet(writer, issues_list):
    """
    Creates a new sheet for clue issues if issues_list is not empty.
    """
    if issues_list:
        issues_df = pd.DataFrame([
            {'序号': i + 1, '受理线索编码': item.get('受理线索编码', ''), '问题': item.get('问题描述', ''), '行号': item.get('行号', ''), '列名': item.get('列名', '')}
            for i, item in enumerate(issues_list)
        ])
        issues_df.to_excel(writer, sheet_name='问题列表', index=False)

def create_case_issues_sheet(writer, issues_list):
    """
    Creates a new sheet for case issues if issues_list is not empty.
    """
    if issues_list:
        issues_df = pd.DataFrame([
            {'序号': i + 1, '案件编码': item.get('案件编码', ''), '涉案人员编码': item.get('涉案人员编码', ''), '问题': item.get('问题描述', '')}
            for i, item in enumerate(issues_list)
        ])
        issues_df.to_excel(writer, sheet_name='问题列表', index=False)
        logger.info(f"Issues written to '问题列表' sheet.")
    else:
        no_issues_df = pd.DataFrame({'提示': ['未发现任何问题。']})
        no_issues_df.to_excel(writer, sheet_name='问题列表', index=False)
        logger.info(f"No issues found for case data.")