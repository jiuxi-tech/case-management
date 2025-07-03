import xlsxwriter
import pandas as pd
from config import Config

def get_column_letter(df, column_name):
    # 获取列的索引，并转换为Excel列字母（1-based index）
    col_idx = df.columns.get_loc(column_name) + 1
    return xlsxwriter.utility.xl_col_to_name(col_idx - 1)

def apply_format(worksheet, row_idx, col_letter, value, condition_met, format_obj):
    """
    根据条件判断是否应用格式。
    row_idx 是 DataFrame 的索引 (0-based)，Excel 行号是 row_idx + 2 (因为有表头和Pandas的默认0行)
    """
    excel_row = row_idx + 2     # Excel行号从1开始，且有表头，所以加2
    if condition_met:
        # 如果条件满足，写入带格式的值
        worksheet.write(f'{col_letter}{excel_row}', value if pd.notna(value) else '', format_obj)
    else:
        # 如果条件不满足，写入不带格式的值
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
                     # --- START OF NEW PARAMETERS (Trial Report) ---
                     trial_report_non_representative_mismatch_indices=set(), 
                     trial_report_detention_mismatch_indices=set(),
                     # --- END OF NEW PARAMETERS (Trial Report) ---
                     confiscation_amount_indices=set(), # 添加 confiscation_amount_indices
                     confiscation_of_property_amount_indices=set(), # <--- 新增没收金额（万元）的索引参数
                     # 【新增】这里添加 compensation_amount_highlight_indices
                     compensation_amount_highlight_indices=set() 
                     ):
    """
    格式化Excel文件，根据验证问题对单元格进行着色。
    df: 原始DataFrame
    mismatch_indices: 包含不一致行的索引的集合
    output_path: 输出Excel文件的路径
    issues_list: 从 validation_core.py 或 case_validators.py 获取的问题列表，
                  每个元素可能是 (original_df_index, clue_code_value, issue_description) (3个值)
                  或 (original_df_index, case_code_value, person_code_value, issue_description) (4个值)
    其他索引集合: 用于标红或标黄特定字段
    confiscation_amount_indices: 收缴金额（万元）需要高亮的行索引集合。
    confiscation_of_property_amount_indices: 没收金额需要高亮的行索引集合。
    compensation_amount_highlight_indices: 责令退赔金额需要高亮的行索引集合。 # <--- 【新增】参数说明
    """
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
        if issues_list and len(issues_list[0]) == 4:
            is_case_table_issues = True

        for idx in range(len(df)):
            row = df.iloc[idx]

            # --- 应用黄色格式的检查 (主要针对线索表) ---
            # 受理时间 (线索表)
            # 修正了 Config.COLUMN_MAPPings 为 Config.COLUMN_MAPPINGS
            if Config.COLUMN_MAPPINGS.get("acceptance_time") in df.columns: 
                value = row.get(Config.COLUMN_MAPPINGS["acceptance_time"])
                if is_case_table_issues:
                    condition = any(issue_desc == Config.VALIDATION_RULES["confirm_acceptance_time"] for i, _, _, issue_desc in issues_list if i == idx)
                else: # 线索表的情况
                    condition = any(issue_desc == Config.VALIDATION_RULES["confirm_acceptance_time"] for i, _, issue_desc in issues_list if i == idx)
                apply_format(worksheet, idx, get_column_letter(df, Config.COLUMN_MAPPINGS["acceptance_time"]), value, condition, yellow_format)

            # 处置情况报告为空 (线索表)
            report_text = row.get("处置情况报告", '')
            if "处置情况报告" in df.columns and (pd.isna(report_text) or report_text == ''):
                if is_case_table_issues:
                    condition = any(issue_desc == Config.VALIDATION_RULES["empty_report"] for i, _, _, issue_desc in issues_list if i == idx)
                else: # 线索表的情况
                    condition = any(issue_desc == Config.VALIDATION_RULES["empty_report"] for i, _, issue_desc in issues_list if i == idx)
                apply_format(worksheet, idx, get_column_letter(df, "处置情况报告"), report_text, condition, yellow_format)

            # 检查金额字段 (线索表) - **【核心修改：移除 "没收金额" 从此循环】**
            for field, rule in [
                ("收缴金额（万元）", Config.VALIDATION_RULES["highlight_collection_amount"]),
                # ("没收金额", Config.VALIDATION_RULES["highlight_confiscation_amount"]), # <--- 移除这一行
                ("责令退赔金额", Config.VALIDATION_RULES["highlight_compensation_amount"]),
                ("登记上交金额", Config.VALIDATION_RULES["highlight_registration_amount"])
            ]:
                if field in df.columns:
                    col_letter = get_column_letter(df, field)
                    value = row.get(field)
                    if is_case_table_issues:
                        condition = any(issue_desc == rule for i, _, _, issue_desc in issues_list if i == idx)
                    else: # 线索表的情况
                        condition = any(issue_desc == rule for i, _, issue_desc in issues_list if i == idx)
                    apply_format(worksheet, idx, col_letter, value, condition, yellow_format)

            # --- 应用红色格式的检查 (主要针对线索表) ---
            # agency (线索表)
            if "填报单位名称" in df.columns and "办理机关" in df.columns and \
               (is_case_table_issues and any(issue_desc == Config.VALIDATION_RULES["inconsistent_agency"] for i, _, _, issue_desc in issues_list if i == idx) or \
               (not is_case_table_issues and any(issue_desc == Config.VALIDATION_RULES["inconsistent_agency"] for i, _, issue_desc in issues_list if i == idx))):
                apply_format(worksheet, idx, get_column_letter(df, "填报单位名称"), row.get("填报单位名称"), True, red_format)
                apply_format(worksheet, idx, get_column_letter(df, "办理机关"), row.get("办理机关"), True, red_format)

            # 被反映人 (线索表)
            if "被反映人" in df.columns and \
               (is_case_table_issues and any(issue_desc == Config.VALIDATION_RULES["inconsistent_name"] for i, _, _, issue_desc in issues_list if i == idx) or \
               (not is_case_table_issues and any(issue_desc == Config.VALIDATION_RULES["inconsistent_name"] for i, _, issue_desc in issues_list if i == idx))):
                apply_format(worksheet, idx, get_column_letter(df, "被反映人"), row.get("被反映人"), True, red_format)

            # 组织措施 (线索表)
            # 修正了 Config.COLUMN_MAPPings 为 Config.COLUMN_MAPPINGS
            if Config.COLUMN_MAPPINGS.get("organization_measure") in df.columns and \
               (is_case_table_issues and any(issue_desc == Config.VALIDATION_RULES["inconsistent_organization_measure"] for i, _, _, issue_desc in issues_list if i == idx) or \
               (not is_case_table_issues and any(issue_desc == Config.VALIDATION_RULES["inconsistent_organization_measure"] for i, _, issue_desc in issues_list if i == idx))):
                col_letter = get_column_letter(df, Config.COLUMN_MAPPINGS["organization_measure"])
                apply_format(worksheet, idx, col_letter, row.get(Config.COLUMN_MAPPINGS["organization_measure"]), True, red_format)

            # 入党时间 (线索表)
            # 修正了 Config.COLUMN_MAPPings 为 Config.COLUMN_MAPPINGS
            if Config.COLUMN_MAPPINGS.get("joining_party_time") in df.columns and \
               (is_case_table_issues and any(issue_desc == Config.VALIDATION_RULES["inconsistent_joining_party_time"] for i, _, _, issue_desc in issues_list if i == idx) or \
               (not is_case_table_issues and any(issue_desc == Config.VALIDATION_RULES["inconsistent_joining_party_time"] for i, _, issue_desc in issues_list if i == idx))):
                col_letter = get_column_letter(df, Config.COLUMN_MAPPINGS["joining_party_time"])
                apply_format(worksheet, idx, col_letter, row.get(Config.COLUMN_MAPPINGS["joining_party_time"]), True, red_format)

            # 民族 (线索表)
            # 修正了 Config.COLUMN_MAPPings 为 Config.COLUMN_MAPPINGS
            if Config.COLUMN_MAPPINGS.get("ethnicity") in df.columns and \
               (is_case_table_issues and any(issue_desc == Config.VALIDATION_RULES["inconsistent_ethnicity"] for i, _, _, issue_desc in issues_list if i == idx) or \
               (not is_case_table_issues and any(issue_desc == Config.VALIDATION_RULES["inconsistent_ethnicity"] for i, _, issue_desc in issues_list if i == idx))):
                col_letter = get_column_letter(df, Config.COLUMN_MAPPINGS["ethnicity"])
                apply_format(worksheet, idx, col_letter, row.get(Config.COLUMN_MAPPINGS["ethnicity"]), True, red_format)

            # 出生年月 (线索表)
            # 修正了 Config.COLUMN_MAPPings 为 Config.COLUMN_MAPPINGS
            if Config.COLUMN_MAPPINGS.get("birth_date") in df.columns and \
               (is_case_table_issues and any(issue_desc == Config.VALIDATION_RULES["highlight_birth_date"] for i, _, _, issue_desc in issues_list if i == idx) or \
               (not is_case_table_issues and any(issue_desc == Config.VALIDATION_RULES["highlight_birth_date"] for i, _, issue_desc in issues_list if i == idx))):
                col_letter = get_column_letter(df, Config.COLUMN_MAPPINGS["birth_date"])
                apply_format(worksheet, idx, col_letter, row.get(Config.COLUMN_MAPPINGS["birth_date"]), True, red_format)

            # 办结时间 (线索表)
            # 修正了 Config.COLUMN_MAPPings 为 Config.COLUMN_MAPPINGS
            if Config.COLUMN_MAPPINGS.get("completion_time") in df.columns and \
               (is_case_table_issues and any(issue_desc == Config.VALIDATION_RULES["highlight_completion_time"] for i, _, _, issue_desc in issues_list if i == idx) or \
               (not is_case_table_issues and any(issue_desc == Config.VALIDATION_RULES["highlight_completion_time"] for i, _, issue_desc in issues_list if i == idx))):
                col_letter = get_column_letter(df, Config.COLUMN_MAPPINGS["completion_time"])
                apply_format(worksheet, idx, col_letter, row.get(Config.COLUMN_MAPPINGS["completion_time"]), True, red_format)

            # 处置方式1二级 (线索表)
            # 修正了 Config.COLUMN_MAPPings 为 Config.COLUMN_MAPPINGS
            if Config.COLUMN_MAPPINGS.get("disposal_method_1") in df.columns and \
               (is_case_table_issues and any(issue_desc == Config.VALIDATION_RULES["highlight_disposal_method_1"] for i, _, _, issue_desc in issues_list if i == idx) or \
               (not is_case_table_issues and any(issue_desc == Config.VALIDATION_RULES["highlight_disposal_method_1"] for i, _, issue_desc in issues_list if i == idx))):
                col_letter = get_column_letter(df, Config.COLUMN_MAPPINGS["disposal_method_1"])
                apply_format(worksheet, idx, col_letter, row.get(Config.COLUMN_MAPPINGS["disposal_method_1"]), True, yellow_format)


            # --- 针对立案登记表的格式化逻辑 ---

            # 被调查人姓名不一致 (red)
            if "被调查人" in df.columns and idx in mismatch_indices:
                apply_format(worksheet, idx, get_column_letter(df, "被调查人"), row.get("被调查人"), True, red_format)
                for report_col_name, rule_key in [
                    ("立案报告", "inconsistent_case_name_report"),
                    ("处分决定", "inconsistent_case_name_decision"),
                    ("审查调查报告", "inconsistent_case_name_investigation"),
                    ("审理报告", "inconsistent_case_name_trial")
                ]:
                    if report_col_name in df.columns and \
                       (is_case_table_issues and any(issue_desc == Config.VALIDATION_RULES[rule_key] for i, _, _, issue_desc in issues_list if i == idx) or \
                       (not is_case_table_issues and any(issue_desc == Config.VALIDATION_RULES[rule_key] for i, _, issue_desc in issues_list if i == idx))):
                        apply_format(worksheet, idx, get_column_letter(df, report_col_name), row.get(report_col_name), True, red_format)

            # 性别不一致 (red)
            if "性别" in df.columns and idx in gender_mismatch_indices:
                apply_format(worksheet, idx, get_column_letter(df, "性别"), row.get("性别"), True, red_format)

            # 年龄不一致 (red)
            if "年龄" in df.columns and idx in age_mismatch_indices:
                apply_format(worksheet, idx, get_column_letter(df, "年龄"), row.get("年龄"), True, red_format)

            # 简要案情不一致 (red)
            if "简要案情" in df.columns and idx in brief_case_details_mismatch_indices:
                apply_format(worksheet, idx, get_column_letter(df, "简要案情"), row.get("简要案情"), True, red_format)

            # 出生年月不一致 (red)
            if "出生年月" in df.columns and idx in birth_date_mismatch_indices:
                apply_format(worksheet, idx, get_column_letter(df, "出生年月"), row.get("出生年月"), True, red_format)

            # 学历不一致 (red)
            if "学历" in df.columns and idx in education_mismatch_indices:
                apply_format(worksheet, idx, get_column_letter(df, "学历"), row.get("学历"), True, red_format)

            # 民族不一致 (red)
            if "民族" in df.columns and idx in ethnicity_mismatch_indices:
                apply_format(worksheet, idx, get_column_letter(df, "民族"), row.get("民族"), True, red_format)

            # 是否中共党员不一致 (red)
            if "是否中共党员" in df.columns and idx in party_member_mismatch_indices:
                apply_format(worksheet, idx, get_column_letter(df, "是否中共党员"), row.get("是否中共党员"), True, red_format)

            # 入党时间不一致 (red)
            if "入党时间" in df.columns and idx in party_joining_date_mismatch_indices:
                apply_format(worksheet, idx, get_column_letter(df, "入党时间"), row.get("入党时间"), True, red_format)

            # 立案时间相关 (red)
            if "立案时间" in df.columns and idx in filing_time_mismatch_indices:
                apply_format(worksheet, idx, get_column_letter(df, "立案时间"), row.get("立案时间"), True, red_format)

            # 纪委立案时间 (red)
            if "纪委立案时间" in df.columns and idx in disciplinary_committee_filing_time_mismatch_indices:
                apply_format(worksheet, idx, get_column_letter(df, "纪委立案时间"), row.get("纪委立案时间"), True, red_format)

            # 纪委立案机关 (red)
            if "纪委立案机关" in df.columns and idx in disciplinary_committee_filing_authority_mismatch_indices:
                apply_format(worksheet, idx, get_column_letter(df, "纪委立案机关"), row.get("纪委立案机关"), True, red_format)

            # 监委立案时间 (red)
            if "监委立案时间" in df.columns and idx in supervisory_committee_filing_time_mismatch_indices:
                apply_format(worksheet, idx, get_column_letter(df, "监委立案时间"), row.get("监委立案时间"), True, red_format)

            # 监委立案机关 (red)
            if "监委立案机关" in df.columns and idx in supervisory_committee_filing_authority_mismatch_indices:
                apply_format(worksheet, idx, get_column_letter(df, "监委立案机关"), row.get("监委立案机关"), True, red_format)

            # 立案报告关键词 (red)
            if "立案报告" in df.columns and idx in case_report_keyword_mismatch_indices:
                apply_format(worksheet, idx, get_column_letter(df, "立案报告"), row.get("立案报告"), True, red_format)

            # 是否违反中央八项规定精神 (red) 
            if "是否违反中央八项规定精神" in df.columns and idx in disposal_spirit_mismatch_indices:
                apply_format(worksheet, idx, get_column_letter(df, "是否违反中央八项规定精神"), row.get("是否违反中央八项规定精神"), True, red_format)

            # 是否主动交代问题 (yellow) 
            if "是否主动交代问题" in df.columns and idx in voluntary_confession_highlight_indices:
                apply_format(worksheet, idx, get_column_letter(df, "是否主动交代问题"), row.get("是否主动交代问题"), True, yellow_format)

            # 结案时间 (red) 
            if "结案时间" in df.columns and idx in closing_time_mismatch_indices:
                apply_format(worksheet, idx, get_column_letter(df, "结案时间"), row.get("结案时间"), True, red_format)

            # 是否属于本应撤销党内职务... (red)
            if "是否属于本应撤销党内职务，但本人没有党内职务而给予严重警告处分" in df.columns and idx in no_party_position_warning_mismatch_indices:
                apply_format(worksheet, idx, get_column_letter(df, "是否属于本应撤销党内职务，但本人没有党内职务而给予严重警告处分"), row.get("是否属于本应撤销党内职务，但本人没有党内职务而给予严重警告处分"), True, red_format)

            # 追缴失职渎职滥用职权造成的损失金额 (yellow) 
            if "追缴失职渎职滥用职权造成的损失金额" in df.columns and idx in recovery_amount_highlight_indices:
                apply_format(worksheet, idx, get_column_letter(df, "追缴失职渎职滥用职权造成的损失金额"), row.get("追缴失职渎职滥用职权造成的损失金额"), True, yellow_format)

            # 审理受理时间 (red)
            if "审理受理时间" in df.columns and idx in trial_acceptance_time_mismatch_indices:
                apply_format(worksheet, idx, get_column_letter(df, "审理受理时间"), row.get("审理受理时间"), True, red_format)

            # 审结时间 (red)
            if "审结时间" in df.columns and idx in trial_closing_time_mismatch_indices:
                apply_format(worksheet, idx, get_column_letter(df, "审结时间"), row.get("审结时间"), True, red_format)

            # 审理机关与填报单位名称不一致 (red)
            if "审理机关" in df.columns and "填报单位名称" in df.columns and idx in trial_authority_agency_mismatch_indices:
                apply_format(worksheet, idx, get_column_letter(df, "审理机关"), row.get("审理机关"), True, red_format)
                apply_format(worksheet, idx, get_column_letter(df, "填报单位名称"), row.get("填报单位名称"), True, red_format)

            # 处分决定关键词高亮 (red)
            if "处分决定" in df.columns and idx in disposal_decision_keyword_mismatch_indices:
                apply_format(worksheet, idx, get_column_letter(df, "处分决定"), row.get("处分决定"), True, red_format)
            
            # --- START OF NEW TRIAL REPORT HIGHLIGHTING ---
            # 审理报告 - 非人大代表/非政协委员/非党委委员/非中共党代表/非纪委委员 (red)
            if "审理报告" in df.columns and idx in trial_report_non_representative_mismatch_indices:
                apply_format(worksheet, idx, get_column_letter(df, "审理报告"), row.get("审理报告"), True, red_format)

            # 审理报告 - 扣押 (red)
            if "审理报告" in df.columns and idx in trial_report_detention_mismatch_indices:
                apply_format(worksheet, idx, get_column_letter(df, "审理报告"), row.get("审理报告"), True, red_format)

            # 收缴金额（万元）需要高亮 (yellow)
            if "收缴金额（万元）" in df.columns and idx in confiscation_amount_indices:
                apply_format(worksheet, idx, get_column_letter(df, "收缴金额（万元）"), row.get("收缴金额（万元）"), True, yellow_format)
            
            # 没收金额需要高亮 (yellow) <--- 新增的没收金额格式化逻辑
            if "没收金额" in df.columns and idx in confiscation_of_property_amount_indices:
                apply_format(worksheet, idx, get_column_letter(df, "没收金额"), row.get("没收金额"), True, yellow_format)

            # 【新增】责令退赔金额需要高亮 (yellow)
            if "责令退赔金额" in df.columns and idx in compensation_amount_highlight_indices:
                apply_format(worksheet, idx, get_column_letter(df, "责令退赔金额"), row.get("责令退赔金额"), True, yellow_format)
            # --- END OF NEW TRIAL REPORT HIGHLIGHTING ---