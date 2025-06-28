import logging
import pandas as pd
from datetime import datetime
from config import Config
import re

# 从 case_validation_helpers 导入核心验证函数
from validation_rules.case_validation_helpers import (
    validate_gender_rules,
    validate_age_rules,
    validate_brief_case_details_rules
)

# 从 case_validation_extended 导入扩展验证函数
from validation_rules.case_validation_extended import (
    validate_birth_date_rules,
    validate_education_rules,
    validate_ethnicity_rules,
    validate_party_member_rules,
    validate_party_joining_date_rules
)

# 从 case_validation_additional 导入其他验证函数
from validation_rules.case_validation_additional import (
    validate_name_rules,
    validate_case_report_keywords_rules,
    validate_voluntary_confession_rules,
    validate_no_party_position_warning_rules
)

# 导入立案时间规则
from validation_rules.case_timestamp_rules import validate_filing_time
# 导入处分和金额相关规则
from validation_rules.case_disposal_amount_rules import validate_disposal_and_amount_rules

logger = logging.getLogger(__name__)

def parse_chinese_date(date_str):
    """
    Parses a Chinese date string (e.g., '2025年3月20日') into a datetime.date object.
    """
    if not isinstance(date_str, str):
        return None
    
    # Remove any extra characters after the date, like "，"
    date_str = date_str.split('，')[0].strip()

    # Regex to capture year, month, day
    match = re.match(r'(\d{4})年(\d{1,2})月(\d{1,2})日', date_str)
    if match:
        try:
            year, month, day = map(int, match.groups())
            return datetime(year, month, day).date()
        except ValueError:
            return None
    return None

def validate_case_relationships(df):
    """Validate relationships between fields in the case registration Excel."""
    mismatch_indices = set()
    gender_mismatch_indices = set()
    age_mismatch_indices = set()
    brief_case_details_mismatch_indices = set()
    birth_date_mismatch_indices = set()
    education_mismatch_indices = set()
    ethnicity_mismatch_indices = set()
    party_member_mismatch_indices = set()
    party_joining_date_mismatch_indices = set()
    filing_time_mismatch_indices = set()
    disciplinary_committee_filing_time_mismatch_indices = set()
    disciplinary_committee_filing_authority_mismatch_indices = set()
    supervisory_committee_filing_time_mismatch_indices = set()
    supervisory_committee_filing_authority_mismatch_indices = set()
    case_report_keyword_mismatch_indices = set()
    disposal_spirit_mismatch_indices = set()
    voluntary_confession_highlight_indices = set()
    closing_time_mismatch_indices = set()
    no_party_position_warning_mismatch_indices = set()
    recovery_amount_highlight_indices = set()
    # 【新增】审理受理时间不一致的行索引集合
    trial_acceptance_time_mismatch_indices = set()

    issues_list = [] 
    
    required_headers = [
        "被调查人", "性别", "年龄", "出生年月", "学历", "民族", 
        "是否中共党员", "入党时间", "立案报告", "处分决定", 
        "审查调查报告", "审理报告", "简要案情",
        "案件编码", "涉案人员编码",
        "立案时间", "立案决定书",
        "纪委立案时间", "纪委立案机关", "监委立案时间", "监委立案机关", "填报单位名称",
        "是否违反中央八项规定精神",
        "是否主动交代问题",
        "结案时间",
        "是否属于本应撤销党内职务，但本人没有党内职务而给予严重警告处分",
        "追缴失职渎职滥用职权造成的损失金额",
        # 【新增】审理受理时间
        "审理受理时间" 
    ]
    if not all(header in df.columns for header in required_headers):
        msg = f"缺少必要的表头: {required_headers}"
        logger.error(msg)
        print(msg)
        return (mismatch_indices, gender_mismatch_indices, age_mismatch_indices, brief_case_details_mismatch_indices, issues_list, 
                birth_date_mismatch_indices, education_mismatch_indices, ethnicity_mismatch_indices, 
                party_member_mismatch_indices, party_joining_date_mismatch_indices, filing_time_mismatch_indices, 
                disciplinary_committee_filing_time_mismatch_indices, disciplinary_committee_filing_authority_mismatch_indices, 
                supervisory_committee_filing_time_mismatch_indices, supervisory_committee_filing_authority_mismatch_indices, 
                case_report_keyword_mismatch_indices, disposal_spirit_mismatch_indices, voluntary_confession_highlight_indices, 
                closing_time_mismatch_indices, no_party_position_warning_mismatch_indices,
                recovery_amount_highlight_indices, trial_acceptance_time_mismatch_indices) # 【新增】返回 trial_acceptance_time_mismatch_indices

    current_year = datetime.now().year

    case_report_keywords_to_check = ["人大代表", "政协委员", "党委委员", "中共党代表", "纪委委员"]

    for index, row in df.iterrows():
        logger.debug(f"Processing row {index + 1}")
        print(f"处理行 {index + 1}")

        investigated_person = str(row.get("被调查人", "")).strip()
        if not investigated_person:
            logger.info(f"行 {index + 1} - '被调查人' 字段为空，跳过此行验证。")
            print(f"行 {index + 1} - '被调查人' 字段为空，跳过此行验证。")
            continue

        excel_voluntary_confession = str(row.get("是否主动交代问题", "")).strip()
        excel_gender = str(row.get("性别", "")).strip()
        
        excel_age = None
        if pd.notna(row.get("年龄")):
            try:
                excel_age = int(row.get("年龄"))
            except ValueError:
                logger.warning(f"行 {index + 1} - Excel '年龄' 字段 '{row.get('年龄')}' 不是有效数字。")
                print(f"行 {index + 1} - Excel '年龄' 字段 '{row.get('年龄')}' 不是有效数字。")
                age_mismatch_indices.add(index)
                issues_list.append((index, row.get("案件编码", ""), row.get("涉案人员编码", ""), "N2年龄字段格式不正确"))

        excel_brief_case_details = str(row.get("简要案情", "")).strip()
        excel_birth_date = str(row.get("出生年月", "")).strip()
        excel_education = str(row.get("学历", "")).strip()
        excel_ethnicity = str(row.get("民族", "")).strip()
        excel_party_member = str(row.get("是否中共党员", "")).strip()
        excel_party_joining_date = str(row.get("入党时间", "")).strip()
        excel_no_party_position_warning = str(row.get("是否属于本应撤销党内职务，但本人没有党内职务而给予严重警告处分", "")).strip()
        excel_recovery_amount = str(row.get("追缴失职渎职滥用职权造成的损失金额", "")).strip()
        # 【新增】获取“审理受理时间”字段值
        excel_trial_acceptance_time = row.get("审理受理时间")

        excel_case_code = str(row.get("案件编码", "")).strip()
        excel_person_code = str(row.get("涉案人员编码", "")).strip()

        report_text_raw = row.get("立案报告", "") if pd.notna(row.get("立案报告")) else ''
        decision_text_raw = row.get("处分决定", "") if pd.notna(row.get("处分决定")) else ''
        investigation_text_raw = row.get("审查调查报告", "") if pd.notna(row.get("审查调查报告")) else ''
        trial_text_raw = row.get("审理报告", "") if pd.notna(row.get("审理报告")) else ''
        filing_decision_doc_raw = row.get("立案决定书", "") if pd.notna(row.get("立案决定书")) else ''

        # --- 调用辅助函数进行验证 ---
        validate_gender_rules(row, index, excel_case_code, excel_person_code, issues_list, gender_mismatch_indices,
                              excel_gender, report_text_raw, decision_text_raw, investigation_text_raw, trial_text_raw)

        validate_age_rules(row, index, excel_case_code, excel_person_code, issues_list, age_mismatch_indices,
                           excel_age, current_year, report_text_raw, decision_text_raw, investigation_text_raw, trial_text_raw)

        validate_brief_case_details_rules(row, index, excel_case_code, excel_person_code, issues_list, brief_case_details_mismatch_indices,
                                          excel_brief_case_details, investigated_person, report_text_raw, decision_text_raw)

        validate_birth_date_rules(row, index, excel_case_code, excel_person_code, issues_list, birth_date_mismatch_indices,
                                  excel_birth_date, report_text_raw, decision_text_raw, investigation_text_raw, trial_text_raw)

        validate_education_rules(row, index, excel_case_code, excel_person_code, issues_list, education_mismatch_indices,
                                 excel_education, report_text_raw)

        validate_ethnicity_rules(row, index, excel_case_code, excel_person_code, issues_list, ethnicity_mismatch_indices,
                                 excel_ethnicity, report_text_raw, decision_text_raw, investigation_text_raw, trial_text_raw)

        validate_party_member_rules(row, index, excel_case_code, excel_person_code, issues_list, party_member_mismatch_indices,
                                    excel_party_member, report_text_raw, decision_text_raw)

        validate_party_joining_date_rules(row, index, excel_case_code, excel_person_code, issues_list, party_joining_date_mismatch_indices,
                                          excel_party_member, excel_party_joining_date, report_text_raw)

        validate_name_rules(row, index, excel_case_code, excel_person_code, issues_list, mismatch_indices,
                            investigated_person, report_text_raw, decision_text_raw, investigation_text_raw, trial_text_raw)

        validate_case_report_keywords_rules(row, index, excel_case_code, excel_person_code, issues_list, case_report_keyword_mismatch_indices,
                                            case_report_keywords_to_check, report_text_raw, decision_text_raw, investigation_text_raw, trial_text_raw)
        
        validate_voluntary_confession_rules(row, index, excel_case_code, excel_person_code, issues_list, voluntary_confession_highlight_indices,
                                            excel_voluntary_confession, trial_text_raw)

        validate_no_party_position_warning_rules(row, index, excel_case_code, excel_person_code, issues_list, no_party_position_warning_mismatch_indices,
                                                 excel_no_party_position_warning, decision_text_raw)

        if pd.notna(row.get("追缴失职渎职滥用职权造成的损失金额")) and str(row.get("追缴失职渎职滥用职权造成的损失金额", "")).strip() != '':
            recovery_amount_highlight_indices.add(index)
            issues_list.append((index, excel_case_code, excel_person_code, Config.VALIDATION_RULES["highlight_recovery_amount"]))

        # 【新增逻辑】审理受理时间与审理报告开头时间内容比对
        if pd.notna(excel_trial_acceptance_time) and pd.notna(trial_text_raw):
            excel_date_obj = None
            if isinstance(excel_trial_acceptance_time, datetime):
                excel_date_obj = excel_trial_acceptance_time.date()
            elif isinstance(excel_trial_acceptance_time, str):
                try:
                    # 尝试解析常见的日期字符串格式
                    excel_date_obj = pd.to_datetime(excel_trial_acceptance_time).date()
                except ValueError:
                    logger.warning(f"行 {index + 1} - '审理受理时间' 字段 '{excel_trial_acceptance_time}' 无法解析为日期。")
                    print(f"行 {index + 1} - '审理受理时间' 字段 '{excel_trial_acceptance_time}' 无法解析为日期。")
                    trial_acceptance_time_mismatch_indices.add(index)
                    issues_list.append((index, excel_case_code, excel_person_code, "CP审理受理时间格式不正确"))
            
            if excel_date_obj:
                # 查找特定字符串及其后的段落
                match_phrase_pattern = r"关于王.+?同志违纪案的审理报告主动交代"
                match_phrase = re.search(match_phrase_pattern, trial_text_raw)

                if match_phrase:
                    # 获取匹配字符串后的内容
                    after_match = trial_text_raw[match_phrase.end():].strip()
                    
                    # 查找第一个换行符，如果存在，则取第一行
                    first_newline_pos = after_match.find('\n')
                    first_line = after_match
                    if first_newline_pos != -1:
                        first_line = after_match[:first_newline_pos].strip()
                    
                    # 提取第一个逗号前的内容作为日期字符串
                    comma_pos = first_line.find('，')
                    extracted_date_str = first_line
                    if comma_pos != -1:
                        extracted_date_str = first_line[:comma_pos].strip()
                    
                    extracted_date_obj = parse_chinese_date(extracted_date_str)

                    if extracted_date_obj:
                        if excel_date_obj != extracted_date_obj:
                            trial_acceptance_time_mismatch_indices.add(index)
                            issues_list.append((index, excel_case_code, excel_person_code, "CP审理受理时间与CY审理报告不一致"))
                            logger.info(f"行 {index + 1} - 审理受理时间不一致：Excel: {excel_date_obj}, 审理报告: {extracted_date_obj}")
                            print(f"行 {index + 1} - 审理受理时间不一致：Excel: {excel_date_obj}, 审理报告: {extracted_date_obj}")
                        else:
                            logger.info(f"行 {index + 1} - 审理受理时间一致：Excel: {excel_date_obj}, 审理报告: {extracted_date_obj}")
                            print(f"行 {index + 1} - 审理受理时间一致：Excel: {excel_date_obj}, 审理报告: {extracted_date_obj}")
                    else:
                        logger.warning(f"行 {index + 1} - 从审理报告中提取的日期 '{extracted_date_str}' 无法解析。")
                        print(f"行 {index + 1} - 从审理报告中提取的日期 '{extracted_date_str}' 无法解析。")
                        trial_acceptance_time_mismatch_indices.add(index)
                        issues_list.append((index, excel_case_code, excel_person_code, "CY审理报告中审理受理时间格式不正确或未找到"))
                else:
                    logger.info(f"行 {index + 1} - 审理报告中未找到匹配的字符串 '关于王xx同志违纪案的审理报告主动交代'。")
                    print(f"行 {index + 1} - 审理报告中未找到匹配的字符串 '关于王xx同志违纪案的审理报告主动交代'。")
                    # 如果审理报告中没有找到特定字符串，也视为不一致，因为无法进行比对
                    trial_acceptance_time_mismatch_indices.add(index)
                    issues_list.append((index, excel_case_code, excel_person_code, "CY审理报告中未找到审理受理时间相关内容"))
        elif pd.notna(excel_trial_acceptance_time) and pd.isna(trial_text_raw):
            logger.info(f"行 {index + 1} - '审理受理时间' 有值但 '审理报告' 为空，无法比对。")
            print(f"行 {index + 1} - '审理受理时间' 有值但 '审理报告' 为空，无法比对。")
            trial_acceptance_time_mismatch_indices.add(index)
            issues_list.append((index, excel_case_code, excel_person_code, "CP审理受理时间有值但CY审理报告为空，无法比对"))
        elif pd.isna(excel_trial_acceptance_time) and pd.notna(trial_text_raw):
            logger.info(f"行 {index + 1} - '审理受理时间' 为空但 '审理报告' 有值，无法比对。")
            print(f"行 {index + 1} - '审理受理时间' 为空但 '审理报告' 有值，无法比对。")
            trial_acceptance_time_mismatch_indices.add(index)
            issues_list.append((index, excel_case_code, excel_person_code, "CP审理受理时间为空但CY审理报告有值，无法比对"))


    # 调用立案时间规则验证函数
    validate_filing_time(df, issues_list, filing_time_mismatch_indices,
                         disciplinary_committee_filing_time_mismatch_indices,
                         disciplinary_committee_filing_authority_mismatch_indices,
                         supervisory_committee_filing_time_mismatch_indices,
                         supervisory_committee_filing_authority_mismatch_indices)

    # 调用处分和金额相关规则验证函数
    validate_disposal_and_amount_rules(df, issues_list, disposal_spirit_mismatch_indices, closing_time_mismatch_indices)

    # 返回所有可能的不一致索引集以及更新后的 issues_list
    # 【修改】在返回列表中加入 recovery_amount_highlight_indices 和 trial_acceptance_time_mismatch_indices
    return (mismatch_indices, gender_mismatch_indices, age_mismatch_indices, brief_case_details_mismatch_indices, issues_list, 
            birth_date_mismatch_indices, education_mismatch_indices, ethnicity_mismatch_indices, 
            party_member_mismatch_indices, party_joining_date_mismatch_indices, filing_time_mismatch_indices, 
            disciplinary_committee_filing_time_mismatch_indices, disciplinary_committee_filing_authority_mismatch_indices, 
            supervisory_committee_filing_time_mismatch_indices, supervisory_committee_filing_authority_mismatch_indices, 
            case_report_keyword_mismatch_indices, disposal_spirit_mismatch_indices, voluntary_confession_highlight_indices, 
            closing_time_mismatch_indices, no_party_position_warning_mismatch_indices,
            recovery_amount_highlight_indices, trial_acceptance_time_mismatch_indices)