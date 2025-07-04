import logging
import pandas as pd
from datetime import datetime
import re
from config import Config  # 导入Config，因为某些验证规则需要用到其中的配置

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

def validate_trial_acceptance_time_vs_report(row, index, excel_case_code, excel_person_code, issues_list, trial_acceptance_time_mismatch_indices):
    """
    验证 '审理受理时间' 与 '审理报告' 开头时间内容的一致性。
    """
    excel_trial_acceptance_time = row.get("审理受理时间")
    trial_text_raw = row.get("审理报告", "") if pd.notna(row.get("审理报告")) else ''

    if pd.notna(excel_trial_acceptance_time) and pd.notna(trial_text_raw):
        excel_date_obj = None
        if isinstance(excel_trial_acceptance_time, datetime):
            excel_date_obj = excel_trial_acceptance_time.date()
        elif isinstance(excel_trial_acceptance_time, str):
            try:
                excel_date_obj = pd.to_datetime(excel_trial_acceptance_time).date()
            except ValueError:
                logger.warning(f"行 {index + 1} - '审理受理时间' 字段 '{excel_trial_acceptance_time}' 无法解析为日期。")
                print(f"行 {index + 1} - '审理受理时间' 字段 '{excel_trial_acceptance_time}' 不是有效日期格式。")
                trial_acceptance_time_mismatch_indices.add(index)
                issues_list.append((index, excel_case_code, excel_person_code, "CP审理受理时间格式不正确"))
        
        if excel_date_obj:
            match_phrase_pattern = r"关于王.+?同志违纪案的审理报告主动交代"
            match_phrase = re.search(match_phrase_pattern, trial_text_raw)

            if match_phrase:
                after_match = trial_text_raw[match_phrase.end():].strip()
                first_newline_pos = after_match.find('\n')
                first_line = after_match
                if first_newline_pos != -1:
                    first_line = after_match[:first_newline_pos].strip()
                
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

def validate_trial_closing_time_vs_report(row, index, excel_case_code, excel_person_code, issues_list, trial_closing_time_mismatch_indices):
    """
    验证 '审结时间' 与 '审理报告' 落款时间的一致性。
    """
    excel_trial_closing_time = row.get("审结时间")
    trial_text_raw = row.get("审理报告", "") if pd.notna(row.get("审理报告")) else ''

    if pd.notna(excel_trial_closing_time) and pd.notna(trial_text_raw):
        excel_closing_date_obj = None
        if isinstance(excel_trial_closing_time, datetime):
            excel_closing_date_obj = excel_trial_closing_time.date()
        elif isinstance(excel_trial_closing_time, str):
            try:
                excel_closing_date_obj = pd.to_datetime(excel_trial_closing_time).date()
            except ValueError:
                logger.warning(f"行 {index + 1} - '审结时间' 字段 '{excel_trial_closing_time}' 无法解析为日期。")
                print(f"行 {index + 1} - '审结时间' 为空或格式不正确，无法解析。")
                trial_closing_time_mismatch_indices.add(index)
                issues_list.append((index, excel_case_code, excel_person_code, "CS审结时间格式不正确"))
        
        if excel_closing_date_obj:
            lines = trial_text_raw.strip().split('\n')
            if lines:
                last_line = lines[-1].strip()
                date_match = re.search(r'(\d{4}年\d{1,2}月\d{1,2}日)', last_line)
                if date_match:
                    extracted_closing_date_str = date_match.group(1)
                    extracted_closing_date_obj = parse_chinese_date(extracted_closing_date_str)

                    if extracted_closing_date_obj:
                        if excel_closing_date_obj != extracted_closing_date_obj:
                            trial_closing_time_mismatch_indices.add(index)
                            issues_list.append((index, excel_case_code, excel_person_code, "CS审结时间与CY审理报告不一致"))
                            logger.info(f"行 {index + 1} - 审结时间不一致：Excel: {excel_closing_date_obj}, 审理报告落款: {extracted_closing_date_obj}")
                            print(f"行 {index + 1} - 审结时间不一致：Excel: {excel_closing_date_obj}, 审理报告落款: {extracted_closing_date_obj}")
                        else:
                            logger.info(f"行 {index + 1} - 审结时间一致：Excel: {excel_closing_date_obj}, 审理报告落款: {extracted_closing_date_obj}")
                            print(f"行 {index + 1} - 审结时间一致：Excel: {excel_closing_date_obj}, 审理报告落款: {extracted_closing_date_obj}")
                    else:
                        logger.warning(f"行 {index + 1} - 从审理报告最后一行 '{last_line}' 中提取的日期 '{extracted_closing_date_str}' 无法解析。")
                        print(f"行 {index + 1} - 从审理报告最后一行 '{last_line}' 中提取的日期 '{extracted_closing_date_str}' 无法解析。")
                        trial_closing_time_mismatch_indices.add(index)
                        issues_list.append((index, excel_case_code, excel_person_code, "CY审理报告落款时间格式不正确或未找到"))
                else:
                    logger.warning(f"行 {index + 1} - 审理报告最后一行 '{last_line}' 未找到日期格式。")
                    print(f"行 {index + 1} - 审理报告最后一行 '{last_line}' 未找到日期格式。")
                    trial_closing_time_mismatch_indices.add(index)
                    issues_list.append((index, excel_case_code, excel_person_code, "CY审理报告落款时间未找到"))
            else:
                logger.info(f"行 {index + 1} - '审理报告' 为空，无法提取落款时间。")
                print(f"行 {index + 1} - '审理报告' 为空，无法提取落款时间。")
                trial_closing_time_mismatch_indices.add(index)
                issues_list.append((index, excel_case_code, excel_person_code, "CY审理报告为空，无法比对审结时间"))
    elif pd.notna(excel_trial_closing_time) and pd.isna(trial_text_raw):
        logger.info(f"行 {index + 1} - '审结时间' 有值但 '审理报告' 为空，无法比对。")
        print(f"行 {index + 1} - '审结时间' 有值但 '审理报告' 为空，无法比对。")
        trial_closing_time_mismatch_indices.add(index)
        issues_list.append((index, excel_case_code, excel_person_code, "CS审结时间有值但CY审理报告为空，无法比对"))
    elif pd.isna(excel_trial_closing_time) and pd.notna(trial_text_raw):
        logger.info(f"行 {index + 1} - '审结时间' 为空但 '审理报告' 有值，无法比对。")
        print(f"行 {index + 1} - '审结时间' 为空但 '审理报告' 有值，无法比对。")
        trial_closing_time_mismatch_indices.add(index)
        issues_list.append((index, excel_case_code, excel_person_code, "CS审结时间为空但CY审理报告有值，无法比对"))

def validate_trial_authority_vs_reporting_agency(row, index, excel_case_code, excel_person_code, issues_list, trial_authority_agency_mismatch_indices, sl_authority_agency_mappings):
    """
    验证 '审理机关' 与 '填报单位名称' 是否与 SL 类别的机关单位字典数据匹配。
    """
    excel_trial_authority = str(row.get("审理机关", "")).strip()
    excel_reporting_agency = str(row.get("填报单位名称", "")).strip()

    if excel_trial_authority and excel_reporting_agency:
        found_match = False
        for mapping in sl_authority_agency_mappings:
            if (mapping["authority"] == excel_trial_authority and 
                mapping["agency"] == excel_reporting_agency):
                found_match = True
                logger.info(f"行 {index + 1} - 审理机关 '{excel_trial_authority}' 和 填报单位名称 '{excel_reporting_agency}' 匹配成功 (Category: SL)。")
                print(f"行 {index + 1} - 审理机关 '{excel_trial_authority}' 和 填报单位名称 '{excel_reporting_agency}' 匹配成功 (Category: SL)。")
                break
        
        if not found_match:
            trial_authority_agency_mismatch_indices.add(index)
            issues_list.append((index, excel_case_code, excel_person_code, "CR审理机关与A填报单位不一致"))
            logger.warning(f"行 {index + 1} - 审理机关 '{excel_trial_authority}' 和 填报单位名称 '{excel_reporting_agency}' 不匹配或Category不为SL。")
            print(f"行 {index + 1} - 审理机关 '{excel_trial_authority}' 和 填报单位名称 '{excel_reporting_agency}' 不匹配或Category不为SL。")
        else:
            logger.info(f"行 {index + 1} - 审理机关 ('{excel_trial_authority}') 和 填报单位名称 ('{excel_reporting_agency}') 在对应表中找到匹配记录。")
            print(f"行 {index + 1} - 审理机关 ('{excel_trial_authority}') 和 填报单位名称 ('{excel_reporting_agency}') 在对应表中找到匹配记录。")
    else:
        logger.info(f"行 {index + 1} - '审理机关' 或 '填报单位名称' 为空，跳过比对。审理机关: '{excel_trial_authority}', 填报单位名称: '{excel_reporting_agency}'")
        print(f"行 {index + 1} - '审理机关' 或 '填报单位名称' 为空，无法比对。")
        trial_authority_agency_mismatch_indices.add(index)
        issues_list.append((index, excel_case_code, excel_person_code, "CR审理机关或A填报单位名称为空，无法比对"))

def validate_disposal_decision_keywords(row, index, excel_case_code, excel_person_code, issues_list, disposal_decision_keyword_mismatch_indices):
    """
    检查 '处分决定' 字段是否包含禁用关键词。
    """
    decision_text_raw = row.get("处分决定", "") if pd.notna(row.get("处分决定")) else ''
    if "处分决定" in row.index and pd.notna(decision_text_raw) and decision_text_raw.strip() != '':
        found_disposal_keyword = False
        for keyword in Config.DISPOSAL_DECISION_KEYWORDS:
            if keyword in decision_text_raw:
                disposal_decision_keyword_mismatch_indices.add(index)
                issues_list.append((index, excel_case_code, excel_person_code, Config.VALIDATION_RULES["disposal_decision_keyword_highlight"]))
                logger.warning(f"行 {index + 1} - '处分决定' 字段包含禁用关键词: '{keyword}'。")
                print(f"行 {index + 1} - '处分决定' 字段包含禁用关键词: '{keyword}'。")
                found_disposal_keyword = True
                break # 找到一个关键词就退出循环，避免重复添加
        if not found_disposal_keyword:
            logger.info(f"行 {index + 1} - '处分决定' 字段不包含禁用关键词。")
            print(f"行 {index + 1} - '处分决定' 字段不包含禁用关键词。")

def validate_trial_report_keywords(row, index, excel_case_code, excel_person_code, issues_list, 
                                   trial_report_non_representative_mismatch_indices, 
                                   trial_report_detention_mismatch_indices, 
                                   compensation_amount_highlight_indices):
    """
    检查 '审理报告' 字段是否包含特定关键词（如非人大代表/政协委员等字样，和扣押字样），
    并检查是否包含“责令退赔”字样。
    """
    trial_text_raw = row.get("审理报告", "") if pd.notna(row.get("审理报告")) else ''
    
    # New keywords for "审理报告"
    trial_report_non_representative_keywords = ["非人大代表", "非政协委员", "非党委委员", "非中共党代表", "非纪委委员"]
    trial_report_detention_keyword = "扣押"
    compensation_keyword = "责令退赔"

    if "审理报告" in row.index and pd.notna(trial_text_raw) and trial_text_raw.strip() != '':
        # Check for non-representative keywords
        for keyword in trial_report_non_representative_keywords:
            if keyword in trial_text_raw:
                trial_report_non_representative_mismatch_indices.add(index)
                issues_list.append((index, excel_case_code, excel_person_code, f"CY审理报告中出现{keyword}等字样"))
                logger.warning(f"行 {index + 1} (案件编码: {excel_case_code}, 涉案人员编码: {excel_person_code})：审理报告中出现非人大代表/政协委员等字样: '{keyword}'。")
                print(f"警告：行 {index + 1} (案件编码: {excel_case_code}, 涉案人员编码: {excel_person_code})：审理报告中出现非人大代表/政协委员等字样: '{keyword}'。")
                
        # Check for "扣押" keyword
        if trial_report_detention_keyword in trial_text_raw:
            trial_report_detention_mismatch_indices.add(index)
            issues_list.append((index, excel_case_code, excel_person_code, "CY审理报告中出现扣押字样"))
            logger.warning(f"行 {index + 1} (案件编码: {excel_case_code}, 涉案人员编码: {excel_person_code})：审理报告中出现“扣押”字样。")
            print(f"警告：行 {index + 1} (案件编码: {excel_case_code}, 涉案人员编码: {excel_person_code})：审理报告中出现“扣押”字样。")

        # 检查“审理报告”是否包含“责令退赔”
        if compensation_keyword in trial_text_raw:
            compensation_amount_highlight_indices.add(index)
            issues_list.append((index, excel_case_code, excel_person_code, Config.VALIDATION_RULES["highlight_compensation_from_trial_report"]))
            logger.warning(f"行 {index + 1} (案件编码: {excel_case_code}, 涉案人员编码: {excel_person_code})：审理报告中出现“责令退赔”字样，请人工再次确认“责令退赔金额”。")
            print(f"警告：行 {index + 1} (案件编码: {excel_case_code}, 涉案人员编码: {excel_person_code})：审理报告中出现“责令退赔”字样，请人工再次确认“责令退赔金额”。")
        else:
            logger.info(f"行 {index + 1} (案件编码: {excel_case_code}, 涉案人员编码: {excel_person_code})：审理报告中未出现“责令退赔”字样。")
            print(f"行 {index + 1} (案件编码: {excel_case_code}, 涉案人员编码: {excel_person_code})：审理报告中未出现“责令退赔”字样。")

def highlight_recovery_amount(row, index, excel_case_code, excel_person_code, issues_list, recovery_amount_highlight_indices):
    """
    标记 '追缴失职渎职滥用职权造成的损失金额' 字段有值的行。
    """
    if pd.notna(row.get("追缴失职渎职滥用职权造成的损失金额")) and str(row.get("追缴失职渎职滥用职权造成的损失金额", "")).strip() != '':
        recovery_amount_highlight_indices.add(index)
        issues_list.append((index, excel_case_code, excel_person_code, Config.VALIDATION_RULES["highlight_recovery_amount"]))