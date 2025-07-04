import logging
import pandas as pd
from datetime import datetime
import re
# from config import Config  # 导入Config，因为某些验证规则需要用到其中的配置，但现在通过 app_config 传递

logger = logging.getLogger(__name__)

def parse_chinese_date(date_str):
    """
    Parses a Chinese date string (e.g., '2025年3月20日') into a datetime.date object.

    Args:
        date_str (str): 中文日期字符串。

    Returns:
        datetime.date or None: 解析后的日期对象，如果无法解析则为None。
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

def validate_trial_acceptance_time_vs_report(row, index, excel_case_code, excel_person_code, issues_list, trial_acceptance_time_mismatch_indices, app_config):
    """
    验证 '审理受理时间' 与 '审理报告' 开头时间内容的一致性。

    Args:
        row (pd.Series): DataFrame 的当前行数据。
        index (int): 当前行的索引。
        excel_case_code (str): Excel 中的案件编码。
        excel_person_code (str): Excel 中的涉案人员编码。
        issues_list (list): 用于收集所有发现问题的列表。
        trial_acceptance_time_mismatch_indices (set): 用于收集审理受理时间不匹配的行索引。
        app_config (dict): Flask 应用的配置字典，包含Config类中的配置。
    """
    excel_trial_acceptance_time = row.get(app_config['COLUMN_MAPPINGS']['trial_acceptance_time'])
    trial_text_raw = row.get(app_config['COLUMN_MAPPINGS']['trial_report'], "") if pd.notna(row.get(app_config['COLUMN_MAPPINGS']['trial_report'])) else ''

    if pd.notna(excel_trial_acceptance_time) and pd.notna(trial_text_raw):
        excel_date_obj = None
        if isinstance(excel_trial_acceptance_time, datetime):
            excel_date_obj = excel_trial_acceptance_time.date()
        elif isinstance(excel_trial_acceptance_time, str):
            try:
                excel_date_obj = pd.to_datetime(excel_trial_acceptance_time).date()
            except ValueError:
                logger.warning(f"行 {index + 1} - '{app_config['COLUMN_MAPPINGS']['trial_acceptance_time']}' 字段 '{excel_trial_acceptance_time}' 无法解析为日期。")
                trial_acceptance_time_mismatch_indices.add(index)
                issues_list.append((index, excel_case_code, excel_person_code, app_config['VALIDATION_RULES'].get("confirm_acceptance_time", "审理受理时间格式不正确"), "中")) # 增加风险等级
        
        if excel_date_obj:
            # 这里的正则表达式需要根据实际报告内容调整，确保能准确捕获日期
            # 原始模式: r"关于王.+?同志违纪案的审理报告主动交代" 看起来是针对特定报告标题的，可能需要更通用
            # 假设审理报告开头会有一个日期，例如 "2025年3月20日，关于王XX同志违纪案的审理报告..."
            match_phrase_pattern = r"(\d{4}年\d{1,2}月\d{1,2}日)" # 更通用的日期匹配模式
            match_phrase = re.search(match_phrase_pattern, trial_text_raw)

            if match_phrase:
                extracted_date_str = match_phrase.group(1)
                extracted_date_obj = parse_chinese_date(extracted_date_str)

                if extracted_date_obj:
                    if excel_date_obj != extracted_date_obj:
                        trial_acceptance_time_mismatch_indices.add(index)
                        issues_list.append((index, excel_case_code, excel_person_code, f"{app_config['COLUMN_MAPPINGS']['trial_acceptance_time']}与{app_config['COLUMN_MAPPINGS']['trial_report']}不一致", "中")) # 增加风险等级
                        logger.info(f"行 {index + 1} - 审理受理时间不一致：Excel: {excel_date_obj}, 审理报告: {extracted_date_obj}")
                    else:
                        logger.info(f"行 {index + 1} - 审理受理时间一致：Excel: {excel_date_obj}, 审理报告: {extracted_date_obj}")
                else:
                    logger.warning(f"行 {index + 1} - 从审理报告中提取的日期 '{extracted_date_str}' 无法解析。")
                    trial_acceptance_time_mismatch_indices.add(index)
                    issues_list.append((index, excel_case_code, excel_person_code, f"{app_config['COLUMN_MAPPINGS']['trial_report']}中审理受理时间格式不正确或未找到", "中")) # 增加风险等级
            else:
                logger.info(f"行 {index + 1} - 审理报告中未找到匹配的日期字符串。")
                trial_acceptance_time_mismatch_indices.add(index)
                issues_list.append((index, excel_case_code, excel_person_code, f"{app_config['COLUMN_MAPPINGS']['trial_report']}中未找到审理受理时间相关内容", "中")) # 增加风险等级
    elif pd.notna(excel_trial_acceptance_time) and pd.isna(trial_text_raw):
        logger.info(f"行 {index + 1} - '{app_config['COLUMN_MAPPINGS']['trial_acceptance_time']}' 有值但 '{app_config['COLUMN_MAPPINGS']['trial_report']}' 为空，无法比对。")
        trial_acceptance_time_mismatch_indices.add(index)
        issues_list.append((index, excel_case_code, excel_person_code, f"{app_config['COLUMN_MAPPINGS']['trial_acceptance_time']}有值但{app_config['COLUMN_MAPPINGS']['trial_report']}为空，无法比对", "中")) # 增加风险等级
    elif pd.isna(excel_trial_acceptance_time) and pd.notna(trial_text_raw):
        logger.info(f"行 {index + 1} - '{app_config['COLUMN_MAPPINGS']['trial_acceptance_time']}' 为空但 '{app_config['COLUMN_MAPPINGS']['trial_report']}' 有值，无法比对。")
        trial_acceptance_time_mismatch_indices.add(index)
        issues_list.append((index, excel_case_code, excel_person_code, f"{app_config['COLUMN_MAPPINGS']['trial_acceptance_time']}为空但{app_config['COLUMN_MAPPINGS']['trial_report']}有值，无法比对", "中")) # 增加风险等级

def validate_trial_closing_time_vs_report(row, index, excel_case_code, excel_person_code, issues_list, trial_closing_time_mismatch_indices, app_config):
    """
    验证 '审结时间' 与 '审理报告' 落款时间的一致性。

    Args:
        row (pd.Series): DataFrame 的当前行数据。
        index (int): 当前行的索引。
        excel_case_code (str): Excel 中的案件编码。
        excel_person_code (str): Excel 中的涉案人员编码。
        issues_list (list): 用于收集所有发现问题的列表。
        trial_closing_time_mismatch_indices (set): 用于收集审结时间不匹配的行索引。
        app_config (dict): Flask 应用的配置字典，包含Config类中的配置。
    """
    excel_trial_closing_time = row.get(app_config['COLUMN_MAPPINGS']['trial_closing_time'])
    trial_text_raw = row.get(app_config['COLUMN_MAPPINGS']['trial_report'], "") if pd.notna(row.get(app_config['COLUMN_MAPPINGS']['trial_report'])) else ''

    if pd.notna(excel_trial_closing_time) and pd.notna(trial_text_raw):
        excel_closing_date_obj = None
        if isinstance(excel_trial_closing_time, datetime):
            excel_closing_date_obj = excel_trial_closing_time.date()
        elif isinstance(excel_trial_closing_time, str):
            try:
                excel_closing_date_obj = pd.to_datetime(excel_trial_closing_time).date()
            except ValueError:
                logger.warning(f"行 {index + 1} - '{app_config['COLUMN_MAPPINGS']['trial_closing_time']}' 字段 '{excel_trial_closing_time}' 无法解析为日期。")
                trial_closing_time_mismatch_indices.add(index)
                issues_list.append((index, excel_case_code, excel_person_code, f"{app_config['COLUMN_MAPPINGS']['trial_closing_time']}格式不正确", "中")) # 增加风险等级
        
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
                            issues_list.append((index, excel_case_code, excel_person_code, f"{app_config['COLUMN_MAPPINGS']['trial_closing_time']}与{app_config['COLUMN_MAPPINGS']['trial_report']}不一致", "中")) # 增加风险等级
                            logger.info(f"行 {index + 1} - 审结时间不一致：Excel: {excel_closing_date_obj}, 审理报告落款: {extracted_closing_date_obj}")
                        else:
                            logger.info(f"行 {index + 1} - 审结时间一致：Excel: {excel_closing_date_obj}, 审理报告落款: {extracted_closing_date_obj}")
                    else:
                        logger.warning(f"行 {index + 1} - 从审理报告最后一行 '{last_line}' 中提取的日期 '{extracted_closing_date_str}' 无法解析。")
                        trial_closing_time_mismatch_indices.add(index)
                        issues_list.append((index, excel_case_code, excel_person_code, f"{app_config['COLUMN_MAPPINGS']['trial_report']}落款时间格式不正确或未找到", "中")) # 增加风险等级
                else:
                    logger.warning(f"行 {index + 1} - 审理报告最后一行 '{last_line}' 未找到日期格式。")
                    trial_closing_time_mismatch_indices.add(index)
                    issues_list.append((index, excel_case_code, excel_person_code, f"{app_config['COLUMN_MAPPINGS']['trial_report']}落款时间未找到", "中")) # 增加风险等级
            else:
                logger.info(f"行 {index + 1} - '{app_config['COLUMN_MAPPINGS']['trial_report']}' 为空，无法提取落款时间。")
                trial_closing_time_mismatch_indices.add(index)
                issues_list.append((index, excel_case_code, excel_person_code, f"{app_config['COLUMN_MAPPINGS']['trial_report']}为空，无法比对审结时间", "中")) # 增加风险等级
    elif pd.notna(excel_trial_closing_time) and pd.isna(trial_text_raw):
        logger.info(f"行 {index + 1} - '{app_config['COLUMN_MAPPINGS']['trial_closing_time']}' 有值但 '{app_config['COLUMN_MAPPINGS']['trial_report']}' 为空，无法比对。")
        trial_closing_time_mismatch_indices.add(index)
        issues_list.append((index, excel_case_code, excel_person_code, f"{app_config['COLUMN_MAPPINGS']['trial_closing_time']}有值但{app_config['COLUMN_MAPPINGS']['trial_report']}为空，无法比对", "中")) # 增加风险等级
    elif pd.isna(excel_trial_closing_time) and pd.notna(trial_text_raw):
        logger.info(f"行 {index + 1} - '{app_config['COLUMN_MAPPINGS']['trial_closing_time']}' 为空但 '{app_config['COLUMN_MAPPINGS']['trial_report']}' 有值，无法比对。")
        trial_closing_time_mismatch_indices.add(index)
        issues_list.append((index, excel_case_code, excel_person_code, f"{app_config['COLUMN_MAPPINGS']['trial_closing_time']}为空但{app_config['COLUMN_MAPPINGS']['trial_report']}有值，无法比对", "中")) # 增加风险等级

def validate_trial_authority_vs_reporting_agency(row, index, excel_case_code, excel_person_code, issues_list, trial_authority_agency_mismatch_indices, sl_authority_agency_mappings, app_config):
    """
    验证 '审理机关' 与 '填报单位名称' 是否与 SL 类别的机关单位字典数据匹配。

    Args:
        row (pd.Series): DataFrame 的当前行数据。
        index (int): 当前行的索引。
        excel_case_code (str): Excel 中的案件编码。
        excel_person_code (str): Excel 中的涉案人员编码。
        issues_list (list): 用于收集所有发现问题的列表。
        trial_authority_agency_mismatch_indices (set): 用于收集审理机关与填报单位名称不匹配的行索引。
        sl_authority_agency_mappings (list): 从数据库获取的 SL 类别的机关单位映射列表。
        app_config (dict): Flask 应用的配置字典，包含Config类中的配置。
    """
    excel_trial_authority = str(row.get(app_config['COLUMN_MAPPINGS']['trial_authority'], "")).strip()
    excel_reporting_agency = str(row.get(app_config['COLUMN_MAPPINGS']['reporting_agency'], "")).strip()

    if excel_trial_authority and excel_reporting_agency:
        found_match = False
        for mapping in sl_authority_agency_mappings:
            if (mapping["authority"] == excel_trial_authority and 
                mapping["agency"] == excel_reporting_agency):
                found_match = True
                logger.info(f"行 {index + 1} - 审理机关 '{excel_trial_authority}' 和 填报单位名称 '{excel_reporting_agency}' 匹配成功 (Category: SL)。")
                break
        
        if not found_match:
            trial_authority_agency_mismatch_indices.add(index)
            # Directly construct the message using column mappings to avoid KeyError
            issues_list.append((index, excel_case_code, excel_person_code, 
                                f"{app_config['COLUMN_MAPPINGS']['trial_authority']}与{app_config['COLUMN_MAPPINGS']['reporting_agency']}不一致", 
                                "高")) # 增加风险等级
            logger.warning(f"行 {index + 1} - 审理机关 '{excel_trial_authority}' 和 填报单位名称 '{excel_reporting_agency}' 不匹配或Category不为SL。")
        else:
            logger.info(f"行 {index + 1} - 审理机关 ('{excel_trial_authority}') 和 填报单位名称 ('{excel_reporting_agency}') 在对应表中找到匹配记录。")
    else:
        logger.info(f"行 {index + 1} - '{app_config['COLUMN_MAPPINGS']['trial_authority']}' 或 '{app_config['COLUMN_MAPPINGS']['reporting_agency']}' 为空，跳过比对。审理机关: '{excel_trial_authority}', 填报单位名称: '{excel_reporting_agency}'")
        trial_authority_agency_mismatch_indices.add(index)
        issues_list.append((index, excel_case_code, excel_person_code, f"{app_config['COLUMN_MAPPINGS']['trial_authority']}或{app_config['COLUMN_MAPPINGS']['reporting_agency']}为空，无法比对", "中")) # 增加风险等级

def validate_disposal_decision_keywords(row, index, excel_case_code, excel_person_code, issues_list, disposal_decision_keyword_mismatch_indices, app_config):
    """
    检查 '处分决定' 字段是否包含禁用关键词 (Config.DISPOSAL_DECISION_KEYWORDS)。

    Args:
        row (pd.Series): DataFrame 的当前行数据。
        index (int): 当前行的索引。
        excel_case_code (str): Excel 中的案件编码。
        excel_person_code (str): Excel 中的涉案人员编码。
        issues_list (list): 用于收集所有发现问题的列表。
        disposal_decision_keyword_mismatch_indices (set): 用于收集处分决定包含禁用关键词的行索引。
        app_config (dict): Flask 应用的配置字典，包含Config类中的配置。
    """
    decision_text_raw = row.get(app_config['COLUMN_MAPPINGS']['disciplinary_decision'], "") if pd.notna(row.get(app_config['COLUMN_MAPPINGS']['disciplinary_decision'])) else ''
    
    if app_config['COLUMN_MAPPINGS']['disciplinary_decision'] in row.index and pd.notna(decision_text_raw) and decision_text_raw.strip() != '':
        found_disposal_keyword = False
        for keyword in app_config['DISPOSAL_DECISION_KEYWORDS']: # 从 app_config 获取关键词
            if keyword in decision_text_raw:
                disposal_decision_keyword_mismatch_indices.add(index)
                issues_list.append((index, excel_case_code, excel_person_code, app_config['VALIDATION_RULES'].get("disposal_decision_keyword_highlight", "处分决定中出现非人大代表、非政协委员、非党委委员、非中共党代表、非纪委委员等字样"), "高")) # 增加风险等级
                logger.warning(f"行 {index + 1} - '{app_config['COLUMN_MAPPINGS']['disciplinary_decision']}' 字段包含禁用关键词: '{keyword}'。")
                found_disposal_keyword = True
                break # 找到一个关键词就退出循环，避免重复添加
        if not found_disposal_keyword:
            logger.info(f"行 {index + 1} - '{app_config['COLUMN_MAPPINGS']['disciplinary_decision']}' 字段不包含禁用关键词。")

def validate_trial_report_keywords(row, index, excel_case_code, excel_person_code, issues_list, 
                                   trial_report_non_representative_mismatch_indices, 
                                   trial_report_detention_mismatch_indices, 
                                   compensation_amount_highlight_indices, app_config):
    """
    检查 '审理报告' 字段是否包含特定关键词（如非人大代表/政协委员等字样，和扣押字样），
    并检查是否包含“责令退赔”字样。

    Args:
        row (pd.Series): DataFrame 的当前行数据。
        index (int): 当前行的索引。
        excel_case_code (str): Excel 中的案件编码。
        excel_person_code (str): Excel 中的涉案人员编码。
        issues_list (list): 用于收集所有发现问题的列表。
        trial_report_non_representative_mismatch_indices (set): 用于收集审理报告中包含非代表关键词的行索引。
        trial_report_detention_mismatch_indices (set): 用于收集审理报告中包含“扣押”关键词的行索引。
        compensation_amount_highlight_indices (set): 用于收集审理报告中包含“责令退赔”关键词的行索引。
        app_config (dict): Flask 应用的配置字典，包含Config类中的配置。
    """
    trial_text_raw = row.get(app_config['COLUMN_MAPPINGS']['trial_report'], "") if pd.notna(row.get(app_config['COLUMN_MAPPINGS']['trial_report'])) else ''
    
    # New keywords for "审理报告" - 从 app_config 获取
    trial_report_non_representative_keywords = app_config['DISPOSAL_DECISION_KEYWORDS'] # 这里的关键词与处分决定关键词相同
    trial_report_detention_keyword = "扣押"
    compensation_keyword = "责令退赔"

    if app_config['COLUMN_MAPPINGS']['trial_report'] in row.index and pd.notna(trial_text_raw) and trial_text_raw.strip() != '':
        # Check for non-representative keywords
        for keyword in trial_report_non_representative_keywords:
            if keyword in trial_text_raw:
                trial_report_non_representative_mismatch_indices.add(index)
                issues_list.append((index, excel_case_code, excel_person_code, f"{app_config['COLUMN_MAPPINGS']['trial_report']}中出现{keyword}等字样", "中")) # 增加风险等级
                logger.warning(f"行 {index + 1} (案件编码: {excel_case_code}, 涉案人员编码: {excel_person_code})：审理报告中出现非人大代表/政协委员等字样: '{keyword}'。")
                
        # Check for "扣押" keyword
        if trial_report_detention_keyword in trial_text_raw:
            trial_report_detention_mismatch_indices.add(index)
            issues_list.append((index, excel_case_code, excel_person_code, f"{app_config['COLUMN_MAPPINGS']['trial_report']}中出现扣押字样", "中")) # 增加风险等级
            logger.warning(f"行 {index + 1} (案件编码: {excel_case_code}, 涉案人员编码: {excel_person_code})：审理报告中出现“扣押”字样。")

        # 检查“审理报告”是否包含“责令退赔”
        if compensation_keyword in trial_text_raw:
            compensation_amount_highlight_indices.add(index)
            issues_list.append((index, excel_case_code, excel_person_code, app_config['VALIDATION_RULES'].get("highlight_compensation_from_trial_report", "审理报告中含有责令退赔四字，请人工再次确认责令退赔金额"), "中")) # 增加风险等级
            logger.warning(f"行 {index + 1} (案件编码: {excel_case_code}, 涉案人员编码: {excel_person_code})：审理报告中出现“责令退赔”字样，请人工再次确认“责令退赔金额”。")
        else:
            logger.info(f"行 {index + 1} (案件编码: {excel_case_code}, 涉案人员编码: {excel_person_code})：审理报告中未出现“责令退赔”字样。")

def highlight_recovery_amount(row, index, excel_case_code, excel_person_code, issues_list, recovery_amount_highlight_indices, app_config):
    """
    标记 '追缴失职渎职滥用职权造成的损失金额' 字段有值的行。

    Args:
        row (pd.Series): DataFrame 的当前行数据。
        index (int): 当前行的索引。
        excel_case_code (str): Excel 中的案件编码。
        excel_person_code (str): Excel 中的涉案人员编码。
        issues_list (list): 用于收集所有发现问题的列表。
        recovery_amount_highlight_indices (set): 用于收集追缴金额有值的行索引。
        app_config (dict): Flask 应用的配置字典，包含Config类中的配置。
    """
    recovery_amount_col = app_config['COLUMN_MAPPINGS']['recovery_amount']
    if pd.notna(row.get(recovery_amount_col)) and str(row.get(recovery_amount_col, "")).strip() != '':
        recovery_amount_highlight_indices.add(index)
        issues_list.append((index, excel_case_code, excel_person_code, app_config['VALIDATION_RULES'].get("highlight_recovery_amount", "追缴失职渎职滥用职权造成的损失金额请再次确认"), "低")) # 增加风险等级
        logger.info(f"行 {index + 1} - '{recovery_amount_col}' 字段有值，已标记。")

