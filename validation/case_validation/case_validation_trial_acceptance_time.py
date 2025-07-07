import logging
import pandas as pd
from datetime import datetime
import re
from .case_document_validators import parse_chinese_date

logger = logging.getLogger(__name__)

def validate_trial_acceptance_time_rules(row, index, excel_case_code, excel_person_code, issues_list, trial_acceptance_time_mismatch_indices, excel_trial_acceptance_time, excel_trial_report, app_config):
    """
    验证审理受理时间字段。
    
    比较 Excel 中的审理受理时间与审理报告中的时间内容的一致性。
    使用 CR 作为字段标识，参考年龄规则的日志风格和编号表字段结构。
    
    参数:
    row (pd.Series): 当前行的数据。
    index (int): 当前行的索引。
    excel_case_code (str): 案件编码。
    excel_person_code (str): 涉案人员编码。
    issues_list (list): 包含所有问题的列表，每个问题是一个字典。
    trial_acceptance_time_mismatch_indices (set): 收集所有"审理受理时间"不匹配的行索引。
    excel_trial_acceptance_time (str): 审理受理时间字段的值。
    excel_trial_report (str): 审理报告字段的值。
    app_config (dict): Flask 应用的配置字典，包含Config类中的配置。
    
    返回:
    None (issues_list 和 trial_acceptance_time_mismatch_indices 会在函数内部被修改)。
    """
    # 规则1: 审理受理时间与审理报告比对
    if pd.notna(excel_trial_acceptance_time) and pd.notna(excel_trial_report):
        excel_date_obj = None
        if isinstance(excel_trial_acceptance_time, datetime):
            excel_date_obj = excel_trial_acceptance_time.date()
        elif isinstance(excel_trial_acceptance_time, str):
            try:
                excel_date_obj = pd.to_datetime(excel_trial_acceptance_time).date()
            except ValueError:
                trial_acceptance_time_mismatch_indices.add(index)
                issues_list.append({
                    '案件编码': excel_case_code,
                    '涉案人员编码': excel_person_code,
                    '行号': index + 2,
                    '比对字段': f"CP{app_config['COLUMN_MAPPINGS']['trial_acceptance_time']}",
                    '被比对字段': f"CY{app_config['COLUMN_MAPPINGS']['trial_report']}",
                    '问题描述': f"CP{index + 2}{app_config['COLUMN_MAPPINGS']['trial_acceptance_time']}格式不正确",
                    '列名': app_config['COLUMN_MAPPINGS']['trial_acceptance_time']
                })
                logger.warning(f"<立案 - （1.审理受理时间格式）> - 行 {index + 2} - 审理受理时间 '{excel_trial_acceptance_time}' 格式不正确")
                return
        
        if excel_date_obj:
            # 从审理报告中提取日期
            match_phrase_pattern = r"(\d{4}年\d{1,2}月\d{1,2}日)"
            match_phrase = re.search(match_phrase_pattern, excel_trial_report)

            if match_phrase:
                extracted_date_str = match_phrase.group(1)
                extracted_date_obj = parse_chinese_date(extracted_date_str)

                if extracted_date_obj:
                    if excel_date_obj != extracted_date_obj:
                        trial_acceptance_time_mismatch_indices.add(index)
                        issues_list.append({
                            '案件编码': excel_case_code,
                            '涉案人员编码': excel_person_code,
                            '行号': index + 2,
                            '比对字段': f"CP{app_config['COLUMN_MAPPINGS']['trial_acceptance_time']}",
                    '被比对字段': f"CY{app_config['COLUMN_MAPPINGS']['trial_report']}",
                    '问题描述': f"CP{index + 2}{app_config['COLUMN_MAPPINGS']['trial_acceptance_time']}与CY{index + 2}审理报告不一致",
                            '列名': app_config['COLUMN_MAPPINGS']['trial_acceptance_time']
                        })
                        logger.warning(f"<立案 - （1.审理受理时间与审理报告）> - 行 {index + 2} - 审理受理时间 '{excel_date_obj}' 与审理报告时间 '{extracted_date_obj}' 不一致")
                    else:
                        logger.info(f"行 {index + 1} (案件编码: {excel_case_code}, 涉案人员编码: {excel_person_code})：审理受理时间与审理报告时间一致。")
                else:
                    trial_acceptance_time_mismatch_indices.add(index)
                    issues_list.append({
                        '案件编码': excel_case_code,
                        '涉案人员编码': excel_person_code,
                        '行号': index + 2,
                        '比对字段': f"CP{app_config['COLUMN_MAPPINGS']['trial_acceptance_time']}",
                        '被比对字段': f"CY{app_config['COLUMN_MAPPINGS']['trial_report']}",
                        '问题描述': f"CY{index + 2}审理报告中审理受理时间格式不正确或未找到",
                        '列名': app_config['COLUMN_MAPPINGS']['trial_acceptance_time']
                    })
                    logger.warning(f"<立案 - （1.审理受理时间与审理报告）> - 行 {index + 2} - 审理报告中提取的日期 '{extracted_date_str}' 无法解析")
            else:
                trial_acceptance_time_mismatch_indices.add(index)
                issues_list.append({
                    '案件编码': excel_case_code,
                    '涉案人员编码': excel_person_code,
                    '行号': index + 2,
                    '比对字段': f"CP{app_config['COLUMN_MAPPINGS']['trial_acceptance_time']}",
                    '被比对字段': f"CY{app_config['COLUMN_MAPPINGS']['trial_report']}",
                    '问题描述': f"CY{index + 2}审理报告中未找到审理受理时间相关内容",
                    '列名': app_config['COLUMN_MAPPINGS']['trial_acceptance_time']
                })
                logger.warning(f"<立案 - （1.审理受理时间与审理报告）> - 行 {index + 2} - 审理报告中未找到匹配的日期字符串")
    elif pd.notna(excel_trial_acceptance_time) and pd.isna(excel_trial_report):
        trial_acceptance_time_mismatch_indices.add(index)
        issues_list.append({
            '案件编码': excel_case_code,
            '涉案人员编码': excel_person_code,
            '行号': index + 2,
            '比对字段': f"CP{app_config['COLUMN_MAPPINGS']['trial_acceptance_time']}",
            '被比对字段': f"CY{app_config['COLUMN_MAPPINGS']['trial_report']}",
            '问题描述': f"CP{index + 2}审理受理时间有值但CY{index + 2}审理报告为空，无法比对",
            '列名': app_config['COLUMN_MAPPINGS']['trial_acceptance_time']
        })
        logger.warning(f"<立案 - （1.审理受理时间与审理报告）> - 行 {index + 2} - 审理受理时间有值但审理报告为空，无法比对")
    elif pd.isna(excel_trial_acceptance_time) and pd.notna(excel_trial_report):
        trial_acceptance_time_mismatch_indices.add(index)
        issues_list.append({
            '案件编码': excel_case_code,
            '涉案人员编码': excel_person_code,
            '行号': index + 2,
            '比对字段': f"CP{app_config['COLUMN_MAPPINGS']['trial_acceptance_time']}",
            '被比对字段': f"CY{app_config['COLUMN_MAPPINGS']['trial_report']}",
            '问题描述': f"CP{index + 2}审理受理时间为空但CY{index + 2}审理报告有值，无法比对",
            '列名': app_config['COLUMN_MAPPINGS']['trial_acceptance_time']
        })
        logger.warning(f"<立案 - （1.审理受理时间与审理报告）> - 行 {index + 2} - 审理受理时间为空但审理报告有值，无法比对")
    else:
        logger.info(f"行 {index + 1} (案件编码: {excel_case_code}, 涉案人员编码: {excel_person_code})：审理受理时间和审理报告均为空，跳过验证。")
    
    logger.info(f"第 {index + 1} 行的审理受理时间相关规则验证完成。")