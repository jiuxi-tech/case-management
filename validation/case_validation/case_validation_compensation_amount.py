import pandas as pd
import logging

logger = logging.getLogger(__name__)

def validate_compensation_amount_rules(row, index, excel_case_code, excel_person_code, issues_list, compensation_amount_mismatch_indices, excel_compensation_amount, excel_trial_report, app_config):
    """
    验证责令退赔金额字段。
    
    当审理报告中包含"责令退赔"关键词时，将该行索引添加到 compensation_amount_mismatch_indices 集合中，
    并在 issues_list 中添加问题描述。
    
    参数:
    row (pd.Series): 当前行的数据。
    index (int): 当前行的索引。
    excel_case_code (str): 案件编码。
    excel_person_code (str): 涉案人员编码。
    issues_list (list): 包含所有问题的列表，每个问题是一个字典。
    compensation_amount_mismatch_indices (set): 收集所有"责令退赔金额"需要标黄的行索引。
    excel_compensation_amount (str): 责令退赔金额字段的值。
    excel_trial_report (str): 审理报告字段的值。
    app_config (dict): Flask 应用的配置字典，包含Config类中的配置。
    
    返回:
    None (issues_list 和 compensation_amount_mismatch_indices 会在函数内部被修改)。
    """
    # 规则1: 责令退赔金额与审理报告比对
    # 当审理报告中包含"责令退赔"关键词时，标记为需要人工确认
    if excel_trial_report and "责令退赔" in excel_trial_report:
        compensation_amount_mismatch_indices.add(index)
        issues_list.append({
            '案件编码': excel_case_code,
            '涉案人员编码': excel_person_code,
            '行号': index + 2,
            '比对字段': f"CH{app_config['COLUMN_MAPPINGS']['compensation_amount']}",
            '被比对字段': f"CY{app_config['COLUMN_MAPPINGS']['trial_report']}",
            '问题描述': f"CH{index + 2}{app_config['COLUMN_MAPPINGS']['compensation_amount']}与CY{index + 2}审理报告不一致",
            '列名': app_config['COLUMN_MAPPINGS']['compensation_amount']
        })
        logger.warning(f"<立案 - （1.责令退赔金额与审理报告）> - 行 {index + 2} - 审理报告中含有责令退赔关键词，请人工再次确认责令退赔金额 '{excel_compensation_amount}'")
    else:
        logger.info(f"行 {index + 1} (案件编码: {excel_case_code}, 涉案人员编码: {excel_person_code})：审理报告中未出现\"责令退赔\"字样。")
    
    logger.info(f"第 {index + 1} 行的责令退赔金额相关规则验证完成。")