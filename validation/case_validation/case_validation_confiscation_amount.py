import logging
import pandas as pd

logger = logging.getLogger(__name__)

def validate_confiscation_amount_rules(row, index, excel_case_code, excel_person_code, issues_list, confiscation_amount_mismatch_indices,
                                     excel_confiscation_amount, trial_text_raw, app_config):
    """
    验证收缴金额（万元）相关规则。
    比较 Excel 中的收缴金额与审理报告中是否包含"收缴"关键词。

    参数:
        row (pd.Series): DataFrame 的当前行数据。
        index (int): 当前行的索引。
        excel_case_code (str): Excel 中的案件编码。
        excel_person_code (str): Excel 中的涉案人员编码。
        issues_list (list): 用于收集所有发现问题的列表。
        confiscation_amount_mismatch_indices (set): 用于收集收缴金额不匹配的行索引。
        excel_confiscation_amount (str): Excel 中提取的收缴金额。
        trial_text_raw (str): 审理报告的原始文本。
        app_config (dict): Flask 应用的配置字典。
    """
    
    # 规则1: 收缴金额与审理报告比对
    # 当审理报告中包含"收缴"关键词时，标记为需要人工确认
    if trial_text_raw and "收缴" in trial_text_raw:
        confiscation_amount_mismatch_indices.add(index)
        issues_list.append({
            '案件编码': excel_case_code,
            '涉案人员编码': excel_person_code,
            '行号': index + 2,
            '比对字段': f"CF{app_config['COLUMN_MAPPINGS']['confiscation_amount']}",
            '被比对字段': f"CY{app_config['COLUMN_MAPPINGS']['trial_report']}",
            '问题描述': f"CF{index + 2}{app_config['COLUMN_MAPPINGS']['confiscation_amount']}与CY{index + 2}审理报告不一致",
            '列名': app_config['COLUMN_MAPPINGS']['confiscation_amount']
        })
        logger.warning(f"<立案 - （1.收缴金额与审理报告）> - 行 {index + 2} - 审理报告中含有收缴二字，请人工再次确认收缴金额 '{excel_confiscation_amount}'")