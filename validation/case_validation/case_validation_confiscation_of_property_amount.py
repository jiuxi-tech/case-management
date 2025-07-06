import logging

logger = logging.getLogger(__name__)

def validate_confiscation_of_property_amount_rules(row, index, excel_case_code, excel_person_code, issues_list, confiscation_of_property_amount_mismatch_indices,
                                                  excel_confiscation_of_property_amount, trial_text_raw, app_config):
    """
    验证没收金额字段规则。
    与"审理报告"字段内容进行对比，查找字符串"没收金额"，
    若出现"没收金额"四字，将该行索引添加到confiscation_of_property_amount_mismatch_indices集合中，
    并在issues_list中添加问题描述。

    参数:
        row (pd.Series): DataFrame 的当前行数据。
        index (int): 当前行的索引。
        excel_case_code (str): Excel 中的案件编码。
        excel_person_code (str): Excel 中的涉案人员编码。
        issues_list (list): 用于收集所有发现问题的列表。
        confiscation_of_property_amount_mismatch_indices (set): 用于收集没收金额不匹配的行索引。
        excel_confiscation_of_property_amount (str): Excel 中提取的没收金额。
        trial_text_raw (str): 审理报告的原始文本。
        app_config (dict): Flask 应用的配置字典。
    """
    
    # 检查审理报告中是否包含"没收金额"四字
    if "没收金额" in trial_text_raw:
        confiscation_of_property_amount_mismatch_indices.add(index)
        
        # 添加问题到issues_list，格式与年龄规则保持一致
        issues_list.append({
            '案件编码': excel_case_code,
            '涉案人员编码': excel_person_code,
            '行号': index + 2,
            '比对字段': f"CG{app_config['COLUMN_MAPPINGS']['confiscation_of_property_amount']}",
            '被比对字段': f"CY{app_config['COLUMN_MAPPINGS']['trial_report']}",
            '问题描述': app_config['VALIDATION_RULES'].get('inconsistent_case_confiscation_of_property_amount_with_trial_report', f"CG{index + 2}{app_config['COLUMN_MAPPINGS']['confiscation_of_property_amount']}与CY{index + 2}审理报告不一致"),
            '列名': app_config['COLUMN_MAPPINGS']['confiscation_of_property_amount']
        })
        logger.warning(f"<立案 - （1.没收金额与审理报告）> - 行 {index + 2} - 审理报告中含有没收金额四字，请人工再次确认没收金额 '{excel_confiscation_of_property_amount}'")