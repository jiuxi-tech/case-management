import logging

logger = logging.getLogger(__name__)

def validate_trial_report_rules(row, index, excel_case_code, excel_person_code, issues_list, trial_report_mismatch_indices, excel_trial_report, app_config):
    """
    验证审理报告字段。
    
    当审理报告中包含指定关键词时，将该行索引添加到 trial_report_mismatch_indices 集合中，
    并在 issues_list 中添加问题描述。
    
    参数:
    row (pd.Series): 当前行的数据。
    index (int): 当前行的索引。
    excel_case_code (str): 案件编码。
    excel_person_code (str): 涉案人员编码。
    issues_list (list): 包含所有问题的列表，每个问题是一个字典。
    trial_report_mismatch_indices (set): 收集所有"审理报告"需要标黄的行索引。
    excel_trial_report (str): 审理报告字段的值。
    app_config (dict): Flask 应用的配置字典，包含Config类中的配置。
    
    返回:
    None (issues_list 和 trial_report_mismatch_indices 会在函数内部被修改)。
    """
    # 规则1: 审理报告关键字处理
    # 关键字包括："非人大代表、非政协委员、非党委委员、非中共党代表、非纪委委员，扣押"
    keywords = ["非人大代表", "非政协委员", "非党委委员", "非中共党代表", "非纪委委员", "扣押"]
    
    if excel_trial_report:
        found_keywords = []
        for keyword in keywords:
            if keyword in excel_trial_report:
                found_keywords.append(keyword)
        
        if found_keywords:
            trial_report_mismatch_indices.add(index)
            issues_list.append({
                '案件编码': excel_case_code,
                '涉案人员编码': excel_person_code,
                '行号': index + 2,
                '比对字段': f"CY{app_config['COLUMN_MAPPINGS']['trial_report']}",
                '被比对字段': f"CY{app_config['COLUMN_MAPPINGS']['trial_report']}",
                '问题描述': f"CY{index + 2}{app_config['COLUMN_MAPPINGS']['trial_report']}审理报告包含关键词",
                '列名': app_config['COLUMN_MAPPINGS']['trial_report']
            })
            logger.warning(f"<立案 - （1.审理报告关键词检查）> - 行 {index + 2} - 审理报告中包含关键词: {', '.join(found_keywords)}")
        else:
            logger.info(f"行 {index + 1} (案件编码: {excel_case_code}, 涉案人员编码: {excel_person_code})：审理报告中未发现指定关键词。")
    else:
        logger.info(f"行 {index + 1} (案件编码: {excel_case_code}, 涉案人员编码: {excel_person_code})：审理报告字段为空。")
    
    logger.info(f"第 {index + 1} 行的审理报告相关规则验证完成。")