import logging

logger = logging.getLogger(__name__)

def validate_trial_authority_rules(row, index, excel_case_code, excel_person_code, issues_list, trial_authority_mismatch_indices, excel_trial_authority, excel_reporting_agency, sl_authority_agency_mappings, app_config):
    """
    验证审理机关字段。
    
    当审理机关与填报单位名称不匹配时，将该行索引添加到 trial_authority_mismatch_indices 集合中，
    并在 issues_list 中添加相应的问题记录。
    
    Args:
        row (pd.Series): DataFrame 的当前行数据。
        index (int): 当前行的索引。
        excel_case_code (str): Excel 中的案件编码。
        excel_person_code (str): Excel 中的涉案人员编码。
        issues_list (list): 用于收集所有发现问题的列表。
        trial_authority_mismatch_indices (set): 收集所有"审理机关"不匹配的行索引。
        excel_trial_authority (str): 审理机关字段的值。
        excel_reporting_agency (str): 填报单位名称字段的值。
        sl_authority_agency_mappings (list): 从数据库获取的 SL 类别的机关单位映射列表。
        app_config (dict): Flask 应用的配置字典，包含Config类中的配置。
    
    Returns:
        None (issues_list 和 trial_authority_mismatch_indices 会在函数内部被修改)。
    """
    # 规则1: 审理机关与填报单位名称比对
    if excel_trial_authority and excel_reporting_agency:
        found_match = False
        for mapping in sl_authority_agency_mappings:
            if (mapping["authority"] == excel_trial_authority and 
                mapping["agency"] == excel_reporting_agency):
                found_match = True
                logger.info(f"行 {index + 2} - 审理机关 '{excel_trial_authority}' 和 填报单位名称 '{excel_reporting_agency}' 匹配成功 (Category: SL)。")
                break
        
        if not found_match:
            trial_authority_mismatch_indices.add(index)
            # 使用与年龄规则一致的日志格式
            issues_list.append({
                '行号': index + 2,
                '案件编码': excel_case_code,
                '涉案人员编码': excel_person_code,
                '比对字段': "CR审理机关",
                '被比对字段': f"A{app_config['COLUMN_MAPPINGS']['reporting_agency']}",
                '问题描述': f"CR{index + 2}审理机关与A填报单位不一致",
                '列名': app_config['COLUMN_MAPPINGS']['trial_authority']
            })
            logger.warning(f"<立案 - （1.审理机关与填报单位名称）> - 行 {index + 2} - 审理机关 '{excel_trial_authority}' 和 填报单位名称 '{excel_reporting_agency}' 不匹配或Category不为SL。")
    else:
        # 处理空值情况
        trial_authority_mismatch_indices.add(index)
        missing_field = []
        if not excel_trial_authority:
            missing_field.append(app_config['COLUMN_MAPPINGS']['trial_authority'])
        if not excel_reporting_agency:
            missing_field.append(app_config['COLUMN_MAPPINGS']['reporting_agency'])
        
        issues_list.append({
            '行号': index + 2,
            '案件编码': excel_case_code,
            '涉案人员编码': excel_person_code,
            '比对字段': "CR审理机关",
            '被比对字段': f"A{app_config['COLUMN_MAPPINGS']['reporting_agency']}",
            '问题描述': f"CR{index + 2}审理机关或A填报单位为空，无法比对",
            '列名': app_config['COLUMN_MAPPINGS']['trial_authority']
        })
        logger.info(f"行 {index + 2} - '{app_config['COLUMN_MAPPINGS']['trial_authority']}' 或 '{app_config['COLUMN_MAPPINGS']['reporting_agency']}' 为空，跳过比对。审理机关: '{excel_trial_authority}', 填报单位名称: '{excel_reporting_agency}'")
    
    logger.info(f"第 {index + 1} 行的审理机关相关规则验证完成。")