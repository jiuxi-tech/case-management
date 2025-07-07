import logging
import pandas as pd

logger = logging.getLogger(__name__)

def validate_disciplinary_decision_rules(row, index, excel_case_code, excel_person_code, issues_list, disciplinary_decision_mismatch_indices, excel_disciplinary_decision, app_config):
    """
    验证处分决定字段的关键字规则。
    
    检查处分决定字段是否包含指定的关键字：
    "非人大代表"、"非政协委员"、"非党委委员"、"非中共党代表"、"非纪委委员"
    
    Args:
        row (pd.Series): DataFrame 的当前行数据。
        index (int): 当前行的索引。
        excel_case_code (str): Excel 中的案件编码。
        excel_person_code (str): Excel 中的涉案人员编码。
        issues_list (list): 用于收集所有发现问题的列表。
        disciplinary_decision_mismatch_indices (set): 用于收集处分决定包含关键字的行索引。
        excel_disciplinary_decision (str): Excel 中的处分决定内容。
        app_config (dict): Flask 应用的配置字典，包含Config类中的配置。
    """
    
    # 关键字包括："非人大代表、非政协委员、非党委委员、非中共党代表、非纪委委员"
    keywords = ["非人大代表", "非政协委员", "非党委委员", "非中共党代表", "非纪委委员"]
    
    # 检查处分决定字段是否为空
    if pd.isna(excel_disciplinary_decision) or excel_disciplinary_decision.strip() == '':
        logger.info(f"行 {index + 2} - 处分决定字段为空，跳过关键字检查")
        return
    
    # 检查是否包含关键字
    found_keyword = False
    for keyword in keywords:
        if keyword in excel_disciplinary_decision:
            disciplinary_decision_mismatch_indices.add(index)
            
            # 添加问题到issues_list，参考年龄规则的格式
            issues_list.append({
                '案件编码': excel_case_code,
                '涉案人员编码': excel_person_code,
                '行号': index + 2,
                '比对字段': f"CU处分决定",
                '被比对字段': f"CU处分决定",
                '问题描述': f"CU{index + 2}处分决定包含关键字'{keyword}'",
                '列名': app_config['COLUMN_MAPPINGS']['disciplinary_decision']
            })
            
            # 记录警告日志，参考年龄规则的日志格式
            logger.warning(f"<立案 - （处分决定关键字检查）> - 行 {index + 2} - 处分决定字段包含关键字: '{keyword}'")
            
            found_keyword = True
            break  # 找到一个关键字就退出循环，避免重复添加
    
    if not found_keyword:
        logger.info(f"行 {index + 2} - 处分决定字段不包含指定关键字")