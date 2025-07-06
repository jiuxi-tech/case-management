import logging
import pandas as pd

logger = logging.getLogger(__name__)

def validate_administrative_sanction_rules(row, index, excel_case_code, excel_person_code, issues_list, administrative_sanction_mismatch_indices,
                                         excel_administrative_sanction, decision_text_raw, app_config):
    """
    验证政务处分相关规则。
    比较 Excel 中的政务处分与处分决定中提取的政务处分关键词。

    参数:
        row (pd.Series): DataFrame 的当前行数据。
        index (int): 当前行的索引。
        excel_case_code (str): Excel 中的案件编码。
        excel_person_code (str): Excel 中的涉案人员编码。
        issues_list (list): 用于收集所有发现问题的列表。
        administrative_sanction_mismatch_indices (set): 用于收集政务处分不匹配的行索引。
        excel_administrative_sanction (str): Excel 中提取的政务处分。
        decision_text_raw (str): 处分决定的原始文本。
        app_config (dict): Flask 应用的配置字典。
    """
    
    # 获取政务处分关键词列表
    administrative_sanction_keywords = app_config['ADMINISTRATIVE_SANCTION_KEYWORDS']
    
    # 规则1: 政务处分与处分决定比对
    # 只有当政务处分有值，但处分决定中不包含任何政务处分关键词时，才标记为不一致
    if excel_administrative_sanction and excel_administrative_sanction.strip() != '':
        if not any(kw in decision_text_raw for kw in administrative_sanction_keywords):
            administrative_sanction_mismatch_indices.add(index)
            issues_list.append({
                '案件编码': excel_case_code,
                '涉案人员编码': excel_person_code,
                '行号': index + 2,
                '比对字段': f"BR{app_config['COLUMN_MAPPINGS']['administrative_sanction']}",
                '被比对字段': f"CU{app_config['COLUMN_MAPPINGS']['disciplinary_decision']}",
                '问题描述': f"BR{index + 2}{app_config['COLUMN_MAPPINGS']['administrative_sanction']}与CU{index + 2}处分决定不一致",
                '列名': app_config['COLUMN_MAPPINGS']['administrative_sanction']
            })
            logger.warning(f"<立案 - （1.政务处分验证）> - 行 {index + 2} - 政务处分 '{excel_administrative_sanction}' 与处分决定内容不一致")