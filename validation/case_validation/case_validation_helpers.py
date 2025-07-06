import logging
import pandas as pd
import re

# 导入必要的提取器函数
from .case_extractors_names import (
    extract_name_from_case_report,
    extract_name_from_decision,
    extract_name_from_trial_report
)


from .case_extractors_demographics import (
    extract_suspected_violation_from_case_report,
    extract_suspected_violation_from_decision
)

logger = logging.getLogger(__name__)


def validate_brief_case_details_rules(row, index, excel_case_code, excel_person_code, issues_list, brief_case_details_mismatch_indices,
                                      excel_brief_case_details, investigated_person, report_text_raw, decision_text_raw, app_config):
    """
    验证简要案情相关规则。
    根据“处分决定”是否为空，从“立案报告”或“处分决定”中提取简要案情，并与Excel中的简要案情进行比较。

    参数:
        row (pd.Series): DataFrame 的当前行数据。
        index (int): 当前行的索引。
        excel_case_code (str): Excel 中的案件编码。
        excel_person_code (str): Excel 中的涉案人员编码。
        issues_list (list): 用于收集所有发现问题的列表。
        brief_case_details_mismatch_indices (set): 用于收集简要案情不匹配的行索引。
        excel_brief_case_details (str): Excel 中提取的简要案情。
        investigated_person (str): 被调查人姓名。
        report_text_raw (str): 立案报告的原始文本。
        decision_text_raw (str): 处分决定的原始文本。
        app_config (dict): Flask 应用的配置字典。
    """
    is_brief_case_details_mismatch = False
    extracted_brief_case_details = None
    source_document_for_issue = "" 

    if pd.isna(row[app_config['COLUMN_MAPPINGS']["disciplinary_decision"]]) or decision_text_raw == '':
        extracted_brief_case_details = extract_suspected_violation_from_case_report(report_text_raw)
        source_document_for_issue = app_config['COLUMN_MAPPINGS']['case_report'] # "立案报告"
        logger.info(f"行 {index + 1} - 处分决定为空，从立案报告提取简要案情：'{extracted_brief_case_details}'")
    else:
        extracted_brief_case_details = extract_suspected_violation_from_decision(decision_text_raw, investigated_person)
        source_document_for_issue = app_config['COLUMN_MAPPINGS']['disciplinary_decision'] # "处分决定"
        logger.info(f"行 {index + 1} - 处分决定不为空，从处分决定提取简要案情：'{extracted_brief_case_details}'")

    if extracted_brief_case_details is None:
        if excel_brief_case_details:
            is_brief_case_details_mismatch = True
            issues_list.append((index, excel_case_code, excel_person_code, 
                                f"{app_config['COLUMN_MAPPINGS']['brief_case_details']}与{source_document_for_issue}不一致（未能提取到内容）", "中")) # 增加风险等级
            logger.info(f"行 {index + 1} - 简要案情不匹配: Excel有值 ('{excel_brief_case_details}') 但未能从{source_document_for_issue}中提取。")
    else:
        cleaned_excel_brief_case_details = re.sub(r'\s+', '', excel_brief_case_details)
        
        if cleaned_excel_brief_case_details != extracted_brief_case_details:
            is_brief_case_details_mismatch = True
            issues_list.append((index, excel_case_code, excel_person_code, 
                                f"{app_config['COLUMN_MAPPINGS']['brief_case_details']}与{source_document_for_issue}不一致", "中")) # 增加风险等级
            logger.info(f"行 {index + 1} - 简要案情不匹配: Excel简要案情 ('{cleaned_excel_brief_case_details}') vs 提取简要案情 ('{extracted_brief_case_details}') (来源: {source_document_for_issue})")

    if is_brief_case_details_mismatch:
        brief_case_details_mismatch_indices.add(index)

