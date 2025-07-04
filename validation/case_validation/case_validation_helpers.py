import logging
import pandas as pd
import re

# 导入必要的提取器函数
from .case_extractors_names import (
    extract_name_from_case_report,
    extract_name_from_decision,
    extract_name_from_trial_report
)
from .case_extractors_gender import (
    extract_gender_from_case_report,
    extract_gender_from_decision_report,
    extract_gender_from_investigation_report,
    extract_gender_from_trial_report
)
from .case_extractors_birth_info import (
    extract_birth_year_from_case_report,
    extract_birth_year_from_decision_report,
    extract_birth_year_from_investigation_report,
    extract_birth_year_from_trial_report
)
from .case_extractors_demographics import (
    extract_suspected_violation_from_case_report,
    extract_suspected_violation_from_decision
)

logger = logging.getLogger(__name__)

def validate_gender_rules(row, index, excel_case_code, excel_person_code, issues_list, gender_mismatch_indices,
                          excel_gender, report_text_raw, decision_text_raw, investigation_text_raw, trial_text_raw, app_config):
    """
    验证性别相关规则。
    比较 Excel 中的性别与立案报告、处分决定、审查调查报告、审理报告中提取的性别。

    参数:
        row (pd.Series): DataFrame 的当前行数据。
        index (int): 当前行的索引。
        excel_case_code (str): Excel 中的案件编码。
        excel_person_code (str): Excel 中的涉案人员编码。
        issues_list (list): 用于收集所有发现问题的列表。
        gender_mismatch_indices (set): 用于收集性别不匹配的行索引。
        excel_gender (str): Excel 中提取的性别。
        report_text_raw (str): 立案报告的原始文本。
        decision_text_raw (str): 处分决定的原始文本。
        investigation_text_raw (str): 审查调查报告的原始文本。
        trial_text_raw (str): 审理报告的原始文本。
        app_config (dict): Flask 应用的配置字典。
    """
    
    extracted_gender_from_report = extract_gender_from_case_report(report_text_raw)
    if extracted_gender_from_report is None or (excel_gender and excel_gender != extracted_gender_from_report):
        gender_mismatch_indices.add(index)
        issues_list.append((index, excel_case_code, excel_person_code, app_config['VALIDATION_RULES']["inconsistent_name"].replace("E2被反映人", app_config['COLUMN_MAPPINGS']['gender'] + "与BF2立案报告不一致"), "中")) # 增加风险等级
        logger.info(f"行 {index + 1} - 性别不匹配: Excel性别 ('{excel_gender}') vs 立案报告提取性别 ('{extracted_gender_from_report}')")

    extracted_gender_from_decision = extract_gender_from_decision_report(decision_text_raw)
    if extracted_gender_from_decision is None or (excel_gender and excel_gender != extracted_gender_from_decision):
        gender_mismatch_indices.add(index) 
        issues_list.append((index, excel_case_code, excel_person_code, app_config['VALIDATION_RULES']["inconsistent_name"].replace("E2被反映人", app_config['COLUMN_MAPPINGS']['gender'] + "与CU2处分决定不一致"), "中")) # 增加风险等级
        logger.info(f"行 {index + 1} - 性别不匹配: Excel性别 ('{excel_gender}') vs 处分决定提取性别 ('{extracted_gender_from_decision}')")

    extracted_gender_from_investigation = extract_gender_from_investigation_report(investigation_text_raw)
    if extracted_gender_from_investigation is None or (excel_gender and excel_gender != extracted_gender_from_investigation):
        gender_mismatch_indices.add(index)
        issues_list.append((index, excel_case_code, excel_person_code, app_config['VALIDATION_RULES']["inconsistent_name"].replace("E2被反映人", app_config['COLUMN_MAPPINGS']['gender'] + "与CX2审查调查报告不一致"), "中")) # 增加风险等级
        logger.info(f"行 {index + 1} - 性别不匹配: Excel性别 ('{excel_gender}') vs 审查调查报告提取性别 ('{extracted_gender_from_investigation}')")

    extracted_gender_from_trial = extract_gender_from_trial_report(trial_text_raw)
    if extracted_gender_from_trial is None or (excel_gender and excel_gender != extracted_gender_from_trial):
        gender_mismatch_indices.add(index)
        issues_list.append((index, excel_case_code, excel_person_code, app_config['VALIDATION_RULES']["inconsistent_name"].replace("E2被反映人", app_config['COLUMN_MAPPINGS']['gender'] + "与CY2审理报告不一致"), "中")) # 增加风险等级
        logger.info(f"行 {index + 1} - 性别不匹配: Excel性别 ('{excel_gender}') vs 审理报告提取性别 ('{extracted_gender_from_trial}')")

def validate_age_rules(row, index, excel_case_code, excel_person_code, issues_list, age_mismatch_indices,
                       excel_age, current_year, report_text_raw, decision_text_raw, investigation_text_raw, trial_text_raw, app_config):
    """
    验证年龄相关规则。
    比较 Excel 中的年龄与立案报告、处分决定、审查调查报告、审理报告中计算的年龄。

    参数:
        row (pd.Series): DataFrame 的当前行数据。
        index (int): 当前行的索引。
        excel_case_code (str): Excel 中的案件编码。
        excel_person_code (str): Excel 中的涉案人员编码。
        issues_list (list): 用于收集所有发现问题的列表。
        age_mismatch_indices (set): 用于收集年龄不匹配的行索引。
        excel_age (int or None): Excel 中提取的年龄。
        current_year (int): 当前年份。
        report_text_raw (str): 立案报告的原始文本。
        decision_text_raw (str): 处分决定的原始文本。
        investigation_text_raw (str): 审查调查报告的原始文本。
        trial_text_raw (str): 审理报告的原始文本。
        app_config (dict): Flask 应用的配置字典。
    """
    
    extracted_birth_year_from_report = extract_birth_year_from_case_report(report_text_raw)
    calculated_age_from_report = None
    if extracted_birth_year_from_report is not None:
        calculated_age_from_report = current_year - extracted_birth_year_from_report
    if (calculated_age_from_report is None) or \
       (excel_age is not None and calculated_age_from_report is not None and excel_age != calculated_age_from_report):
        age_mismatch_indices.add(index)
        issues_list.append((index, excel_case_code, excel_person_code, app_config['VALIDATION_RULES']["highlight_birth_date"].replace("X2出生年月", app_config['COLUMN_MAPPINGS']['age'] + "与BF2立案报告不一致"), "中")) # 增加风险等级
        logger.info(f"行 {index + 1} - 年龄不匹配: Excel年龄 ('{excel_age}') vs 立案报告计算年龄 ('{calculated_age_from_report}')")

    extracted_birth_year_from_decision = extract_birth_year_from_decision_report(decision_text_raw)
    calculated_age_from_decision = None
    if extracted_birth_year_from_decision is not None:
        calculated_age_from_decision = current_year - extracted_birth_year_from_decision
    if (calculated_age_from_decision is None) or \
       (excel_age is not None and calculated_age_from_decision is not None and excel_age != calculated_age_from_decision):
        age_mismatch_indices.add(index)
        issues_list.append((index, excel_case_code, excel_person_code, app_config['VALIDATION_RULES']["highlight_birth_date"].replace("X2出生年月", app_config['COLUMN_MAPPINGS']['age'] + "与CU2处分决定不一致"), "中")) # 增加风险等级
        logger.info(f"行 {index + 1} - 年龄不匹配: Excel年龄 ('{excel_age}') vs 处分决定计算年龄 ('{calculated_age_from_decision}')")

    extracted_birth_year_from_investigation = extract_birth_year_from_investigation_report(investigation_text_raw)
    calculated_age_from_investigation = None
    if extracted_birth_year_from_investigation is not None:
        calculated_age_from_investigation = current_year - extracted_birth_year_from_investigation
    if (calculated_age_from_investigation is None) or \
       (excel_age is not None and calculated_age_from_investigation is not None and excel_age != calculated_age_from_investigation):
        age_mismatch_indices.add(index)
        issues_list.append((index, excel_case_code, excel_person_code, app_config['VALIDATION_RULES']["highlight_birth_date"].replace("X2出生年月", app_config['COLUMN_MAPPINGS']['age'] + "与CX2审查调查报告不一致"), "中")) # 增加风险等级
        logger.info(f"行 {index + 1} - 年龄不匹配: Excel年龄 ('{excel_age}') vs 审查调查报告计算年龄 ('{calculated_age_from_investigation}')")

    extracted_birth_year_from_trial = extract_birth_year_from_trial_report(trial_text_raw)
    calculated_age_from_trial = None
    if extracted_birth_year_from_trial is not None:
        calculated_age_from_trial = current_year - extracted_birth_year_from_trial
    if (calculated_age_from_trial is None) or \
       (excel_age is not None and calculated_age_from_trial is not None and excel_age != calculated_age_from_trial):
        age_mismatch_indices.add(index)
        issues_list.append((index, excel_case_code, excel_person_code, app_config['VALIDATION_RULES']["highlight_birth_date"].replace("X2出生年月", app_config['COLUMN_MAPPINGS']['age'] + "与CY2审理报告不一致"), "中")) # 增加风险等级
        logger.info(f"行 {index + 1} - 年龄不匹配: Excel年龄 ('{excel_age}') vs 审理报告计算年龄 ('{calculated_age_from_trial}')")

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

