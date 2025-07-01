import logging
import pandas as pd
import re

# 导入必要的提取器函数
from validation_rules.case_extractors_names import (
    extract_name_from_case_report,
    extract_name_from_decision,
    extract_name_from_trial_report
)
from validation_rules.case_extractors_gender import (
    extract_gender_from_case_report,
    extract_gender_from_decision_report,
    extract_gender_from_investigation_report,
    extract_gender_from_trial_report
)
from validation_rules.case_extractors_birth_info import (
    extract_birth_year_from_case_report,
    extract_birth_year_from_decision_report,
    extract_birth_year_from_investigation_report,
    extract_birth_year_from_trial_report
)
from validation_rules.case_extractors_demographics import (
    extract_suspected_violation_from_case_report,
    extract_suspected_violation_from_decision
)

logger = logging.getLogger(__name__)

def validate_gender_rules(row, index, excel_case_code, excel_person_code, issues_list, gender_mismatch_indices,
                          excel_gender, report_text_raw, decision_text_raw, investigation_text_raw, trial_text_raw):
    """验证性别相关规则。"""
    
    extracted_gender_from_report = extract_gender_from_case_report(report_text_raw)
    if extracted_gender_from_report is None or (excel_gender and excel_gender != extracted_gender_from_report):
        gender_mismatch_indices.add(index)
        issues_list.append((index, excel_case_code, excel_person_code, "M2性别与BF2立案报告不一致"))
        logger.info(f"行 {index + 1} - 性别不匹配: Excel性别 ('{excel_gender}') vs 立案报告提取性别 ('{extracted_gender_from_report}')")
        print(f"行 {index + 1} - 性别不匹配: Excel性别 ('{excel_gender}') vs 立案报告提取性别 ('{extracted_gender_from_report}')")

    extracted_gender_from_decision = extract_gender_from_decision_report(decision_text_raw)
    if extracted_gender_from_decision is None or (excel_gender and excel_gender != extracted_gender_from_decision):
        gender_mismatch_indices.add(index) 
        issues_list.append((index, excel_case_code, excel_person_code, "M2性别与CU2处分决定不一致"))
        logger.info(f"行 {index + 1} - 性别不匹配: Excel性别 ('{excel_gender}') vs 处分决定提取性别 ('{extracted_gender_from_decision}')")
        print(f"行 {index + 1} - 性别不匹配: Excel性别 ('{excel_gender}') vs 处分决定提取性别 ('{extracted_gender_from_decision}')")

    extracted_gender_from_investigation = extract_gender_from_investigation_report(investigation_text_raw)
    if extracted_gender_from_investigation is None or (excel_gender and excel_gender != extracted_gender_from_investigation):
        gender_mismatch_indices.add(index)
        issues_list.append((index, excel_case_code, excel_person_code, "M2性别与CX2审查调查报告不一致"))
        logger.info(f"行 {index + 1} - 性别不匹配: Excel性别 ('{excel_gender}') vs 审查调查报告提取性别 ('{extracted_gender_from_investigation}')")
        print(f"行 {index + 1} - 性别不匹配: Excel性别 ('{excel_gender}') vs 审查调查报告提取性别 ('{extracted_gender_from_investigation}')")

    extracted_gender_from_trial = extract_gender_from_trial_report(trial_text_raw)
    if extracted_gender_from_trial is None or (excel_gender and excel_gender != extracted_gender_from_trial):
        gender_mismatch_indices.add(index)
        issues_list.append((index, excel_case_code, excel_person_code, "M2性别与CY2审理报告不一致"))
        logger.info(f"行 {index + 1} - 性别不匹配: Excel性别 ('{excel_gender}') vs 审理报告提取性别 ('{extracted_gender_from_trial}')")
        print(f"行 {index + 1} - 性别不匹配: Excel性别 ('{excel_gender}') vs 审理报告提取性别 ('{extracted_gender_from_trial}')")

def validate_age_rules(row, index, excel_case_code, excel_person_code, issues_list, age_mismatch_indices,
                       excel_age, current_year, report_text_raw, decision_text_raw, investigation_text_raw, trial_text_raw):
    """验证年龄相关规则。"""
    
    extracted_birth_year_from_report = extract_birth_year_from_case_report(report_text_raw)
    calculated_age_from_report = None
    if extracted_birth_year_from_report is not None:
        calculated_age_from_report = current_year - extracted_birth_year_from_report
    if (calculated_age_from_report is None) or \
       (excel_age is not None and calculated_age_from_report is not None and excel_age != calculated_age_from_report):
        age_mismatch_indices.add(index)
        issues_list.append((index, excel_case_code, excel_person_code, "N2年龄与BF2立案报告不一致"))
        logger.info(f"行 {index + 1} - 年龄不匹配: Excel年龄 ('{excel_age}') vs 立案报告计算年龄 ('{calculated_age_from_report}')")
        print(f"行 {index + 1} - 年龄不匹配: Excel年龄 ('{excel_age}') vs 立案报告计算年龄 ('{calculated_age_from_report}')")

    extracted_birth_year_from_decision = extract_birth_year_from_decision_report(decision_text_raw)
    calculated_age_from_decision = None
    if extracted_birth_year_from_decision is not None:
        calculated_age_from_decision = current_year - extracted_birth_year_from_decision
    if (calculated_age_from_decision is None) or \
       (excel_age is not None and calculated_age_from_decision is not None and excel_age != calculated_age_from_decision):
        age_mismatch_indices.add(index)
        issues_list.append((index, excel_case_code, excel_person_code, "N2年龄与CU2处分决定不一致"))
        logger.info(f"行 {index + 1} - 年龄不匹配: Excel年龄 ('{excel_age}') vs 处分决定计算年龄 ('{calculated_age_from_decision}')")
        print(f"行 {index + 1} - 年龄不匹配: Excel年龄 ('{excel_age}') vs 处分决定计算年龄 ('{calculated_age_from_decision}')")

    extracted_birth_year_from_investigation = extract_birth_year_from_investigation_report(investigation_text_raw)
    calculated_age_from_investigation = None
    if extracted_birth_year_from_investigation is not None:
        calculated_age_from_investigation = current_year - extracted_birth_year_from_investigation
    if (calculated_age_from_investigation is None) or \
       (excel_age is not None and calculated_age_from_investigation is not None and excel_age != calculated_age_from_investigation):
        age_mismatch_indices.add(index)
        issues_list.append((index, excel_case_code, excel_person_code, "N2年龄与CX2审查调查报告不一致"))
        logger.info(f"行 {index + 1} - 年龄不匹配: Excel年龄 ('{excel_age}') vs 审查调查报告计算年龄 ('{calculated_age_from_investigation}')")
        print(f"行 {index + 1} - 年龄不匹配: Excel年龄 ('{excel_age}') vs 审查调查报告计算年龄 ('{calculated_age_from_investigation}')")

    extracted_birth_year_from_trial = extract_birth_year_from_trial_report(trial_text_raw)
    calculated_age_from_trial = None
    if extracted_birth_year_from_trial is not None:
        calculated_age_from_trial = current_year - extracted_birth_year_from_trial
    if (calculated_age_from_trial is None) or \
       (excel_age is not None and calculated_age_from_trial is not None and excel_age != calculated_age_from_trial):
        age_mismatch_indices.add(index)
        issues_list.append((index, excel_case_code, excel_person_code, "N2年龄与CY2审理报告不一致"))
        logger.info(f"行 {index + 1} - 年龄不匹配: Excel年龄 ('{excel_age}') vs 审理报告计算年龄 ('{calculated_age_from_trial}')")
        print(f"行 {index + 1} - 年龄不匹配: Excel年龄 ('{excel_age}') vs 审理报告计算年龄 ('{calculated_age_from_trial}')")

def validate_brief_case_details_rules(row, index, excel_case_code, excel_person_code, issues_list, brief_case_details_mismatch_indices,
                                      excel_brief_case_details, investigated_person, report_text_raw, decision_text_raw):
    """验证简要案情相关规则。"""
    is_brief_case_details_mismatch = False
    extracted_brief_case_details = None
    # 新增变量用于记录简要案情提取的来源文档，以便动态生成问题描述
    source_document_for_issue = "" 

    if pd.isna(row["处分决定"]) or decision_text_raw == '':
        extracted_brief_case_details = extract_suspected_violation_from_case_report(report_text_raw)
        # 当“处分决定”为空时，来源文档是“立案报告”
        source_document_for_issue = "BF立案报告" 
        logger.info(f"行 {index + 1} - 处分决定为空，从立案报告提取简要案情：'{extracted_brief_case_details}'")
        print(f"行 {index + 1} - 处分决定为空，从立案报告提取简要案情：'{extracted_brief_case_details}'")
    else:
        extracted_brief_case_details = extract_suspected_violation_from_decision(decision_text_raw, investigated_person)
        # 当“处分决定”不为空时，来源文档是“处分决定”
        source_document_for_issue = "CU处分决定" 
        logger.info(f"行 {index + 1} - 处分决定不为空，从处分决定提取简要案情：'{extracted_brief_case_details}'")
        print(f"行 {index + 1} - 处分决定不为空，从处分决定提取简要案情：'{extracted_brief_case_details}'")

    if extracted_brief_case_details is None:
        if excel_brief_case_details:
            is_brief_case_details_mismatch = True
            # 根据 source_document_for_issue 动态生成问题描述
            issues_list.append((index, excel_case_code, excel_person_code, f"BE简要案情与{source_document_for_issue}不一致（未能提取到内容）"))
            logger.info(f"行 {index + 1} - 简要案情不匹配: Excel有值 ('{excel_brief_case_details}') 但未能从{source_document_for_issue}中提取。")
            print(f"行 {index + 1} - 简要案情不匹配: Excel有值 ('{excel_brief_case_details}') 但未能从{source_document_for_issue}中提取。")
    else:
        cleaned_excel_brief_case_details = re.sub(r'\s+', '', excel_brief_case_details)
        
        if cleaned_excel_brief_case_details != extracted_brief_case_details:
            is_brief_case_details_mismatch = True
            # 根据 source_document_for_issue 动态生成问题描述
            issues_list.append((index, excel_case_code, excel_person_code, f"BE简要案情与{source_document_for_issue}不一致"))
            logger.info(f"行 {index + 1} - 简要案情不匹配: Excel简要案情 ('{cleaned_excel_brief_case_details}') vs 提取简要案情 ('{extracted_brief_case_details}') (来源: {source_document_for_issue})")
            print(f"行 {index + 1} - 简要案情不匹配: Excel简要案情 ('{cleaned_excel_brief_case_details}') vs 提取简要案情 ('{extracted_brief_case_details}') (来源: {source_document_for_issue})")

    if is_brief_case_details_mismatch:
        brief_case_details_mismatch_indices.add(index)