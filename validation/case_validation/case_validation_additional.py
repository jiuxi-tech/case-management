import logging
import pandas as pd
import re
from .case_extractors_names import (
    extract_name_from_case_report,
    extract_name_from_decision,
    extract_name_from_trial_report
)

logger = logging.getLogger(__name__)

def validate_name_rules(row, index, excel_case_code, excel_person_code, issues_list, mismatch_indices,
                        investigated_person, report_text_raw, decision_text_raw, investigation_text_raw, trial_text_raw, app_config):
    """验证姓名相关规则。
    统一日志风格和编号表字段结构，与线索表保持一致。
    """
    
    # 规则1: 被调查人与立案报告比对
    report_name = extract_name_from_case_report(report_text_raw)
    if report_name and investigated_person != report_name:
        mismatch_indices.add(index)
        issues_list.append({
            '案件编码': excel_case_code,
            '涉案人员编码': excel_person_code,
            '行号': index + 2,
            '比对字段': "C被调查人",
            '被比对字段': "BF立案报告",
            '问题描述': f"C{index + 2}被调查人与BF{index + 2}立案报告不一致",
            '列名': "被调查人"
        })
        logger.warning(f"<立案 - （1.被调查人与立案报告）> - 行 {index + 2} - 被调查人 '{investigated_person}' 与立案报告姓名 '{report_name}' 不一致")

    # 规则2: 被调查人与处分决定比对
    decision_name = extract_name_from_decision(decision_text_raw)
    if not decision_name or (decision_name and investigated_person != decision_name):
        mismatch_indices.add(index)
        issues_list.append({
            '案件编码': excel_case_code,
            '涉案人员编码': excel_person_code,
            '行号': index + 2,
            '比对字段': "C被调查人",
            '被比对字段': "CU处分决定",
            '问题描述': f"C{index + 2}被调查人与CU{index + 2}处分决定不一致",
            '列名': "被调查人"
        })
        logger.warning(f"<立案 - （2.被调查人与处分决定）> - 行 {index + 2} - 被调查人 '{investigated_person}' 与处分决定姓名 '{decision_name}' 不一致")

    # 规则3: 被调查人与审查调查报告比对
    investigation_name = extract_name_from_case_report(investigation_text_raw)
    if investigation_name and investigated_person != investigation_name:
        mismatch_indices.add(index)
        issues_list.append({
            '案件编码': excel_case_code,
            '涉案人员编码': excel_person_code,
            '行号': index + 2,
            '比对字段': "C被调查人",
            '被比对字段': "CX审查调查报告",
            '问题描述': f"C{index + 2}被调查人与CX{index + 2}审查调查报告不一致",
            '列名': "被调查人"
        })
        logger.warning(f"<立案 - （3.被调查人与审查调查报告）> - 行 {index + 2} - 被调查人 '{investigated_person}' 与审查调查报告姓名 '{investigation_name}' 不一致")

    # 规则4: 被调查人与审理报告比对
    trial_name = extract_name_from_trial_report(trial_text_raw)
    if not trial_name or (trial_name and investigated_person != trial_name):
        mismatch_indices.add(index)
        issues_list.append({
            '案件编码': excel_case_code,
            '涉案人员编码': excel_person_code,
            '行号': index + 2,
            '比对字段': "C被调查人",
            '被比对字段': "CY审理报告",
            '问题描述': f"C{index + 2}被调查人与CY{index + 2}审理报告不一致",
            '列名': "被调查人"
        })
        logger.warning(f"<立案 - （4.被调查人与审理报告）> - 行 {index + 2} - 被调查人 '{investigated_person}' 与审理报告姓名 '{trial_name}' 不一致")

def validate_gender_rules(row, index, excel_case_code, excel_person_code, issues_list, gender_mismatch_indices,
                         excel_gender, report_text_raw, decision_text_raw, investigation_text_raw, trial_text_raw, app_config):
    """验证性别相关规则。
    统一日志风格和编号表字段结构，与被调查人规则保持一致。
    """
    
    # 规则1: 性别与立案报告比对
    from .case_extractors_gender import extract_gender_from_case_report
    extracted_gender_from_report = extract_gender_from_case_report(report_text_raw)
    if extracted_gender_from_report is None or (excel_gender and excel_gender != extracted_gender_from_report):
        gender_mismatch_indices.add(index)
        issues_list.append({
            '案件编码': excel_case_code,
            '涉案人员编码': excel_person_code,
            '行号': index + 2,
            '比对字段': f"M{app_config['COLUMN_MAPPINGS']['gender']}",
            '被比对字段': f"BF{app_config['COLUMN_MAPPINGS']['case_report']}",
            '问题描述': f"M{index + 2}{app_config['COLUMN_MAPPINGS']['gender']}与BF{index + 2}立案报告不一致",
            '列名': app_config['COLUMN_MAPPINGS']['gender']
        })
        logger.warning(f"<立案 - （1.性别与立案报告）> - 行 {index + 2} - 性别 '{excel_gender}' 与立案报告性别 '{extracted_gender_from_report}' 不一致")

    # 规则2: 性别与处分决定比对
    from .case_extractors_gender import extract_gender_from_decision_report
    extracted_gender_from_decision = extract_gender_from_decision_report(decision_text_raw)
    if extracted_gender_from_decision is None or (excel_gender and excel_gender != extracted_gender_from_decision):
        gender_mismatch_indices.add(index)
        issues_list.append({
            '案件编码': excel_case_code,
            '涉案人员编码': excel_person_code,
            '行号': index + 2,
            '比对字段': f"M{app_config['COLUMN_MAPPINGS']['gender']}",
            '被比对字段': f"CU{app_config['COLUMN_MAPPINGS']['disciplinary_decision']}",
            '问题描述': f"M{index + 2}{app_config['COLUMN_MAPPINGS']['gender']}与CU{index + 2}处分决定不一致",
            '列名': app_config['COLUMN_MAPPINGS']['gender']
        })
        logger.warning(f"<立案 - （2.性别与处分决定）> - 行 {index + 2} - 性别 '{excel_gender}' 与处分决定性别 '{extracted_gender_from_decision}' 不一致")

    # 规则3: 性别与审查调查报告比对
    from .case_extractors_gender import extract_gender_from_investigation_report
    extracted_gender_from_investigation = extract_gender_from_investigation_report(investigation_text_raw)
    if extracted_gender_from_investigation is None or (excel_gender and excel_gender != extracted_gender_from_investigation):
        gender_mismatch_indices.add(index)
        issues_list.append({
            '案件编码': excel_case_code,
            '涉案人员编码': excel_person_code,
            '行号': index + 2,
            '比对字段': f"M{app_config['COLUMN_MAPPINGS']['gender']}",
            '被比对字段': f"CX{app_config['COLUMN_MAPPINGS']['investigation_report']}",
            '问题描述': f"M{index + 2}{app_config['COLUMN_MAPPINGS']['gender']}与CX{index + 2}审查调查报告不一致",
            '列名': app_config['COLUMN_MAPPINGS']['gender']
        })
        logger.warning(f"<立案 - （3.性别与审查调查报告）> - 行 {index + 2} - 性别 '{excel_gender}' 与审查调查报告性别 '{extracted_gender_from_investigation}' 不一致")

    # 规则4: 性别与审理报告比对
    from .case_extractors_gender import extract_gender_from_trial_report
    extracted_gender_from_trial = extract_gender_from_trial_report(trial_text_raw)
    if extracted_gender_from_trial is None or (excel_gender and excel_gender != extracted_gender_from_trial):
        gender_mismatch_indices.add(index)
        issues_list.append({
            '案件编码': excel_case_code,
            '涉案人员编码': excel_person_code,
            '行号': index + 2,
            '比对字段': f"M{app_config['COLUMN_MAPPINGS']['gender']}",
            '被比对字段': f"CY{app_config['COLUMN_MAPPINGS']['trial_report']}",
            '问题描述': f"M{index + 2}{app_config['COLUMN_MAPPINGS']['gender']}与CY{index + 2}审理报告不一致",
            '列名': app_config['COLUMN_MAPPINGS']['gender']
        })
        logger.warning(f"<立案 - （4.性别与审理报告）> - 行 {index + 2} - 性别 '{excel_gender}' 与审理报告性别 '{extracted_gender_from_trial}' 不一致")

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
    
    # 规则1: 年龄与立案报告比对
    from .case_extractors_birth_info import extract_birth_year_from_case_report
    extracted_birth_year_from_report = extract_birth_year_from_case_report(report_text_raw)
    calculated_age_from_report = None
    if extracted_birth_year_from_report is not None:
        calculated_age_from_report = current_year - extracted_birth_year_from_report
    if (calculated_age_from_report is None) or \
       (excel_age is not None and calculated_age_from_report is not None and excel_age != calculated_age_from_report):
        age_mismatch_indices.add(index)
        issues_list.append({
            '案件编码': excel_case_code,
            '涉案人员编码': excel_person_code,
            '行号': index + 2,
            '比对字段': f"N{app_config['COLUMN_MAPPINGS']['age']}",
            '被比对字段': f"BF{app_config['COLUMN_MAPPINGS']['case_report']}",
            '问题描述': f"N{index + 2}{app_config['COLUMN_MAPPINGS']['age']}与BF{index + 2}立案报告不一致",
            '列名': app_config['COLUMN_MAPPINGS']['age']
        })
        logger.warning(f"<立案 - （1.年龄与立案报告）> - 行 {index + 2} - 年龄 '{excel_age}' 与立案报告计算年龄 '{calculated_age_from_report}' 不一致")

    # 规则2: 年龄与处分决定比对
    from .case_extractors_birth_info import extract_birth_year_from_decision_report
    extracted_birth_year_from_decision = extract_birth_year_from_decision_report(decision_text_raw)
    calculated_age_from_decision = None
    if extracted_birth_year_from_decision is not None:
        calculated_age_from_decision = current_year - extracted_birth_year_from_decision
    if (calculated_age_from_decision is None) or \
       (excel_age is not None and calculated_age_from_decision is not None and excel_age != calculated_age_from_decision):
        age_mismatch_indices.add(index)
        issues_list.append({
            '案件编码': excel_case_code,
            '涉案人员编码': excel_person_code,
            '行号': index + 2,
            '比对字段': f"N{app_config['COLUMN_MAPPINGS']['age']}",
            '被比对字段': f"CU{app_config['COLUMN_MAPPINGS']['disciplinary_decision']}",
            '问题描述': f"N{index + 2}{app_config['COLUMN_MAPPINGS']['age']}与CU{index + 2}处分决定不一致",
            '列名': app_config['COLUMN_MAPPINGS']['age']
        })
        logger.warning(f"<立案 - （2.年龄与处分决定）> - 行 {index + 2} - 年龄 '{excel_age}' 与处分决定计算年龄 '{calculated_age_from_decision}' 不一致")

    # 规则3: 年龄与审查调查报告比对
    from .case_extractors_birth_info import extract_birth_year_from_investigation_report
    extracted_birth_year_from_investigation = extract_birth_year_from_investigation_report(investigation_text_raw)
    calculated_age_from_investigation = None
    if extracted_birth_year_from_investigation is not None:
        calculated_age_from_investigation = current_year - extracted_birth_year_from_investigation
    if (calculated_age_from_investigation is None) or \
       (excel_age is not None and calculated_age_from_investigation is not None and excel_age != calculated_age_from_investigation):
        age_mismatch_indices.add(index)
        issues_list.append({
            '案件编码': excel_case_code,
            '涉案人员编码': excel_person_code,
            '行号': index + 2,
            '比对字段': f"N{app_config['COLUMN_MAPPINGS']['age']}",
            '被比对字段': f"CX{app_config['COLUMN_MAPPINGS']['investigation_report']}",
            '问题描述': f"N{index + 2}{app_config['COLUMN_MAPPINGS']['age']}与CX{index + 2}审查调查报告不一致",
            '列名': app_config['COLUMN_MAPPINGS']['age']
        })
        logger.warning(f"<立案 - （3.年龄与审查调查报告）> - 行 {index + 2} - 年龄 '{excel_age}' 与审查调查报告计算年龄 '{calculated_age_from_investigation}' 不一致")

    # 规则4: 年龄与审理报告比对
    from .case_extractors_birth_info import extract_birth_year_from_trial_report
    extracted_birth_year_from_trial = extract_birth_year_from_trial_report(trial_text_raw)
    calculated_age_from_trial = None
    if extracted_birth_year_from_trial is not None:
        calculated_age_from_trial = current_year - extracted_birth_year_from_trial
    if (calculated_age_from_trial is None) or \
       (excel_age is not None and calculated_age_from_trial is not None and excel_age != calculated_age_from_trial):
        age_mismatch_indices.add(index)
        issues_list.append({
            '案件编码': excel_case_code,
            '涉案人员编码': excel_person_code,
            '行号': index + 2,
            '比对字段': f"N{app_config['COLUMN_MAPPINGS']['age']}",
            '被比对字段': f"CY{app_config['COLUMN_MAPPINGS']['trial_report']}",
            '问题描述': f"N{index + 2}{app_config['COLUMN_MAPPINGS']['age']}与CY{index + 2}审理报告不一致",
            '列名': app_config['COLUMN_MAPPINGS']['age']
        })
        logger.warning(f"<立案 - （4.年龄与审理报告）> - 行 {index + 2} - 年龄 '{excel_age}' 与审理报告计算年龄 '{calculated_age_from_trial}' 不一致")

def validate_birth_date_rules(row, index, excel_case_code, excel_person_code, issues_list, birth_date_mismatch_indices,
                              excel_birth_date, report_text_raw, decision_text_raw, investigation_text_raw, trial_text_raw, app_config):
    """
    验证出生年月相关规则。
    比较 Excel 中的出生年月与立案报告、处分决定、审查调查报告、审理报告中提取的出生年月。

    参数:
        row (pd.Series): DataFrame 的当前行数据。
        index (int): 当前行的索引。
        excel_case_code (str): Excel 中的案件编码。
        excel_person_code (str): Excel 中的涉案人员编码。
        issues_list (list): 用于收集所有发现问题的列表。
        birth_date_mismatch_indices (set): 用于收集出生年月不匹配的行索引。
        excel_birth_date (str): Excel 中提取的出生年月。
        report_text_raw (str): 立案报告的原始文本。
        decision_text_raw (str): 处分决定的原始文本。
        investigation_text_raw (str): 审查调查报告的原始文本。
        trial_text_raw (str): 审理报告的原始文本。
        app_config (dict): Flask 应用的配置字典。
    """
    
    # 规则1: 出生年月与立案报告比对
    from .case_extractors_birth_info import extract_birth_date_from_case_report
    extracted_birth_date_from_report = extract_birth_date_from_case_report(report_text_raw)
    if (excel_birth_date and excel_birth_date.strip() != '' and 
        (extracted_birth_date_from_report is None or excel_birth_date != extracted_birth_date_from_report)):
        birth_date_mismatch_indices.add(index)
        issues_list.append({
            '案件编码': excel_case_code,
            '涉案人员编码': excel_person_code,
            '行号': index + 2,
            '比对字段': f"O{app_config['COLUMN_MAPPINGS']['birth_date']}",
            '被比对字段': f"BF{app_config['COLUMN_MAPPINGS']['case_report']}",
            '问题描述': f"O{index + 2}{app_config['COLUMN_MAPPINGS']['birth_date']}与BF{index + 2}立案报告不一致",
            '列名': app_config['COLUMN_MAPPINGS']['birth_date']
        })
        logger.warning(f"<立案 - （1.出生年月与立案报告）> - 行 {index + 2} - 出生年月 '{excel_birth_date}' 与立案报告提取出生年月 '{extracted_birth_date_from_report}' 不一致")

    # 规则2: 出生年月与处分决定比对
    from .case_extractors_birth_info import extract_birth_date_from_decision_report
    extracted_birth_date_from_decision = extract_birth_date_from_decision_report(decision_text_raw)
    if (excel_birth_date and excel_birth_date.strip() != '' and 
        (extracted_birth_date_from_decision is None or excel_birth_date != extracted_birth_date_from_decision)):
         birth_date_mismatch_indices.add(index)
         issues_list.append({
             '案件编码': excel_case_code,
             '涉案人员编码': excel_person_code,
             '行号': index + 2,
             '比对字段': f"O{app_config['COLUMN_MAPPINGS']['birth_date']}",
             '被比对字段': f"CU{app_config['COLUMN_MAPPINGS']['disciplinary_decision']}",
             '问题描述': f"O{index + 2}{app_config['COLUMN_MAPPINGS']['birth_date']}与CU{index + 2}处分决定不一致",
             '列名': app_config['COLUMN_MAPPINGS']['birth_date']
         })
         logger.warning(f"<立案 - （2.出生年月与处分决定）> - 行 {index + 2} - 出生年月 '{excel_birth_date}' 与处分决定提取出生年月 '{extracted_birth_date_from_decision}' 不一致")

    # 规则3: 出生年月与审查调查报告比对
    from .case_extractors_birth_info import extract_birth_date_from_investigation_report
    extracted_birth_date_from_investigation = extract_birth_date_from_investigation_report(investigation_text_raw)
    if (excel_birth_date and excel_birth_date.strip() != '' and 
        (extracted_birth_date_from_investigation is None or excel_birth_date != extracted_birth_date_from_investigation)):
         birth_date_mismatch_indices.add(index)
         issues_list.append({
             '案件编码': excel_case_code,
             '涉案人员编码': excel_person_code,
             '行号': index + 2,
             '比对字段': f"O{app_config['COLUMN_MAPPINGS']['birth_date']}",
             '被比对字段': f"CX{app_config['COLUMN_MAPPINGS']['investigation_report']}",
             '问题描述': f"O{index + 2}{app_config['COLUMN_MAPPINGS']['birth_date']}与CX{index + 2}审查调查报告不一致",
             '列名': app_config['COLUMN_MAPPINGS']['birth_date']
         })
         logger.warning(f"<立案 - （3.出生年月与审查调查报告）> - 行 {index + 2} - 出生年月 '{excel_birth_date}' 与审查调查报告提取出生年月 '{extracted_birth_date_from_investigation}' 不一致")

    # 规则4: 出生年月与审理报告比对
    from .case_extractors_birth_info import extract_birth_date_from_trial_report
    extracted_birth_date_from_trial = extract_birth_date_from_trial_report(trial_text_raw)
    if (excel_birth_date and excel_birth_date.strip() != '' and 
        (extracted_birth_date_from_trial is None or excel_birth_date != extracted_birth_date_from_trial)):
         birth_date_mismatch_indices.add(index)
         issues_list.append({
             '案件编码': excel_case_code,
             '涉案人员编码': excel_person_code,
             '行号': index + 2,
             '比对字段': f"O{app_config['COLUMN_MAPPINGS']['birth_date']}",
             '被比对字段': f"CY{app_config['COLUMN_MAPPINGS']['trial_report']}",
             '问题描述': f"O{index + 2}{app_config['COLUMN_MAPPINGS']['birth_date']}与CY{index + 2}审理报告不一致",
             '列名': app_config['COLUMN_MAPPINGS']['birth_date']
         })
         logger.warning(f"<立案 - （4.出生年月与审理报告）> - 行 {index + 2} - 出生年月 '{excel_birth_date}' 与审理报告提取出生年月 '{extracted_birth_date_from_trial}' 不一致")

def validate_education_rules(row, index, excel_case_code, excel_person_code, issues_list, education_mismatch_indices,
                            excel_education, report_text_raw, decision_text_raw, investigation_text_raw, trial_text_raw, app_config):
    """
    验证学历相关规则。
    
    参数:
    row: 当前行数据
    index: 行索引
    excel_case_code: Excel中的案件编码
    excel_person_code: Excel中的涉案人员编码
    issues_list: 问题列表
    education_mismatch_indices (set): 用于收集学历不匹配的行索引。
    excel_education: Excel中的学历
    report_text_raw: 立案报告原始文本
    decision_text_raw: 处分决定原始文本
    investigation_text_raw: 审查调查报告原始文本
    trial_text_raw: 审理报告原始文本
    app_config: 应用配置
    """
    
    # 规则1: 学历与立案报告比对
    from .case_extractors_demographics import extract_education_from_case_report
    extracted_education_from_report = extract_education_from_case_report(report_text_raw)
    
    # 标准化学历名称（处理"大学本科"与"本科"的匹配）
    excel_education_normalized = excel_education
    if excel_education == "大学本科":
        excel_education_normalized = "本科"
    extracted_education_normalized = extracted_education_from_report
    if extracted_education_from_report == "大学本科":
        extracted_education_normalized = "本科"
    
    if (excel_education and excel_education.strip() != '' and 
        (extracted_education_from_report is None or excel_education_normalized != extracted_education_normalized)):
        education_mismatch_indices.add(index)
        issues_list.append({
            '案件编码': excel_case_code,
            '涉案人员编码': excel_person_code,
            '行号': index + 2,
            '比对字段': f"P{app_config['COLUMN_MAPPINGS']['education']}",
            '被比对字段': f"BF{app_config['COLUMN_MAPPINGS']['case_report']}",
            '问题描述': f"P{index + 2}{app_config['COLUMN_MAPPINGS']['education']}与BF{index + 2}立案报告不一致",
            '列名': app_config['COLUMN_MAPPINGS']['education']
        })
        logger.warning(f"<立案 - （1.学历与立案报告）> - 行 {index + 2} - 学历 '{excel_education}' 与立案报告提取学历 '{extracted_education_from_report}' 不一致")

    # 规则2: 学历与处分决定比对
    # 注意：这里假设有对应的提取函数，如果没有可以使用相同的提取逻辑或创建新函数
    extracted_education_from_decision = extract_education_from_case_report(decision_text_raw)  # 可以复用或创建专门函数
    extracted_education_decision_normalized = extracted_education_from_decision
    if extracted_education_from_decision == "大学本科":
        extracted_education_decision_normalized = "本科"
    
    if (excel_education and excel_education.strip() != '' and 
        (extracted_education_from_decision is None or excel_education_normalized != extracted_education_decision_normalized)):
        education_mismatch_indices.add(index)
        issues_list.append({
            '案件编码': excel_case_code,
            '涉案人员编码': excel_person_code,
            '行号': index + 2,
            '比对字段': f"P{app_config['COLUMN_MAPPINGS']['education']}",
            '被比对字段': f"CU{app_config['COLUMN_MAPPINGS']['disciplinary_decision']}",
            '问题描述': f"P{index + 2}{app_config['COLUMN_MAPPINGS']['education']}与CU{index + 2}处分决定不一致",
            '列名': app_config['COLUMN_MAPPINGS']['education']
        })
        logger.warning(f"<立案 - （2.学历与处分决定）> - 行 {index + 2} - 学历 '{excel_education}' 与处分决定提取学历 '{extracted_education_from_decision}' 不一致")

    # 规则3: 学历与审查调查报告比对
    extracted_education_from_investigation = extract_education_from_case_report(investigation_text_raw)  # 可以复用或创建专门函数
    extracted_education_investigation_normalized = extracted_education_from_investigation
    if extracted_education_from_investigation == "大学本科":
        extracted_education_investigation_normalized = "本科"
    
    if (excel_education and excel_education.strip() != '' and 
        (extracted_education_from_investigation is None or excel_education_normalized != extracted_education_investigation_normalized)):
        education_mismatch_indices.add(index)
        issues_list.append({
            '案件编码': excel_case_code,
            '涉案人员编码': excel_person_code,
            '行号': index + 2,
            '比对字段': f"P{app_config['COLUMN_MAPPINGS']['education']}",
            '被比对字段': f"CX{app_config['COLUMN_MAPPINGS']['investigation_report']}",
            '问题描述': f"P{index + 2}{app_config['COLUMN_MAPPINGS']['education']}与CX{index + 2}审查调查报告不一致",
            '列名': app_config['COLUMN_MAPPINGS']['education']
        })
        logger.warning(f"<立案 - （3.学历与审查调查报告）> - 行 {index + 2} - 学历 '{excel_education}' 与审查调查报告提取学历 '{extracted_education_from_investigation}' 不一致")

    # 规则4: 学历与审理报告比对
    extracted_education_from_trial = extract_education_from_case_report(trial_text_raw)  # 可以复用或创建专门函数
    extracted_education_trial_normalized = extracted_education_from_trial
    if extracted_education_from_trial == "大学本科":
        extracted_education_trial_normalized = "本科"
    
    if (excel_education and excel_education.strip() != '' and 
        (extracted_education_from_trial is None or excel_education_normalized != extracted_education_trial_normalized)):
        education_mismatch_indices.add(index)
        issues_list.append({
            '案件编码': excel_case_code,
            '涉案人员编码': excel_person_code,
            '行号': index + 2,
            '比对字段': f"P{app_config['COLUMN_MAPPINGS']['education']}",
            '被比对字段': f"CY{app_config['COLUMN_MAPPINGS']['trial_report']}",
            '问题描述': f"P{index + 2}{app_config['COLUMN_MAPPINGS']['education']}与CY{index + 2}审理报告不一致",
            '列名': app_config['COLUMN_MAPPINGS']['education']
        })
        logger.warning(f"<立案 - （4.学历与审理报告）> - 行 {index + 2} - 学历 '{excel_education}' 与审理报告提取学历 '{extracted_education_from_trial}' 不一致")

def validate_ethnicity_rules(row, index, excel_case_code, excel_person_code, issues_list, ethnicity_mismatch_indices,
                             excel_ethnicity, report_text_raw, decision_text_raw, investigation_text_raw, trial_text_raw, app_config):
    """
    验证民族相关规则。
    
    参数:
    row: 当前行数据
    index: 行索引
    excel_case_code: Excel中的案件编码
    excel_person_code: Excel中的涉案人员编码
    issues_list: 问题列表
    ethnicity_mismatch_indices (set): 用于收集民族不匹配的行索引。
    excel_ethnicity: Excel中的民族
    report_text_raw: 立案报告原始文本
    decision_text_raw: 处分决定原始文本
    investigation_text_raw: 审查调查报告原始文本
    trial_text_raw: 审理报告原始文本
    app_config: 应用配置
    """
    
    # 规则1: 民族与立案报告比对
    from .case_extractors_demographics import extract_ethnicity_from_case_report
    extracted_ethnicity_from_report = extract_ethnicity_from_case_report(report_text_raw)
    
    if (excel_ethnicity and excel_ethnicity.strip() != '' and 
        (extracted_ethnicity_from_report is None or excel_ethnicity != extracted_ethnicity_from_report)):
        ethnicity_mismatch_indices.add(index)
        issues_list.append({
            '案件编码': excel_case_code,
            '涉案人员编码': excel_person_code,
            '行号': index + 2,
            '比对字段': f"Q{app_config['COLUMN_MAPPINGS']['ethnicity']}",
            '被比对字段': f"BF{app_config['COLUMN_MAPPINGS']['case_report']}",
            '问题描述': f"Q{index + 2}{app_config['COLUMN_MAPPINGS']['ethnicity']}与BF{index + 2}立案报告不一致",
            '列名': app_config['COLUMN_MAPPINGS']['ethnicity']
        })
        logger.warning(f"<立案 - （1.民族与立案报告）> - 行 {index + 2} - 民族 '{excel_ethnicity}' 与立案报告提取民族 '{extracted_ethnicity_from_report}' 不一致")

    # 规则2: 民族与处分决定比对
    from .case_extractors_demographics import extract_ethnicity_from_decision_report
    extracted_ethnicity_from_decision = extract_ethnicity_from_decision_report(decision_text_raw)
    
    if (excel_ethnicity and excel_ethnicity.strip() != '' and 
        (extracted_ethnicity_from_decision is None or excel_ethnicity != extracted_ethnicity_from_decision)):
        ethnicity_mismatch_indices.add(index)
        issues_list.append({
            '案件编码': excel_case_code,
            '涉案人员编码': excel_person_code,
            '行号': index + 2,
            '比对字段': f"Q{app_config['COLUMN_MAPPINGS']['ethnicity']}",
            '被比对字段': f"CU{app_config['COLUMN_MAPPINGS']['disciplinary_decision']}",
            '问题描述': f"Q{index + 2}{app_config['COLUMN_MAPPINGS']['ethnicity']}与CU{index + 2}处分决定不一致",
            '列名': app_config['COLUMN_MAPPINGS']['ethnicity']
        })
        logger.warning(f"<立案 - （2.民族与处分决定）> - 行 {index + 2} - 民族 '{excel_ethnicity}' 与处分决定提取民族 '{extracted_ethnicity_from_decision}' 不一致")

    # 规则3: 民族与审查调查报告比对
    from .case_extractors_demographics import extract_ethnicity_from_investigation_report
    extracted_ethnicity_from_investigation = extract_ethnicity_from_investigation_report(investigation_text_raw)
    
    if (excel_ethnicity and excel_ethnicity.strip() != '' and 
        (extracted_ethnicity_from_investigation is None or excel_ethnicity != extracted_ethnicity_from_investigation)):
        ethnicity_mismatch_indices.add(index)
        issues_list.append({
            '案件编码': excel_case_code,
            '涉案人员编码': excel_person_code,
            '行号': index + 2,
            '比对字段': f"Q{app_config['COLUMN_MAPPINGS']['ethnicity']}",
            '被比对字段': f"CX{app_config['COLUMN_MAPPINGS']['investigation_report']}",
            '问题描述': f"Q{index + 2}{app_config['COLUMN_MAPPINGS']['ethnicity']}与CX{index + 2}审查调查报告不一致",
            '列名': app_config['COLUMN_MAPPINGS']['ethnicity']
        })
        logger.warning(f"<立案 - （3.民族与审查调查报告）> - 行 {index + 2} - 民族 '{excel_ethnicity}' 与审查调查报告提取民族 '{extracted_ethnicity_from_investigation}' 不一致")

    # 规则4: 民族与审理报告比对
    from .case_extractors_demographics import extract_ethnicity_from_trial_report
    extracted_ethnicity_from_trial = extract_ethnicity_from_trial_report(trial_text_raw)
    
    if (excel_ethnicity and excel_ethnicity.strip() != '' and 
        (extracted_ethnicity_from_trial is None or excel_ethnicity != extracted_ethnicity_from_trial)):
        ethnicity_mismatch_indices.add(index)
        issues_list.append({
            '案件编码': excel_case_code,
            '涉案人员编码': excel_person_code,
            '行号': index + 2,
            '比对字段': f"Q{app_config['COLUMN_MAPPINGS']['ethnicity']}",
            '被比对字段': f"CY{app_config['COLUMN_MAPPINGS']['trial_report']}",
            '问题描述': f"Q{index + 2}{app_config['COLUMN_MAPPINGS']['ethnicity']}与CY{index + 2}审理报告不一致",
            '列名': app_config['COLUMN_MAPPINGS']['ethnicity']
        })
        logger.warning(f"<立案 - （4.民族与审理报告）> - 行 {index + 2} - 民族 '{excel_ethnicity}' 与审理报告提取民族 '{extracted_ethnicity_from_trial}' 不一致")

def validate_party_member_rules(row, index, excel_case_code, excel_person_code, issues_list, party_member_mismatch_indices,
                               excel_party_member, report_text_raw, decision_text_raw, app_config):
    """验证是否中共党员相关规则。
    """
    from .case_extractors_party_info import extract_party_member_from_case_report, extract_party_member_from_decision_report
    
    # 1. 是否中共党员与立案报告比对
    extracted_party_member_from_report = extract_party_member_from_case_report(report_text_raw)
    is_party_member_mismatch_report = False
    if not excel_party_member:
        if extracted_party_member_from_report == "是":
            is_party_member_mismatch_report = True
            issues_list.append({
                '案件编码': excel_case_code,
                '涉案人员编码': excel_person_code,
                '行号': index + 2,
                '比对字段': f"T{app_config['COLUMN_MAPPINGS']['party_member']}",
                '被比对字段': f"BF{app_config['COLUMN_MAPPINGS']['case_report']}",
                '问题描述': f"T{index + 2}{app_config['COLUMN_MAPPINGS']['party_member']}与BF{index + 2}立案报告不一致",
                '列名': app_config['COLUMN_MAPPINGS']['party_member']
            })
            logger.warning(f"<立案 - （1.是否中共党员与立案报告）> - 行 {index + 2} - 是否中共党员 '{excel_party_member}' 与立案报告提取党员信息 '是' 不一致")
    elif extracted_party_member_from_report is None:
        is_party_member_mismatch_report = True
        issues_list.append({
            '案件编码': excel_case_code,
            '涉案人员编码': excel_person_code,
            '行号': index + 2,
            '比对字段': f"T{app_config['COLUMN_MAPPINGS']['party_member']}",
            '被比对字段': f"BF{app_config['COLUMN_MAPPINGS']['case_report']}",
            '问题描述': f"T{index + 2}{app_config['COLUMN_MAPPINGS']['party_member']}与BF{index + 2}立案报告不一致",
            '列名': app_config['COLUMN_MAPPINGS']['party_member']
        })
        logger.warning(f"<立案 - （1.是否中共党员与立案报告）> - 行 {index + 2} - 是否中共党员 '{excel_party_member}' 与立案报告提取党员信息 '未明确' 不一致")
    elif excel_party_member != extracted_party_member_from_report:
        is_party_member_mismatch_report = True
        issues_list.append({
            '案件编码': excel_case_code,
            '涉案人员编码': excel_person_code,
            '行号': index + 2,
            '比对字段': f"T{app_config['COLUMN_MAPPINGS']['party_member']}",
            '被比对字段': f"BF{app_config['COLUMN_MAPPINGS']['case_report']}",
            '问题描述': f"T{index + 2}{app_config['COLUMN_MAPPINGS']['party_member']}与BF{index + 2}立案报告不一致",
            '列名': app_config['COLUMN_MAPPINGS']['party_member']
        })
        logger.warning(f"<立案 - （1.是否中共党员与立案报告）> - 行 {index + 2} - 是否中共党员 '{excel_party_member}' 与立案报告提取党员信息 '{extracted_party_member_from_report}' 不一致")
    if is_party_member_mismatch_report:
        party_member_mismatch_indices.add(index)

    # 2. 是否中共党员与处分决定比对
    extracted_party_member_from_decision = extract_party_member_from_decision_report(decision_text_raw)
    is_party_member_mismatch_decision = False
    if not excel_party_member:
        if extracted_party_member_from_decision == "是":
            is_party_member_mismatch_decision = True
            issues_list.append({
                '案件编码': excel_case_code,
                '涉案人员编码': excel_person_code,
                '行号': index + 2,
                '比对字段': f"T{app_config['COLUMN_MAPPINGS']['party_member']}",
                '被比对字段': f"CU{app_config['COLUMN_MAPPINGS']['disciplinary_decision']}",
                '问题描述': f"T{index + 2}{app_config['COLUMN_MAPPINGS']['party_member']}与CU{index + 2}处分决定不一致",
                '列名': app_config['COLUMN_MAPPINGS']['party_member']
            })
            logger.warning(f"<立案 - （2.是否中共党员与处分决定）> - 行 {index + 2} - 是否中共党员 '{excel_party_member}' 与处分决定提取党员信息 '是' 不一致")
        elif extracted_party_member_from_decision == "否":
            pass  # 如果Excel为空且处分决定提取为否，则认为一致
    elif extracted_party_member_from_decision is None:
        is_party_member_mismatch_decision = True
        issues_list.append({
            '案件编码': excel_case_code,
            '涉案人员编码': excel_person_code,
            '行号': index + 2,
            '比对字段': f"T{app_config['COLUMN_MAPPINGS']['party_member']}",
            '被比对字段': f"CU{app_config['COLUMN_MAPPINGS']['disciplinary_decision']}",
            '问题描述': f"T{index + 2}{app_config['COLUMN_MAPPINGS']['party_member']}与CU{index + 2}处分决定不一致",
            '列名': app_config['COLUMN_MAPPINGS']['party_member']
        })
        logger.warning(f"<立案 - （2.是否中共党员与处分决定）> - 行 {index + 2} - 是否中共党员 '{excel_party_member}' 与处分决定提取党员信息 '未明确' 不一致")
    elif excel_party_member != extracted_party_member_from_decision:
        is_party_member_mismatch_decision = True
        issues_list.append({
            '案件编码': excel_case_code,
            '涉案人员编码': excel_person_code,
            '行号': index + 2,
            '比对字段': f"T{app_config['COLUMN_MAPPINGS']['party_member']}",
            '被比对字段': f"CU{app_config['COLUMN_MAPPINGS']['disciplinary_decision']}",
            '问题描述': f"T{index + 2}{app_config['COLUMN_MAPPINGS']['party_member']}与CU{index + 2}处分决定不一致",
            '列名': app_config['COLUMN_MAPPINGS']['party_member']
        })
        logger.warning(f"<立案 - （2.是否中共党员与处分决定）> - 行 {index + 2} - 是否中共党员 '{excel_party_member}' 与处分决定提取党员信息 '{extracted_party_member_from_decision}' 不一致")
    if is_party_member_mismatch_decision:
        party_member_mismatch_indices.add(index)

def validate_party_joining_date_rules(row, index, excel_case_code, excel_person_code, issues_list, party_joining_date_mismatch_indices,
                                      excel_party_member, excel_party_joining_date, report_text_raw, app_config):
    """验证入党时间相关规则。"""
    from .case_extractors_party_info import extract_party_joining_date_from_case_report
    
    extracted_party_joining_date_from_report = extract_party_joining_date_from_case_report(report_text_raw)
    is_party_joining_date_mismatch = False

    if excel_party_member == "是":
        if not excel_party_joining_date:
            if extracted_party_joining_date_from_report is not None:
                is_party_joining_date_mismatch = True
                issues_list.append({
                    '案件编码': excel_case_code,
                    '涉案人员编码': excel_person_code,
                    '行号': index + 2,
                    '比对字段': f"AC{app_config['COLUMN_MAPPINGS']['party_joining_date']}",
                    '被比对字段': f"BF{app_config['COLUMN_MAPPINGS']['case_report']}",
                    '问题描述': f"AC{index + 2}{app_config['COLUMN_MAPPINGS']['party_joining_date']}与BF{index + 2}立案报告不一致",
                    '列名': app_config['COLUMN_MAPPINGS']['party_joining_date']
                })
                logger.warning(f"<立案 - （1.入党时间与立案报告）> - 行 {index + 2} - 入党时间 '{excel_party_joining_date}' 与立案报告提取入党时间 '{extracted_party_joining_date_from_report}' 不一致")
        elif extracted_party_joining_date_from_report is None:
            is_party_joining_date_mismatch = True
            issues_list.append({
                '案件编码': excel_case_code,
                '涉案人员编码': excel_person_code,
                '行号': index + 2,
                '比对字段': f"AC{app_config['COLUMN_MAPPINGS']['party_joining_date']}",
                '被比对字段': f"BF{app_config['COLUMN_MAPPINGS']['case_report']}",
                '问题描述': f"AC{index + 2}{app_config['COLUMN_MAPPINGS']['party_joining_date']}与BF{index + 2}立案报告不一致",
                '列名': app_config['COLUMN_MAPPINGS']['party_joining_date']
            })
            logger.warning(f"<立案 - （1.入党时间与立案报告）> - 行 {index + 2} - 入党时间 '{excel_party_joining_date}' 与立案报告提取入党时间 '未提取到' 不一致")
        elif excel_party_joining_date != extracted_party_joining_date_from_report:
            is_party_joining_date_mismatch = True
            issues_list.append({
                '案件编码': excel_case_code,
                '涉案人员编码': excel_person_code,
                '行号': index + 2,
                '比对字段': f"AC{app_config['COLUMN_MAPPINGS']['party_joining_date']}",
                '被比对字段': f"BF{app_config['COLUMN_MAPPINGS']['case_report']}",
                '问题描述': f"AC{index + 2}{app_config['COLUMN_MAPPINGS']['party_joining_date']}与BF{index + 2}立案报告不一致",
                '列名': app_config['COLUMN_MAPPINGS']['party_joining_date']
            })
            logger.warning(f"<立案 - （1.入党时间与立案报告）> - 行 {index + 2} - 入党时间 '{excel_party_joining_date}' 与立案报告提取入党时间 '{extracted_party_joining_date_from_report}' 不一致")
    elif excel_party_member == "否":
        if excel_party_joining_date:
            is_party_joining_date_mismatch = True
            issues_list.append({
                '案件编码': excel_case_code,
                '涉案人员编码': excel_person_code,
                '行号': index + 2,
                '比对字段': f"AC{app_config['COLUMN_MAPPINGS']['party_joining_date']}",
                '被比对字段': f"T{app_config['COLUMN_MAPPINGS']['party_member']}",
                '问题描述': f"AC{index + 2}{app_config['COLUMN_MAPPINGS']['party_joining_date']}与T{index + 2}是否中共党员不一致",
                '列名': app_config['COLUMN_MAPPINGS']['party_joining_date']
            })
            logger.warning(f"<立案 - （2.入党时间与党员身份）> - 行 {index + 2} - 入党时间 '{excel_party_joining_date}' 与是否中共党员 '否' 不一致")

    if is_party_joining_date_mismatch:
        party_joining_date_mismatch_indices.add(index)

def validate_brief_case_details_rules(row, index, excel_case_code, excel_person_code, issues_list, brief_case_details_mismatch_indices,
                                      excel_brief_case_details, investigated_person, report_text_raw, decision_text_raw, app_config):
    """
    验证简要案情相关规则。
    根据"处分决定"是否为空，从"立案报告"或"处分决定"中提取简要案情，并与Excel中的简要案情进行比较。

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
    from .case_extractors_demographics import extract_suspected_violation_from_case_report, extract_suspected_violation_from_decision
    
    is_brief_case_details_mismatch = False
    extracted_brief_case_details = None
    source_document_for_issue = ""

    # 规则1: 简要案情与立案报告或处分决定比对
    if pd.isna(row[app_config['COLUMN_MAPPINGS']["disciplinary_decision"]]) or decision_text_raw == '':
        extracted_brief_case_details = extract_suspected_violation_from_case_report(report_text_raw)
        source_document_for_issue = "立案报告"
        if extracted_brief_case_details is None:
            if excel_brief_case_details:
                is_brief_case_details_mismatch = True
                brief_case_details_mismatch_indices.add(index)
                issues_list.append({
                    '案件编码': excel_case_code,
                    '涉案人员编码': excel_person_code,
                    '行号': index + 2,
                    '比对字段': f"BE{app_config['COLUMN_MAPPINGS']['brief_case_details']}",
                    '被比对字段': f"BF{app_config['COLUMN_MAPPINGS']['case_report']}",
                    '问题描述': f"BE{index + 2}{app_config['COLUMN_MAPPINGS']['brief_case_details']}与BF{index + 2}立案报告不一致（未能提取到内容）",
                    '列名': app_config['COLUMN_MAPPINGS']['brief_case_details']
                })
                logger.warning(f"<立案 - （1.简要案情与立案报告）> - 行 {index + 2} - 简要案情 '{excel_brief_case_details}' 与立案报告提取简要案情 '未提取到' 不一致")
        else:
            cleaned_excel_brief_case_details = re.sub(r'\s+', '', excel_brief_case_details) if excel_brief_case_details else ''
            if cleaned_excel_brief_case_details != extracted_brief_case_details:
                is_brief_case_details_mismatch = True
                brief_case_details_mismatch_indices.add(index)
                issues_list.append({
                    '案件编码': excel_case_code,
                    '涉案人员编码': excel_person_code,
                    '行号': index + 2,
                    '比对字段': f"BE{app_config['COLUMN_MAPPINGS']['brief_case_details']}",
                    '被比对字段': f"BF{app_config['COLUMN_MAPPINGS']['case_report']}",
                    '问题描述': f"BE{index + 2}{app_config['COLUMN_MAPPINGS']['brief_case_details']}与BF{index + 2}立案报告不一致",
                    '列名': app_config['COLUMN_MAPPINGS']['brief_case_details']
                })
                logger.warning(f"<立案 - （1.简要案情与立案报告）> - 行 {index + 2} - 简要案情 '{cleaned_excel_brief_case_details}' 与立案报告提取简要案情 '{extracted_brief_case_details}' 不一致")
    else:
        extracted_brief_case_details = extract_suspected_violation_from_decision(decision_text_raw, investigated_person)
        source_document_for_issue = "处分决定"
        if extracted_brief_case_details is None:
            if excel_brief_case_details:
                is_brief_case_details_mismatch = True
                brief_case_details_mismatch_indices.add(index)
                issues_list.append({
                    '案件编码': excel_case_code,
                    '涉案人员编码': excel_person_code,
                    '行号': index + 2,
                    '比对字段': f"BE{app_config['COLUMN_MAPPINGS']['brief_case_details']}",
                    '被比对字段': f"CU{app_config['COLUMN_MAPPINGS']['disciplinary_decision']}",
                    '问题描述': f"BE{index + 2}{app_config['COLUMN_MAPPINGS']['brief_case_details']}与CU{index + 2}处分决定不一致（未能提取到内容）",
                    '列名': app_config['COLUMN_MAPPINGS']['brief_case_details']
                })
                logger.warning(f"<立案 - （2.简要案情与处分决定）> - 行 {index + 2} - 简要案情 '{excel_brief_case_details}' 与处分决定提取简要案情 '未提取到' 不一致")
        else:
            cleaned_excel_brief_case_details = re.sub(r'\s+', '', excel_brief_case_details) if excel_brief_case_details else ''
            if cleaned_excel_brief_case_details != extracted_brief_case_details:
                is_brief_case_details_mismatch = True
                brief_case_details_mismatch_indices.add(index)
                issues_list.append({
                    '案件编码': excel_case_code,
                    '涉案人员编码': excel_person_code,
                    '行号': index + 2,
                    '比对字段': f"BE{app_config['COLUMN_MAPPINGS']['brief_case_details']}",
                    '被比对字段': f"CU{app_config['COLUMN_MAPPINGS']['disciplinary_decision']}",
                    '问题描述': f"BE{index + 2}{app_config['COLUMN_MAPPINGS']['brief_case_details']}与CU{index + 2}处分决定不一致",
                    '列名': app_config['COLUMN_MAPPINGS']['brief_case_details']
                })
                logger.warning(f"<立案 - （2.简要案情与处分决定）> - 行 {index + 2} - 简要案情 '{cleaned_excel_brief_case_details}' 与处分决定提取简要案情 '{extracted_brief_case_details}' 不一致")

def validate_filing_time_rules(row, index, excel_case_code, excel_person_code, issues_list, filing_time_mismatch_indices,
                               excel_filing_time, excel_filing_decision_doc, app_config):
    """
    验证立案时间相关规则。
    比较 Excel 中的立案时间与立案决定书中提取的落款时间。

    参数:
        row (pd.Series): DataFrame 的当前行数据。
        index (int): 当前行的索引。
        excel_case_code (str): Excel 中的案件编码。
        excel_person_code (str): Excel 中的涉案人员编码。
        issues_list (list): 用于收集所有发现问题的列表。
        filing_time_mismatch_indices (set): 用于收集立案时间不匹配的行索引。
        excel_filing_time (str or None): Excel 中提取的立案时间。
        excel_filing_decision_doc (str or None): Excel 中的立案决定书内容。
        app_config (dict): Flask 应用的配置字典。
    """
    
    # 规则1: 立案时间与立案决定书落款时间比对
    from .case_extractors_timestamp import extract_filing_decision_signature_time
    extracted_signature_time = extract_filing_decision_signature_time(excel_filing_decision_doc)
    
    if (extracted_signature_time is None) or \
       (excel_filing_time is not None and extracted_signature_time is not None and excel_filing_time != extracted_signature_time):
        filing_time_mismatch_indices.add(index)
        issues_list.append({
            '案件编码': excel_case_code,
            '涉案人员编码': excel_person_code,
            '行号': index + 2,
            '比对字段': f"AR{app_config['COLUMN_MAPPINGS']['filing_time']}",
            '被比对字段': f"BG{app_config['COLUMN_MAPPINGS']['filing_decision_doc']}",
            '问题描述': f"AR{index + 2}{app_config['COLUMN_MAPPINGS']['filing_time']}与BG{index + 2}立案决定书落款时间不一致",
            '列名': app_config['COLUMN_MAPPINGS']['filing_time']
        })
        logger.warning(f"<立案 - （1.立案时间与立案决定书）> - 行 {index + 2} - 立案时间 '{excel_filing_time}' 与立案决定书落款时间 '{extracted_signature_time}' 不一致")

def validate_disciplinary_committee_filing_time_rules(row, index, excel_case_code, excel_person_code, issues_list, disciplinary_committee_filing_time_mismatch_indices,
                                                      excel_disciplinary_committee_filing_time, excel_filing_decision_doc, app_config):
    """
    验证纪委立案时间相关规则。
    比较 Excel 中的纪委立案时间与立案决定书中提取的落款时间。

    参数:
        row (pd.Series): DataFrame 的当前行数据。
        index (int): 当前行的索引。
        excel_case_code (str): Excel 中的案件编码。
        excel_person_code (str): Excel 中的涉案人员编码。
        issues_list (list): 用于收集所有发现问题的列表。
        disciplinary_committee_filing_time_mismatch_indices (set): 用于收集纪委立案时间不匹配的行索引。
        excel_disciplinary_committee_filing_time (str or None): Excel 中提取的纪委立案时间。
        excel_filing_decision_doc (str or None): Excel 中的立案决定书内容。
        app_config (dict): Flask 应用的配置字典。
    """
    
    # 规则1: 纪委立案时间与立案决定书落款时间比对
    from .case_extractors_timestamp import extract_filing_decision_signature_time
    extracted_signature_time = extract_filing_decision_signature_time(excel_filing_decision_doc)
    
    if (extracted_signature_time is None) or \
       (excel_disciplinary_committee_filing_time is not None and extracted_signature_time is not None and excel_disciplinary_committee_filing_time != extracted_signature_time):
        disciplinary_committee_filing_time_mismatch_indices.add(index)
        issues_list.append({
            '案件编码': excel_case_code,
            '涉案人员编码': excel_person_code,
            '行号': index + 2,
            '比对字段': f"AW{app_config['COLUMN_MAPPINGS']['disciplinary_committee_filing_time']}",
            '被比对字段': f"BG{app_config['COLUMN_MAPPINGS']['filing_decision_doc']}",
            '问题描述': f"AW{index + 2}{app_config['COLUMN_MAPPINGS']['disciplinary_committee_filing_time']}与BG{index + 2}立案决定书落款时间不一致",
            '列名': app_config['COLUMN_MAPPINGS']['disciplinary_committee_filing_time']
        })
        logger.warning(f"<立案 - （1.纪委立案时间与立案决定书）> - 行 {index + 2} - 纪委立案时间 '{excel_disciplinary_committee_filing_time}' 与立案决定书落款时间 '{extracted_signature_time}' 不一致")

def validate_supervisory_committee_filing_time_rules(row, index, excel_case_code, excel_person_code, issues_list, supervisory_committee_filing_time_mismatch_indices,
                                                     excel_supervisory_committee_filing_time, excel_filing_decision_doc, app_config):
    """
    验证监委立案时间相关规则。
    比较 Excel 中的监委立案时间与立案决定书中提取的落款时间。

    参数:
        row (pd.Series): DataFrame 的当前行数据。
        index (int): 当前行的索引。
        excel_case_code (str): Excel 中的案件编码。
        excel_person_code (str): Excel 中的涉案人员编码。
        issues_list (list): 用于收集所有发现问题的列表。
        supervisory_committee_filing_time_mismatch_indices (set): 用于收集监委立案时间不匹配的行索引。
        excel_supervisory_committee_filing_time (str or None): Excel 中提取的监委立案时间。
        excel_filing_decision_doc (str or None): Excel 中的立案决定书内容。
        app_config (dict): Flask 应用的配置字典。
    """
    
    # 规则1: 监委立案时间与立案决定书落款时间比对
    from .case_extractors_timestamp import extract_filing_decision_signature_time
    extracted_signature_time = extract_filing_decision_signature_time(excel_filing_decision_doc)
    
    if (extracted_signature_time is None) or \
       (excel_supervisory_committee_filing_time is not None and extracted_signature_time is not None and excel_supervisory_committee_filing_time != extracted_signature_time):
        supervisory_committee_filing_time_mismatch_indices.add(index)
        issues_list.append({
            '案件编码': excel_case_code,
            '涉案人员编码': excel_person_code,
            '行号': index + 2,
            '比对字段': f"AZ{app_config['COLUMN_MAPPINGS']['supervisory_committee_filing_time']}",
            '被比对字段': f"BG{app_config['COLUMN_MAPPINGS']['filing_decision_doc']}",
            '问题描述': f"AZ{index + 2}{app_config['COLUMN_MAPPINGS']['supervisory_committee_filing_time']}与BG{index + 2}立案决定书落款时间不一致",
            '列名': app_config['COLUMN_MAPPINGS']['supervisory_committee_filing_time']
        })
        logger.warning(f"<立案 - （1.监委立案时间与立案决定书）> - 行 {index + 2} - 监委立案时间 '{excel_supervisory_committee_filing_time}' 与立案决定书落款时间 '{extracted_signature_time}' 不一致")

def validate_disciplinary_committee_filing_authority_rules(row, index, excel_case_code, excel_person_code, issues_list, disciplinary_committee_filing_authority_mismatch_indices,
                                                           excel_disciplinary_committee_filing_authority, excel_reporting_unit_name, authority_agency_lookup, app_config):
    """
    验证纪委立案机关相关规则。
    检查纪委立案机关与填报单位名称是否在机关单位对应表中匹配。

    参数:
        row (pd.Series): DataFrame 的当前行数据。
        index (int): 当前行的索引。
        excel_case_code (str): Excel 中的案件编码。
        excel_person_code (str): Excel 中的涉案人员编码。
        issues_list (list): 用于收集所有发现问题的列表。
        disciplinary_committee_filing_authority_mismatch_indices (set): 用于收集纪委立案机关不匹配的行索引。
        excel_disciplinary_committee_filing_authority (str or None): Excel 中提取的纪委立案机关。
        excel_reporting_unit_name (str or None): Excel 中的填报单位名称。
        authority_agency_lookup (set): 机关单位对应表的查询集合。
        app_config (dict): Flask 应用的配置字典。
    """
    
    # 规则1: 纪委立案机关与填报单位名称匹配检查
    found_match_disciplinary = False
    if (excel_disciplinary_committee_filing_authority, excel_reporting_unit_name, "NSL") in authority_agency_lookup:
        found_match_disciplinary = True
    
    if not found_match_disciplinary:
        disciplinary_committee_filing_authority_mismatch_indices.add(index)
        issues_list.append({
            '案件编码': excel_case_code,
            '涉案人员编码': excel_person_code,
            '行号': index + 2,
            '比对字段': f"AV{app_config['COLUMN_MAPPINGS']['disciplinary_committee_filing_authority']}",
            '被比对字段': f"A{app_config['COLUMN_MAPPINGS']['reporting_agency']}",
            '问题描述': f"AV{index + 2}{app_config['COLUMN_MAPPINGS']['disciplinary_committee_filing_authority']}与A{index + 2}填报单位名称不一致",
            '列名': app_config['COLUMN_MAPPINGS']['disciplinary_committee_filing_authority']
        })
        logger.warning(f"<立案 - （1.纪委立案机关与填报单位名称）> - 行 {index + 2} - 纪委立案机关 '{excel_disciplinary_committee_filing_authority}' 与填报单位名称 '{excel_reporting_unit_name}' 不匹配")

def validate_supervisory_committee_filing_authority_rules(row, index, excel_case_code, excel_person_code, issues_list, supervisory_committee_filing_authority_mismatch_indices,
                                                          excel_supervisory_committee_filing_authority, excel_reporting_unit_name, authority_agency_lookup, app_config):
    """
    验证监委立案机关相关规则。
    检查监委立案机关与填报单位名称是否在机关单位对应表中匹配。

    参数:
        row (pd.Series): DataFrame 的当前行数据。
        index (int): 当前行的索引。
        excel_case_code (str): Excel 中的案件编码。
        excel_person_code (str): Excel 中的涉案人员编码。
        issues_list (list): 用于收集所有发现问题的列表。
        supervisory_committee_filing_authority_mismatch_indices (set): 用于收集监委立案机关不匹配的行索引。
        excel_supervisory_committee_filing_authority (str or None): Excel 中提取的监委立案机关。
        excel_reporting_unit_name (str or None): Excel 中的填报单位名称。
        authority_agency_lookup (set): 机关单位对应表的查询集合。
        app_config (dict): Flask 应用的配置字典。
    """
    
    # 规则1: 监委立案机关与填报单位名称匹配检查
    found_match_supervisory = False
    if (excel_supervisory_committee_filing_authority, excel_reporting_unit_name, "NSL") in authority_agency_lookup:
        found_match_supervisory = True
    
    if not found_match_supervisory:
        supervisory_committee_filing_authority_mismatch_indices.add(index)
        issues_list.append({
            '案件编码': excel_case_code,
            '涉案人员编码': excel_person_code,
            '行号': index + 2,
            '比对字段': f"AY{app_config['COLUMN_MAPPINGS']['supervisory_committee_filing_authority']}",
            '被比对字段': f"A{app_config['COLUMN_MAPPINGS']['reporting_agency']}",
            '问题描述': f"AY{index + 2}{app_config['COLUMN_MAPPINGS']['supervisory_committee_filing_authority']}与A{index + 2}填报单位名称不一致",
            '列名': app_config['COLUMN_MAPPINGS']['supervisory_committee_filing_authority']
        })
        logger.warning(f"<立案 - （1.监委立案机关与填报单位名称）> - 行 {index + 2} - 监委立案机关 '{excel_supervisory_committee_filing_authority}' 与填报单位名称 '{excel_reporting_unit_name}' 不匹配")

def validate_case_report_rules(row, index, excel_case_code, excel_person_code, issues_list, case_report_mismatch_indices,
                               case_report_keywords_to_check, report_text_raw, decision_text_raw, investigation_text_raw, trial_text_raw, app_config):
    """
    验证立案报告相关规则。
    检查立案报告中的关键字是否在处分决定、审理报告、审查调查报告中存在。

    参数:
        row (pd.Series): DataFrame 的当前行数据。
        index (int): 当前行的索引。
        excel_case_code (str): Excel 中的案件编码。
        excel_person_code (str): Excel 中的涉案人员编码。
        issues_list (list): 用于收集所有发现问题的列表。
        case_report_mismatch_indices (set): 用于收集立案报告不匹配的行索引。
        case_report_keywords_to_check (list): 需要检查的关键字列表。
        report_text_raw (str): 立案报告的原始文本。
        decision_text_raw (str): 处分决定的原始文本。
        investigation_text_raw (str): 审查调查报告的原始文本。
        trial_text_raw (str): 审理报告的原始文本。
        app_config (dict): Flask 应用的配置字典。
    """
    
    # 规则1: 立案报告关键字与其他报告的一致性检查
    found_keywords_in_case_report = [kw for kw in case_report_keywords_to_check if kw in report_text_raw]
    
    if found_keywords_in_case_report:
        keyword_mismatch_in_other_reports = False
        for keyword in found_keywords_in_case_report:
            if not (keyword in decision_text_raw and keyword in trial_text_raw and keyword in investigation_text_raw):
                keyword_mismatch_in_other_reports = True
                break

        if keyword_mismatch_in_other_reports:
            case_report_mismatch_indices.add(index)
            issues_list.append({
                '案件编码': excel_case_code,
                '涉案人员编码': excel_person_code,
                '行号': index + 2,
                '比对字段': f"BF{app_config['COLUMN_MAPPINGS']['case_report']}",
                '被比对字段': f"CU{app_config['COLUMN_MAPPINGS']['disciplinary_decision']}",
                '问题描述': f"BF{index + 2}{app_config['COLUMN_MAPPINGS']['case_report']}与CU{index + 2}处分决定、CY{index + 2}审理报告、CX{index + 2}审查调查报告不一致",
                '列名': app_config['COLUMN_MAPPINGS']['case_report']
            })
            logger.warning(f"<立案 - （1.立案报告与其他报告）> - 行 {index + 2} - 立案报告中关键字与处分决定、审理报告、审查调查报告不一致")

def validate_central_eight_provisions_rules(row, index, excel_case_code, excel_person_code, issues_list, central_eight_provisions_mismatch_indices,
                                           excel_central_eight_provisions, decision_text_raw, app_config):
    """
    验证是否违反中央八项规定精神相关规则。
    检查是否违反中央八项规定精神字段与处分决定内容的一致性。

    参数:
        row (pd.Series): DataFrame 的当前行数据。
        index (int): 当前行的索引。
        excel_case_code (str): Excel 中的案件编码。
        excel_person_code (str): Excel 中的涉案人员编码。
        issues_list (list): 用于收集所有发现问题的列表。
        central_eight_provisions_mismatch_indices (set): 用于收集是否违反中央八项规定精神不匹配的行索引。
        excel_central_eight_provisions (str): Excel 中的是否违反中央八项规定精神字段值。
        decision_text_raw (str): 处分决定的原始文本。
        app_config (dict): Flask 应用的配置字典。
    """
    
    # 规则1: 是否违反中央八项规定精神与处分决定的一致性检查
    decision_contains_violation_phrase = "违反中央八项规定精神" in decision_text_raw
    expected_central_eight_provisions = "是" if decision_contains_violation_phrase else "否"
    
    # 处理Excel值，确保空值、NaN或'nan'视为"否"
    if not excel_central_eight_provisions or excel_central_eight_provisions.lower() in ['nan', 'none', '']:
        excel_central_eight_provisions = "否"
    
    if excel_central_eight_provisions != expected_central_eight_provisions:
        central_eight_provisions_mismatch_indices.add(index)
        issues_list.append({
            '案件编码': excel_case_code,
            '涉案人员编码': excel_person_code,
            '行号': index + 2,
            '比对字段': f"BI{app_config['COLUMN_MAPPINGS']['central_eight_provisions']}",
            '被比对字段': f"CU{app_config['COLUMN_MAPPINGS']['disciplinary_decision']}",
            '问题描述': f"BI{index + 2}{app_config['COLUMN_MAPPINGS']['central_eight_provisions']}与CU{index + 2}处分决定不一致",
            '列名': app_config['COLUMN_MAPPINGS']['central_eight_provisions']
        })
        logger.warning(f"<立案 - （1.是否违反中央八项规定精神与处分决定）> - 行 {index + 2} - 是否违反中央八项规定精神 '{excel_central_eight_provisions}' 与处分决定内容不一致，预期为 '{expected_central_eight_provisions}'")

def validate_case_report_keywords_rules(row, index, excel_case_code, excel_person_code, issues_list, case_report_keyword_mismatch_indices,
                                        case_report_keywords_to_check, report_text_raw, decision_text_raw, investigation_text_raw, trial_text_raw, app_config):
    """验证立案报告关键字规则。
    新增 app_config 参数以匹配调用方传递的参数数量。
    """
    
    found_keywords_in_case_report = [kw for kw in case_report_keywords_to_check if kw in report_text_raw]
    
    if found_keywords_in_case_report:
        logger.info(f"行 {index + 1} - 立案报告中发现关键字: {found_keywords_in_case_report}")
        print(f"行 {index + 1} - 立案报告中发现关键字: {found_keywords_in_case_report}")

        keyword_mismatch_in_other_reports = False
        for keyword in found_keywords_in_case_report:
            if not (keyword in decision_text_raw and keyword in trial_text_raw and keyword in investigation_text_raw):
                keyword_mismatch_in_other_reports = True
                logger.info(f"行 {index + 1} - 关键字 '{keyword}' 在处分决定、审理报告或审查调查报告中缺失。")
                print(f"行 {index + 1} - 关键字 '{keyword}' 在处分决定、审理报告或审查调查报告中缺失。")
                break

        if keyword_mismatch_in_other_reports:
            case_report_keyword_mismatch_indices.add(index)
            issues_list.append((index, excel_case_code, excel_person_code, "BF立案报告与CU处分决定、CY审理报告、CX审查调查报告不一致"))
            logger.warning(f"行 {index + 1} - 规则违规: 立案报告中关键字与处分决定、审理报告、审查调查报告不一致。")
            print(f"行 {index + 1} - 规则违规: 立案报告中关键字与处分决定、审理报告、审查调查报告不一致。")
        else:
            logger.info(f"行 {index + 1} - 立案报告中所有关键字在处分决定、审理报告和审查调查报告中均存在。")
            print(f"行 {index + 1} - 立案报告中所有关键字在处分决定、审理报告和审查调查报告中均存在。")
    else:
        logger.info(f"行 {index + 1} - 立案报告中未发现指定关键字。")
        print(f"行 {index + 1} - 立案报告中未发现指定关键字。")

def validate_voluntary_confession_rules(row, index, excel_case_code, excel_person_code, issues_list, voluntary_confession_highlight_indices,
                                        excel_voluntary_confession, trial_text_raw, app_config):
    """验证是否主动交代问题规则。
    新增 app_config 参数以匹配调用方传递的参数数量。
    """
    
    trial_report_contains_confession = "主动交代" in trial_text_raw

    # 规则1: 是否主动交代问题与审理报告比对

    if trial_report_contains_confession:
        voluntary_confession_highlight_indices.add(index)
        issues_list.append({
            '案件编码': excel_case_code,
            '涉案人员编码': excel_person_code,
            '行号': index + 2,
            '比对字段': f"BK{app_config['COLUMN_MAPPINGS']['voluntary_confession']}",
            '被比对字段': f"CY{app_config['COLUMN_MAPPINGS']['trial_report']}",
            '问题描述': f"BK{index + 2}{app_config['COLUMN_MAPPINGS']['voluntary_confession']}与CY{index + 2}审理报告不一致",
            '列名': app_config['COLUMN_MAPPINGS']['voluntary_confession']
        })
        logger.warning(f"<立案 - （1.是否主动交代问题与审理报告）> - 行 {index + 2} - 审理报告中发现'主动交代'关键字，需要人工确认是否主动交代问题字段")
        

def validate_disciplinary_sanction_rules(row, index, excel_case_code, excel_person_code, issues_list, disciplinary_sanction_mismatch_indices,
                                         excel_disciplinary_sanction, decision_text_raw, app_config):
    """验证党纪处分规则。
    比较 Excel 中的党纪处分与处分决定中的内容是否一致。

    参数:
        row (pd.Series): DataFrame 的当前行数据。
        index (int): 当前行的索引。
        excel_case_code (str): Excel 中的案件编码。
        excel_person_code (str): Excel 中的涉案人员编码。
        issues_list (list): 用于收集所有发现问题的列表。
        disciplinary_sanction_mismatch_indices (set): 用于收集党纪处分不匹配的行索引。
        excel_disciplinary_sanction (str or None): Excel 中提取的党纪处分。
        decision_text_raw (str): 处分决定的原始文本。
        app_config (dict): Flask 应用的配置字典。
    """
    
    # 检查党纪处分字段是否为空
    if not excel_disciplinary_sanction or str(excel_disciplinary_sanction).strip() == "":
        return
    
    # 检查处分决定字段是否为空
    if not decision_text_raw or str(decision_text_raw).strip() == "":
        logger.warning(f"<立案 - （1.党纪处分验证）> - 行 {index + 2} - 处分决定字段为空，无法进行比对")
        disciplinary_sanction_mismatch_indices.add(index)
        issues_list.append({
             '案件编码': excel_case_code,
             '涉案人员编码': excel_person_code,
             '行号': index + 2,
             '比对字段': f"BO{app_config['COLUMN_MAPPINGS']['disciplinary_sanction']}",
             '被比对字段': f"CU{app_config['COLUMN_MAPPINGS']['disciplinary_decision']}",
             '问题描述': f"BO{index + 2}{app_config['COLUMN_MAPPINGS']['disciplinary_sanction']}与CU{index + 2}处分决定不一致",
             '列名': app_config['COLUMN_MAPPINGS']['disciplinary_sanction']
         })
        return
    
    # 获取党纪处分关键词
    sanction_keywords = app_config.get('DISCIPLINARY_SANCTION_KEYWORDS', [])
    
    # 规则1: 党纪处分与处分决定内容一致性验证
    excel_disciplinary_sanction_str = str(excel_disciplinary_sanction).strip()
    decision_text_str = str(decision_text_raw).strip()
    
    # 检查党纪处分字段中的内容是否在处分决定中出现
    sanction_found = False
    
    # 如果有配置的关键词，检查关键词是否在处分决定中
    if sanction_keywords:
        for keyword in sanction_keywords:
            if keyword in excel_disciplinary_sanction_str and keyword in decision_text_str:
                sanction_found = True
                break
    else:
        # 如果没有配置关键词，直接检查党纪处分内容是否在处分决定中
        if excel_disciplinary_sanction_str in decision_text_str:
            sanction_found = True
    
    if not sanction_found:
        logger.warning(f"<立案 - （1.党纪处分验证）> - 行 {index + 2} - 党纪处分 '{excel_disciplinary_sanction}' 与处分决定内容不一致")
        disciplinary_sanction_mismatch_indices.add(index)
        issues_list.append({
             '案件编码': excel_case_code,
             '涉案人员编码': excel_person_code,
             '行号': index + 2,
             '比对字段': f"BO{app_config['COLUMN_MAPPINGS']['disciplinary_sanction']}",
             '被比对字段': f"CU{app_config['COLUMN_MAPPINGS']['disciplinary_decision']}",
             '问题描述': f"BO{index + 2}{app_config['COLUMN_MAPPINGS']['disciplinary_sanction']}与CU{index + 2}处分决定不一致",
             '列名': app_config['COLUMN_MAPPINGS']['disciplinary_sanction']
         })

def validate_case_closing_time_rules(row, index, excel_case_code, excel_person_code, issues_list, closing_time_mismatch_indices,
                                     excel_closing_time, decision_text_raw, app_config):
    """验证结案时间规则。
    检查结案时间字段与处分决定中的生效日期是否一致。
    
    参数:
        row (pd.Series): DataFrame 的当前行数据。
        index (int): 当前行的索引。
        excel_case_code (str): Excel 中的案件编码。
        excel_person_code (str): Excel 中的涉案人员编码。
        issues_list (list): 用于收集所有发现问题的列表。
        closing_time_highlight_indices (set): 用于收集结案时间不匹配的行索引。
        excel_closing_time (str): Excel 中的结案时间字段值。
        decision_text_raw (str): 处分决定的原始文本。
        app_config (dict): Flask 应用的配置字典。
    """
    import re
    from datetime import datetime
    import pandas as pd
    
    logger.info(f"开始结案时间验证 - 行{index+1}: excel_closing_time='{excel_closing_time}', decision_text_raw长度={len(decision_text_raw)}")
    
    # 规则1: 结案时间与处分决定比对
    excel_closing_time_obj = None
    if pd.notna(excel_closing_time) and excel_closing_time:
        try:
            # 尝试将Excel中的结案时间转换为日期对象，忽略时间部分
            if isinstance(excel_closing_time, datetime):
                excel_closing_time_obj = excel_closing_time.date()
            else:
                # pd.to_datetime 可以很好地处理多种日期格式
                excel_closing_time_obj = pd.to_datetime(excel_closing_time).date()
        except Exception as e:
            logger.warning(f"<立案 - （1.结案时间格式）> - 行 {index + 2} - 无法解析结案时间字段 '{excel_closing_time}': {e}")
            closing_time_mismatch_indices.add(index)
            issues_list.append({
                '案件编码': excel_case_code,
                '涉案人员编码': excel_person_code,
                '行号': index + 2,
                '比对字段': f"BN{app_config['COLUMN_MAPPINGS']['closing_time']}",
                '被比对字段': f"BN{app_config['COLUMN_MAPPINGS']['closing_time']}",
                '问题描述': f"BN{index + 2}{app_config['COLUMN_MAPPINGS']['closing_time']}格式不正确",
                '列名': app_config['COLUMN_MAPPINGS']['closing_time']
            })
            return
    
    # 使用正则表达式从处分决定中提取生效日期
    match = re.search(r"本处分决定自(\d{4}年\d{1,2}月\d{1,2}日)起生效", decision_text_raw)
    extracted_disposal_date = None
    
    if match:
        date_str = match.group(1)
        try:
            # 转换提取到的日期字符串为 datetime.date 对象
            formatted_date_str = date_str.replace('年', '-').replace('月', '-').replace('日', '')
            extracted_disposal_date = datetime.strptime(formatted_date_str, "%Y-%m-%d").date()
            logger.info(f"行 {index + 2} - 从处分决定中提取的生效日期: '{date_str}' (格式化后: {extracted_disposal_date})")
        except ValueError as ve:
            logger.warning(f"行 {index + 2} - 无法解析处分决定中提取的日期 '{date_str}': {ve}")
    else:
        logger.info(f"行 {index + 2} - 处分决定中未找到匹配的生效日期模式")
    
    # 进行结案时间与提取日期之间的比对
    if excel_closing_time_obj and extracted_disposal_date:
        if excel_closing_time_obj != extracted_disposal_date:
            closing_time_mismatch_indices.add(index)
            issues_list.append({
                '案件编码': excel_case_code,
                '涉案人员编码': excel_person_code,
                '行号': index + 2,
                '比对字段': f"BN{app_config['COLUMN_MAPPINGS']['closing_time']}",
                '被比对字段': f"CU{app_config['COLUMN_MAPPINGS']['disciplinary_decision']}",
                '问题描述': f"BN{index + 2}{app_config['COLUMN_MAPPINGS']['closing_time']}与CU{index + 2}处分决定不一致",
                '列名': app_config['COLUMN_MAPPINGS']['closing_time']
            })
            logger.warning(f"<立案 - （1.结案时间与处分决定）> - 行 {index + 2} - 结案时间 '{excel_closing_time_obj}' 与处分决定中提取的生效日期 '{extracted_disposal_date}' 不一致")
        else:
            logger.info(f"<立案 - （1.结案时间与处分决定）> - 行 {index + 2} - 结案时间与处分决定中提取的生效日期一致")
    elif excel_closing_time_obj is None:
        logger.info(f"行 {index + 2} - 结案时间字段为空，跳过比对")
    elif extracted_disposal_date is None:
        logger.info(f"行 {index + 2} - 未能从处分决定中提取到生效日期，跳过比对")

def validate_no_party_position_warning_rules(row, index, excel_case_code, excel_person_code, issues_list, no_party_position_warning_mismatch_indices,
                                             excel_no_party_position_warning, decision_text_raw, app_config):
    """
    验证是否属于本应撤销党内职务，但本人没有党内职务而给予严重警告处分规则。
    比较 Excel 中的BP字段与处分决定中的相关内容。

    参数:
        row (pd.Series): DataFrame 的当前行数据。
        index (int): 当前行的索引。
        excel_case_code (str): Excel 中的案件编码。
        excel_person_code (str): Excel 中的涉案人员编码。
        issues_list (list): 用于收集所有发现问题的列表。
        no_party_position_warning_mismatch_indices (set): 用于收集BP字段不匹配的行索引。
        excel_no_party_position_warning (str): Excel 中的BP字段值。
        decision_text_raw (str): 处分决定的原始文本。
        app_config (dict): Flask 应用的配置字典。
    """
    
    target_string = "属于本应撤销党内职务，但本人没有党内职务而给予严重警告处分"
    decision_contains_warning = target_string in decision_text_raw
    extracted_no_party_position_warning = "是" if decision_contains_warning else "否"
    
    # 处理 Excel 值，确保空值、NaN 或 'nan' 视为"否"
    excel_value = str(excel_no_party_position_warning or "").strip()
    excel_no_party_position_warning = "否" if not excel_value or excel_value.lower() == 'nan' else excel_value

    if excel_no_party_position_warning != extracted_no_party_position_warning:
        no_party_position_warning_mismatch_indices.add(index)
        issues_list.append({
            '案件编码': excel_case_code,
            '涉案人员编码': excel_person_code,
            '行号': index + 2,
            '比对字段': f"BP{app_config['COLUMN_MAPPINGS']['no_party_position_warning']}",
            '被比对字段': f"CU{app_config['COLUMN_MAPPINGS']['disciplinary_decision']}",
            '问题描述': f"BP{index + 2}{app_config['COLUMN_MAPPINGS']['no_party_position_warning']}与CU{index + 2}处分决定不一致",
            '列名': app_config['COLUMN_MAPPINGS']['no_party_position_warning']
        })
        logger.warning(f"<立案 - （1.BP字段与处分决定）> - 行 {index + 2} - BP字段 '{excel_no_party_position_warning}' 与处分决定提取值 '{extracted_no_party_position_warning}' 不一致")
