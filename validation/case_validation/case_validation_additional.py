import logging
import pandas as pd
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

    logger.info(f"行 {index + 1} - 字段 '是否主动交代问题' Excel值: '{excel_voluntary_confession}'。审理报告中'主动交代'匹配结果: {trial_report_contains_confession}。")
    print(f"行 {index + 1} - 字段 '是否主动交代问题' Excel值: '{excel_voluntary_confession}'。审理报告中'主动交代'匹配结果: {trial_report_contains_confession}。")

    if trial_report_contains_confession:
        voluntary_confession_highlight_indices.add(index)
        issues_list.append((index, excel_case_code, excel_person_code, "请基于CY审理报告进行人工确认主动交代"))
        logger.warning(f"行 {index + 1} - 规则触发: 审理报告中发现“主动交代”，已标记“是否主动交代问题”字段为黄色并添加问题描述。")
        print(f"行 {index + 1} - 规则触发: 审理报告中发现“主动交代”，已标记“是否主动交代问题”字段为黄色并添加问题描述。")

def validate_no_party_position_warning_rules(row, index, excel_case_code, excel_person_code, issues_list, no_party_position_warning_mismatch_indices,
                                             excel_no_party_position_warning, decision_text_raw, app_config):
    """验证是否属于本应撤销党内职务，但本人没有党内职务而给予严重警告处分规则。
    新增 app_config 参数以匹配调用方传递的参数数量。
    """
    
    target_string = "属于本应撤销党内职务，但本人没有党内职务而给予严重警告处分"
    decision_contains_warning = target_string in decision_text_raw
    extracted_no_party_position_warning = "是" if decision_contains_warning else "否"
    
    # 处理 Excel 值，确保空值、NaN 或 'nan' 视为“否”
    excel_value = str(row.get("是否属于本应撤销党内职务，但本人没有党内职务而给予严重警告处分", "")).strip()
    excel_no_party_position_warning = "否" if not excel_value or excel_value.lower() == 'nan' else excel_value

    logger.info(f"行 {index + 1} - 字段 '是否属于本应撤销党内职务，但本人没有党内职务而给予严重警告处分' Excel值: '{excel_no_party_position_warning}'。处分决定中匹配结果: {decision_contains_warning} (提取值: '{extracted_no_party_position_warning}')。")
    print(f"行 {index + 1} - 字段 '是否属于本应撤销党内职务，但本人没有党内职务而给予严重警告处分' Excel值: '{excel_no_party_position_warning}'。处分决定中匹配结果: {decision_contains_warning} (提取值: '{extracted_no_party_position_warning}')。")

    if excel_no_party_position_warning != extracted_no_party_position_warning:
        no_party_position_warning_mismatch_indices.add(index)
        issues_list.append((index, excel_case_code, excel_person_code, "BP是否属于本应撤销党内职务，但本人没有党内职务而给予严重警告处分与CU处分决定不一致"))
        logger.warning(f"行 {index + 1} - 规则违规: Excel值 ('{excel_no_party_position_warning}') 与处分决定提取值 ('{extracted_no_party_position_warning}') 不一致，已标记字段为红色。")
        print(f"行 {index + 1} - 规则违规: Excel值 ('{excel_no_party_position_warning}') 与处分决定提取值 ('{extracted_no_party_position_warning}') 不一致，已标记字段为红色。")
