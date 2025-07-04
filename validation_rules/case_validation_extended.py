import logging
import pandas as pd
from validation_rules.case_extractors_birth_info import (
    extract_birth_date_from_case_report,
    extract_birth_date_from_decision_report,
    extract_birth_date_from_investigation_report,
    extract_birth_date_from_trial_report
)
from validation_rules.case_extractors_demographics import (
    extract_education_from_case_report,
    extract_ethnicity_from_case_report,
    extract_ethnicity_from_decision_report,
    extract_ethnicity_from_investigation_report,
    extract_ethnicity_from_trial_report
)
from validation_rules.case_extractors_party_info import (
    extract_party_member_from_case_report,
    extract_party_member_from_decision_report,
    extract_party_joining_date_from_case_report
)

logger = logging.getLogger(__name__)

def validate_birth_date_rules(row, index, excel_case_code, excel_person_code, issues_list, birth_date_mismatch_indices,
                              excel_birth_date, report_text_raw, decision_text_raw, investigation_text_raw, trial_text_raw, app_config):
    """验证出生年月相关规则。
    新增 app_config 参数以匹配调用方传递的参数数量。
    """
    
    extracted_birth_date_from_report = extract_birth_date_from_case_report(report_text_raw)
    is_birth_date_mismatch_report = False
    if pd.isna(row["出生年月"]) or excel_birth_date == '':
        if extracted_birth_date_from_report is not None:
            is_birth_date_mismatch_report = True
    elif extracted_birth_date_from_report is None:
        is_birth_date_mismatch_report = True
    elif excel_birth_date != extracted_birth_date_from_report:
        is_birth_date_mismatch_report = True
    if is_birth_date_mismatch_report:
        birth_date_mismatch_indices.add(index)
        issues_list.append((index, excel_case_code, excel_person_code, "O2出生年月与BF2立案报告不一致"))
        logger.info(f"行 {index + 1} - 出生年月不匹配: Excel出生年月 ('{excel_birth_date}') vs 立案报告提取出生年月 ('{extracted_birth_date_from_report}')")
        print(f"行 {index + 1} - 出生年月不匹配: Excel出生年月 ('{excel_birth_date}') vs 立案报告提取出生年月 ('{extracted_birth_date_from_report}')")

    extracted_birth_date_from_decision = extract_birth_date_from_decision_report(decision_text_raw)
    is_birth_date_mismatch_decision = False
    if pd.isna(row["出生年月"]) or excel_birth_date == '':
        if extracted_birth_date_from_decision is not None:
            is_birth_date_mismatch_decision = True
    elif extracted_birth_date_from_decision is None:
        is_birth_date_mismatch_decision = True
    elif excel_birth_date != extracted_birth_date_from_decision:
        is_birth_date_mismatch_decision = True
    if is_birth_date_mismatch_decision:
        birth_date_mismatch_indices.add(index)
        issues_list.append((index, excel_case_code, excel_person_code, "O2出生年月与CU2处分决定不一致"))
        logger.info(f"行 {index + 1} - 出生年月不匹配: Excel出生年月 ('{excel_birth_date}') vs 处分决定提取出生年月 ('{extracted_birth_date_from_decision}')")
        print(f"行 {index + 1} - 出生年月不匹配: Excel出生年月 ('{excel_birth_date}') vs 处分决定提取出生年月 ('{extracted_birth_date_from_decision}')")

    extracted_birth_date_from_investigation = extract_birth_date_from_investigation_report(investigation_text_raw)
    is_birth_date_mismatch_investigation = False
    if pd.isna(row["出生年月"]) or excel_birth_date == '':
        if extracted_birth_date_from_investigation is not None:
            is_birth_date_mismatch_investigation = True
    elif extracted_birth_date_from_investigation is None:
        is_birth_date_mismatch_investigation = True
    elif excel_birth_date != extracted_birth_date_from_investigation:
        is_birth_date_mismatch_investigation = True
    if is_birth_date_mismatch_investigation:
        birth_date_mismatch_indices.add(index)
        issues_list.append((index, excel_case_code, excel_person_code, "O2出生年月与CX2审查调查报告不一致"))
        logger.info(f"行 {index + 1} - 出生年月不匹配: Excel出生年月 ('{excel_birth_date}') vs 审查调查报告提取出生年月 ('{extracted_birth_date_from_investigation}')")
        print(f"行 {index + 1} - 出生年月不匹配: Excel出生年月 ('{excel_birth_date}') vs 审查调查报告提取出生年月 ('{extracted_birth_date_from_investigation}')")

    extracted_birth_date_from_trial = extract_birth_date_from_trial_report(trial_text_raw)
    is_birth_date_mismatch_trial = False
    if pd.isna(row["出生年月"]) or excel_birth_date == '':
        if extracted_birth_date_from_trial is not None:
            is_birth_date_mismatch_trial = True
    elif extracted_birth_date_from_trial is None:
        is_birth_date_mismatch_trial = True
    elif excel_birth_date != extracted_birth_date_from_trial:
        is_birth_date_mismatch_trial = True
    if is_birth_date_mismatch_trial:
        birth_date_mismatch_indices.add(index)
        issues_list.append((index, excel_case_code, excel_person_code, "O2出生年月与CY2审理报告不一致"))
        logger.info(f"行 {index + 1} - 出生年月不匹配: Excel出生年月 ('{excel_birth_date}') vs 审理报告提取出生年月 ('{extracted_birth_date_from_trial}')")
        print(f"行 {index + 1} - 出生年月不匹配: Excel出生年月 ('{excel_birth_date}') vs 审理报告提取出生年月 ('{extracted_birth_date_from_trial}')")

def validate_education_rules(row, index, excel_case_code, excel_person_code, issues_list, education_mismatch_indices,
                             excel_education, report_text_raw, app_config):
    """验证学历相关规则。
    新增 app_config 参数以匹配调用方传递的参数数量。
    """
    
    extracted_education_from_report = extract_education_from_case_report(report_text_raw)
    is_education_mismatch_report = False
    excel_education_normalized = excel_education
    if excel_education == "大学本科":
        excel_education_normalized = "本科"
    extracted_education_normalized = extracted_education_from_report
    if extracted_education_from_report == "大学本科":
        extracted_education_normalized = "本科"

    if not excel_education:
        if extracted_education_from_report is not None:
            is_education_mismatch_report = True
            issues_list.append((index, excel_case_code, excel_person_code, "P2学历与BF2立案报告不一致"))
            logger.info(f"行 {index + 1} - 学历不匹配: Excel学历为空，但立案报告中提取到学历 ('{extracted_education_from_report}')。")
            print(f"行 {index + 1} - 学历不匹配: Excel学历为空，但立案报告中提取到学历 ('{extracted_education_from_report}')。")
    else:
        if extracted_education_from_report is None:
            is_education_mismatch_report = True
            issues_list.append((index, excel_case_code, excel_person_code, "P2学历与BF2立案报告不一致"))
            logger.info(f"行 {index + 1} - 学历不匹配: Excel学历 ('{excel_education}') 有值，但立案报告中未提取到学历。")
            print(f"行 {index + 1} - 学历不匹配: Excel学历 ('{excel_education}') 有值，但立案报告中未提取到学历。")
        elif excel_education_normalized != extracted_education_normalized:
            is_education_mismatch_report = True
            issues_list.append((index, excel_case_code, excel_person_code, "P2学历与BF2立案报告不一致"))
            logger.info(f"行 {index + 1} - 学历不匹配: Excel学历 ('{excel_education}') vs 立案报告提取学历 ('{extracted_education_from_report}')。")
            print(f"行 {index + 1} - 学历不匹配: Excel学历 ('{excel_education}') vs 立案报告提取学历 ('{extracted_education_from_report}')")
    if is_education_mismatch_report:
        education_mismatch_indices.add(index)

def validate_ethnicity_rules(row, index, excel_case_code, excel_person_code, issues_list, ethnicity_mismatch_indices,
                             excel_ethnicity, report_text_raw, decision_text_raw, investigation_text_raw, trial_text_raw, app_config):
    """验证民族相关规则。
    新增 app_config 参数以匹配调用方传递的参数数量。
    """
    
    extracted_ethnicity_from_report = extract_ethnicity_from_case_report(report_text_raw)
    is_ethnicity_mismatch_report = False
    if not excel_ethnicity:
        if extracted_ethnicity_from_report is not None:
            is_ethnicity_mismatch_report = True
            issues_list.append((index, excel_case_code, excel_person_code, "Q2民族与BF2立案报告不一致"))
            logger.info(f"行 {index + 1} - 民族不匹配: Excel民族为空，但立案报告中提取到民族 ('{extracted_ethnicity_from_report}')。")
            print(f"行 {index + 1} - 民族不匹配: Excel民族为空，但立案报告中提取到民族 ('{extracted_ethnicity_from_report}'))。")
    elif extracted_ethnicity_from_report is None:
        is_ethnicity_mismatch_report = True
        issues_list.append((index, excel_case_code, excel_person_code, "Q2民族与BF2立案报告不一致"))
        logger.info(f"行 {index + 1} - 民族不匹配: Excel民族 ('{excel_ethnicity}') 有值，但立案报告中未提取到民族。")
        print(f"行 {index + 1} - 民族不匹配: Excel民族 ('{excel_ethnicity}') 有值，但立案报告中未提取到民族。")
    elif excel_ethnicity != extracted_ethnicity_from_report:
        is_ethnicity_mismatch_report = True
        issues_list.append((index, excel_case_code, excel_person_code, "Q2民族与BF2立案报告不一致"))
        logger.info(f"行 {index + 1} - 民族不匹配: Excel民族 ('{excel_ethnicity}') vs 立案报告提取民族 ('{extracted_ethnicity_from_report}')。")
        print(f"行 {index + 1} - 民族不匹配: Excel民族 ('{excel_ethnicity}') vs 立案报告提取民族 ('{extracted_ethnicity_from_report}')")
    if is_ethnicity_mismatch_report:
        ethnicity_mismatch_indices.add(index)
    
    extracted_ethnicity_from_decision = extract_ethnicity_from_decision_report(decision_text_raw)
    is_ethnicity_mismatch_decision = False
    if not excel_ethnicity:
        if extracted_ethnicity_from_decision is not None:
            is_ethnicity_mismatch_decision = True
            issues_list.append((index, excel_case_code, excel_person_code, "Q2民族与CU2处分决定不一致"))
            logger.info(f"行 {index + 1} - 民族不匹配: Excel民族为空，但处分决定中提取到民族 ('{extracted_ethnicity_from_decision}')。")
            print(f"行 {index + 1} - 民族不匹配: Excel民族为空，但处分决定中提取到民族 ('{extracted_ethnicity_from_decision}')。")
    elif extracted_ethnicity_from_decision is None:
        is_ethnicity_mismatch_decision = True
        issues_list.append((index, excel_case_code, excel_person_code, "Q2民族与CU2处分决定不一致"))
        logger.info(f"行 {index + 1} - 民族不匹配: Excel民族 ('{excel_ethnicity}') 有值，但处分决定中未提取到民族。")
        print(f"行 {index + 1} - 民族不匹配: Excel民族 ('{excel_ethnicity}') 有值，但处分决定中未提取到民族。")
    elif excel_ethnicity != extracted_ethnicity_from_decision:
        is_ethnicity_mismatch_decision = True
        issues_list.append((index, excel_case_code, excel_person_code, "Q2民族与CU2处分决定不一致"))
        logger.info(f"行 {index + 1} - 民族不匹配: Excel民族 ('{excel_ethnicity}') vs 处分决定提取民族 ('{extracted_ethnicity_from_decision}')。")
        print(f"行 {index + 1} - 民族不匹配: Excel民族 ('{excel_ethnicity}') vs 处分决定提取民族 ('{extracted_ethnicity_from_decision}')")
    if is_ethnicity_mismatch_decision:
        ethnicity_mismatch_indices.add(index)

    extracted_ethnicity_from_investigation = extract_ethnicity_from_investigation_report(investigation_text_raw)
    is_ethnicity_mismatch_investigation = False
    if not excel_ethnicity:
        if extracted_ethnicity_from_investigation is not None:
            is_ethnicity_mismatch_investigation = True
            issues_list.append((index, excel_case_code, excel_person_code, "Q2民族与CX2审查调查报告不一致"))
            logger.info(f"行 {index + 1} - 民族不匹配: Excel民族为空，但审查调查报告中提取到民族 ('{extracted_ethnicity_from_investigation}')。")
            print(f"行 {index + 1} - 民族不匹配: Excel民族为空，但审查调查报告中提取到民族 ('{extracted_ethnicity_from_investigation}')。")
    elif extracted_ethnicity_from_investigation is None:
        is_ethnicity_mismatch_investigation = True
        issues_list.append((index, excel_case_code, excel_person_code, "Q2民族与CX2审查调查报告不一致"))
        logger.info(f"行 {index + 1} - 民族不匹配: Excel民族 ('{excel_ethnicity}') 有值，但审查调查报告中未提取到民族。")
        print(f"行 {index + 1} - 民族不匹配: Excel民族 ('{excel_ethnicity}') 有值，但审查调查报告中未提取到民族。")
    elif excel_ethnicity != extracted_ethnicity_from_investigation:
        is_ethnicity_mismatch_investigation = True
        issues_list.append((index, excel_case_code, excel_person_code, "Q2民族与CX2审查调查报告不一致"))
        logger.info(f"行 {index + 1} - 民族不匹配: Excel民族 ('{excel_ethnicity}') vs 审查调查报告提取民族 ('{extracted_ethnicity_from_investigation}')。")
        print(f"行 {index + 1} - 民族不匹配: Excel民族 ('{excel_ethnicity}') vs 审查调查报告提取民族 ('{extracted_ethnicity_from_investigation}')")
    if is_ethnicity_mismatch_investigation:
        ethnicity_mismatch_indices.add(index)

    extracted_ethnicity_from_trial = extract_ethnicity_from_trial_report(trial_text_raw)
    is_ethnicity_mismatch_trial = False
    if not excel_ethnicity:
        if extracted_ethnicity_from_trial is not None:
            is_ethnicity_mismatch_trial = True
            issues_list.append((index, excel_case_code, excel_person_code, "Q2民族与CY2审理报告不一致"))
            logger.info(f"行 {index + 1} - 民族不匹配: Excel民族为空，但审理报告中提取到民族 ('{extracted_ethnicity_from_trial}')。")
            print(f"行 {index + 1} - 民族不匹配: Excel民族为空，但审理报告中提取到民族 ('{extracted_ethnicity_from_trial}')。")
    elif extracted_ethnicity_from_trial is None:
        is_ethnicity_mismatch_trial = True
        issues_list.append((index, excel_case_code, excel_person_code, "Q2民族与CY2审理报告不一致"))
        logger.info(f"行 {index + 1} - 民族不匹配: Excel民族 ('{excel_ethnicity}') 有值，但审理报告中未提取到民族。")
        print(f"行 {index + 1} - 民族不匹配: Excel民族 ('{excel_ethnicity}') 有值，但审理报告中未提取到民族。")
    elif excel_ethnicity != extracted_ethnicity_from_trial:
        is_ethnicity_mismatch_trial = True
        issues_list.append((index, excel_case_code, excel_person_code, "Q2民族与CY2审理报告不一致"))
        logger.info(f"行 {index + 1} - 民族不匹配: Excel民族 ('{excel_ethnicity}') vs 审理报告提取民族 ('{extracted_ethnicity_from_trial}')。")
        print(f"行 {index + 1} - 民族不匹配: Excel民族 ('{excel_ethnicity}') vs 审理报告提取民族 ('{extracted_ethnicity_from_trial}')")
    if is_ethnicity_mismatch_trial:
        ethnicity_mismatch_indices.add(index)

def validate_party_member_rules(row, index, excel_case_code, excel_person_code, issues_list, party_member_mismatch_indices,
                                 excel_party_member, report_text_raw, decision_text_raw, app_config):
    """验证是否中共党员相关规则。
    新增 app_config 参数以匹配调用方传递的参数数量。
    """
    
    extracted_party_member_from_report = extract_party_member_from_case_report(report_text_raw)
    is_party_member_mismatch_report = False
    if not excel_party_member:
        if extracted_party_member_from_report == "是":
            is_party_member_mismatch_report = True
            issues_list.append((index, excel_case_code, excel_person_code, "T2是否中共党员与BF2立案报告不一致"))
            logger.info(f"行 {index + 1} - 是否中共党员不匹配: Excel字段为空，但立案报告中提取到“是”。")
            print(f"行 {index + 1} - 是否中共党员不匹配: Excel字段为空，但立案报告中提取到“是”。")
    elif extracted_party_member_from_report is None:
        is_party_member_mismatch_report = True
        issues_list.append((index, excel_case_code, excel_person_code, "T2是否中共党员与BF2立案报告不一致"))
        logger.info(f"行 {index + 1} - 是否中共党员不匹配: Excel字段 ('{excel_party_member}') 有值，但立案报告中未明确提取到党员信息。")
        print(f"行 {index + 1} - 是否中共党员不匹配: Excel字段 ('{excel_party_member}') 有值，但立案报告中未明确提取到党员信息。")
    elif excel_party_member != extracted_party_member_from_report:
        is_party_member_mismatch_report = True
        issues_list.append((index, excel_case_code, excel_person_code, "T2是否中共党员与BF2立案报告不一致"))
        logger.info(f"行 {index + 1} - 是否中共党员不匹配: Excel字段 ('{excel_party_member}') vs 立案报告提取 ('{extracted_party_member_from_report}')。")
        print(f"行 {index + 1} - 是否中共党员不匹配: Excel字段 ('{excel_party_member}') vs 立案报告提取 ('{extracted_party_member_from_report}')。")
    if is_party_member_mismatch_report:
        party_member_mismatch_indices.add(index)

    extracted_party_member_from_decision = extract_party_member_from_decision_report(decision_text_raw)
    is_party_member_mismatch_decision = False
    if not excel_party_member:
        if extracted_party_member_from_decision == "是":
            is_party_member_mismatch_decision = True
            issues_list.append((index, excel_case_code, excel_person_code, "T2是否中共党员与CU2处分决定不一致"))
            logger.info(f"行 {index + 1} - 是否中共党员不匹配: Excel字段为空，但处分决定中提取到“是”。")
            print(f"行 {index + 1} - 是否中共党员不匹配: Excel字段为空，但处分决定中提取到“是”。")
        elif extracted_party_member_from_decision == "否":
            pass # 如果Excel为空且处分决定提取为否，则认为一致
    elif extracted_party_member_from_decision is None:
        is_party_member_mismatch_decision = True
        issues_list.append((index, excel_case_code, excel_person_code, "T2是否中共党员与CU2处分决定不一致"))
        logger.info(f"行 {index + 1} - 是否中共党员不匹配: Excel字段 ('{excel_party_member}') 有值，但处分决定中未明确提取到党员信息。")
        print(f"行 {index + 1} - 是否中共党员不匹配: Excel字段 ('{excel_party_member}') 有值，但处分决定中未明确提取到党员信息。")
    elif excel_party_member != extracted_party_member_from_decision:
        is_party_member_mismatch_decision = True
        issues_list.append((index, excel_case_code, excel_person_code, "T2是否中共党员与CU2处分决定不一致"))
        logger.info(f"行 {index + 1} - 是否中共党员不匹配: Excel字段 ('{excel_party_member}') vs 处分决定提取 ('{extracted_party_member_from_decision}')。")
        print(f"行 {index + 1} - 是否中共党员不匹配: Excel字段 ('{excel_party_member}') vs 处分决定提取 ('{extracted_party_member_from_decision}')。")
    if is_party_member_mismatch_decision:
        party_member_mismatch_indices.add(index)

def validate_party_joining_date_rules(row, index, excel_case_code, excel_person_code, issues_list, party_joining_date_mismatch_indices,
                                      excel_party_member, excel_party_joining_date, report_text_raw, app_config):
    """验证入党时间相关规则。
    新增 app_config 参数以匹配调用方传递的参数数量。
    """
    
    extracted_party_joining_date_from_report = extract_party_joining_date_from_case_report(report_text_raw)
    is_party_joining_date_mismatch = False

    if excel_party_member == "是":
        if not excel_party_joining_date:
            if extracted_party_joining_date_from_report is not None:
                is_party_joining_date_mismatch = True
                issues_list.append((index, excel_case_code, excel_person_code, "V2入党时间与BF2立案报告不一致"))
                logger.info(f"行 {index + 1} - 入党时间不匹配: Excel入党时间为空，但立案报告中提取到 ('{extracted_party_joining_date_from_report}')。")
                print(f"行 {index + 1} - 入党时间不匹配: Excel入党时间为空，但立案报告中提取到 ('{extracted_party_joining_date_from_report}')。")
        elif extracted_party_joining_date_from_report is None:
            is_party_joining_date_mismatch = True
            issues_list.append((index, excel_case_code, excel_person_code, "V2入党时间与BF2立案报告不一致"))
            logger.info(f"行 {index + 1} - 入党时间不匹配: Excel入党时间 ('{excel_party_joining_date}') 有值，但立案报告中未提取到。")
            print(f"行 {index + 1} - 入党时间不匹配: Excel入党时间 ('{excel_party_joining_date}') 有值，但立案报告中未提取到。")
        elif excel_party_joining_date != extracted_party_joining_date_from_report:
            is_party_joining_date_mismatch = True
            issues_list.append((index, excel_case_code, excel_person_code, "V2入党时间与BF2立案报告不一致"))
            logger.info(f"行 {index + 1} - 入党时间不匹配: Excel入党时间 ('{excel_party_joining_date}') vs 立案报告提取 ('{extracted_party_joining_date_from_report}')。")
            print(f"行 {index + 1} - 入党时间不匹配: Excel入党时间 ('{excel_party_joining_date}') vs 立案报告提取 ('{extracted_party_joining_date_from_report}')。")
    elif excel_party_member == "否":
        if excel_party_joining_date:
            is_party_joining_date_mismatch = True
            issues_list.append((index, excel_case_code, excel_person_code, "V2入党时间与BF2立案报告不一致"))
            logger.info(f"行 {index + 1} - 入党时间不匹配: Excel是否中共党员为“否”，但入党时间字段不为空 ('{excel_party_joining_date}')。")
            print(f"行 {index + 1} - 入党时间不匹配: Excel是否中共党员为“否”，但入党时间字段不为空 ('{excel_party_joining_date}')。")

    if is_party_joining_date_mismatch:
        party_joining_date_mismatch_indices.add(index)
