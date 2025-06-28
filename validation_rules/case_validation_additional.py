import logging
import pandas as pd
from validation_rules.case_extractors_names import (
    extract_name_from_case_report,
    extract_name_from_decision,
    extract_name_from_trial_report
)

logger = logging.getLogger(__name__)

def validate_name_rules(row, index, excel_case_code, excel_person_code, issues_list, mismatch_indices,
                       investigated_person, report_text_raw, decision_text_raw, investigation_text_raw, trial_text_raw):
    """验证姓名相关规则。"""
    
    report_name = extract_name_from_case_report(report_text_raw)
    if report_name and investigated_person != report_name:
        mismatch_indices.add(index)
        issues_list.append((index, excel_case_code, excel_person_code, "C2被调查人与BF2立案报告不一致"))
        logger.info(f"行 {index + 1} - 姓名不匹配: C2被调查人 ('{investigated_person}') vs BF2立案报告 ('{report_name}')")
        print(f"行 {index + 1} - 姓名不匹配: C2被调查人 ('{investigated_person}') vs BF2立案报告 ('{report_name}')")

    decision_name = extract_name_from_decision(decision_text_raw)
    if not decision_name or (decision_name and investigated_person != decision_name):
        mismatch_indices.add(index)
        issues_list.append((index, excel_case_code, excel_person_code, "C2被调查人与CU2处分决定不一致"))
        logger.info(f"行 {index + 1} - 姓名不匹配: C2被调查人 ('{investigated_person}') vs CU2处分决定 ('{decision_name}')")
        print(f"行 {index + 1} - 姓名不匹配: C2被调查人 ('{investigated_person}') vs CU2处分决定 ('{decision_name}')")

    investigation_name = extract_name_from_case_report(investigation_text_raw)
    if investigation_name and investigated_person != investigation_name:
        mismatch_indices.add(index)
        issues_list.append((index, excel_case_code, excel_person_code, "C2被调查人与CX2审查调查报告不一致"))
        logger.info(f"行 {index + 1} - 姓名不匹配: C2被调查人 ('{investigated_person}') vs CX2审查调查报告 ('{investigation_name}')")
        print(f"行 {index + 1} - 姓名不匹配: C2被调查人 ('{investigated_person}') vs CX2审查调查报告 ('{investigation_name}')")

    trial_name = extract_name_from_trial_report(trial_text_raw)
    if not trial_name or (trial_name and investigated_person != trial_name):
        mismatch_indices.add(index)
        issues_list.append((index, excel_case_code, excel_person_code, "C2被调查人与CY2审理报告不一致"))
        logger.info(f"行 {index + 1} - 姓名不匹配: C2被调查人 ('{investigated_person}') vs CY2审理报告 ('{trial_name}')")
        print(f"行 {index + 1} - 姓名不匹配: C2被调查人 ('{investigated_person}') vs CY2审理报告 ('{trial_name}')")

def validate_case_report_keywords_rules(row, index, excel_case_code, excel_person_code, issues_list, case_report_keyword_mismatch_indices,
                                       case_report_keywords_to_check, report_text_raw, decision_text_raw, investigation_text_raw, trial_text_raw):
    """验证立案报告关键字规则。"""
    
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
                                       excel_voluntary_confession, trial_text_raw):
    """验证是否主动交代问题规则。"""
    
    trial_report_contains_confession = "主动交代" in trial_text_raw

    logger.info(f"行 {index + 1} - 字段 '是否主动交代问题' Excel值: '{excel_voluntary_confession}'。审理报告中'主动交代'匹配结果: {trial_report_contains_confession}。")
    print(f"行 {index + 1} - 字段 '是否主动交代问题' Excel值: '{excel_voluntary_confession}'。审理报告中'主动交代'匹配结果: {trial_report_contains_confession}。")

    if trial_report_contains_confession:
        voluntary_confession_highlight_indices.add(index)
        issues_list.append((index, excel_case_code, excel_person_code, "请基于CY审理报告进行人工确认主动交代"))
        logger.warning(f"行 {index + 1} - 规则触发: 审理报告中发现“主动交代”，已标记“是否主动交代问题”字段为黄色并添加问题描述。")
        print(f"行 {index + 1} - 规则触发: 审理报告中发现“主动交代”，已标记“是否主动交代问题”字段为黄色并添加问题描述。")