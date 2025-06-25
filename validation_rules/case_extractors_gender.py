##性别提取相关的函数

import re
import logging

logger = logging.getLogger(__name__)

def extract_gender_from_case_report(report_text):
    """
    从立案报告中提取性别。
    性别位于“一、XXX同志基本情况”后第一个和第二个逗号之间。
    例如：“王xx，男，汉族”中的“男”。
    """
    if not report_text or not isinstance(report_text, str):
        msg = f"extract_gender_from_case_report: report_text 为空或无效: {report_text}"
        logger.info(msg)
        print(msg)
        return None
    
    pattern = r"一、.+?同志基本情况.*?，([^，]+)，"
    match = re.search(pattern, report_text, re.DOTALL)
    if match:
        gender = match.group(1).strip()
        msg = f"提取性别 (立案报告): {gender} from case report"
        logger.info(msg)
        print(msg)
        return gender
    else:
        msg = f"未找到性别信息 in case report: {report_text[:100]}..."
        logger.warning(msg)
        print(msg)
        return None

def extract_gender_from_decision_report(decision_text):
    """
    从处分决定中提取性别。
    性别位于“关于给予...同志党内警告处分的决定”后的段落里，第一个和第二个逗号之间。
    例如：“王xx，男，汉族”中的“男”。
    """
    if not decision_text or not isinstance(decision_text, str):
        msg = f"extract_gender_from_decision_report: decision_text 为空或无效: {decision_text}"
        logger.info(msg)
        print(msg)
        return None
    
    title_pattern = r"关于给予.+?同志党内警告处分的决定"
    title_match = re.search(title_pattern, decision_text, re.DOTALL)

    if title_match:
        start_pos = title_match.end()
        gender_pattern = r".*?，([^，]+)，"
        search_area = decision_text[start_pos : start_pos + 200] 
        gender_match = re.search(gender_pattern, search_area, re.DOTALL)

        if gender_match:
            gender = gender_match.group(1).strip()
            msg = f"提取性别 (处分决定): {gender} from decision report"
            logger.info(msg)
            print(msg)
            return gender
        else:
            msg = f"在处分决定标题后未找到性别信息: {search_area[:50]}..."
            logger.warning(msg)
            print(msg)
            return None
    else:
        msg = f"未找到 '关于给予...同志党内警告处分的决定' 标记，无法提取性别: {decision_text[:100]}..."
        logger.warning(msg)
        print(msg)
        return None

def extract_gender_from_investigation_report(investigation_text):
    """
    从审查调查报告中提取性别。
    性别位于“一、XXX同志基本情况”后第一个和第二个逗号之间。
    例如：“王xx，男，汉族”中的“男”。
    """
    if not investigation_text or not isinstance(investigation_text, str):
        msg = f"extract_gender_from_investigation_report: investigation_text 为空或无效: {investigation_text}"
        logger.info(msg)
        print(msg)
        return None
    
    pattern = r"一、.+?同志基本情况.*?，([^，]+)，"
    match = re.search(pattern, investigation_text, re.DOTALL)
    if match:
        gender = match.group(1).strip()
        msg = f"提取性别 (审查调查报告): {gender} from investigation report"
        logger.info(msg)
        print(msg)
        return gender
    else:
        msg = f"未找到性别信息 in investigation report: {investigation_text[:100]}..."
        logger.warning(msg)
        print(msg)
        return None

def extract_gender_from_trial_report(trial_text):
    """
    从审理报告中提取性别。
    性别位于“现将具体情况报告如下”这样字符串下面段落里第一个逗号和第二个逗号中间的位置。
    例如：“王xx，男，汉族”中的“男”。
    """
    if not trial_text or not isinstance(trial_text, str):
        msg = f"extract_gender_from_trial_report: trial_text 为空或无效: {trial_text}"
        logger.info(msg)
        print(msg)
        return None

    title_marker = "现将具体情况报告如下"
    marker_pos = trial_text.find(title_marker)

    if marker_pos != -1:
        start_pos = marker_pos + len(title_marker)
        gender_pattern = r".*?，([^，]+)，"
        search_area = trial_text[start_pos : start_pos + 200]
        gender_match = re.search(gender_pattern, search_area, re.DOTALL)

        if gender_match:
            gender = gender_match.group(1).strip()
            msg = f"提取性别 (审理报告): {gender} from trial report"
            logger.info(msg)
            print(msg)
            return gender
        else:
            msg = f"在审理报告中 '现将具体情况报告如下' 后未找到性别信息: {search_area[:50]}..."
            logger.warning(msg)
            print(msg)
            return None
    else:
        msg = f"未找到 '现将具体情况报告如下' 标记，无法提取审理报告性别: {trial_text[:100]}..."
        logger.warning(msg)
        print(msg)
        return None
