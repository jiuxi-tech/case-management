# 包含党员身份和入党时间提取相关的函数

import re
import logging

logger = logging.getLogger(__name__)

def extract_party_member_from_case_report(report_text):
    """
    从立案报告中提取是否为中共党员。
    若报告中存在“加入中国共产党”则返回“是”，否则返回“否”。
    """
    if not report_text or not isinstance(report_text, str):
        msg = f"extract_party_member_from_case_report: report_text 为空或无效: {report_text}"
        logger.info(msg)
        print(msg)
        return None

    if re.search(r"加入中国共产党", report_text, re.IGNORECASE | re.DOTALL):
        msg = f"提取是否中共党员 (立案报告): '是' (找到 '加入中国共产党') from text: {report_text[:100]}..."
        logger.info(msg)
        print(msg)
        return "是"
    else:
        msg = f"提取是否中共党员 (立案报告): '否' (未找到 '加入中国共产党') from text: {report_text[:100]}..."
        logger.info(msg)
        print(msg)
        return "否"

def extract_party_member_from_decision_report(decision_text):
    """
    从处分决定中提取是否为中共党员。
    若报告中存在“加入中国共产党”则返回“是”，若存在“群众”则返回“否”，否则返回 None。
    """
    if not decision_text or not isinstance(decision_text, str):
        msg = f"extract_party_member_from_decision_report: decision_text 为空或无效: {decision_text}"
        logger.info(msg)
        print(msg)
        return None

    if re.search(r"加入中国共产党", decision_text, re.IGNORECASE | re.DOTALL):
        msg = f"提取是否中共党员 (处分决定): '是' (找到 '加入中国共产党') from text: {decision_text[:100]}..."
        logger.info(msg)
        print(msg)
        return "是"
    elif re.search(r"群众", decision_text, re.IGNORECASE | re.DOTALL):
        msg = f"提取是否中共党员 (处分决定): '否' (找到 '群众') from text: {decision_text[:100]}..."
        logger.info(msg)
        print(msg)
        return "否"
    else:
        msg = f"提取是否中共党员 (处分决定): 未明确找到党员或群众信息 from text: {decision_text[:100]}..."
        logger.info(msg)
        print(msg)
        return None

def extract_party_joining_date_from_case_report(report_text):
    """
    从立案报告中提取入党时间，并格式化为“YYYY/MM”。
    例如：“1990年1月加入中国共产党”中的“1990年1月”。
    """
    if not report_text or not isinstance(report_text, str):
        msg = f"extract_party_joining_date_from_case_report: report_text 为空或无效: {report_text}"
        logger.info(msg)
        print(msg)
        return None

    pattern = r"(\d{4})年(\d{1,2})月加入中国共产党"
    match = re.search(pattern, report_text)

    if match:
        year = match.group(1)
        month = match.group(2).zfill(2)
        formatted_date = f"{year}/{month}"
        msg = f"提取入党时间 (立案报告): '{formatted_date}' from text: '{report_text[:100]}...'"
        logger.info(msg)
        print(msg)
        return formatted_date
    else:
        msg = f"在立案报告中未找到“加入中国共产党”及其前面的入党时间信息: {report_text[:100]}..."
        logger.info(msg)
        print(msg)
        return None
