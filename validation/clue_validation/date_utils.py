import logging
import re
import pandas as pd
from config import Config

logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)

def parse_date(date_str):
    if not date_str or pd.isna(date_str):
        msg = f"日期字符串为空或无效: {date_str}"
        logger.info(msg)
        print(msg)
        return None
    date_str = str(date_str).strip()
    patterns = [
        r'(\d{4})[/-](\d{1,2})[/-](\d{1,2})',
        r'(\d{4})[/-](\d{1,2})',
        r'(\d{4})年(\d{1,2})月'
    ]
    for pattern in patterns:
        match = re.search(pattern, date_str)
        if match:
            year = match.group(1)
            month = match.group(2).zfill(2)
            if len(match.groups()) == 3:
                day = match.group(3).zfill(2)
                parsed_date = f"{year}-{month}-{day}"
            else:
                parsed_date = f"{year}-{month}"
            msg = f"解析日期 {date_str} 为 {parsed_date}"
            logger.info(msg)
            print(msg)
            return parsed_date
    msg = f"无法解析日期格式: {date_str}"
    logger.warning(msg)
    print(msg)
    return None

def extract_joining_party_time(report_text):
    if not report_text or pd.isna(report_text):
        msg = f"report_text 为空或无效: {report_text}"
        logger.info(msg)
        print(msg)
        return None
    pattern = r'(\d{4}年\d{1,2}月|\d{4}[/-]\d{1,2}(?:[/-]\d{1,2})?)\s*加入中国共产党'
    match = re.search(pattern, str(report_text))
    if match:
        date_str = match.group(1)
        parsed_time = parse_date(date_str)
        msg = f"从报告中提取 '加入中国共产党' 时间: {date_str} -> {parsed_time}"
        logger.info(msg)
        print(msg)
        return parsed_time
    msg = f"未找到 '加入中国共产党' 相关时间: {report_text[:500]}..."
    logger.warning(msg)
    print(msg)
    return None

def validate_joining_party_time(joining_party_time, report_text):
    report_time = extract_joining_party_time(report_text)
    if not joining_party_time or not report_time:
        msg = f"入党时间或报告时间为空: joining_party_time={joining_party_time}, report_time={report_time}"
        logger.info(msg)
        print(msg)
        return False
    parsed_joining_time = parse_date(joining_party_time)
    msg = f"比较入党时间: {joining_party_time} -> {parsed_joining_time} vs 报告时间: {report_time}"
    logger.info(msg)
    print(msg)
    if '-' in parsed_joining_time and len(parsed_joining_time.split('-')) == 3:
        return parsed_joining_time == report_time
    return parsed_joining_time == report_time