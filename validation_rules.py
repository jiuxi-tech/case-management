import logging
import pandas as pd
import re
from datetime import datetime
from config import Config
from db_utils import get_db

logger = logging.getLogger(__name__)

def validate_agency(agency, authority, db_dict):
    return not authority or not agency or (authority, agency) not in db_dict

def validate_name(reported_person, report_name):
    return reported_person and report_name and reported_person != report_name

def validate_acceptance_time(acceptance_time):
    return pd.notna(acceptance_time)

def validate_organization_measure(measure, report_text, measures_list):
    return not measure or measure not in measures_list or (measure and measure not in report_text)

def extract_name_from_report(report_text):
    if not report_text or pd.isna(report_text):
        logger.info(f"report_text 为空或无效: {report_text}")
        return None
    start_marker = "（一）被反映人基本情况"
    start_idx = report_text.find(start_marker)
    if start_idx == -1:
        logger.warning(f"未找到 '（一）被反映人基本情况' 标记: {report_text}")
        return None
    start_idx += len(start_marker)
    logger.debug(f"原始报告文本: {report_text}")
    next_newline_idx = report_text.find("\n", start_idx)
    if next_newline_idx == -1:
        next_newline_idx = len(report_text)
    paragraph = report_text[start_idx:next_newline_idx].strip()
    if not paragraph:
        next_newline_idx = report_text.find("\n", start_idx)
        if next_newline_idx == -1:
            next_newline_idx = len(report_text)
        next_paragraph_start = report_text.find("\n", next_newline_idx + 1)
        if next_paragraph_start == -1:
            next_paragraph_start = len(report_text)
        paragraph = report_text[next_newline_idx + 1:next_paragraph_start].strip()
    logger.debug(f"提取的段落: {paragraph}")
    end_idx = -1
    for char in [",", "，"]:
        temp_idx = paragraph.find(char)
        if temp_idx != -1 and (end_idx == -1 or temp_idx < end_idx):
            end_idx = temp_idx
    if end_idx == -1:
        logger.warning(f"未找到逗号: {paragraph}")
        return None
    name = paragraph[:end_idx].strip()
    logger.info(f"提取姓名: {name} from paragraph: {paragraph}")
    return name if name else None

def parse_date(date_str):
    """解析多种日期格式，区分年月和完整日期"""
    if not date_str or pd.isna(date_str):
        logger.info(f"日期字符串为空或无效: {date_str}")
        print(f"日期字符串为空或无效: {date_str}")
        return None
    date_str = str(date_str).strip()
    # 支持的格式：YYYY/MM/DD, YYYY-MM-DD, YYYY/M/DD, YYYY-M-DD, YYYY年M月, YYYY/M, YYYY-M
    patterns = [
        r'(\d{4})[/-](\d{1,2})[/-](\d{1,2})',  # YYYY/MM/DD, YYYY-MM-DD
        r'(\d{4})[/-](\d{1,2})',               # YYYY/MM, YYYY-MM
        r'(\d{4})年(\d{1,2})月'               # YYYY年M月
    ]
    for pattern in patterns:
        match = re.search(pattern, date_str)
        if match:
            year = match.group(1)
            month = match.group(2).zfill(2)  # 确保月份为两位
            if len(match.groups()) == 3:  # 包含日，解析为 YYYY-MM-DD
                day = match.group(3).zfill(2)  # 确保日为两位
                parsed_date = f"{year}-{month}-{day}"
            else:  # 仅年月，解析为 YYYY-MM
                parsed_date = f"{year}-{month}"
            logger.info(f"解析日期 {date_str} 为 {parsed_date}")
            print(f"解析日期 {date_str} 为 {parsed_date}")
            return parsed_date
    logger.warning(f"无法解析日期格式: {date_str}")
    print(f"无法解析日期格式: {date_str}")
    return None

def extract_joining_party_time(report_text):
    """从处置情况报告中提取加入中国共产党的时间"""
    if not report_text or pd.isna(report_text):
        logger.info(f"report_text 为空或无效: {report_text}")
        print(f"report_text 为空或无效: {report_text}")
        return None
    pattern = r'(\d{4}年\d{1,2}月|\d{4}[/-]\d{1,2}(?:[/-]\d{1,2})?)\s*加入中国共产党'
    match = re.search(pattern, str(report_text))
    if match:
        date_str = match.group(1)
        parsed_time = parse_date(date_str)
        logger.info(f"从报告中提取 '加入中国共产党' 时间: {date_str} -> {parsed_time}")
        print(f"从报告中提取 '加入中国共产党' 时间: {date_str} -> {parsed_time}")
        return parsed_time
    logger.warning(f"未找到 '加入中国共产党' 相关时间: {report_text[:500]}...")
    print(f"未找到 '加入中国共产党' 相关时间: {report_text[:500]}...")
    return None

def validate_joining_party_time(joining_party_time, report_text):
    """验证入党时间与报告中时间是否一致，考虑完整日期"""
    report_time = extract_joining_party_time(report_text)
    if not joining_party_time or not report_time:
        logger.info(f"入党时间或报告时间为空: joining_party_time={joining_party_time}, report_time={report_time}")
        print(f"入党时间或报告时间为空: joining_party_time={joining_party_time}, report_time={report_time}")
        return False
    parsed_joining_time = parse_date(joining_party_time)
    logger.info(f"比较入党时间: {joining_party_time} -> {parsed_joining_time} vs 报告时间: {report_time}")
    print(f"比较入党时间: {joining_party_time} -> {parsed_joining_time} vs 报告时间: {report_time}")
    # 若入党时间包含完整日期（YYYY-MM-DD），需完全匹配
    if '-' in parsed_joining_time and len(parsed_joining_time.split('-')) == 3:
        return parsed_joining_time == report_time
    # 若仅年月，比较年月
    return parsed_joining_time == report_time

def get_validation_issues(df):
    mismatch_indices = set()
    issues_list = []
    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute('SELECT authority, agency FROM authority_agency_dict WHERE category = ?', ('NSL',))
        db_records = cursor.fetchall()
        db_dict = {(str(row['authority']).strip().lower(), str(row['agency']).strip().lower()) for row in db_records}

    for index, row in df.iterrows():
        logger.debug(f"处理行 {index + 1}")
        # 填报单位名称 vs 办理机关
        agency = str(row["填报单位名称"]).strip().lower() if pd.notna(row["填报单位名称"]) else ''
        authority = str(row["办理机关"]).strip().lower() if pd.notna(row["办理机关"]) else ''
        if validate_agency(agency, authority, db_dict):
            mismatch_indices.add(index)
            issues_list.append((index, Config.VALIDATION_RULES["inconsistent_agency"]))
            logger.info(f"行 {index + 1} - 填报单位名称与办理机关不一致")

        # 被反映人 vs 处置情况报告
        reported_person = str(row["被反映人"]).strip() if pd.notna(row["被反映人"]) else ''
        report_text = row["处置情况报告"]
        report_name = extract_name_from_report(report_text)
        logger.info(f"行 {index + 1} - 被反映人: {reported_person}, 提取姓名: {report_name}")
        if pd.isna(report_text):
            mismatch_indices.add(index)
            issues_list.append((index, Config.VALIDATION_RULES["empty_report"]))
            logger.info(f"行 {index + 1} - 处置情况报告为空")
        elif validate_name(reported_person, report_name):
            mismatch_indices.add(index)
            issues_list.append((index, Config.VALIDATION_RULES["inconsistent_name"]))
            logger.info(f"行 {index + 1} - 被反映人与处置情况报告姓名不一致")

        # 受理时间标记
        if Config.COLUMN_MAPPINGS["acceptance_time"] in df.columns and validate_acceptance_time(row[Config.COLUMN_MAPPINGS["acceptance_time"]]):
            mismatch_indices.add(index)
            issues_list.append((index, Config.VALIDATION_RULES["confirm_acceptance_time"]))
            logger.info(f"行 {index + 1} - 受理时间需确认")

        # 组织措施 vs 处置情况报告
        organization_measure = str(row[Config.COLUMN_MAPPINGS["organization_measure"]]).strip() if pd.notna(row[Config.COLUMN_MAPPINGS["organization_measure"]]) else ''
        report_text = str(report_text).strip() if pd.notna(report_text) else ''
        logger.debug(f"行 {index + 1} - 组织措施: {organization_measure}, 处置情况报告: {report_text[:500]}...")
        if validate_organization_measure(organization_measure, report_text, Config.ORGANIZATION_MEASURES):
            mismatch_indices.add(index)
            issues_list.append((index, Config.VALIDATION_RULES["confirm_organization_measure"]))
            logger.info(f"行 {index + 1} - 组织措施与处置情况报告不一致: {organization_measure}, 原因: 为空或不在关键词列表或未在报告中找到")

        # 入党时间 vs 处置情况报告
        joining_party_time = str(row[Config.COLUMN_MAPPINGS["joining_party_time"]]).strip() if pd.notna(row[Config.COLUMN_MAPPINGS["joining_party_time"]]) else ''
        if not pd.isna(report_text) and not validate_joining_party_time(joining_party_time, report_text):
            mismatch_indices.add(index)
            issues_list.append((index, Config.VALIDATION_RULES["inconsistent_joining_party_time"]))
            logger.info(f"行 {index + 1} - 入党时间与处置情况报告不一致: {joining_party_time}")

    return mismatch_indices, issues_list