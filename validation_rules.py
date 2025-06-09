import logging
import pandas as pd
import re
from datetime import datetime
from config import Config
from db_utils import get_db

logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)  # 确保调试日志输出

# 定义组织措施的精准匹配关键词
ORGANIZATION_MEASURE_KEYWORDS = [
    "谈话提醒", "提醒谈话", "批评教育", "责令检查", "责令其做出书面检查",
    "责令其做出检查", "诫勉", "警示谈话", "通报批评", "责令公开道歉（检查）",
    "责令具结悔过"
]

def validate_agency(authority, agency, db_dict):
    if not authority or not agency:
        msg = "空值检测: authority 或 agency 为空"
        logger.debug(msg)
        print(msg)
        return True  # 空值视为不匹配
    norm_authority = ''.join(authority.split()).lower()
    norm_agency = ''.join(agency.split()).lower()
    key = (norm_authority, norm_agency)
    msg = f"Matching rule: (authority, agency) = ({norm_authority}, {norm_agency}) must be in {db_dict}"
    print(msg)
    logger.debug(msg)
    msg = f"校验: key = {key}, 是否在 db_dict: {key in db_dict}"
    print(msg)
    logger.debug(msg)
    return key not in db_dict

def validate_name(reported_person, report_name):
    return reported_person and report_name and reported_person != report_name

def validate_acceptance_time(acceptance_time):
    return pd.notna(acceptance_time)

def validate_organization_measure(measure, report_text, measures_list):
    if not measure or pd.isna(measure):
        msg = f"组织措施为空或无效: {measure}"
        logger.info(msg)
        print(msg)
        return True
    report_text = str(report_text).strip() if pd.notna(report_text) else ''
    # 按回车分割多行
    lines = measure.split('\n')
    has_mismatch = False
    for line in lines:
        line = line.strip()
        if not line:
            continue
        # 移除标号（如 "1、"、"2、"）及其后的空格
        cleaned_line = re.sub(r'^\d+[、]\s*', '', line)
        msg = f"处理组织措施行: 原始 '{line}' -> 清理后 '{cleaned_line}'"
        logger.debug(msg)
        print(msg)
        # 如果清理后为空，跳过此行
        if not cleaned_line:
            continue
        # 精准匹配：必须在 ORGANIZATION_MEASURE_KEYWORDS 中且在 report_text 中找到
        if cleaned_line not in ORGANIZATION_MEASURE_KEYWORDS or cleaned_line not in report_text:
            msg = f"组织措施 '{cleaned_line}' 不匹配: 不在关键词列表或未在报告中找到"
            logger.info(msg)
            print(msg)
            has_mismatch = True
    return has_mismatch

def extract_name_from_report(report_text):
    if not report_text or pd.isna(report_text):
        msg = f"report_text 为空或无效: {report_text}"
        logger.info(msg)
        print(msg)
        return None
    start_marker = "（一）被反映人基本情况"
    start_idx = report_text.find(start_marker)
    if start_idx == -1:
        msg = f"未找到 '（一）被反映人基本情况' 标记: {report_text}"
        logger.warning(msg)
        print(msg)
        return None
    start_idx += len(start_marker)
    msg = f"原始报告文本: {report_text}"
    logger.debug(msg)
    print(msg)
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
    msg = f"提取的段落: {paragraph}"
    logger.debug(msg)
    print(msg)
    end_idx = -1
    for char in [",", "，"]:
        temp_idx = paragraph.find(char)
        if temp_idx != -1 and (end_idx == -1 or temp_idx < end_idx):
            end_idx = temp_idx
    if end_idx == -1:
        msg = f"未找到逗号: {paragraph}"
        logger.warning(msg)
        print(msg)
        return None
    name = paragraph[:end_idx].strip()
    msg = f"提取姓名: {name} from paragraph: {paragraph}"
    logger.info(msg)
    print(msg)
    return name if name else None

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

def get_validation_issues(df):
    mismatch_indices = set()  # 仅用于 agency 不匹配的索引
    issues_list = []  # 记录所有验证问题
    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute('SELECT authority, agency FROM authority_agency_dict WHERE category = ?', ('NSL',))
        db_records = cursor.fetchall()
        db_dict = {(str(row['authority']).strip().lower(), str(row['agency']).strip().lower()) for row in db_records}
        msg = f"Valid NSL records: {db_dict}"
        logger.info(msg)
        print(msg)

    for index, row in df.iterrows():
        msg = f"处理行 {index + 1}"
        logger.debug(msg)
        print(msg)
        authority = str(row["办理机关"]).strip().lower() if pd.notna(row["办理机关"]) else ''
        agency = str(row["填报单位名称"]).strip().lower() if pd.notna(row["填报单位名称"]) else ''
        msg = f"传入 validate_agency: authority={authority}, agency={agency}"
        logger.debug(msg)
        print(msg)
        if validate_agency(authority, agency, db_dict):  # 仅当 agency 不匹配时标红
            mismatch_indices.add(index)
            issues_list.append((index, Config.VALIDATION_RULES["inconsistent_agency"]))
            msg = f"行 {index + 1} - 办理机关: {authority}, 填报单位名称: {agency} 不匹配 NSL 记录"
            logger.info(msg)
            print(msg)
        else:
            msg = f"行 {index + 1} - 办理机关: {authority}, 填报单位名称: {agency} 匹配成功"
            logger.info(msg)
            print(msg)

        # 其他验证规则，仅记录到 issues_list，不影响 mismatch_indices
        reported_person = str(row["被反映人"]).strip() if pd.notna(row["被反映人"]) else ''
        report_text = row["处置情况报告"]
        report_name = extract_name_from_report(report_text)
        msg = f"行 {index + 1} - 被反映人: {reported_person}, 提取姓名: {report_name}"
        logger.info(msg)
        print(msg)
        if pd.isna(report_text):
            issues_list.append((index, Config.VALIDATION_RULES["empty_report"]))
            msg = f"行 {index + 1} - 处置情况报告为空"
            logger.info(msg)
            print(msg)
        elif validate_name(reported_person, report_name):
            issues_list.append((index, Config.VALIDATION_RULES["inconsistent_name"]))
            msg = f"行 {index + 1} - 被反映人与处置情况报告姓名不一致"
            logger.info(msg)
            print(msg)

        if Config.COLUMN_MAPPINGS["acceptance_time"] in df.columns and validate_acceptance_time(row[Config.COLUMN_MAPPINGS["acceptance_time"]]):
            issues_list.append((index, Config.VALIDATION_RULES["confirm_acceptance_time"]))
            msg = f"行 {index + 1} - 受理时间需确认"
            logger.info(msg)
            print(msg)

        organization_measure = str(row[Config.COLUMN_MAPPINGS["organization_measure"]]).strip() if pd.notna(row[Config.COLUMN_MAPPINGS["organization_measure"]]) else ''
        report_text = str(report_text).strip() if pd.notna(report_text) else ''
        msg = f"行 {index + 1} - 组织措施: {organization_measure}, 处置情况报告: {report_text[:500]}..."
        logger.debug(msg)
        print(msg)
        if validate_organization_measure(organization_measure, report_text, Config.ORGANIZATION_MEASURES):
            issues_list.append((index, "组织措施跟处置报告不一致"))
            msg = f"行 {index + 1} - 组织措施与处置情况报告不一致: {organization_measure}"
            logger.info(msg)
            print(msg)

        joining_party_time = str(row[Config.COLUMN_MAPPINGS["joining_party_time"]]).strip() if pd.notna(row[Config.COLUMN_MAPPINGS["joining_party_time"]]) else ''
        if not pd.isna(report_text) and not validate_joining_party_time(joining_party_time, report_text):
            issues_list.append((index, Config.VALIDATION_RULES["inconsistent_joining_party_time"]))
            msg = f"行 {index + 1} - 入党时间与处置情况报告不一致: {joining_party_time}"
            logger.info(msg)
            print(msg)

    return mismatch_indices, issues_list