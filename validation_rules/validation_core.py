import logging
import pandas as pd
import re
from config import Config
from db_utils import get_db
from validation_rules.name_extraction import extract_name_from_report
from validation_rules.date_utils import validate_joining_party_time  # 确认导入

logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)

def validate_agency(authority, agency, db_dict):
    if not authority or not agency:
        msg = "空值检测: authority 或 agency 为空"
        logger.debug(msg)
        print(msg)
        return True
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
    return acceptance_time is not None and not pd.isna(acceptance_time)

def validate_organization_measure(measure, report_text):
    if not measure or pd.isna(measure):
        msg = f"组织措施为空或无效: {measure}"
        logger.info(msg)
        print(msg)
        return True
    report_text = str(report_text).strip() if pd.notna(report_text) else ''
    lines = measure.split('\n')
    has_mismatch = False
    for line in lines:
        line = line.strip()
        if not line:
            continue
        cleaned_line = re.sub(r'^\d+[、]\s*', '', line)
        msg = f"处理组织措施行: 原始 '{line}' -> 清理后 '{cleaned_line}'"
        logger.debug(msg)
        print(msg)
        if not cleaned_line:
            continue
        if cleaned_line not in Config.ORGANIZATION_MEASURE_KEYWORDS or cleaned_line not in report_text:
            msg = f"组织措施 '{cleaned_line}' 不匹配: 不在关键词列表或未在报告中找到"
            logger.info(msg)
            print(msg)
            has_mismatch = True
    return has_mismatch

def validate_collection_amount(report_text):
    """检查处置情况报告是否包含'收缴'，返回是否需要高亮收缴金额"""
    if pd.isna(report_text):
        return False
    return "收缴" in str(report_text)

def get_validation_issues(df):
    mismatch_indices = set()
    issues_list = []
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
        if validate_agency(authority, agency, db_dict):
            mismatch_indices.add(index)
            issues_list.append((index, Config.VALIDATION_RULES["inconsistent_agency"]))
            msg = f"行 {index + 1} - 办理机关: {authority}, 填报单位名称: {agency} 不匹配 NSL 记录"
            logger.info(msg)
            print(msg)
        else:
            msg = f"行 {index + 1} - 办理机关: {authority}, 填报单位名称: {agency} 匹配成功"
            logger.info(msg)
            print(msg)

        reported_person = str(row["被反映人"]).strip() if pd.notna(row["被反映人"]) else ''
        report_text = row["处置情况报告"]
        report_name = extract_name_from_report(report_text)
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

        organization_measure = str(row[Config.COLUMN_MAPPINGS["organization_measure"]].strip()) if pd.notna(row[Config.COLUMN_MAPPINGS["organization_measure"]]) else ''
        if validate_organization_measure(organization_measure, report_text):
            issues_list.append((index, Config.VALIDATION_RULES["inconsistent_organization_measure"]))
            msg = f"行 {index + 1} - {Config.VALIDATION_RULES['inconsistent_organization_measure']}: {organization_measure}"
            logger.info(msg)
            print(msg)

        joining_party_time = str(row[Config.COLUMN_MAPPINGS["joining_party_time"]].strip()) if pd.notna(row[Config.COLUMN_MAPPINGS["joining_party_time"]]) else ''
        if not pd.isna(report_text) and not validate_joining_party_time(joining_party_time, report_text):
            issues_list.append((index, Config.VALIDATION_RULES["inconsistent_joining_party_time"]))
            msg = f"行 {index + 1} - {Config.VALIDATION_RULES['inconsistent_joining_party_time']}: {joining_party_time}"
            logger.info(msg)
            print(msg)

        # 新增：检查“收缴金额”相关逻辑
        if "收缴金额（万元）" in df.columns and validate_collection_amount(report_text):
            issues_list.append((index, Config.VALIDATION_RULES["highlight_collection_amount"]))

    return mismatch_indices, issues_list