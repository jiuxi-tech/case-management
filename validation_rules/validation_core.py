import logging
import pandas as pd
import re
from config import Config
from db_utils import get_db
from validation_rules.name_extraction import extract_name_from_report

# 初始化 logger
logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)

try:
    from validation_rules.date_utils import validate_joining_party_time
    msg = "Successfully imported validate_joining_party_time from date_utils"
    logger.debug(msg)
    print(msg)
except ImportError as e:
    msg = f"Failed to import validate_joining_party_time: {str(e)}"
    logger.error(msg)
    print(msg)

def normalize_date(date_str):
    """标准化日期格式为 YYYY-MM"""
    if pd.isna(date_str):
        return None
    date_str = str(date_str).strip()
    msg = f"标准化日期: 原始 '{date_str}'"
    logger.debug(msg)
    print(msg)
    # 匹配 YYYY年M月 或 YYYY/M 格式
    match = re.match(r'(\d{4})[年/](\d{1,2})[月]?', date_str)
    if match:
        year, month = match.groups()
        normalized = f"{year}-{month.zfill(2)}"
        msg = f"标准化结果: {normalized}"
        logger.debug(msg)
        print(msg)
        return normalized
    return date_str

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

def validate_confiscation_amount(report_text):
    """检查处置情况报告是否包含'没收'，返回是否需要高亮没收金额"""
    if pd.isna(report_text):
        return False
    return "没收" in str(report_text)

def validate_compensation_amount(report_text):
    """检查处置情况报告是否包含'责令退赔'，返回是否需要高亮责令退赔金额"""
    if pd.isna(report_text):
        return False
    return "责令退赔" in str(report_text)

def validate_registration_amount(report_text):
    """检查处置情况报告是否包含'登记上交金额'，返回是否需要高亮登记上交金额"""
    if pd.isna(report_text):
        return False
    return "登记上交金额" in str(report_text)

def validate_recovery_amount(report_text):
    """检查处置情况报告是否包含'追缴'，返回是否需要高亮追缴失职渎职滥用职权造成的损失金额"""
    if pd.isna(report_text):
        return None
    return "追缴" in str(report_text)

def extract_ethnicity_from_report(report_text):
    """从处置情况报告中提取民族，基于'（二）相关人员基本情况'段落的逗号规则"""
    if not report_text or pd.isna(report_text):
        msg = f"report_text 为空或无效: {report_text}"
        logger.info(msg)
        print(msg)
        return None
    start_marker = "（二）相关人员基本情况"
    start_idx = report_text.find(start_marker)
    if start_idx == -1:
        msg = f"未找到 '（二）相关人员基本情况' 标记: {report_text}"
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
    if not paragraph:
        msg = f"段落为空: {paragraph}"
        logger.warning(msg)
        print(msg)
        return None
    # 找到第二个和第三个逗号之间的内容
    commas = [i for i, char in enumerate(paragraph) if char in [",", "，"]]
    msg = f"逗号位置: {commas}"
    logger.debug(msg)
    print(msg)
    if len(commas) < 3:
        msg = f"逗号数量少于3: {paragraph}"
        logger.warning(msg)
        print(msg)
        return None
    start_idx = commas[1] + 1
    end_idx = commas[2]
    ethnicity = paragraph[start_idx:end_idx].strip()
    msg = f"提取民族: {ethnicity} from paragraph: {paragraph}"
    logger.info(msg)
    print(msg)
    return ethnicity if ethnicity else None

def validate_ethnicity(ethnicity, report_text):
    """验证民族字段与处置情况报告中提取的民族是否一致"""
    report_ethnicity = extract_ethnicity_from_report(report_text)
    msg = f"民族字段值: '{ethnicity}', 报告中提取的民族: '{report_ethnicity}'"
    logger.debug(msg)
    print(msg)
    if pd.isna(ethnicity) or not ethnicity or not report_ethnicity:
        msg = "民族字段为空或报告中无有效民族，视为不一致"
        logger.debug(msg)
        print(msg)
        return True  # 为空或不存在视为不一致
    result = str(ethnicity).strip() != report_ethnicity
    msg = f"比较结果: {result}"
    logger.debug(msg)
    print(msg)
    return result

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
        normalized_jt = normalize_date(joining_party_time)
        report_jt = None
        if not pd.isna(report_text):
            # 修正正则表达式，匹配加入中国共产党前的紧邻日期
            join_match = re.search(r'(\d{4}年\d{1,2}月)(?:\s*,?\s*加入中国共产党)', report_text)
            if join_match:
                report_jt = normalize_date(join_match.group(1))
        msg = f"比较入党时间: {joining_party_time} -> {normalized_jt} vs 报告时间: {report_jt}"
        logger.debug(msg)
        print(msg)
        if normalized_jt and report_jt and normalized_jt != report_jt:
            issues_list.append((index, Config.VALIDATION_RULES["inconsistent_joining_party_time"]))
            msg = f"行 {index + 1} - {Config.VALIDATION_RULES['inconsistent_joining_party_time']}: {joining_party_time}"
            logger.info(msg)
            print(msg)

        # 检查“收缴金额”相关逻辑
        if "收缴金额（万元）" in df.columns and validate_collection_amount(report_text):
            issues_list.append((index, Config.VALIDATION_RULES["highlight_collection_amount"]))

        # 检查“没收金额”相关逻辑
        if "没收金额" in df.columns and validate_confiscation_amount(report_text):
            issues_list.append((index, Config.VALIDATION_RULES["highlight_confiscation_amount"]))

        # 检查“责令退赔金额”相关逻辑
        if "责令退赔金额" in df.columns and validate_compensation_amount(report_text):
            issues_list.append((index, Config.VALIDATION_RULES["highlight_compensation_amount"]))

        # 检查“登记上交金额”相关逻辑
        if "登记上交金额" in df.columns and validate_registration_amount(report_text):
            issues_list.append((index, Config.VALIDATION_RULES["highlight_registration_amount"]))

        # 检查“追缴失职渎职滥用职权造成的损失金额”相关逻辑
        if "追缴失职渎职滥用职权造成的损失金额" in df.columns and validate_recovery_amount(report_text):
            issues_list.append((index, Config.VALIDATION_RULES["highlight_recovery_amount"]))

        # 检查“民族”相关逻辑
        if Config.COLUMN_MAPPINGS["ethnicity"] in df.columns:
            ethnicity = row[Config.COLUMN_MAPPINGS["ethnicity"]]
            if validate_ethnicity(ethnicity, report_text):
                issues_list.append((index, Config.VALIDATION_RULES["inconsistent_ethnicity"]))
                msg = f"行 {index + 1} - 民族不一致: 字段值 '{ethnicity}', 报告提取值 '{extract_ethnicity_from_report(report_text)}'"
                logger.info(msg)
                print(msg)
        msg = f"行 {index + 1} issues_list: {[(i, issue) for i, issue in issues_list if i == index]}"
        logger.debug(msg)
        print(msg)

    return mismatch_indices, issues_list