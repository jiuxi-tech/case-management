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

def normalize_date(date_str, full_date=False):
    """标准化日期格式，full_date=True返回YYYY-MM-DD，否则返回YYYY-MM"""
    if pd.isna(date_str):
        return None
    date_str = str(date_str).strip()
    msg = f"标准化日期: 原始 '{date_str}'"
    logger.debug(msg)
    print(msg)
    # 匹配YYYY/M、YYYY年M月、YYYY年M月D日、YYYY年M月生
    match = re.match(r'(\d{4})[年/-](\d{1,2})(?:[月/-](\d{1,2})[日]?)?(?:生)?', date_str)
    if match:
        year, month, day = match.groups()
        if full_date and day:  # 办结时间需完整日期
            normalized = f"{year}-{month.zfill(2)}-{day.zfill(2)}"
        else:  # 入党时间和出生年月仅需年月
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
    msg = f"校验: key = {key}, 是否在 db_dict: {key in db_dict}"
    logger.debug(msg)
    print(msg)
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
            msg = f"组织措施 '{cleaned_line}' 不匹配"
            logger.info(msg)
            print(msg)
            has_mismatch = True
    return has_mismatch

def validate_collection_amount(report_text):
    return pd.notna(report_text) and "收缴" in str(report_text)

def validate_confiscation_amount(report_text):
    return pd.notna(report_text) and "没收" in str(report_text)

def validate_compensation_amount(report_text):
    return pd.notna(report_text) and "责令退赔" in str(report_text)

def validate_registration_amount(report_text):
    return pd.notna(report_text) and "登记上交金额" in str(report_text)

def validate_recovery_amount(report_text):
    return pd.notna(report_text) and "追缴" in str(report_text)

def extract_ethnicity_from_report(report_text):
    if not report_text or pd.isna(report_text):
        msg = f"report_text 为空或无效: {report_text}"
        logger.info(msg)
        print(msg)
        return None
    start_marker = "（二）相关人员基本情况"
    start_idx = report_text.find(start_marker)
    if start_idx == -1:
        msg = f"未找到 '（二）相关人员基本情况' 标记"
        logger.warning(msg)
        print(msg)
        return None
    start_idx += len(start_marker)
    next_newline_idx = report_text.find("\n", start_idx)
    paragraph = report_text[start_idx:next_newline_idx].strip() if next_newline_idx != -1 else report_text[start_idx:].strip()
    if not paragraph:
        next_newline_idx = report_text.find("\n", start_idx)
        next_paragraph_start = report_text.find("\n", next_newline_idx + 1) if next_newline_idx != -1 else len(report_text)
        paragraph = report_text[next_newline_idx + 1:next_paragraph_start].strip() if next_newline_idx != -1 else ""
    if not paragraph:
        msg = f"段落为空: {paragraph}"
        logger.warning(msg)
        print(msg)
        return None
    commas = [i for i, char in enumerate(paragraph) if char in [",", "，"]]
    if len(commas) < 3:
        msg = f"逗号数量少于3: {paragraph}"
        logger.warning(msg)
        print(msg)
        return None
    start_idx = commas[1] + 1
    end_idx = commas[2]
    ethnicity = paragraph[start_idx:end_idx].strip()
    msg = f"提取民族: {ethnicity}"
    logger.info(msg)
    print(msg)
    return ethnicity if ethnicity else None

def extract_birth_date_from_report(report_text):
    if not report_text or pd.isna(report_text):
        msg = f"report_text 为空或无效: {report_text}"
        logger.info(msg)
        print(msg)
        return None
    start_marker = "（一）被反映人基本情况"
    start_idx = report_text.find(start_marker)
    if start_idx == -1:
        msg = f"未找到 '（一）被反映人基本情况' 标记"
        logger.warning(msg)
        print(msg)
        return None
    start_idx += len(start_marker)
    next_newline_idx = report_text.find("\n", start_idx)
    paragraph = report_text[start_idx:next_newline_idx].strip() if next_newline_idx != -1 else report_text[start_idx:].strip()
    if not paragraph:
        next_newline_idx = report_text.find("\n", start_idx)
        next_paragraph_start = report_text.find("\n", next_newline_idx + 1) if next_newline_idx != -1 else len(report_text)
        paragraph = report_text[next_newline_idx + 1:next_paragraph_start].strip() if next_newline_idx != -1 else ""
    if not paragraph:
        msg = f"段落为空: {paragraph}"
        logger.warning(msg)
        print(msg)
        return None
    commas = [i for i, char in enumerate(paragraph) if char in [",", "，"]]
    if len(commas) < 4:
        msg = f"逗号数量少于4: {paragraph}"
        logger.warning(msg)
        print(msg)
        return None
    start_idx = commas[2] + 1
    end_idx = commas[3] if len(commas) > 3 else len(paragraph)
    birth_date = paragraph[start_idx:end_idx].strip()
    msg = f"提取出生年月: {birth_date}"
    logger.info(msg)
    print(msg)
    return normalize_date(birth_date)

def extract_completion_date_from_report(report_text):
    """从处置情况报告最后一行的日期"""
    if not report_text or pd.isna(report_text):
        msg = f"report_text 为空或无效: {report_text}"
        logger.info(msg)
        print(msg)
        return None
    lines = report_text.strip().split('\n')
    last_line = lines[-1].strip() if lines else ""
    msg = f"最后一行的内容: '{last_line}'"
    logger.debug(msg)
    print(msg)
    match = re.search(r'(\d{4})[年/-](\d{1,2})[月/-](\d{1,2})[日]?', last_line)
    if match:
        year, month, day = match.groups()
        completion_date = f"{year}-{month.zfill(2)}-{day.zfill(2)}"
        msg = f"提取落款日期: {completion_date}"
        logger.debug(msg)
        print(msg)
        return completion_date
    msg = "未找到有效的落款日期"
    logger.warning(msg)
    print(msg)
    return None

def validate_ethnicity(ethnicity, report_text):
    report_ethnicity = extract_ethnicity_from_report(report_text)
    msg = f"民族字段值: '{ethnicity}', 报告中提取的民族: '{report_ethnicity}'"
    logger.debug(msg)
    print(msg)
    if pd.isna(ethnicity) or not ethnicity or not report_ethnicity:
        msg = "民族字段为空或报告中无有效民族，视为不一致"
        logger.debug(msg)
        print(msg)
        return True
    result = str(ethnicity).strip() != report_ethnicity
    msg = f"比较结果: {result}"
    logger.debug(msg)
    print(msg)
    return result

def get_validation_issues(df):
    mismatch_indices = set()
    issues_list = []
    
    # 获取受理线索编码的列名，如果Config中没有定义，则默认为"受理线索编码"
    accepted_clue_code_col = Config.COLUMN_MAPPINGS.get("accepted_clue_code", "受理线索编码")

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
        
        # 获取当前行的受理线索编码
        clue_code = row[accepted_clue_code_col] if accepted_clue_code_col in df.columns else "N/A"
        
        authority = str(row["办理机关"]).strip().lower() if pd.notna(row["办理机关"]) else ''
        agency = str(row["填报单位名称"]).strip().lower() if pd.notna(row["填报单位名称"]) else ''
        if validate_agency(authority, agency, db_dict):
            mismatch_indices.add(index)
            # 在 issues_list 中添加受理线索编码
            issues_list.append((index, clue_code, Config.VALIDATION_RULES["inconsistent_agency"]))
            msg = f"行 {index + 1} - 办理机关: {authority}, 填报单位名称: {agency} 不匹配"
            logger.info(msg)
            print(msg)

        reported_person = str(row["被反映人"]).strip() if pd.notna(row["被反映人"]) else ''
        report_text = row["处置情况报告"] if "处置情况报告" in df.columns else ''
        report_name = extract_name_from_report(report_text)
        if pd.isna(report_text):
            issues_list.append((index, clue_code, Config.VALIDATION_RULES["empty_report"]))
            msg = f"行 {index + 1} - 处置情况报告为空"
            logger.info(msg)
            print(msg)
        elif validate_name(reported_person, report_name):
            issues_list.append((index, clue_code, Config.VALIDATION_RULES["inconsistent_name"]))
            msg = f"行 {index + 1} - 被反映人与报告姓名不一致"
            logger.info(msg)
            print(msg)

        if Config.COLUMN_MAPPINGS["acceptance_time"] in df.columns and validate_acceptance_time(row[Config.COLUMN_MAPPINGS["acceptance_time"]]):
            issues_list.append((index, clue_code, Config.VALIDATION_RULES["confirm_acceptance_time"]))
            msg = f"行 {index + 1} - 受理时间需确认"
            logger.info(msg)
            print(msg)

        organization_measure = str(row[Config.COLUMN_MAPPINGS["organization_measure"]]).strip() if pd.notna(row[Config.COLUMN_MAPPINGS["organization_measure"]]) else ''
        if validate_organization_measure(organization_measure, report_text):
            issues_list.append((index, clue_code, Config.VALIDATION_RULES["inconsistent_organization_measure"]))
            msg = f"行 {index + 1} - 组织措施与报告不一致: {organization_measure}"
            logger.info(msg)
            print(msg)

        joining_party_time = str(row[Config.COLUMN_MAPPINGS["joining_party_time"]]).strip() if pd.notna(row[Config.COLUMN_MAPPINGS["joining_party_time"]]) else ''
        normalized_jt = normalize_date(joining_party_time)
        report_jt = None
        if not pd.isna(report_text):
            join_match = re.search(r'(\d{4}年\d{1,2}月)(?:\s*,?\s*加入中国共产党)', report_text)
            if join_match:
                report_jt = normalize_date(join_match.group(1))
        if normalized_jt and report_jt and normalized_jt != report_jt:
            issues_list.append((index, clue_code, Config.VALIDATION_RULES["inconsistent_joining_party_time"]))
            msg = f"行 {index + 1} - 入党时间与报告不一致: {joining_party_time}"
            logger.info(msg)
            print(msg)

        if Config.COLUMN_MAPPINGS["birth_date"] in df.columns:
            birth_date = str(row[Config.COLUMN_MAPPINGS["birth_date"]]).strip() if pd.notna(row[Config.COLUMN_MAPPINGS["birth_date"]]) else ''
            normalized_bd = normalize_date(birth_date)
            report_bd = extract_birth_date_from_report(report_text)
            if not normalized_bd or not report_bd or normalized_bd != report_bd:
                issues_list.append((index, clue_code, Config.VALIDATION_RULES["highlight_birth_date"]))
                msg = f"行 {index + 1} - 出生年月与报告不一致: {birth_date}"
                logger.info(msg)
                print(msg)

        if Config.COLUMN_MAPPINGS["completion_time"] in df.columns:
            completion_time = str(row[Config.COLUMN_MAPPINGS["completion_time"]]).strip() if pd.notna(row[Config.COLUMN_MAPPINGS["completion_time"]]) else ''
            normalized_ct = normalize_date(completion_time, full_date=True)
            report_ct = extract_completion_date_from_report(report_text)
            msg = f"比较办结时间: {completion_time} -> {normalized_ct} vs 报告落款时间: {report_ct}"
            logger.debug(msg)
            print(msg)
            if normalized_ct and report_ct and normalized_ct != report_ct:
                issues_list.append((index, clue_code, Config.VALIDATION_RULES["highlight_completion_time"]))
                msg = f"行 {index + 1} - 办结时间与报告落款不一致: {completion_time}"
                logger.info(msg)
                print(msg)

        if Config.COLUMN_MAPPINGS["disposal_method_1"] in df.columns:
            disposal_method = row[Config.COLUMN_MAPPINGS["disposal_method_1"]]
            issues_list.append((index, clue_code, Config.VALIDATION_RULES["highlight_disposal_method_1"]))
            msg = f"行 {index + 1} - 处置方式1二级需确认: {disposal_method}"
            logger.info(msg)
            print(msg)

        # 收缴金额规则已移动到clue_validation.py文件中

        if "没收金额" in df.columns and validate_confiscation_amount(report_text):
            # 构建比对字段和被比对字段的描述
            compared_field = f"R{index + 2}没收金额"
            being_compared_field = f"AB{index + 2}处置情况报告"
            issues_list.append({
                "受理线索编码": clue_code,
                "受理人员编码": personnel_code if personnel_code else "",
                "行号": index + 2,
                "比对字段": compared_field,
                "被比对字段": being_compared_field,
                "问题描述": f"R{index + 2}没收金额与AB{index + 2}处置情况报告对比中出现没收二字",
                "列名": "没收金额" # 添加列名用于标黄
            })

        if "责令退赔金额" in df.columns and validate_compensation_amount(report_text):
            # 构建比对字段和被比对字段的描述
            compared_field = f"S{index + 2}责令退赔金额"
            being_compared_field = f"AB{index + 2}处置情况报告"
            issues_list.append({
                "受理线索编码": clue_code,
                "受理人员编码": personnel_code if personnel_code else "",
                "行号": index + 2,
                "比对字段": compared_field,
                "被比对字段": being_compared_field,
                "问题描述": f"S{index + 2}责令退赔金额与AB{index + 2}处置情况报告对比中出现责令退赔三字",
                "列名": "责令退赔金额" # 添加列名用于标黄
            })

        if "登记上交金额" in df.columns and validate_registration_amount(report_text):
            # 构建比对字段和被比对字段的描述
            compared_field = f"T{index + 2}登记上交金额"
            being_compared_field = f"AB{index + 2}处置情况报告"
            issues_list.append({
                "受理线索编码": clue_code,
                "受理人员编码": personnel_code if personnel_code else "",
                "行号": index + 2,
                "比对字段": compared_field,
                "被比对字段": being_compared_field,
                "问题描述": f"T{index + 2}登记上交金额与AB{index + 2}处置情况报告对比中出现登记上交三字",
                "列名": "登记上交金额" # 添加列名用于标黄
            })

        if "追缴失职渎职滥用职权造成的损失金额" in df.columns and validate_recovery_amount(report_text):
            # 构建比对字段和被比对字段的描述
            compared_field = f"CJ{index + 2}追缴失职渎职滥用职权造成的损失金额"
            being_compared_field = f"AB{index + 2}处置情况报告"
            issues_list.append({
                "受理线索编码": clue_code,
                "受理人员编码": personnel_code if personnel_code else "",
                "行号": index + 2,
                "比对字段": compared_field,
                "被比对字段": being_compared_field,
                "问题描述": f"CJ{index + 2}追缴失职渎职滥用职权造成的损失金额与AB{index + 2}处置情况报告对比中出现追缴二字",
                "列名": "追缴失职渎职滥用职权造成的损失金额" # 添加列名用于标黄
            })

        if Config.COLUMN_MAPPINGS["ethnicity"] in df.columns:
            ethnicity = row[Config.COLUMN_MAPPINGS["ethnicity"]]
            if validate_ethnicity(ethnicity, report_text):
                issues_list.append((index, clue_code, Config.VALIDATION_RULES["inconsistent_ethnicity"]))
                msg = f"行 {index + 1} - 民族不一致: 字段值 '{ethnicity}', 报告值 '{extract_ethnicity_from_report(report_text)}'"
                logger.info(msg)
                print(msg)

        msg = f"行 {index + 1} issues_list: {[(i, c, issue) for i, c, issue in issues_list if i == index]}"
        logger.debug(msg)
        print(msg)

    return mismatch_indices, issues_list
