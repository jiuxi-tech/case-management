import logging
import pandas as pd
import os
import re
from datetime import datetime
from config import Config
import xlsxwriter
from validation_rules.case_name_extraction import extract_name_from_case_report

logger = logging.getLogger(__name__)

def extract_name_from_decision(decision_text):
    """Extract name from decision text based on '关于给予...同志党内警告处分的决定' marker."""
    if not decision_text or not isinstance(decision_text, str):
        msg = f"decision_text 为空或无效: {decision_text}"
        logger.info(msg)
        return None
    
    # 定义姓名的正则表达式，匹配“关于给予...同志党内警告处分的决定”中的姓名
    pattern = r"关于给予(.+?)同志党内警告处分的决定"
    match = re.search(pattern, decision_text)
    if match:
        name = match.group(1).strip()
        msg = f"提取姓名: {name} from decision: {decision_text}"
        logger.info(msg)
        return name
    else:
        msg = f"未找到 '关于给予...同志党内警告处分的决定' 标记: {decision_text}"
        logger.warning(msg)
        return None

def extract_name_from_trial_report(trial_text):
    """Extract name from trial report text based on '关于...同志违纪案的审理报告' marker."""
    if not trial_text or not isinstance(trial_text, str):
        msg = f"trial_text 为空或无效: {trial_text}"
        logger.info(msg)
        return None
    
    # 定义姓名的正则表达式，匹配“关于...同志违纪案的审理报告”中的姓名
    pattern = r"关于(.+?)同志违纪案的审理报告"
    match = re.search(pattern, trial_text)
    if match:
        name = match.group(1).strip()
        msg = f"提取姓名: {name} from trial report: {trial_text}"
        logger.info(msg)
        return name
    else:
        msg = f"未找到 '关于...同志违纪案的审理报告' 标记: {trial_text}"
        logger.warning(msg)
        return None

def validate_case_relationships(df):
    """Validate relationships between fields in the case registration Excel."""
    mismatch_indices = set()
    issues_list = []
    
    # Define required headers specific to case registration
    required_headers = ["被调查人", "立案报告", "处分决定", "审查调查报告", "审理报告"]
    if not all(header in df.columns for header in required_headers):
        logger.error(f"Missing required headers for case registration: {required_headers}")
        return mismatch_indices, issues_list

    for index, row in df.iterrows():
        logger.debug(f"Processing row {index + 1}")
        # Extract investigated person
        investigated_person = str(row["被调查人"]).strip() if pd.notna(row["被调查人"]) else ''
        if not investigated_person:
            continue

        # 1) Match with "立案报告"
        report_text = row["立案报告"] if pd.notna(row["立案报告"]) else ''
        report_name = extract_name_from_case_report(report_text)
        if report_name and investigated_person != report_name:
            mismatch_indices.add(index)
            issues_list.append((index, "C2被调查人与BF2立案报告不一致"))
            logger.info(f"Row {index + 1} - Name mismatch: C2被调查人 ({investigated_person}) vs BF2立案报告 ({report_name})")

        # 2) Match with "处分决定"
        decision_text = row["处分决定"] if pd.notna(row["处分决定"]) else ''
        decision_name = extract_name_from_decision(decision_text)
        if not decision_name or (decision_name and investigated_person != decision_name):
            mismatch_indices.add(index)
            issues_list.append((index, "C2被调查人与CU2处分决定不一致"))
            logger.info(f"Row {index + 1} - Name mismatch: C2被调查人 ({investigated_person}) vs CU2处分决定 ({decision_name})")

        # 3) Match with "审查调查报告"
        investigation_text = row["审查调查报告"] if pd.notna(row["审查调查报告"]) else ''
        investigation_name = extract_name_from_case_report(investigation_text)
        if investigation_name and investigated_person != investigation_name:
            mismatch_indices.add(index)
            issues_list.append((index, "C2被调查人与CX2审查调查报告不一致"))
            logger.info(f"Row {index + 1} - Name mismatch: C2被调查人 ({investigated_person}) vs CX2审查调查报告 ({investigation_name})")

        # 4) Match with "审理报告"
        trial_text = row["审理报告"] if pd.notna(row["审理报告"]) else ''
        trial_name = extract_name_from_trial_report(trial_text)
        if not trial_name or (trial_name and investigated_person != trial_name):
            mismatch_indices.add(index)
            issues_list.append((index, "C2被调查人与CY2审理报告不一致"))
            logger.info(f"Row {index + 1} - Name mismatch: C2被调查人 ({investigated_person}) vs CY2审理报告 ({trial_name})")

    return mismatch_indices, issues_list

def generate_case_files(df, original_filename, upload_dir, mismatch_indices, issues_list):
    """Generate copy and case number Excel files based on analysis."""
    today = datetime.now().strftime('%Y%m%d')
    case_dir = os.path.join(upload_dir, today, 'case')
    os.makedirs(case_dir, exist_ok=True)

    # Generate copy file with formatting
    copy_filename = original_filename.replace('.xlsx', '_副本.xlsx').replace('.xls', '_副本.xlsx')
    copy_path = os.path.join(case_dir, copy_filename)
    with pd.ExcelWriter(copy_path, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='Sheet1', index=False)
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']
        red_format = workbook.add_format({'bg_color': Config.FORMATS["red"]})

        for idx in range(len(df)):
            if idx in mismatch_indices:  # Highlight mismatched rows
                worksheet.write(f'C{idx + 2}', df.iloc[idx]["被调查人"] if pd.notna(df.iloc[idx]["被调查人"]) else '', red_format)

    logger.info(f"Generated copy file with highlights: {copy_path}")

    # Generate case number file
    case_num_filename = f"立案编号{today}.xlsx"
    case_num_path = os.path.join(case_dir, case_num_filename)
    issues_df = pd.DataFrame(columns=['序号', '问题'])
    if not issues_list:
        issues_df = pd.DataFrame({'序号': [1], '问题': ['无问题']})
    else:
        current_index = 1
        for i, (index, issue) in enumerate(issues_list, 1):
            issues_df = pd.concat([issues_df, pd.DataFrame({'序号': [current_index], '问题': [issue]})], ignore_index=True)
            current_index += 1
    issues_df.to_excel(case_num_path, index=False)
    logger.info(f"Generated case number file: {case_num_path}")

    return copy_path, case_num_path