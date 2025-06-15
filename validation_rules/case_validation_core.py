import logging
import pandas as pd
import os
from datetime import datetime
from config import Config

logger = logging.getLogger(__name__)

def validate_case_relationships(df):
    """Validate relationships between fields in the case registration Excel."""
    mismatch_indices = set()
    issues_list = []
    
    required_headers = Config.REQUIRED_HEADERS
    if not all(header in df.columns for header in required_headers):
        logger.error(f"Missing required headers: {required_headers}")
        return mismatch_indices, issues_list

    for index, row in df.iterrows():
        logger.debug(f"Processing row {index + 1}")
        # Validate agency consistency
        authority = str(row["办理机关"]).strip() if pd.notna(row["办理机关"]) else ''
        agency = str(row["填报单位名称"]).strip() if pd.notna(row["填报单位名称"]) else ''
        if authority != agency:
            mismatch_indices.add(index)
            issues_list.append((index, "办理机关与填报单位名称不一致"))
            logger.info(f"Row {index + 1} - Agency mismatch: {authority} vs {agency}")

        # Validate name consistency with report
        reported_person = str(row["被反映人"]).strip() if pd.notna(row["被反映人"]) else ''
        report_text = row["处置情况报告"] if "处置情况报告" in df.columns else ''
        if reported_person and report_text and reported_person not in report_text:
            issues_list.append((index, "被反映人姓名与处置情况报告不一致"))
            logger.info(f"Row {index + 1} - Name mismatch: {reported_person}")

        # Validate acceptance time
        acceptance_time = row[Config.COLUMN_MAPPINGS["acceptance_time"]] if pd.notna(row[Config.COLUMN_MAPPINGS["acceptance_time"]]) else None
        if acceptance_time and isinstance(acceptance_time, str) and not re.match(r'^\d{4}-\d{2}-\d{2}$', acceptance_time):
            issues_list.append((index, "受理时间格式不正确，请确认"))

    return mismatch_indices, issues_list

def generate_case_files(df, original_filename, upload_dir):
    """Generate copy and case number Excel files based on analysis."""
    today = datetime.now().strftime('%Y%m%d')
    case_dir = os.path.join(upload_dir, today, 'case')
    os.makedirs(case_dir, exist_ok=True)

    # Generate copy file
    copy_filename = original_filename.replace('.xlsx', '_副本.xlsx').replace('.xls', '_副本.xlsx')
    copy_path = os.path.join(case_dir, copy_filename)
    df.to_excel(copy_path, index=False)
    logger.info(f"Generated copy file: {copy_path}")

    # Generate case number file
    case_num_filename = f"立案编号{today}.xlsx"
    case_num_path = os.path.join(case_dir, case_num_filename)
    issues_df = pd.DataFrame(columns=['序号', '问题'])
    if issues_df.empty:
        issues_df = pd.DataFrame({'序号': [1], '问题': ['无问题']})
    issues_df.to_excel(case_num_path, index=False)
    logger.info(f"Generated case number file: {case_num_path}")

    return copy_path, case_num_path