import logging
import pandas as pd
import os
import re
from datetime import datetime
import xlsxwriter

# Assume Config class exists in config.py
# If not, you need to provide Config class definition, or directly define FORMATS dict
try:
    from config import Config
except ImportError:
    class Config:
        FORMATS = {
            "red": "#FFC7CE"  # Assuming red highlight color code
        }
        # Supplement some paths that may be used elsewhere, ensure no error due to missing
        BASE_UPLOAD_FOLDER = "uploads" # Example path
        CASE_FOLDER = os.path.join(BASE_UPLOAD_FOLDER, "case_files") # Example path


# Assume extract_name_from_case_report exists in validation_rules.case_name_extraction
# If not, please ensure to provide its implementation, here's a sample implementation
try:
    from validation_rules.case_name_extraction import extract_name_from_case_report
except ImportError:
    # If import fails, provide a simple placeholder implementation
    def extract_name_from_case_report(report_text):
        """
        This is a placeholder function, you need to implement it according to your actual situation.
        It is usually used to extract the name of the person being investigated from the case report.
        """
        if not report_text or not isinstance(report_text, str):
            return None
        # Example: Assume the name is in "一、XXX同志基本情况"
        match = re.search(r"一、(.+?)同志基本情况", report_text)
        if match:
            return match.group(1).strip()
        return None

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
    
    # 查找“同志基本情况”后面的第一个和第二个逗号之间的内容
    # 使用 re.DOTALL 确保 '.' 匹配换行符
    # 匹配模式： "一、" + 任意内容 + "同志基本情况" + 任意内容（非贪婪）+ "，" + 捕获组（非逗号字符）+ "，"
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
    
    # 首先找到处分决定的标题标记
    title_pattern = r"关于给予.+?同志党内警告处分的决定"
    title_match = re.search(title_pattern, decision_text, re.DOTALL)

    if title_match:
        # 从标题匹配的结束位置开始搜索性别信息
        start_pos = title_match.end()
        # 假设性别信息在标题后的第一个非空行或段落中
        # 查找标题后，第一个逗号和第二个逗号之间的内容
        # 匹配标题结束后的任意内容（非贪婪），直到第一个逗号，然后捕获非逗号内容，再到第二个逗号
        gender_pattern = r".*?，([^，]+)，"
        # 限制搜索范围，例如只搜索标题后的前200个字符，防止匹配到无关内容
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
    
    # 查找“同志基本情况”后面的第一个和第二个逗号之间的内容
    # 使用 re.DOTALL 确保 '.' 匹配换行符
    # 匹配模式： "一、" + 任意内容 + "同志基本情况" + 任意内容（非贪婪）+ "，" + 捕获组（非逗号字符）+ "，"
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

    # 首先找到“现将具体情况报告如下”标记
    title_marker = "现将具体情况报告如下"
    marker_pos = trial_text.find(title_marker)

    if marker_pos != -1:
        # 从标记的结束位置开始搜索性别信息
        start_pos = marker_pos + len(title_marker)
        # 查找标记后，第一个逗号和第二个逗号之间的内容
        # 匹配标记结束后的任意内容（非贪婪），直到第一个逗号，然后捕获非逗号内容，再到第二个逗号
        gender_pattern = r".*?，([^，]+)，"
        # 限制搜索范围，例如只搜索标记后的前200个字符，防止匹配到无关内容
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

def extract_birth_year_from_case_report(report_text):
    """
    从立案报告中提取出生年份。
    年龄位置在“一、王xx同志基本情况”这样字符串下面段落里第三个逗号和第四个逗号中间的位置。
    例如：“王xx，男，汉族，1966年12月生，山东省平度市xx镇xx村人”中的“1966”。
    """
    if not report_text or not isinstance(report_text, str):
        msg = f"extract_birth_year_from_case_report: report_text 为空或无效: {report_text}"
        logger.info(msg)
        print(msg)
        return None

    # Find the "一、XXX同志基本情况" marker
    marker_pattern = r"一、.+?同志基本情况"
    marker_match = re.search(marker_pattern, report_text, re.DOTALL)

    if marker_match:
        start_pos = marker_match.end()
        # Search area after the marker, limit to prevent matching irrelevant text further down
        search_area = report_text[start_pos : start_pos + 300] # Adjust length as needed

        # Split by comma and find the 4th segment (0-indexed 3rd segment) which contains year
        # "王xx，男，汉族，1966年12月生，山东省平度市xx镇xx村人"
        # 0: 王xx
        # 1: 男
        # 2: 汉族
        # 3: 1966年12月生
        parts = [p.strip() for p in search_area.split('，')]
        
        if len(parts) > 3:
            # The 4th part (index 3) should contain the birth year
            birth_info_segment = parts[3]
            year_match = re.search(r'(\d{4})年', birth_info_segment)
            if year_match:
                birth_year = int(year_match.group(1))
                msg = f"提取出生年份 (立案报告): {birth_year} from case report"
                logger.info(msg)
                print(msg)
                return birth_year
            else:
                msg = f"在立案报告的第4个逗号分隔段中未找到年份信息: '{birth_info_segment}'"
                logger.warning(msg)
                print(msg)
                return None
        else:
            msg = f"立案报告中 '一、同志基本情况' 后面的逗号分隔段不足，无法提取年份: {search_area[:50]}..."
            logger.warning(msg)
            print(msg)
            return None
    else:
        msg = f"未找到 '一、XXX同志基本情况' 标记，无法提取立案报告出生年份: {report_text[:100]}..."
        logger.warning(msg)
        print(msg)
        return None

def extract_name_from_decision(decision_text):
    """从处分决定中提取姓名，基于'关于给予...同志党内警告处分的决定'标记。"""
    if not decision_text or not isinstance(decision_text, str):
        msg = f"extract_name_from_decision: decision_text 为空或无效: {decision_text}"
        logger.info(msg)
        print(msg) # Added print for console output
        return None
    
    # 定义姓名的正则表达式，匹配“关于给予...同志党内警告处分的决定”中的姓名
    pattern = r"关于给予(.+?)同志党内警告处分的决定"
    match = re.search(pattern, decision_text)
    if match:
        name = match.group(1).strip()
        msg = f"提取姓名: {name} from decision: {decision_text[:50]}..."
        logger.info(msg)
        print(msg) # Added print for console output
        return name
    else:
        msg = f"未找到 '关于给予...同志党内警告处分的决定' 标记: {decision_text[:50]}..."
        logger.warning(msg)
        print(msg) # Added print for console output
        return None

def extract_name_from_trial_report(trial_text):
    """从审理报告中提取姓名，基于'关于...同志违纪案的审理报告'标记。"""
    if not trial_text or not isinstance(trial_text, str):
        msg = f"extract_name_from_trial_report: trial_text 为空或无效: {trial_text}"
        logger.info(msg)
        print(msg) # Added print for console output
        return None
    
    # 定义姓名的正则表达式，匹配“关于...同志违纪案的审理报告”中的姓名
    pattern = r"关于(.+?)同志违纪案的审理报告"
    match = re.search(pattern, trial_text)
    if match:
        name = match.group(1).strip()
        msg = f"提取姓名: {name} from trial report: {trial_text[:50]}..."
        logger.info(msg)
        print(msg) # Added print for console output
        return name
    else:
        msg = f"未找到 '关于...同志违纪案的审理报告' 标记: {trial_text[:50]}..."
        logger.warning(msg)
        print(msg) # Added print for console output
        return None

def validate_case_relationships(df):
    """Validate relationships between fields in the case registration Excel."""
    mismatch_indices = set() # For name mismatches
    gender_mismatch_indices = set() # For gender mismatches
    age_mismatch_indices = set() # New: For age mismatches
    issues_list = []
    
    # Define required headers specific to case registration
    # Added "年龄" to required headers
    required_headers = ["被调查人", "性别", "年龄", "立案报告", "处分决定", "审查调查报告", "审理报告"]
    if not all(header in df.columns for header in required_headers):
        logger.error(f"Missing required headers for case registration: {required_headers}")
        print(f"缺少必要的表头: {required_headers}") # Added print for console output
        return mismatch_indices, gender_mismatch_indices, age_mismatch_indices, issues_list # Return four values

    current_year = datetime.now().year

    for index, row in df.iterrows():
        logger.debug(f"Processing row {index + 1}")
        print(f"处理行 {index + 1}") # Added print for console output

        # Extract investigated person
        investigated_person = str(row["被调查人"]).strip() if pd.notna(row["被调查人"]) else ''
        if not investigated_person:
            logger.info(f"行 {index + 1} - '被调查人' 字段为空，跳过此行验证。")
            print(f"行 {index + 1} - '被调查人' 字段为空，跳过此行验证。")
            continue

        excel_gender = str(row["性别"]).strip() if pd.notna(row["性别"]) else ''
        excel_age = None
        if pd.notna(row["年龄"]):
            try:
                excel_age = int(row["年龄"])
            except ValueError:
                logger.warning(f"行 {index + 1} - Excel '年龄' 字段 '{row['年龄']}' 不是有效数字。")
                print(f"行 {index + 1} - Excel '年龄' 字段 '{row['年龄']}' 不是有效数字。")
                age_mismatch_indices.add(index)
                issues_list.append((index, "N2年龄字段格式不正确"))


        # --- Gender matching rules ---
        # 1) 性别与“立案报告”匹配
        report_text_for_gender = row["立案报告"] if pd.notna(row["立案报告"]) else ''
        extracted_gender_from_report = extract_gender_from_case_report(report_text_for_gender)
        if extracted_gender_from_report is None or (excel_gender and excel_gender != extracted_gender_from_report):
            gender_mismatch_indices.add(index)
            issues_list.append((index, "M2性别与BF2立案报告不一致"))
            logger.info(f"行 {index + 1} - 性别不匹配: Excel性别 ('{excel_gender}') vs 立案报告提取性别 ('{extracted_gender_from_report}')")
            print(f"行 {index + 1} - 性别不匹配: Excel性别 ('{excel_gender}') vs 立案报告提取性别 ('{extracted_gender_from_report}')")

        # 2) 性别与“处分决定”匹配
        decision_text_for_gender = row["处分决定"] if pd.notna(row["处分决定"]) else ''
        extracted_gender_from_decision = extract_gender_from_decision_report(decision_text_for_gender)
        if extracted_gender_from_decision is None or (excel_gender and excel_gender != extracted_gender_from_decision):
            gender_mismatch_indices.add(index) 
            issues_list.append((index, "M2性别与CU2处分决定不一致"))
            logger.info(f"行 {index + 1} - 性别不匹配: Excel性别 ('{excel_gender}') vs 处分决定提取性别 ('{extracted_gender_from_decision}')")
            print(f"行 {index + 1} - 性别不匹配: Excel性别 ('{excel_gender}') vs 处分决定提取性别 ('{extracted_gender_from_decision}')")

        # 3) 性别与“审查调查报告”匹配
        investigation_text_for_gender = row["审查调查报告"] if pd.notna(row["审查调查报告"]) else ''
        extracted_gender_from_investigation = extract_gender_from_investigation_report(investigation_text_for_gender)
        if extracted_gender_from_investigation is None or (excel_gender and excel_gender != extracted_gender_from_investigation):
            gender_mismatch_indices.add(index)
            issues_list.append((index, "M2性别与CX2审查调查报告不一致"))
            logger.info(f"行 {index + 1} - 性别不匹配: Excel性别 ('{excel_gender}') vs 审查调查报告提取性别 ('{extracted_gender_from_investigation}')")
            print(f"行 {index + 1} - 性别不匹配: Excel性别 ('{excel_gender}') vs 审查调查报告提取性别 ('{extracted_gender_from_investigation}')")

        # 4) 性别与“审理报告”匹配
        trial_text_for_gender = row["审理报告"] if pd.notna(row["审理报告"]) else ''
        extracted_gender_from_trial = extract_gender_from_trial_report(trial_text_for_gender)
        if extracted_gender_from_trial is None or (excel_gender and excel_gender != extracted_gender_from_trial):
            gender_mismatch_indices.add(index)
            issues_list.append((index, "M2性别与CY2审理报告不一致"))
            logger.info(f"行 {index + 1} - 性别不匹配: Excel性别 ('{excel_gender}') vs 审理报告提取性别 ('{extracted_gender_from_trial}')")
            print(f"行 {index + 1} - 性别不匹配: Excel性别 ('{excel_gender}') vs 审理报告提取性别 ('{extracted_gender_from_trial}')")

        # --- Age matching rules ---
        # 年龄与“立案报告”匹配
        report_text_for_age = row["立案报告"] if pd.notna(row["立案报告"]) else ''
        extracted_birth_year = extract_birth_year_from_case_report(report_text_for_age)
        
        calculated_age = None
        if extracted_birth_year is not None:
            calculated_age = current_year - extracted_birth_year
            logger.info(f"行 {index + 1} - 计算年龄: {current_year} - {extracted_birth_year} = {calculated_age}")
            print(f"行 {index + 1} - 计算年龄: {current_year} - {extracted_birth_year} = {calculated_age}")

        # Check for age mismatch
        if (calculated_age is None) or \
           (excel_age is not None and calculated_age is not None and excel_age != calculated_age):
            age_mismatch_indices.add(index)
            issues_list.append((index, "N2年龄与BF2立案报告不一致"))
            logger.info(f"行 {index + 1} - 年龄不匹配: Excel年龄 ('{excel_age}') vs 立案报告计算年龄 ('{calculated_age}')")
            print(f"行 {index + 1} - 年龄不匹配: Excel年龄 ('{excel_age}') vs 立案报告计算年龄 ('{calculated_age}')")


        # --- Name matching rules (remain unchanged) ---
        # 1) Match with "立案报告" (Name part)
        report_text = row["立案报告"] if pd.notna(row["立案报告"]) else ''
        report_name = extract_name_from_case_report(report_text)
        if report_name and investigated_person != report_name:
            mismatch_indices.add(index)
            issues_list.append((index, "C2被调查人与BF2立案报告不一致"))
            logger.info(f"行 {index + 1} - 姓名不匹配: C2被调查人 ('{investigated_person}') vs BF2立案报告 ('{report_name}')")
            print(f"行 {index + 1} - 姓名不匹配: C2被调查人 ('{investigated_person}') vs BF2立案报告 ('{report_name}')")


        # 2) Match with "处分决定" (Name part)
        decision_text = row["处分决定"] if pd.notna(row["处分决定"]) else ''
        decision_name = extract_name_from_decision(decision_text)
        if not decision_name or (decision_name and investigated_person != decision_name):
            mismatch_indices.add(index)
            issues_list.append((index, "C2被调查人与CU2处分决定不一致"))
            logger.info(f"行 {index + 1} - 姓名不匹配: C2被调查人 ('{investigated_person}') vs CU2处分决定 ('{decision_name}')")
            print(f"行 {index + 1} - 姓名不匹配: C2被调查人 ('{investigated_person}') vs CU2处分决定 ('{decision_name}')")


        # 3) Match with "审查调查报告" (Name part)
        investigation_text = row["审查调查报告"] if pd.notna(row["审查调查报告"]) else ''
        investigation_name = extract_name_from_case_report(investigation_text)
        if investigation_name and investigated_person != investigation_name:
            mismatch_indices.add(index)
            issues_list.append((index, "C2被调查人与CX2审查调查报告不一致"))
            logger.info(f"行 {index + 1} - 姓名不匹配: C2被调查人 ('{investigated_person}') vs CX2审查调查报告 ('{investigation_name}')")
            print(f"行 {index + 1} - 姓名不匹配: C2被调查人 ('{investigated_person}') vs CX2审查调查报告 ('{investigation_name}')")


        # 4) Match with "审理报告" (Name part)
        trial_text = row["审理报告"] if pd.notna(row["审理报告"]) else ''
        trial_name = extract_name_from_trial_report(trial_text)
        if not trial_name or (trial_name and investigated_person != trial_name):
            mismatch_indices.add(index)
            issues_list.append((index, "C2被调查人与CY2审理报告不一致"))
            logger.info(f"行 {index + 1} - 姓名不匹配: C2被调查人 ('{investigated_person}') vs CY2审理报告 ('{trial_name}')")
            print(f"行 {index + 1} - 姓名不匹配: C2被调查人 ('{investigated_person}') vs CY2审理报告 ('{trial_name}')")

    # Important: The return statement now includes age_mismatch_indices
    return mismatch_indices, gender_mismatch_indices, age_mismatch_indices, issues_list

def generate_case_files(df, original_filename, upload_dir, mismatch_indices, gender_mismatch_indices, issues_list, age_mismatch_indices):
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

        # Get actual column indices for "被调查人", "性别", and "年龄"
        try:
            col_index_investigated_person = df.columns.get_loc("被调查人")
            col_index_gender = df.columns.get_loc("性别")
            col_index_age = df.columns.get_loc("年龄") # New: Age column index
        except KeyError as e:
            logger.error(f"Excel 文件缺少必要的列: {e}")
            print(f"Excel 文件缺少必要的列: {e}")
            # Ensure proper return if columns are missing
            return None, None 

        for idx in range(len(df)):
            # Highlight "被调查人" column
            if idx in mismatch_indices:  # Highlight mismatched rows for '被调查人'
                # idx + 1 is because Excel rows are 1-indexed, while pandas is 0-indexed
                worksheet.write(idx + 1, col_index_investigated_person, 
                                df.iloc[idx]["被调查人"] if pd.notna(df.iloc[idx]["被调查人"]) else '', red_format)
            
            # Highlight "性别" column
            if idx in gender_mismatch_indices: # Highlight mismatched rows for '性别'
                worksheet.write(idx + 1, col_index_gender,
                                df.iloc[idx]["性别"] if pd.notna(df.iloc[idx]["性别"]) else '', red_format)

            # Highlight "年龄" column (New)
            if idx in age_mismatch_indices: # Highlight mismatched rows for '年龄'
                worksheet.write(idx + 1, col_index_age,
                                df.iloc[idx]["年龄"] if pd.notna(df.iloc[idx]["年龄"]) else '', red_format)


    logger.info(f"Generated copy file with highlights: {copy_path}")
    print(f"生成高亮后的副本文件: {copy_path}") # Added print for console output


    # Generate case number file
    case_num_filename = f"立案编号{today}.xlsx"
    case_num_path = os.path.join(case_dir, case_num_filename)
    issues_df = pd.DataFrame(columns=['序号', '问题'])
    if not issues_list:
        issues_df = pd.DataFrame({'序号': [1], '问题': ['无问题']})
    else:
        # Use list comprehension to build data, then create DataFrame at once for better efficiency
        data = [{'序号': i + 1, '问题': issue} for i, (index, issue) in enumerate(issues_list)]
        issues_df = pd.DataFrame(data)

    issues_df.to_excel(case_num_path, index=False)
    logger.info(f"Generated case number file: {case_num_path}")
    print(f"生成立案编号表: {case_num_path}") # Added print for console output

    return copy_path, case_num_path