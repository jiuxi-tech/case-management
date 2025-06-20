import logging
import pandas as pd
import os
import re
from datetime import datetime
import xlsxwriter

# 假设 Config 类存在于 config.py 中
# 如果没有，你需要提供 Config 类的定义，或者直接定义 FORMATS 字典
# 为了让这个文件独立运行（用于测试或作为示例），这里提供一个简单的占位 Config
try:
    from config import Config
except ImportError:
    class Config:
        FORMATS = {
            "red": "#FFC7CE"  # 假设红色高亮的颜色代码
        }
        # 补充一些可能在其他地方用到的路径，确保不会因为缺失而报错
        BASE_UPLOAD_FOLDER = "uploads" # 示例路径
        CASE_FOLDER = os.path.join(BASE_UPLOAD_FOLDER, "case_files") # 示例路径


# 假设 extract_name_from_case_report 存在于 validation_rules.case_name_extraction 中
# 如果没有，请确保提供此函数的实现，这里提供一个示例实现
try:
    from validation_rules.case_name_extraction import extract_name_from_case_report
except ImportError:
    # 如果导入失败，提供一个简单的占位实现
    def extract_name_from_case_report(report_text):
        """
        这是一个占位函数，你需要根据实际情况实现它。
        通常用于从立案报告中提取被调查人姓名。
        """
        if not report_text or not isinstance(report_text, str):
            return None
        # 示例：假设名字在“一、XXX同志基本情况”中
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
        msg = f"提取性别: {gender} from case report"
        logger.info(msg)
        print(msg)
        return gender
    else:
        msg = f"未找到性别信息 in case report: {report_text[:100]}..."
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
    mismatch_indices = set()
    gender_mismatch_indices = set() # 新增：用于存储性别不匹配的行索引
    issues_list = []
    
    # Define required headers specific to case registration
    # 新增“性别”到必要表头中
    required_headers = ["被调查人", "性别", "立案报告", "处分决定", "审查调查报告", "审理报告"]
    if not all(header in df.columns for header in required_headers):
        logger.error(f"Missing required headers for case registration: {required_headers}")
        print(f"缺少必要的表头: {required_headers}") # Added print for console output
        return mismatch_indices, gender_mismatch_indices, issues_list # 返回三个值

    for index, row in df.iterrows():
        logger.debug(f"Processing row {index + 1}")
        print(f"处理行 {index + 1}") # Added print for console output

        # Extract investigated person
        investigated_person = str(row["被调查人"]).strip() if pd.notna(row["被调查人"]) else ''
        if not investigated_person:
            logger.info(f"行 {index + 1} - '被调查人' 字段为空，跳过此行验证。")
            print(f"行 {index + 1} - '被调查人' 字段为空，跳过此行验证。")
            continue

        # 新增：性别验证
        excel_gender = str(row["性别"]).strip() if pd.notna(row["性别"]) else ''
        report_text_for_gender = row["立案报告"] if pd.notna(row["立案报告"]) else ''
        extracted_gender = extract_gender_from_case_report(report_text_for_gender)

        # 检查提取的性别是否为空，或者与Excel中的性别不一致
        if extracted_gender is None or (excel_gender and excel_gender != extracted_gender):
            gender_mismatch_indices.add(index)
            issues_list.append((index, "M2性别与BF2立案报告不一致"))
            logger.info(f"行 {index + 1} - 性别不匹配: Excel性别 ('{excel_gender}') vs 立案报告提取性别 ('{extracted_gender}')")
            print(f"行 {index + 1} - 性别不匹配: Excel性别 ('{excel_gender}') vs 立案报告提取性别 ('{extracted_gender}')")


        # 1) Match with "立案报告" (姓名部分)
        report_text = row["立案报告"] if pd.notna(row["立案报告"]) else ''
        report_name = extract_name_from_case_report(report_text)
        if report_name and investigated_person != report_name:
            mismatch_indices.add(index)
            issues_list.append((index, "C2被调查人与BF2立案报告不一致"))
            logger.info(f"行 {index + 1} - 姓名不匹配: C2被调查人 ('{investigated_person}') vs BF2立案报告 ('{report_name}')")
            print(f"行 {index + 1} - 姓名不匹配: C2被调查人 ('{investigated_person}') vs BF2立案报告 ('{report_name}')")


        # 2) Match with "处分决定"
        decision_text = row["处分决定"] if pd.notna(row["处分决定"]) else ''
        decision_name = extract_name_from_decision(decision_text)
        if not decision_name or (decision_name and investigated_person != decision_name):
            mismatch_indices.add(index)
            issues_list.append((index, "C2被调查人与CU2处分决定不一致"))
            logger.info(f"行 {index + 1} - 姓名不匹配: C2被调查人 ('{investigated_person}') vs CU2处分决定 ('{decision_name}')")
            print(f"行 {index + 1} - 姓名不匹配: C2被调查人 ('{investigated_person}') vs CU2处分决定 ('{decision_name}')")


        # 3) Match with "审查调查报告"
        investigation_text = row["审查调查报告"] if pd.notna(row["审查调查报告"]) else ''
        investigation_name = extract_name_from_case_report(investigation_text)
        if investigation_name and investigated_person != investigation_name:
            mismatch_indices.add(index)
            issues_list.append((index, "C2被调查人与CX2审查调查报告不一致"))
            logger.info(f"行 {index + 1} - 姓名不匹配: C2被调查人 ('{investigated_person}') vs CX2审查调查报告 ('{investigation_name}')")
            print(f"行 {index + 1} - 姓名不匹配: C2被调查人 ('{investigated_person}') vs CX2审查调查报告 ('{investigation_name}')")


        # 4) Match with "审理报告"
        trial_text = row["审理报告"] if pd.notna(row["审理报告"]) else ''
        trial_name = extract_name_from_trial_report(trial_text)
        if not trial_name or (trial_name and investigated_person != trial_name):
            mismatch_indices.add(index)
            issues_list.append((index, "C2被调查人与CY2审理报告不一致"))
            logger.info(f"行 {index + 1} - 姓名不匹配: C2被调查人 ('{investigated_person}') vs CY2审理报告 ('{trial_name}')")
            print(f"行 {index + 1} - 姓名不匹配: C2被调查人 ('{investigated_person}') vs CY2审理报告 ('{trial_name}')")

    return mismatch_indices, gender_mismatch_indices, issues_list # 返回三个值

def generate_case_files(df, original_filename, upload_dir, mismatch_indices, gender_mismatch_indices, issues_list):
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

        # 获取“被调查人”和“性别”列的实际索引
        try:
            col_index_investigated_person = df.columns.get_loc("被调查人")
            col_index_gender = df.columns.get_loc("性别")
        except KeyError as e:
            logger.error(f"Excel 文件缺少必要的列: {e}")
            print(f"Excel 文件缺少必要的列: {e}")
            return None, None # 如果列不存在，则无法进行高亮和后续操作


        for idx in range(len(df)):
            # 高亮“被调查人”列
            if idx in mismatch_indices:  # Highlight mismatched rows for '被调查人'
                # idx + 1 是因为 Excel 行是 1-indexed，而 pandas 是 0-indexed
                worksheet.write(idx + 1, col_index_investigated_person, 
                                df.iloc[idx]["被调查人"] if pd.notna(df.iloc[idx]["被调查人"]) else '', red_format)
            
            # 高亮“性别”列
            if idx in gender_mismatch_indices: # Highlight mismatched rows for '性别'
                worksheet.write(idx + 1, col_index_gender,
                                df.iloc[idx]["性别"] if pd.notna(df.iloc[idx]["性别"]) else '', red_format)


    logger.info(f"Generated copy file with highlights: {copy_path}")
    print(f"生成高亮后的副本文件: {copy_path}") # Added print for console output


    # Generate case number file
    case_num_filename = f"立案编号{today}.xlsx"
    case_num_path = os.path.join(case_dir, case_num_filename)
    issues_df = pd.DataFrame(columns=['序号', '问题'])
    if not issues_list:
        issues_df = pd.DataFrame({'序号': [1], '问题': ['无问题']})
    else:
        # 使用列表推导式构建数据，然后一次性创建DataFrame，效率更高
        data = [{'序号': i + 1, '问题': issue} for i, (index, issue) in enumerate(issues_list)]
        issues_df = pd.DataFrame(data)

    issues_df.to_excel(case_num_path, index=False)
    logger.info(f"Generated case number file: {case_num_path}")
    print(f"生成立案编号表: {case_num_path}") # Added print for console output

    return copy_path, case_num_path