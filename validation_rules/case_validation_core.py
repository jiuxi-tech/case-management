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
        # 匹配标题结束后的任意内容（非贪婪），直到第一个逗号，然后捕获非逗号内容，再到第二个逗号
        gender_pattern = r".*?，([^，]+)，"
        # 限制搜索范围，例如只搜索标题后的前200个字符，防止匹配到无关内容
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

def extract_birth_year_from_decision_report(decision_text):
    """
    从处分决定中提取出生年份。
    年龄位置在“关于给予王xx同志党内警告处分的决定”这样字符串下面段落里
    第三个逗号和第四个逗号中间的位置。
    例如：“王xx，男，汉族，1966年12月生，山东省平度市xx镇xx村人”中的“1966”。
    """
    if not decision_text or not isinstance(decision_text, str):
        msg = f"extract_birth_year_from_decision_report: decision_text 为空或无效: {decision_text}"
        logger.info(msg)
        print(msg)
        return None
    
    # 首先找到处分决定的标题标记
    title_pattern = r"关于给予.+?同志党内警告处分的决定"
    title_match = re.search(title_pattern, decision_text, re.DOTALL)

    if title_match:
        start_pos = title_match.end()
        # Search area after the marker, limit to prevent matching irrelevant text further down
        search_area = decision_text[start_pos : start_pos + 300] # Adjust length as needed

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
                msg = f"提取出生年份 (处分决定): {birth_year} from decision report"
                logger.info(msg)
                print(msg)
                return birth_year
            else:
                msg = f"在处分决定的第4个逗号分隔段中未找到年份信息: '{birth_info_segment}'"
                logger.warning(msg)
                print(msg)
                return None
        else:
            msg = f"处分决定中 '关于给予...同志党内警告处分的决定' 后面的逗号分隔段不足，无法提取年份: {search_area[:50]}..."
            logger.warning(msg)
            print(msg)
            return None
    else:
        msg = f"未找到 '关于给予...同志党内警告处分的决定' 标记，无法提取处分决定出生年份: {decision_text[:100]}..."
        logger.warning(msg)
        print(msg)
        return None

def extract_birth_year_from_investigation_report(investigation_text):
    """
    从审查调查报告中提取出生年份。
    年龄位置在“一、王xx同志基本情况”这样字符串下面段落里第三个逗号和第四个逗号中间的位置。
    例如：“王xx，男，汉族，1966年12月生，山东省平度市xx镇xx村人”中的“1966”。
    """
    if not investigation_text or not isinstance(investigation_text, str):
        msg = f"extract_birth_year_from_investigation_report: investigation_text 为空或无效: {investigation_text}"
        logger.info(msg)
        print(msg)
        return None

    # Find the "一、XXX同志基本情况" marker
    marker_pattern = r"一、.+?同志基本情况"
    marker_match = re.search(marker_pattern, investigation_text, re.DOTALL)

    if marker_match:
        start_pos = marker_match.end()
        # Search area after the marker, limit to prevent matching irrelevant text further down
        search_area = investigation_text[start_pos : start_pos + 300] # Adjust length as needed

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
                msg = f"提取出生年份 (审查调查报告): {birth_year} from investigation report"
                logger.info(msg)
                print(msg)
                return birth_year
            else:
                msg = f"在审查调查报告的第4个逗号分隔段中未找到年份信息: '{birth_info_segment}'"
                logger.warning(msg)
                print(msg)
                return None
        else:
            msg = f"审查调查报告中 '一、同志基本情况' 后面的逗号分隔段不足，无法提取年份: {search_area[:50]}..."
            logger.warning(msg)
            print(msg)
            return None
    else:
        msg = f"未找到 '一、XXX同志基本情况' 标记，无法提取审查调查报告出生年份: {investigation_text[:100]}..."
        logger.warning(msg)
        print(msg)
        return None

def extract_birth_year_from_trial_report(trial_text):
    """
    从审理报告中提取出生年份。
    年龄位置在“现将具体情况报告如下”这样字符串下面段落里第三个逗号和第四个逗号中间的位置。
    例如：“王xx，男，汉族，1966年12月生，山东省平度市xx镇xx村人”中的“1966”。
    """
    if not trial_text or not isinstance(trial_text, str):
        msg = f"extract_birth_year_from_trial_report: trial_text 为空或无效: {trial_text}"
        logger.info(msg)
        print(msg)
        return None

    # 首先找到“现将具体情况报告如下”标记
    marker = "现将具体情况报告如下"
    marker_pos = trial_text.find(marker)

    if marker_pos != -1:
        start_pos = marker_pos + len(marker)
        # Search area after the marker, limit to prevent matching irrelevant text further down
        search_area = trial_text[start_pos : start_pos + 300] # Adjust length as needed

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
                msg = f"提取出生年份 (审理报告): {birth_year} from trial report"
                logger.info(msg)
                print(msg)
                return birth_year
            else:
                msg = f"在审理报告的第4个逗号分隔段中未找到年份信息: '{birth_info_segment}'"
                logger.warning(msg)
                print(msg)
                return None
        else:
            msg = f"审理报告中 '现将具体情况报告如下' 后面的逗号分隔段不足，无法提取年份: {search_area[:50]}..."
            logger.warning(msg)
            print(msg)
            return None
    else:
        msg = f"未找到 '现将具体情况报告如下' 标记，无法提取审理报告出生年份: {trial_text[:100]}..."
        logger.warning(msg)
        print(msg)
        return None

def extract_birth_date_from_case_report(report_text):
    """
    从立案报告中提取出生年月，并格式化为“YYYY/MM”。
    出生年月位置在“一、王xx同志基本情况”这样字符串下面段落里第三个逗号和第四个逗号中间的位置。
    例如：“王xx，男，汉族，1966年12月生，山东省平度市xx镇xx村人”中的“1966年12月”。
    """
    if not report_text or not isinstance(report_text, str):
        msg = f"extract_birth_date_from_case_report: report_text 为空或无效: {report_text}"
        logger.info(msg)
        print(msg)
        return None

    marker_pattern = r"一、.+?同志基本情况"
    marker_match = re.search(marker_pattern, report_text, re.DOTALL)

    if marker_match:
        start_pos = marker_match.end()
        search_area = report_text[start_pos : start_pos + 300] 
        parts = [p.strip() for p in search_area.split('，')]
        
        if len(parts) > 3:
            birth_info_segment = parts[3]
            # 匹配“YYYY年MM月”或“YYYY年M月”的格式
            date_match = re.search(r'(\d{4})年(\d{1,2})月', birth_info_segment)
            if date_match:
                year = date_match.group(1)
                month = date_match.group(2).zfill(2) # 补齐两位月份
                formatted_date = f"{year}/{month}"
                msg = f"提取出生年月 (立案报告): {formatted_date} from case report"
                logger.info(msg)
                print(msg)
                return formatted_date
            else:
                msg = f"在立案报告的第4个逗号分隔段中未找到出生年月信息: '{birth_info_segment}'"
                logger.warning(msg)
                print(msg)
                return None
        else:
            msg = f"立案报告中 '一、同志基本情况' 后面的逗号分隔段不足，无法提取出生年月: {search_area[:50]}..."
            logger.warning(msg)
            print(msg)
            return None
    else:
        msg = f"未找到 '一、XXX同志基本情况' 标记，无法提取立案报告出生年月: {report_text[:100]}..."
        logger.warning(msg)
        print(msg)
        return None

def extract_birth_date_from_decision_report(decision_text):
    """
    从处分决定中提取出生年月，并格式化为“YYYY/MM”。
    出生年月位置在“关于给予王xx同志党内警告处分的决定”这样字符串（王xx是变量）下面段落里
    第三个逗号和第四个逗号中间的位置。
    例如：“王xx，男，汉族，1966年12月生，山东省平度市xx镇xx村人”中的“1966年12月”。
    """
    if not decision_text or not isinstance(decision_text, str):
        msg = f"extract_birth_date_from_decision_report: decision_text 为空或无效: {decision_text}"
        logger.info(msg)
        print(msg)
        return None

    title_pattern = r"关于给予.+?同志党内警告处分的决定"
    title_match = re.search(title_pattern, decision_text, re.DOTALL)

    if title_match:
        start_pos = title_match.end()
        search_area = decision_text[start_pos : start_pos + 300] 
        parts = [p.strip() for p in search_area.split('，')]
        
        if len(parts) > 3:
            birth_info_segment = parts[3]
            date_match = re.search(r'(\d{4})年(\d{1,2})月', birth_info_segment)
            if date_match:
                year = date_match.group(1)
                month = date_match.group(2).zfill(2)
                formatted_date = f"{year}/{month}"
                msg = f"提取出生年月 (处分决定): {formatted_date} from decision report"
                logger.info(msg)
                print(msg)
                return formatted_date
            else:
                msg = f"在处分决定的第4个逗号分隔段中未找到出生年月信息: '{birth_info_segment}'"
                logger.warning(msg)
                print(msg)
                return None
        else:
            msg = f"处分决定中 '关于给予...同志党内警告处分的决定' 后面的逗号分隔段不足，无法提取出生年月: {search_area[:50]}..."
            logger.warning(msg)
            print(msg)
            return None
    else:
        msg = f"未找到 '关于给予...同志党内警告处分的决定' 标记，无法提取处分决定出生年月: {decision_text[:100]}..."
        logger.warning(msg)
        print(msg)
        return None

def extract_birth_date_from_investigation_report(investigation_text):
    """
    从审查调查报告中提取出生年月，并格式化为“YYYY/MM”。
    出生年月位置在“一、王xx同志基本情况”这样字符串（王xx是变量）下面段落里
    第三个逗号和第四个逗号中间的位置。
    例如：“王xx，男，汉族，1966年12月生，山东省平度市xx镇xx村人”中的“1966年12月”。
    """
    if not investigation_text or not isinstance(investigation_text, str):
        msg = f"extract_birth_date_from_investigation_report: investigation_text 为空或无效: {investigation_text}"
        logger.info(msg)
        print(msg)
        return None

    marker_pattern = r"一、.+?同志基本情况"
    marker_match = re.search(marker_pattern, investigation_text, re.DOTALL)

    if marker_match:
        start_pos = marker_match.end()
        search_area = investigation_text[start_pos : start_pos + 300]
        parts = [p.strip() for p in search_area.split('，')]
        
        if len(parts) > 3:
            birth_info_segment = parts[3]
            date_match = re.search(r'(\d{4})年(\d{1,2})月', birth_info_segment)
            if date_match:
                year = date_match.group(1)
                month = date_match.group(2).zfill(2)
                formatted_date = f"{year}/{month}"
                msg = f"提取出生年月 (审查调查报告): {formatted_date} from investigation report"
                logger.info(msg)
                print(msg)
                return formatted_date
            else:
                msg = f"在审查调查报告的第4个逗号分隔段中未找到出生年月信息: '{birth_info_segment}'"
                logger.warning(msg)
                print(msg)
                return None
        else:
            msg = f"审查调查报告中 '一、同志基本情况' 后面的逗号分隔段不足，无法提取出生年月: {search_area[:50]}..."
            logger.warning(msg)
            print(msg)
            return None
    else:
        msg = f"未找到 '一、XXX同志基本情况' 标记，无法提取审查调查报告出生年月: {investigation_text[:100]}..."
        logger.warning(msg)
        print(msg)
        return None

def extract_birth_date_from_trial_report(trial_text):
    """
    从审理报告中提取出生年月，并格式化为“YYYY/MM”。
    出生年月位置在“现将具体情况报告如下”这样字符串下面段落里第三个逗号和第四个逗号中间的位置。
    例如：“王xx，男，汉族，1966年12月生，山东省平度市xx镇xx村人”中的“1966年12月”。
    """
    if not trial_text or not isinstance(trial_text, str):
        msg = f"extract_birth_date_from_trial_report: trial_text 为空或无效: {trial_text}"
        logger.info(msg)
        print(msg)
        return None

    marker = "现将具体情况报告如下"
    marker_pos = trial_text.find(marker)

    if marker_pos != -1:
        start_pos = marker_pos + len(marker)
        search_area = trial_text[start_pos : start_pos + 300]
        parts = [p.strip() for p in search_area.split('，')]
        
        if len(parts) > 3:
            birth_info_segment = parts[3]
            date_match = re.search(r'(\d{4})年(\d{1,2})月', birth_info_segment)
            if date_match:
                year = date_match.group(1)
                month = date_match.group(2).zfill(2)
                formatted_date = f"{year}/{month}"
                msg = f"提取出生年月 (审理报告): {formatted_date} from trial report"
                logger.info(msg)
                print(msg)
                return formatted_date
            else:
                msg = f"在审理报告的第4个逗号分隔段中未找到出生年月信息: '{birth_info_segment}'"
                logger.warning(msg)
                print(msg)
                return None
        else:
            msg = f"审理报告中 '现将具体情况报告如下' 后面的逗号分隔段不足，无法提取出生年月: {search_area[:50]}..."
            logger.warning(msg)
            print(msg)
            return None
    else:
        msg = f"未找到 '现将具体情况报告如下' 标记，无法提取审理报告出生年月: {trial_text[:100]}..."
        logger.warning(msg)
        print(msg)
        return None

def extract_education_from_case_report(report_text):
    """
    从立案报告中提取学历。
    学历位置可能在“一、王xx同志基本情况”这样字符串下面段落里某个位置。
    会优先匹配更具体的学历词汇，并能处理“大学本科”与“本科”的匹配。
    """
    if not report_text or not isinstance(report_text, str):
        msg = f"extract_education_from_case_report: report_text 为空或无效: {report_text}"
        logger.info(msg)
        print(msg)
        return None

    marker_pattern = r"一、.+?同志基本情况"
    marker_match = re.search(marker_pattern, report_text, re.DOTALL)

    if marker_match:
        start_pos = marker_match.end()
        # Search a sufficiently large area after the marker for education info
        search_area = report_text[start_pos : start_pos + 1000].lower() # Convert to lowercase early for matching

        # Define education terms and their preferred return values
        # Order them from most specific/full term to less specific/shorter term
        # Prioritize full terms like "大学本科" before "本科"
        education_mappings = {
            "大学本科": "大学本科",
            "本科": "本科",
            "研究生": "研究生",
            "硕士": "硕士",
            "博士": "博士",
            "大专": "大专",
            "高中": "高中",
            "中专": "中专",
            "初中": "初中",
            "小学": "小学"
        }
        
        # Iterate through ordered keys to find the first match
        for term_in_list, return_value in education_mappings.items():
            # Create a regex pattern that looks for the term, optionally followed by "学历", "学位", "毕业"
            # Using \b for word boundary at the start to prevent partial matches
            pattern = r'\b' + re.escape(term_in_list).lower() + r'(?:学历)?(?:学位)?(?:毕业)?'
            if re.search(pattern, search_area):
                msg = f"提取学历 (立案报告): '{return_value}' from text: '{search_area[:100]}...'"
                logger.info(msg)
                print(msg)
                return return_value # Return the standardized value
        
        msg = f"在立案报告的基本情况段落中未找到已知学历信息: {search_area[:100]}..."
        logger.warning(msg)
        print(msg)
        return None
    else:
        msg = f"未找到 '一、XXX同志基本情况' 标记，无法提取立案报告学历: {report_text[:100]}..."
        logger.warning(msg)
        print(msg)
        return None

def extract_ethnicity_from_case_report(report_text):
    """
    从立案报告中提取民族。
    民族位于“一、XXX同志基本情况”后第二个和第三个逗号之间。
    例如：“王xx，男，汉族，1966年12月生”中的“汉族”。
    """
    if not report_text or not isinstance(report_text, str):
        msg = f"extract_ethnicity_from_case_report: report_text 为空或无效: {report_text}"
        logger.info(msg)
        print(msg)
        return None

    marker_pattern = r"一、.+?同志基本情况"
    marker_match = re.search(marker_pattern, report_text, re.DOTALL)

    if marker_match:
        start_pos = marker_match.end()
        # Search area after the marker, limit to prevent matching irrelevant text further down
        # Need to capture enough text to include the ethnicity part
        search_area = report_text[start_pos : start_pos + 300] # Adjust length as needed

        # Split by comma and find the 3rd segment (0-indexed 2nd segment) which contains ethnicity
        # "王xx，男，汉族，1966年12月生，..."
        # 0: 王xx
        # 1: 男
        # 2: 汉族 <-- This is the one
        parts = [p.strip() for p in search_area.split('，')] # Use full-width comma
        
        if len(parts) > 2: # Ensure there are at least 3 parts (index 0, 1, 2)
            ethnicity = parts[2]
            msg = f"提取民族 (立案报告): '{ethnicity}' from text: '{search_area[:50]}...'"
            logger.info(msg)
            print(msg)
            return ethnicity
        else:
            msg = f"立案报告中 '一、同志基本情况' 后面的逗号分隔段不足，无法提取民族: {search_area[:50]}..."
            logger.warning(msg)
            print(msg)
            return None
    else:
        msg = f"未找到 '一、XXX同志基本情况' 标记，无法提取立案报告民族: {report_text[:100]}..."
        logger.warning(msg)
        print(msg)
        return None

def extract_ethnicity_from_decision_report(decision_text):
    """
    从处分决定中提取民族。
    民族位于“关于给予王xx同志党内警告处分的决定”这样字符串（王xx是变量）下面段落里
    第二个逗号和第三个逗号中间的位置。
    例如：“王xx，男，汉族，1966年12月生，山东省平度市xx镇xx村人”中的“汉族”。
    """
    if not decision_text or not isinstance(decision_text, str):
        msg = f"extract_ethnicity_from_decision_report: decision_text 为空或无效: {decision_text}"
        logger.info(msg)
        print(msg)
        return None
    
    title_pattern = r"关于给予.+?同志党内警告处分的决定"
    title_match = re.search(title_pattern, decision_text, re.DOTALL)

    if title_match:
        start_pos = title_match.end()
        search_area = decision_text[start_pos : start_pos + 300] # Adjust length as needed

        # Split by comma and find the 3rd segment (0-indexed 2nd segment) which contains ethnicity
        # "王xx，男，汉族，1966年12月生，..."
        # 0: 王xx
        # 1: 男
        # 2: 汉族 <-- This is the one
        parts = [p.strip() for p in search_area.split('，')]
        
        if len(parts) > 2:
            ethnicity = parts[2]
            msg = f"提取民族 (处分决定): '{ethnicity}' from text: '{search_area[:50]}...'"
            logger.info(msg)
            print(msg)
            return ethnicity
        else:
            msg = f"处分决定中 '关于给予...同志党内警告处分的决定' 后面的逗号分隔段不足，无法提取民族: {search_area[:50]}..."
            logger.warning(msg)
            print(msg)
            return None
    else:
        msg = f"未找到 '关于给予...同志党内警告处分的决定' 标记，无法提取处分决定民族: {decision_text[:100]}..."
        logger.warning(msg)
        print(msg)
        return None

def extract_ethnicity_from_investigation_report(investigation_text):
    """
    从审查调查报告中提取民族。
    民族位于“一、XXX同志基本情况”这样字符串（王xx是变量）下面段落里
    第二个逗号和第三个逗号中间的位置。
    例如：“王xx，男，汉族，1966年12月生，山东省平度市xx镇xx村人”中的“汉族”。
    """
    if not investigation_text or not isinstance(investigation_text, str):
        msg = f"extract_ethnicity_from_investigation_report: investigation_text 为空或无效: {investigation_text}"
        logger.info(msg)
        print(msg)
        return None

    marker_pattern = r"一、.+?同志基本情况"
    marker_match = re.search(marker_pattern, investigation_text, re.DOTALL)

    if marker_match:
        start_pos = marker_match.end()
        search_area = investigation_text[start_pos : start_pos + 300] # Adjust length as needed

        # Split by comma and find the 3rd segment (0-indexed 2nd segment) which contains ethnicity
        # "王xx，男，汉族，1966年12月生，..."
        # 0: 王xx
        # 1: 男
        # 2: 汉族 <-- This is the one
        parts = [p.strip() for p in search_area.split('，')]
        
        if len(parts) > 2:
            ethnicity = parts[2]
            msg = f"提取民族 (审查调查报告): '{ethnicity}' from text: '{search_area[:50]}...'"
            logger.info(msg)
            print(msg)
            return ethnicity
        else:
            msg = f"审查调查报告中 '一、同志基本情况' 后面的逗号分隔段不足，无法提取民族: {search_area[:50]}..."
            logger.warning(msg)
            print(msg)
            return None
    else:
        msg = f"未找到 '一、XXX同志基本情况' 标记，无法提取审查调查报告民族: {investigation_text[:100]}..."
        logger.warning(msg)
        print(msg)
        return None

def extract_ethnicity_from_trial_report(trial_text):
    """
    从审理报告中提取民族。
    民族位于“现将具体情况报告如下”这样字符串下面段落里
    第二个逗号和第三个逗号中间的位置。
    例如：“王xx，男，汉族，1966年12月生，山东省平度市xx镇xx村人”中的“汉族”。
    """
    if not trial_text or not isinstance(trial_text, str):
        msg = f"extract_ethnicity_from_trial_report: trial_text 为空或无效: {trial_text}"
        logger.info(msg)
        print(msg)
        return None
    
    marker = "现将具体情况报告如下"
    marker_pos = trial_text.find(marker)

    if marker_pos != -1:
        start_pos = marker_pos + len(marker)
        search_area = trial_text[start_pos : start_pos + 300] # Adjust length as needed

        # Split by comma and find the 3rd segment (0-indexed 2nd segment) which contains ethnicity
        # "王xx，男，汉族，1966年12月生，..."
        # 0: 王xx
        # 1: 男
        # 2: 汉族 <-- This is the one
        parts = [p.strip() for p in search_area.split('，')]
        
        if len(parts) > 2:
            ethnicity = parts[2]
            msg = f"提取民族 (审理报告): '{ethnicity}' from text: '{search_area[:50]}...'"
            logger.info(msg)
            print(msg)
            return ethnicity
        else:
            msg = f"审理报告中 '现将具体情况报告如下' 后面的逗号分隔段不足，无法提取民族: {search_area[:50]}..."
            logger.warning(msg)
            print(msg)
            return None
    else:
        msg = f"未找到 '现将具体情况报告如下' 标记，无法提取审理报告民族: {trial_text[:100]}..."
        logger.warning(msg)
        print(msg)
        return None

def extract_party_member_from_case_report(report_text):
    """
    从立案报告中提取是否为中共党员。
    若报告中存在“加入中国共产党”则返回“是”，否则返回“否”。
    """
    if not report_text or not isinstance(report_text, str):
        msg = f"extract_party_member_from_case_report: report_text 为空或无效: {report_text}"
        logger.info(msg)
        print(msg)
        return None # Or "否", depending on desired default for empty text

    # 使用正则表达式查找“加入中国共产党”
    # re.IGNORECASE 使得匹配不区分大小写，虽然中文一般不需要
    # re.DOTALL 使得 '.' 匹配包括换行符在内的任何字符
    if re.search(r"加入中国共产党", report_text, re.IGNORECASE | re.DOTALL):
        msg = f"提取是否中共党员 (立案报告): '是' (找到 '加入中国共产党') from text: {report_text[:100]}..."
        logger.info(msg)
        print(msg)
        return "是"
    else:
        msg = f"提取是否中共党员 (立案报告): '否' (未找到 '加入中国共产党') from text: {report_text[:100]}..."
        logger.info(msg)
        print(msg)
        return "否"

def extract_party_member_from_decision_report(decision_text):
    """
    从处分决定中提取是否为中共党员。
    若报告中存在“加入中国共产党”则返回“是”，若存在“群众”则返回“否”，否则返回 None。
    """
    if not decision_text or not isinstance(decision_text, str):
        msg = f"extract_party_member_from_decision_report: decision_text 为空或无效: {decision_text}"
        logger.info(msg)
        print(msg)
        return None

    # 优先匹配“加入中国共产党”
    if re.search(r"加入中国共产党", decision_text, re.IGNORECASE | re.DOTALL):
        msg = f"提取是否中共党员 (处分决定): '是' (找到 '加入中国共产党') from text: {decision_text[:100]}..."
        logger.info(msg)
        print(msg)
        return "是"
    # 如果没有找到“加入中国共产党”，则检查是否为“群众”
    elif re.search(r"群众", decision_text, re.IGNORECASE | re.DOTALL):
        msg = f"提取是否中共党员 (处分决定): '否' (找到 '群众') from text: {decision_text[:100]}..."
        logger.info(msg)
        print(msg)
        return "否"
    else:
        msg = f"提取是否中共党员 (处分决定): 未明确找到党员或群众信息 from text: {decision_text[:100]}..."
        logger.info(msg)
        print(msg)
        return None

def extract_party_joining_date_from_case_report(report_text):
    """
    从立案报告中提取入党时间，并格式化为“YYYY/MM”。
    例如：“1990年1月加入中国共产党”中的“1990年1月”。
    """
    if not report_text or not isinstance(report_text, str):
        msg = f"extract_party_joining_date_from_case_report: report_text 为空或无效: {report_text}"
        logger.info(msg)
        print(msg)
        return None

    # 正则表达式匹配“YYYY年MM月加入中国共产党”或“YYYY年M月加入中国共产党”
    # 捕获组 (d{4}年d{1,2}月)
    pattern = r"(\d{4}年\d{1,2}月)加入中国共产党"
    match = re.search(pattern, report_text)

    if match:
        date_str = match.group(1) # e.g., "1990年1月"
        # Extract year and month
        year_match = re.search(r"(\d{4})年", date_str)
        month_match = re.search(r"(\d{1,2})月", date_str)

        if year_match and month_match:
            year = year_match.group(1)
            month = month_match.group(1).zfill(2) # Pad month with leading zero if single digit
            formatted_date = f"{year}/{month}"
            msg = f"提取入党时间 (立案报告): '{formatted_date}' from text: '{report_text[:100]}...'"
            logger.info(msg)
            print(msg)
            return formatted_date
        else:
            msg = f"在立案报告中找到入党时间段但无法解析年月: '{date_str}'"
            logger.warning(msg)
            print(msg)
            return None
    else:
        msg = f"在立案报告中未找到“加入中国共产党”及其前面的入党时间信息: {report_text[:100]}..."
        logger.info(msg)
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
        # Corrected the syntax error here
        msg = f"未找到 '关于...同志违纪案的审理报告' 标记: {trial_text[:50]}..."
        logger.warning(msg) # Corrected from trial_warning(msg)
        print(msg) # Added print for console output
        return None

def validate_case_relationships(df):
    """Validate relationships between fields in the case registration Excel."""
    mismatch_indices = set() # For name mismatches
    gender_mismatch_indices = set() # For gender mismatches
    age_mismatch_indices = set() # For age mismatches
    birth_date_mismatch_indices = set() # For birth date mismatches
    education_mismatch_indices = set() # For education mismatches
    ethnicity_mismatch_indices = set() # For ethnicity mismatches
    party_member_mismatch_indices = set() # For party member mismatches
    party_joining_date_mismatch_indices = set() # For party joining date mismatches
    issues_list = []
    
    # Define required headers specific to case registration
    # Added "年龄", "出生年月", "学历", "民族", "是否中共党员", "入党时间" to required headers
    required_headers = ["被调查人", "性别", "年龄", "出生年月", "学历", "民族", "是否中共党员", "入党时间", "立案报告", "处分决定", "审查调查报告", "审理报告"]
    if not all(header in df.columns for header in required_headers):
        logger.error(f"Missing required headers for case registration: {required_headers}")
        print(f"缺少必要的表头: {required_headers}") # Added print for console output
        # Update return values to include the new set
        return mismatch_indices, gender_mismatch_indices, age_mismatch_indices, issues_list, birth_date_mismatch_indices, education_mismatch_indices, ethnicity_mismatch_indices, party_member_mismatch_indices, party_joining_date_mismatch_indices

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

        excel_birth_date = str(row["出生年月"]).strip() if pd.notna(row["出生年月"]) else ''
        excel_education = str(row["学历"]).strip() if pd.notna(row["学历"]) else '' 
        excel_ethnicity = str(row["民族"]).strip() if pd.notna(row["民族"]) else '' # Extract Excel Ethnicity
        excel_party_member = str(row["是否中共党员"]).strip() if pd.notna(row["是否中共党员"]) else '' # Extract Excel Party Member
        excel_party_joining_date = str(row["入党时间"]).strip() if pd.notna(row["入党时间"]) else '' # Extract Excel Party Joining Date


        # --- Gender matching rules ---
        # 1) 性别与“立案报告”匹配
        report_text_raw = row["立案报告"] if pd.notna(row["立案报告"]) else '' # Use a raw variable for multiple extractions
        extracted_gender_from_report = extract_gender_from_case_report(report_text_raw)
        if extracted_gender_from_report is None or (excel_gender and excel_gender != extracted_gender_from_report):
            gender_mismatch_indices.add(index)
            issues_list.append((index, "M2性别与BF2立案报告不一致"))
            logger.info(f"行 {index + 1} - 性别不匹配: Excel性别 ('{excel_gender}') vs 立案报告提取性别 ('{extracted_gender_from_report}')")
            print(f"行 {index + 1} - 性别不匹配: Excel性别 ('{excel_gender}') vs 立案报告提取性别 ('{extracted_gender_from_report}')")

        # 2) 性别与“处分决定”匹配
        decision_text_raw = row["处分决定"] if pd.notna(row["处分决定"]) else ''
        extracted_gender_from_decision = extract_gender_from_decision_report(decision_text_raw)
        if extracted_gender_from_decision is None or (excel_gender and excel_gender != extracted_gender_from_decision):
            gender_mismatch_indices.add(index) 
            issues_list.append((index, "M2性别与CU2处分决定不一致"))
            logger.info(f"行 {index + 1} - 性别不匹配: Excel性别 ('{excel_gender}') vs 处分决定提取性别 ('{extracted_gender_from_decision}')")
            print(f"行 {index + 1} - 性别不匹配: Excel性别 ('{excel_gender}') vs 处分决定提取性别 ('{extracted_gender_from_decision}')")

        # 3) 性别与“审查调查报告”匹配
        investigation_text_raw = row["审查调查报告"] if pd.notna(row["审查调查报告"]) else ''
        extracted_gender_from_investigation = extract_gender_from_investigation_report(investigation_text_raw)
        if extracted_gender_from_investigation is None or (excel_gender and excel_gender != extracted_gender_from_investigation):
            gender_mismatch_indices.add(index)
            issues_list.append((index, "M2性别与CX2审查调查报告不一致"))
            logger.info(f"行 {index + 1} - 性别不匹配: Excel性别 ('{excel_gender}') vs 审查调查报告提取性别 ('{extracted_gender_from_investigation}')")
            print(f"行 {index + 1} - 性别不匹配: Excel性别 ('{excel_gender}') vs 审查调查报告提取性别 ('{extracted_gender_from_investigation}')")

        # 4) 性别与“审理报告”匹配
        trial_text_raw = row["审理报告"] if pd.notna(row["审理报告"]) else ''
        extracted_gender_from_trial = extract_gender_from_trial_report(trial_text_raw)
        if extracted_gender_from_trial is None or (excel_gender and excel_gender != extracted_gender_from_trial):
            gender_mismatch_indices.add(index)
            issues_list.append((index, "M2性别与CY2审理报告不一致"))
            logger.info(f"行 {index + 1} - 性别不匹配: Excel性别 ('{excel_gender}') vs 审理报告提取性别 ('{extracted_gender_from_trial}')")
            print(f"行 {index + 1} - 性别不匹配: Excel性别 ('{excel_gender}') vs 审理报告提取性别 ('{extracted_gender_from_trial}')")

        # --- Age matching rules ---
        # 1) 年龄与“立案报告”匹配
        extracted_birth_year_from_report = extract_birth_year_from_case_report(report_text_raw)
        
        calculated_age_from_report = None
        if extracted_birth_year_from_report is not None:
            calculated_age_from_report = current_year - extracted_birth_year_from_report
            logger.info(f"行 {index + 1} - 立案报告计算年龄: {current_year} - {extracted_birth_year_from_report} = {calculated_age_from_report}")
            print(f"行 {index + 1} - 立案报告计算年龄: {current_year} - {extracted_birth_year_from_report} = {calculated_age_from_report}")

        # Check for age mismatch with report
        if (calculated_age_from_report is None) or \
           (excel_age is not None and calculated_age_from_report is not None and excel_age != calculated_age_from_report):
            age_mismatch_indices.add(index)
            issues_list.append((index, "N2年龄与BF2立案报告不一致"))
            logger.info(f"行 {index + 1} - 年龄不匹配: Excel年龄 ('{excel_age}') vs 立案报告计算年龄 ('{calculated_age_from_report}')")
            print(f"行 {index + 1} - 年龄不匹配: Excel年龄 ('{excel_age}') vs 立案报告计算年龄 ('{calculated_age_from_report}')")

        # 2) 年龄与“处分决定”匹配
        extracted_birth_year_from_decision = extract_birth_year_from_decision_report(decision_text_raw)

        calculated_age_from_decision = None
        if extracted_birth_year_from_decision is not None:
            calculated_age_from_decision = current_year - extracted_birth_year_from_decision
            logger.info(f"行 {index + 1} - 处分决定计算年龄: {current_year} - {extracted_birth_year_from_decision} = {calculated_age_from_decision}")
            print(f"行 {index + 1} - 处分决定计算年龄: {current_year} - {extracted_birth_year_from_decision} = {calculated_age_from_decision}")
        
        # Check for age mismatch with decision report
        if (calculated_age_from_decision is None) or \
           (excel_age is not None and calculated_age_from_decision is not None and excel_age != calculated_age_from_decision):
            age_mismatch_indices.add(index)
            issues_list.append((index, "N2年龄与CU2处分决定不一致"))
            logger.info(f"行 {index + 1} - 年龄不匹配: Excel年龄 ('{excel_age}') vs 处分决定计算年龄 ('{calculated_age_from_decision}')")
            print(f"行 {index + 1} - 年龄不匹配: Excel年龄 ('{excel_age}') vs 处分决定计算年龄 ('{calculated_age_from_decision}')")

        # 3) 年龄与“审查调查报告”匹配
        extracted_birth_year_from_investigation = extract_birth_year_from_investigation_report(investigation_text_raw)
        
        calculated_age_from_investigation = None
        if extracted_birth_year_from_investigation is not None:
            calculated_age_from_investigation = current_year - extracted_birth_year_from_investigation
            logger.info(f"行 {index + 1} - 审查调查报告计算年龄: {current_year} - {extracted_birth_year_from_investigation} = {calculated_age_from_investigation}")
            print(f"行 {index + 1} - 审查调查报告计算年龄: {current_year} - {extracted_birth_year_from_investigation} = {calculated_age_from_investigation}")
        
        # Check for age mismatch with investigation report
        if (calculated_age_from_investigation is None) or \
           (excel_age is not None and calculated_age_from_investigation is not None and excel_age != calculated_age_from_investigation):
            age_mismatch_indices.add(index)
            issues_list.append((index, "N2年龄与CX2审查调查报告不一致"))
            logger.info(f"行 {index + 1} - 年龄不匹配: Excel年龄 ('{excel_age}') vs 审查调查报告计算年龄 ('{calculated_age_from_investigation}')")
            print(f"行 {index + 1} - 年龄不匹配: Excel年龄 ('{excel_age}') vs 审查调查报告计算年龄 ('{calculated_age_from_investigation}')")

        # 4) 年龄与“审理报告”匹配
        extracted_birth_year_from_trial = extract_birth_year_from_trial_report(trial_text_raw)

        calculated_age_from_trial = None
        if extracted_birth_year_from_trial is not None:
            calculated_age_from_trial = current_year - extracted_birth_year_from_trial
            logger.info(f"行 {index + 1} - 审理报告计算年龄: {current_year} - {extracted_birth_year_from_trial} = {calculated_age_from_trial}")
            print(f"行 {index + 1} - 审理报告计算年龄: {current_year} - {extracted_birth_year_from_trial} = {calculated_age_from_trial}")
        
        # Check for age mismatch with trial report
        if (calculated_age_from_trial is None) or \
           (excel_age is not None and calculated_age_from_trial is not None and excel_age != calculated_age_from_trial):
            age_mismatch_indices.add(index)
            issues_list.append((index, "N2年龄与CY2审理报告不一致"))
            logger.info(f"行 {index + 1} - 年龄不匹配: Excel年龄 ('{excel_age}') vs 审理报告计算年龄 ('{calculated_age_from_trial}')")
            print(f"行 {index + 1} - 年龄不匹配: Excel年龄 ('{excel_age}') vs 审理报告计算年龄 ('{calculated_age_from_trial}')")

        # --- Birth Date matching rules ---
        # 1) 出生年月与“立案报告”匹配
        extracted_birth_date_from_report = extract_birth_date_from_case_report(report_text_raw)
        
        is_birth_date_mismatch_report = False
        if pd.isna(row["出生年月"]) or excel_birth_date == '': # Excel字段缺失或为空字符串
            if extracted_birth_date_from_report is not None: # 但报告中提取到了
                is_birth_date_mismatch_report = True
        elif extracted_birth_date_from_report is None: # Excel字段有值，但报告中未提取到
            is_birth_date_mismatch_report = True
        elif excel_birth_date != extracted_birth_date_from_report: # 两者都有值但内容不一致
            is_birth_date_mismatch_report = True

        if is_birth_date_mismatch_report:
            birth_date_mismatch_indices.add(index)
            issues_list.append((index, "O2出生年月与BF2立案报告不一致"))
            logger.info(f"行 {index + 1} - 出生年月不匹配: Excel出生年月 ('{excel_birth_date}') vs 立案报告提取出生年月 ('{extracted_birth_date_from_report}')")
            print(f"行 {index + 1} - 出生年月不匹配: Excel出生年月 ('{excel_birth_date}') vs 立案报告提取出生年月 ('{extracted_birth_date_from_report}')")

        # 2) 出生年月与“处分决定”匹配
        extracted_birth_date_from_decision = extract_birth_date_from_decision_report(decision_text_raw)

        is_birth_date_mismatch_decision = False
        if pd.isna(row["出生年月"]) or excel_birth_date == '': # Excel字段缺失或为空字符串
            if extracted_birth_date_from_decision is not None: # 但报告中提取到了
                is_birth_date_mismatch_decision = True
        elif extracted_birth_date_from_decision is None: # Excel字段有值，但报告中未提取到
            is_birth_date_mismatch_decision = True
        elif excel_birth_date != extracted_birth_date_from_decision: # 两者都有值但内容不一致
            is_birth_date_mismatch_decision = True

        if is_birth_date_mismatch_decision:
            birth_date_mismatch_indices.add(index)
            issues_list.append((index, "O2出生年月与CU2处分决定不一致"))
            logger.info(f"行 {index + 1} - 出生年月不匹配: Excel出生年月 ('{excel_birth_date}') vs 处分决定提取出生年月 ('{extracted_birth_date_from_decision}')")
            print(f"行 {index + 1} - 出生年月不匹配: Excel出生年月 ('{excel_birth_date}') vs 处分决定提取出生年月 ('{extracted_birth_date_from_decision}')")

        # 3) 出生年月与“审查调查报告”匹配
        extracted_birth_date_from_investigation = extract_birth_date_from_investigation_report(investigation_text_raw)
        
        is_birth_date_mismatch_investigation = False
        if pd.isna(row["出生年月"]) or excel_birth_date == '': # Excel字段缺失或为空字符串
            if extracted_birth_date_from_investigation is not None: # 但报告中提取到了
                is_birth_date_mismatch_investigation = True
        elif extracted_birth_date_from_investigation is None: # Excel字段有值，但报告中未提取到
            is_birth_date_mismatch_investigation = True
        elif excel_birth_date != extracted_birth_date_from_investigation: # 两者都有值但内容不一致
            is_birth_date_mismatch_investigation = True

        if is_birth_date_mismatch_investigation:
            birth_date_mismatch_indices.add(index)
            issues_list.append((index, "O2出生年月与CX2审查调查报告不一致"))
            logger.info(f"行 {index + 1} - 出生年月不匹配: Excel出生年月 ('{excel_birth_date}') vs 审查调查报告提取出生年月 ('{extracted_birth_date_from_investigation}')")
            print(f"行 {index + 1} - 出生年月不匹配: Excel出生年月 ('{excel_birth_date}') vs 审查调查报告提取出生年月 ('{extracted_birth_date_from_investigation}')")

        # 4) 出生年月与“审理报告”匹配
        extracted_birth_date_from_trial = extract_birth_date_from_trial_report(trial_text_raw)

        is_birth_date_mismatch_trial = False
        if pd.isna(row["出生年月"]) or excel_birth_date == '': # Excel字段缺失或为空字符串
            if extracted_birth_date_from_trial is not None: # 但报告中提取到了
                is_birth_date_mismatch_trial = True
        elif extracted_birth_date_from_trial is None: # Excel字段有值，但报告中未提取到
            is_birth_date_mismatch_trial = True
        elif excel_birth_date != extracted_birth_date_from_trial: # 两者都有值但内容不一致
            is_birth_date_mismatch_trial = True

        if is_birth_date_mismatch_trial:
            birth_date_mismatch_indices.add(index)
            issues_list.append((index, "O2出生年月与CY2审理报告不一致"))
            logger.info(f"行 {index + 1} - 出生年月不匹配: Excel出生年月 ('{excel_birth_date}') vs 审理报告提取出生年月 ('{extracted_birth_date_from_trial}')")
            print(f"行 {index + 1} - 出生年月不匹配: Excel出生年月 ('{excel_birth_date}') vs 审理报告提取出生年月 ('{extracted_birth_date_from_trial}')")

        # --- Education matching rules ---
        # 1) 学历与“立案报告”匹配
        extracted_education_from_report = extract_education_from_case_report(report_text_raw)

        is_education_mismatch_report = False

        # Normalize excel_education for comparison (e.g., "大学本科" vs "本科")
        excel_education_normalized = excel_education # Keep original for logging
        if excel_education == "大学本科":
            excel_education_normalized = "本科" # For comparison, treat "大学本科" as "本科"

        extracted_education_normalized = extracted_education_from_report # Keep original for logging
        if extracted_education_from_report == "大学本科":
            extracted_education_normalized = "本科"
        # If extracted is None, it remains None for comparison.

        if not excel_education: # Case: Excel's education field is empty
            if extracted_education_from_report is not None: # If Excel is empty but report has an education
                is_education_mismatch_report = True
                logger.info(f"行 {index + 1} - 学历不匹配: Excel学历为空，但立案报告中提取到学历 ('{extracted_education_from_report}')。")
                print(f"行 {index + 1} - 学历不匹配: Excel学历为空，但立案报告中提取到学历 ('{extracted_education_from_report}')。")
        else: # Case: Excel's education field has a value
            if extracted_education_from_report is None: # Excel has value, but report has no recognized education
                is_education_mismatch_report = True
                logger.info(f"行 {index + 1} - 学历不匹配: Excel学历 ('{excel_education}') 有值，但立案报告中未提取到学历。")
                print(f"行 {index + 1} - 学历不匹配: Excel学历 ('{excel_education}') 有值，但立案报告中未提取到学历。")
            elif excel_education_normalized != extracted_education_normalized: # Both have values, compare normalized forms
                is_education_mismatch_report = True
                logger.info(f"行 {index + 1} - 学历不匹配: Excel学历 ('{excel_education}') vs 立案报告提取学历 ('{extracted_education_from_report}')。")
                print(f"行 {index + 1} - 学历不匹配: Excel学历 ('{excel_education}') vs 立案报告提取学历 ('{extracted_education_from_report}')")

        if is_education_mismatch_report:
            education_mismatch_indices.add(index)
            issues_list.append((index, "P2学历与BF2立案报告不一致"))
        
        # --- Ethnicity matching rules ---
        # 1) 民族与“立案报告”匹配 (Existing Rule)
        extracted_ethnicity_from_report = extract_ethnicity_from_case_report(report_text_raw)

        is_ethnicity_mismatch_report = False
        if not excel_ethnicity: # Excel field is empty
            if extracted_ethnicity_from_report is not None: # But report has it
                is_ethnicity_mismatch_report = True
                logger.info(f"行 {index + 1} - 民族不匹配: Excel民族为空，但立案报告中提取到民族 ('{extracted_ethnicity_from_report}')。")
                print(f"行 {index + 1} - 民族不匹配: Excel民族为空，但立案报告中提取到民族 ('{extracted_ethnicity_from_report}'))。")
        elif extracted_ethnicity_from_report is None: # Excel has value, but report has no recognized ethnicity
            is_ethnicity_mismatch_report = True
            logger.info(f"行 {index + 1} - 民族不匹配: Excel民族 ('{excel_ethnicity}') 有值，但立案报告中未提取到民族。")
            print(f"行 {index + 1} - 民族不匹配: Excel民族 ('{excel_ethnicity}') 有值，但立案报告中未提取到民族。")
        elif excel_ethnicity != extracted_ethnicity_from_report: # Both have values, but mismatch (exact match)
            is_ethnicity_mismatch_report = True
            logger.info(f"行 {index + 1} - 民族不匹配: Excel民族 ('{excel_ethnicity}') vs 立案报告提取民族 ('{extracted_ethnicity_from_report}')。")
            print(f"行 {index + 1} - 民族不匹配: Excel民族 ('{excel_ethnicity}') vs 立案报告提取民族 ('{extracted_ethnicity_from_report}')")

        if is_ethnicity_mismatch_report:
            ethnicity_mismatch_indices.add(index)
            issues_list.append((index, "Q2民族与BF2立案报告不一致"))
        
        # 2) 民族与“处分决定”匹配 (Existing Rule)
        extracted_ethnicity_from_decision = extract_ethnicity_from_decision_report(decision_text_raw)

        is_ethnicity_mismatch_decision = False
        if not excel_ethnicity: # Excel field is empty
            if extracted_ethnicity_from_decision is not None: # But decision has it
                is_ethnicity_mismatch_decision = True
                logger.info(f"行 {index + 1} - 民族不匹配: Excel民族为空，但处分决定中提取到民族 ('{extracted_ethnicity_from_decision}')。")
                print(f"行 {index + 1} - 民族不匹配: Excel民族为空，但处分决定中提取到民族 ('{extracted_ethnicity_from_decision}')。")
        elif extracted_ethnicity_from_decision is None: # Excel has value, but decision has no recognized ethnicity
            is_ethnicity_mismatch_decision = True
            logger.info(f"行 {index + 1} - 民族不匹配: Excel民族 ('{excel_ethnicity}') 有值，但处分决定中未提取到民族。")
            print(f"行 {index + 1} - 民族不匹配: Excel民族 ('{excel_ethnicity}') 有值，但处分决定中未提取到民族。")
        elif excel_ethnicity != extracted_ethnicity_from_decision: # Both have values, but mismatch (exact match)
            is_ethnicity_mismatch_decision = True
            logger.info(f"行 {index + 1} - 民族不匹配: Excel民族 ('{excel_ethnicity}') vs 处分决定提取民族 ('{extracted_ethnicity_from_decision}')。")
            print(f"行 {index + 1} - 民族不匹配: Excel民族 ('{excel_ethnicity}') vs 处分决定提取民族 ('{extracted_ethnicity_from_decision}')")

        if is_ethnicity_mismatch_decision:
            ethnicity_mismatch_indices.add(index)
            issues_list.append((index, "Q2民族与CU2处分决定不一致"))

        # 3) 民族与“审查调查报告”匹配 (Existing Rule)
        extracted_ethnicity_from_investigation = extract_ethnicity_from_investigation_report(investigation_text_raw)

        is_ethnicity_mismatch_investigation = False
        if not excel_ethnicity: # Excel field is empty
            if extracted_ethnicity_from_investigation is not None: # But investigation report has it
                is_ethnicity_mismatch_investigation = True
                logger.info(f"行 {index + 1} - 民族不匹配: Excel民族为空，但审查调查报告中提取到民族 ('{extracted_ethnicity_from_investigation}')。")
                print(f"行 {index + 1} - 民族不匹配: Excel民族为空，但审查调查报告中提取到民族 ('{extracted_ethnicity_from_investigation}')。")
        elif extracted_ethnicity_from_investigation is None: # Excel has value, but investigation report has no recognized ethnicity
            is_ethnicity_mismatch_investigation = True
            logger.info(f"行 {index + 1} - 民族不匹配: Excel民族 ('{excel_ethnicity}') 有值，但审查调查报告中未提取到民族。")
            print(f"行 {index + 1} - 民族不匹配: Excel民族 ('{excel_ethnicity}') 有值，但审查调查报告中未提取到民族。")
        elif excel_ethnicity != extracted_ethnicity_from_investigation: # Both have values, but mismatch (exact match)
            is_ethnicity_mismatch_investigation = True
            logger.info(f"行 {index + 1} - 民族不匹配: Excel民族 ('{excel_ethnicity}') vs 审查调查报告提取民族 ('{extracted_ethnicity_from_investigation}')。")
            print(f"行 {index + 1} - 民族不匹配: Excel民族 ('{excel_ethnicity}') vs 审查调查报告提取民族 ('{extracted_ethnicity_from_investigation}')")

        if is_ethnicity_mismatch_investigation:
            ethnicity_mismatch_indices.add(index)
            issues_list.append((index, "Q2民族与CX2审查调查报告不一致"))

        # 4) 民族与“审理报告”匹配 (Existing Rule)
        extracted_ethnicity_from_trial = extract_ethnicity_from_trial_report(trial_text_raw)

        is_ethnicity_mismatch_trial = False
        if not excel_ethnicity: # Excel field is empty
            if extracted_ethnicity_from_trial is not None: # But trial report has it
                is_ethnicity_mismatch_trial = True
                logger.info(f"行 {index + 1} - 民族不匹配: Excel民族为空，但审理报告中提取到民族 ('{extracted_ethnicity_from_trial}')。")
                print(f"行 {index + 1} - 民族不匹配: Excel民族为空，但审理报告中提取到民族 ('{extracted_ethnicity_from_trial}')。")
        elif extracted_ethnicity_from_trial is None: # Excel has value, but trial report has no recognized ethnicity
            is_ethnicity_mismatch_trial = True
            logger.info(f"行 {index + 1} - 民族不匹配: Excel民族 ('{excel_ethnicity}') 有值，但审理报告中未提取到民族。")
            print(f"行 {index + 1} - 民族不匹配: Excel民族 ('{excel_ethnicity}') 有值，但审理报告中未提取到民族。")
        elif excel_ethnicity != extracted_ethnicity_from_trial: # Both have values, but mismatch (exact match)
            is_ethnicity_mismatch_trial = True
            logger.info(f"行 {index + 1} - 民族不匹配: Excel民族 ('{excel_ethnicity}') vs 审理报告提取民族 ('{extracted_ethnicity_from_trial}')。")
            print(f"行 {index + 1} - 民族不匹配: Excel民族 ('{excel_ethnicity}') vs 审理报告提取民族 ('{extracted_ethnicity_from_trial}')")

        if is_ethnicity_mismatch_trial:
            ethnicity_mismatch_indices.add(index)
            issues_list.append((index, "Q2民族与CY2审理报告不一致"))

        # --- Party Member matching rules ---
        # 1) "是否中共党员"与“立案报告”匹配 (Existing Rule)
        extracted_party_member_from_report = extract_party_member_from_case_report(report_text_raw)

        is_party_member_mismatch_report = False
        if not excel_party_member: # Excel field is empty
            if extracted_party_member_from_report is not None: # But report has it (e.g., '是' or '否')
                # If Excel is empty, but report determined '是', it's a mismatch.
                # If Excel is empty, and report determined '否', it's NOT a mismatch (empty implies '否').
                if extracted_party_member_from_report == "是":
                    is_party_member_mismatch_report = True
                    logger.info(f"行 {index + 1} - 是否中共党员不匹配: Excel字段为空，但立案报告中提取到“是”。")
                    print(f"行 {index + 1} - 是否中共党员不匹配: Excel字段为空，但立案报告中提取到“是”。")
        elif extracted_party_member_from_report is None: # Excel has value, but report extraction failed
            is_party_member_mismatch_report = True
            logger.info(f"行 {index + 1} - 是否中共党员不匹配: Excel字段 ('{excel_party_member}') 有值，但立案报告中未明确提取到党员信息。")
            print(f"行 {index + 1} - 是否中共党员不匹配: Excel字段 ('{excel_party_member}') 有值，但立案报告中未明确提取到党员信息。")
        elif excel_party_member != extracted_party_member_from_report: # Both have values, but mismatch (exact match)
            is_party_member_mismatch_report = True
            logger.info(f"行 {index + 1} - 是否中共党员不匹配: Excel字段 ('{excel_party_member}') vs 立案报告提取 ('{extracted_party_member_from_report}')。")
            print(f"行 {index + 1} - 是否中共党员不匹配: Excel字段 ('{excel_party_member}') vs 立案报告提取 ('{extracted_party_member_from_report}')。")

        if is_party_member_mismatch_report:
            party_member_mismatch_indices.add(index)
            issues_list.append((index, "T2是否中共党员与BF2立案报告不一致"))

        # 2) 新增：“是否中共党员”与“处分决定”匹配 (Existing Rule)
        extracted_party_member_from_decision = extract_party_member_from_decision_report(decision_text_raw)

        is_party_member_mismatch_decision = False
        if not excel_party_member: # Excel field is empty
            if extracted_party_member_from_decision == "是": # If Excel is empty, but decision determined '是'
                is_party_member_mismatch_decision = True
                logger.info(f"行 {index + 1} - 是否中共党员不匹配: Excel字段为空，但处分决定中提取到“是”。")
                print(f"行 {index + 1} - 是否中共党员不匹配: Excel字段为空，但处分决定中提取到“是”。")
            elif extracted_party_member_from_decision == "否": # If Excel is empty, and decision determined '否', it's consistent
                pass # No mismatch
        elif extracted_party_member_from_decision is None: # Excel has value, but decision extraction failed
            is_party_member_mismatch_decision = True
            logger.info(f"行 {index + 1} - 是否中共党员不匹配: Excel字段 ('{excel_party_member}') 有值，但处分决定中未明确提取到党员信息。")
            print(f"行 {index + 1} - 是否中共党员不匹配: Excel字段 ('{excel_party_member}') 有值，但处分决定中未明确提取到党员信息。")
        elif excel_party_member != extracted_party_member_from_decision: # Both have values, but mismatch (exact match)
            is_party_member_mismatch_decision = True
            logger.info(f"行 {index + 1} - 是否中共党员不匹配: Excel字段 ('{excel_party_member}') vs 处分决定提取 ('{extracted_party_member_from_decision}')。")
            print(f"行 {index + 1} - 是否中共党员不匹配: Excel字段 ('{excel_party_member}') vs 处分决定提取 ('{extracted_party_member_from_decision}')。")

        if is_party_member_mismatch_decision:
            party_member_mismatch_indices.add(index)
            issues_list.append((index, "T2是否中共党员与CU2处分决定不一致"))

        # --- New: 入党时间 (Party Joining Date) matching rule ---
        extracted_party_joining_date_from_report = extract_party_joining_date_from_case_report(report_text_raw)

        is_party_joining_date_mismatch = False

        if excel_party_member == "是":
            # 如果是党员，检查入党时间是否一致或为空
            if not excel_party_joining_date: # Excel字段为空
                if extracted_party_joining_date_from_report is not None: # 报告中有提取到
                    is_party_joining_date_mismatch = True
                    logger.info(f"行 {index + 1} - 入党时间不匹配: Excel入党时间为空，但立案报告中提取到 ('{extracted_party_joining_date_from_report}')。")
                    print(f"行 {index + 1} - 入党时间不匹配: Excel入党时间为空，但立案报告中提取到 ('{extracted_party_joining_date_from_report}')。")
            elif extracted_party_joining_date_from_report is None: # Excel有值，但报告未提取到
                is_party_joining_date_mismatch = True
                logger.info(f"行 {index + 1} - 入党时间不匹配: Excel入党时间 ('{excel_party_joining_date}') 有值，但立案报告中未提取到。")
                print(f"行 {index + 1} - 入党时间不匹配: Excel入党时间 ('{excel_party_joining_date}') 有值，但立案报告中未提取到。")
            elif excel_party_joining_date != extracted_party_joining_date_from_report: # 两者有值但不一致
                is_party_joining_date_mismatch = True
                logger.info(f"行 {index + 1} - 入党时间不匹配: Excel入党时间 ('{excel_party_joining_date}') vs 立案报告提取 ('{extracted_party_joining_date_from_report}')。")
                print(f"行 {index + 1} - 入党时间不匹配: Excel入党时间 ('{excel_party_joining_date}') vs 立案报告提取 ('{extracted_party_joining_date_from_report}')。")
        elif excel_party_member == "否":
            # 如果不是党员，入党时间必须为空
            if excel_party_joining_date: # Excel字段不为空
                is_party_joining_date_mismatch = True
                logger.info(f"行 {index + 1} - 入党时间不匹配: Excel是否中共党员为“否”，但入党时间字段不为空 ('{excel_party_joining_date}')。")
                print(f"行 {index + 1} - 入党时间不匹配: Excel是否中共党员为“否”，但入党时间字段不为空 ('{excel_party_joining_date}')。")
        # else: # excel_party_member is empty or other unexpected value
        #     # Decision on how to handle ambiguous "是否中共党员" field:
        #     # For now, treat as no mismatch if "是否中共党员" is not explicitly "是" or "否".
        #     # Or, you might want to add a rule for this ambiguity itself.
        #     pass

        if is_party_joining_date_mismatch:
            party_joining_date_mismatch_indices.add(index)
            issues_list.append((index, "V2入党时间与BF2立案报告不一致"))


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

    # Important: The return statement now includes the new party_joining_date_mismatch_indices
    return mismatch_indices, gender_mismatch_indices, age_mismatch_indices, issues_list, birth_date_mismatch_indices, education_mismatch_indices, ethnicity_mismatch_indices, party_member_mismatch_indices, party_joining_date_mismatch_indices

def generate_case_files(df, original_filename, upload_dir, mismatch_indices, gender_mismatch_indices, issues_list, age_mismatch_indices, birth_date_mismatch_indices, education_mismatch_indices, ethnicity_mismatch_indices, party_member_mismatch_indices, party_joining_date_mismatch_indices):
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

        # Get actual column indices for relevant fields
        try:
            col_index_investigated_person = df.columns.get_loc("被调查人")
            col_index_gender = df.columns.get_loc("性别")
            col_index_age = df.columns.get_loc("年龄")
            col_index_birth_date = df.columns.get_loc("出生年月")
            col_index_education = df.columns.get_loc("学历") 
            col_index_ethnicity = df.columns.get_loc("民族") 
            col_index_party_member = df.columns.get_loc("是否中共党员") 
            col_index_party_joining_date = df.columns.get_loc("入党时间") # New: Party Joining Date column index
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

            # Highlight "年龄" column
            if idx in age_mismatch_indices: # Highlight mismatched rows for '年龄'
                worksheet.write(idx + 1, col_index_age,
                                df.iloc[idx]["年龄"] if pd.notna(df.iloc[idx]["年龄"]) else '', red_format)

            # Highlight "出生年月" column
            if idx in birth_date_mismatch_indices: # Highlight mismatched rows for '出生年月'
                worksheet.write(idx + 1, col_index_birth_date,
                                df.iloc[idx]["出生年月"] if pd.notna(df.iloc[idx]["出生年月"]) else '', red_format)

            # Highlight "学历" column
            if idx in education_mismatch_indices: # Highlight mismatched rows for '学历'
                worksheet.write(idx + 1, col_index_education,
                                df.iloc[idx]["学历"] if pd.notna(df.iloc[idx]["学历"]) else '', red_format)

            # Highlight "民族" column
            if idx in ethnicity_mismatch_indices: # Highlight mismatched rows for '民族'
                worksheet.write(idx + 1, col_index_ethnicity,
                                df.iloc[idx]["民族"] if pd.notna(df.iloc[idx]["民族"]) else '', red_format)

            # Highlight "是否中共党员" column
            if idx in party_member_mismatch_indices: # Highlight mismatched rows for '是否中共党员'
                worksheet.write(idx + 1, col_index_party_member,
                                df.iloc[idx]["是否中共党员"] if pd.notna(df.iloc[idx]["是否中共党员"]) else '', red_format)

            # Highlight "入党时间" column
            if idx in party_joining_date_mismatch_indices: # Highlight mismatched rows for '入党时间'
                worksheet.write(idx + 1, col_index_party_joining_date,
                                df.iloc[idx]["入党时间"] if pd.notna(df.iloc[idx]["入党时间"]) else '', red_format)


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