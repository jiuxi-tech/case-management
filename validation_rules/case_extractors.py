import re
import logging
import pandas as pd
from datetime import datetime

logger = logging.getLogger(__name__)

# Placeholder for extract_name_from_case_report
# If validation_rules.case_name_extraction is genuinely a separate module,
# this placeholder should be removed, and the import adjusted in case_validators.py.
# For now, it's included here as it was part of the original large file.
try:
    from validation_rules.case_name_extraction import extract_name_from_case_report
except ImportError:
    def extract_name_from_case_report(report_text):
        """
        This is a placeholder function, you need to implement it according to your actual situation.
        It is usually used to extract the name of the person being investigated from the case report.
        """
        if not report_text or not isinstance(report_text, str):
            return None
        match = re.search(r"一、(.+?)同志基本情况", report_text)
        if match:
            return match.group(1).strip()
        return None

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
    
    title_pattern = r"关于给予.+?同志党内警告处分的决定"
    title_match = re.search(title_pattern, decision_text, re.DOTALL)

    if title_match:
        start_pos = title_match.end()
        gender_pattern = r".*?，([^，]+)，"
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

    title_marker = "现将具体情况报告如下"
    marker_pos = trial_text.find(title_marker)

    if marker_pos != -1:
        start_pos = marker_pos + len(title_marker)
        gender_pattern = r".*?，([^，]+)，"
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

    marker_pattern = r"一、.+?同志基本情况"
    marker_match = re.search(marker_pattern, report_text, re.DOTALL)

    if marker_match:
        start_pos = marker_match.end()
        search_area = report_text[start_pos : start_pos + 300]
        parts = [p.strip() for p in search_area.split('，')]
        
        if len(parts) > 3:
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
    
    title_pattern = r"关于给予.+?同志党内警告处分的决定"
    title_match = re.search(title_pattern, decision_text, re.DOTALL)

    if title_match:
        start_pos = title_match.end()
        search_area = decision_text[start_pos : start_pos + 300]
        parts = [p.strip() for p in search_area.split('，')]
        
        if len(parts) > 3:
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

    marker_pattern = r"一、.+?同志基本情况"
    marker_match = re.search(marker_pattern, investigation_text, re.DOTALL)

    if marker_match:
        start_pos = marker_match.end()
        search_area = investigation_text[start_pos : start_pos + 300]
        parts = [p.strip() for p in search_area.split('，')]
        
        if len(parts) > 3:
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

    marker = "现将具体情况报告如下"
    marker_pos = trial_text.find(marker)

    if marker_pos != -1:
        start_pos = marker_pos + len(marker)
        search_area = trial_text[start_pos : start_pos + 300]
        parts = [p.strip() for p in search_area.split('，')]
        
        if len(parts) > 3:
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
            date_match = re.search(r'(\d{4})年(\d{1,2})月', birth_info_segment)
            if date_match:
                year = date_match.group(1)
                month = date_match.group(2).zfill(2)
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
        search_area = report_text[start_pos : start_pos + 1000].lower()

        education_mappings = {
            "大学本科": "大学本科", "本科": "本科", "研究生": "研究生",
            "硕士": "硕士", "博士": "博士", "大专": "大专",
            "高中": "高中", "中专": "中专", "初中": "初中", "小学": "小学"
        }
        
        for term_in_list, return_value in education_mappings.items():
            pattern = r'\b' + re.escape(term_in_list).lower() + r'(?:学历)?(?:学位)?(?:毕业)?'
            if re.search(pattern, search_area):
                msg = f"提取学历 (立案报告): '{return_value}' from text: '{search_area[:100]}...'"
                logger.info(msg)
                print(msg)
                return return_value
        
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
        search_area = report_text[start_pos : start_pos + 300]
        parts = [p.strip() for p in search_area.split('，')]
        
        if len(parts) > 2:
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
        search_area = decision_text[start_pos : start_pos + 300]
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
        search_area = investigation_text[start_pos : start_pos + 300]
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
        search_area = trial_text[start_pos : start_pos + 300]
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
        return None

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

    if re.search(r"加入中国共产党", decision_text, re.IGNORECASE | re.DOTALL):
        msg = f"提取是否中共党员 (处分决定): '是' (找到 '加入中国共产党') from text: {decision_text[:100]}..."
        logger.info(msg)
        print(msg)
        return "是"
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

    pattern = r"(\d{4})年(\d{1,2})月加入中国共产党"
    match = re.search(pattern, report_text)

    if match:
        year = match.group(1)
        month = match.group(2).zfill(2)
        formatted_date = f"{year}/{month}"
        msg = f"提取入党时间 (立案报告): '{formatted_date}' from text: '{report_text[:100]}...'"
        logger.info(msg)
        print(msg)
        return formatted_date
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
        print(msg)
        return None
    
    pattern = r"关于给予(.+?)同志党内警告处分的决定"
    match = re.search(pattern, decision_text)
    if match:
        name = match.group(1).strip()
        msg = f"提取姓名: {name} from decision: {decision_text[:50]}..."
        logger.info(msg)
        print(msg)
        return name
    else:
        msg = f"未找到 '关于给予...同志党内警告处分的决定' 标记: {decision_text[:50]}..."
        logger.warning(msg)
        print(msg)
        return None

def extract_name_from_trial_report(trial_text):
    """从审理报告中提取姓名，基于'关于...同志违纪案的审理报告'标记。"""
    if not trial_text or not isinstance(trial_text, str):
        msg = f"extract_name_from_trial_report: trial_text 为空或无效: {trial_text}"
        logger.info(msg)
        print(msg)
        return None
    
    pattern = r"关于(.+?)同志违纪案的审理报告"
    match = re.search(pattern, trial_text)
    if match:
        name = match.group(1).strip()
        msg = f"提取姓名: {name} from trial report: {trial_text[:50]}..."
        logger.info(msg)
        print(msg)
        return name
    else:
        msg = f"未找到 '关于...同志违纪案的审理报告' 标记: {trial_text[:50]}..."
        logger.warning(msg)
        print(msg)
        return None
