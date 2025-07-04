# 出生年份和出生年月提取相关的函数

import re
import logging
from datetime import datetime # Required for current_year in original logic, though not directly used in extractors

logger = logging.getLogger(__name__)

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
