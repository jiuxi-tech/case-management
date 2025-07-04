import re
import logging

logger = logging.getLogger(__name__)

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

def extract_suspected_violation_from_case_report(report_text):
    """
    从立案报告中提取“涉嫌违纪问题”段落。
    段落位置在“二、涉嫌违反工作纪律的问题”到“三、意见建议”之间。
    """
    if not report_text or not isinstance(report_text, str):
        msg = f"extract_suspected_violation_from_case_report: report_text 为空或无效: {report_text}"
        logger.info(msg)
        print(msg)
        return None

    # 匹配开始和结束标记，使用 re.DOTALL 确保 . 匹配换行符
    # 匹配 "二、涉嫌违反工作纪律的问题" 或类似开头，到 "三、意见建议" 之间的内容
    pattern = r"二、涉嫌违反[\s\S]+?的问题([\s\S]*?)三、意见建议"
    match = re.search(pattern, report_text, re.DOTALL)

    if match:
        extracted_text = match.group(1).strip()
        # 清理多余的空白符，包括换行符和制表符
        cleaned_text = re.sub(r'\s+', '', extracted_text)
        msg = f"提取涉嫌违纪问题 (立案报告): '{cleaned_text[:100]}...' from case report"
        logger.info(msg)
        print(msg)
        return cleaned_text
    else:
        msg = f"未找到 '涉嫌违纪问题' 段落 (立案报告): {report_text[:100]}..."
        logger.warning(msg)
        print(msg)
        return None

def extract_suspected_violation_from_decision(decision_text, investigated_person_name_from_excel=None):
    """
    从处分决定中提取“涉嫌违纪问题”段落。
    段落位置在“经审查，XXX存在以下违纪问题。”之后，到“XXX同志身为中共党员”之前。
    会动态地从文本中识别出“违纪问题”段落中的姓名进行匹配，而不是强制使用Excel姓名。
    """
    if not decision_text or not isinstance(decision_text, str):
        msg = f"extract_suspected_violation_from_decision: decision_text 为空或无效: {decision_text}"
        logger.info(msg)
        print(msg)
        return None

    # 第一步：尝试从“经审查，XXX存在以下违纪问题。”中提取出实际使用的姓名
    # 这里的 .+? 会匹配任何字符直到“存在以下违纪问题”
    start_name_pattern = r"经审查，(.+?)存在以下违纪问题。"
    start_name_match = re.search(start_name_pattern, decision_text)

    actual_violation_name = None
    if start_name_match:
        actual_violation_name = start_name_match.group(1).strip()
        logger.info(f"extract_suspected_violation_from_decision: 从起始标记中提取到姓名：'{actual_violation_name}'")
        print(f"extract_suspected_violation_from_decision: 从起始标记中提取到姓名：'{actual_violation_name}'")
    else:
        msg = f"extract_suspected_violation_from_decision: 未找到起始标记 '经审查，XXX存在以下违纪问题。'，无法提取简要案情。原始文本前100字: '{decision_text[:100]}...'"
        logger.warning(msg)
        print(msg)
        return None

    # 如果未能从起始标记中提取到姓名，则返回 None
    if not actual_violation_name:
        msg = f"extract_suspected_violation_from_decision: 无法从起始标记中提取有效姓名，返回 None。"
        logger.warning(msg)
        print(msg)
        return None

    # 构建正则表达式：
    # 起始标记：精确匹配 "经审查，[动态捕获到的姓名]存在以下违纪问题。"
    # 内容：([\s\S]*?) 非贪婪匹配所有字符
    # 结束标记：动态匹配 "[动态捕获到的姓名]同志身为中共党员" 或 "本处分决定自" 或 "主送：" 或字符串结束
    # 使用 re.escape 确保姓名中的特殊字符被正确处理
    pattern = rf"经审查，{re.escape(actual_violation_name)}存在以下违纪问题。([\s\S]*?)(?:{re.escape(actual_violation_name)}同志身为中共党员|本处分决定自|主送：|\Z)"
    
    logger.info(f"extract_suspected_violation_from_decision: 最终使用的正则表达式模式：'{pattern}'")
    print(f"extract_suspected_violation_from_decision: 最终使用的正则表达式模式：'{pattern}'")
    # logger.debug(f"extract_suspected_violation_from_decision: 处分决定原始文本：\n{decision_text}") # 打印完整文本以便调试
    # print(f"extract_suspected_violation_from_decision: 处分决定原始文本：\n{decision_text}") # 打印完整文本以便调试

    match = re.search(pattern, decision_text, re.DOTALL)

    if match:
        extracted_text = match.group(1).strip()
        # 清理多余的空白符，包括换行符和制表符
        cleaned_text = re.sub(r'\s+', '', extracted_text)
        msg = f"提取涉嫌违纪问题 (处分决定) 成功: '{cleaned_text[:100]}...' (使用姓名 '{actual_violation_name}')"
        logger.info(msg)
        print(msg)
        return cleaned_text
    else:
        # Debugging: Log if the pattern or text mismatch
        msg = f"未找到 '{actual_violation_name}' 涉嫌违纪问题段落 (处分决定)。\n" \
              f"尝试的模式: '{pattern}'\n" \
              f"原始文本前200字: '{decision_text[:200]}...'"
        logger.warning(msg)
        print(msg)
        return None
