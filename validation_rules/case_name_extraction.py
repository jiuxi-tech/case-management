import logging
import re

logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)

def extract_name_from_case_report(report_text):
    """Extract name from case report based on '一、王xx同志基本情况' marker."""
    if not report_text or not isinstance(report_text, str):
        msg = f"report_text 为空或无效: {report_text}"
        logger.info(msg)
        return None
    
    # 定义姓名的正则表达式，匹配“一、王xx同志基本情况”后的姓名
    pattern = r"一、(.+?)同志基本情况"
    match = re.search(pattern, report_text)
    if match:
        name = match.group(1).strip()
        msg = f"提取姓名: {name} from report: {report_text}"
        logger.info(msg)
        return name
    else:
        msg = f"未找到 '一、...同志基本情况' 标记: {report_text}"
        logger.warning(msg)
        return None