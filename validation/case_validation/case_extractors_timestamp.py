import re
import logging

logger = logging.getLogger(__name__)

def extract_timestamp_from_filing_decision(decision_text):
    """
    从立案决定书内容中提取落款时间，并标准化为“YYYY-MM-DD”格式。
    同时，它返回提取到的原始字符串（用于检查空格）和标准化后的日期。
    """
    if not decision_text or not isinstance(decision_text, str):
        msg = f"extract_timestamp_from_filing_decision: decision_text is empty or invalid: {decision_text}"
        logger.info(msg)
        print(msg) # Added print statement
        return None, None # 返回两个None

    # 匹配“YYYY年M月D日”或“YYYY年MM月DD日”等形式的日期
    # 注意：这里的日期模式需要与您的实际文本相符。
    pattern = r'(\d{4})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日'
    
    match = re.search(pattern, decision_text)
    if match:
        original_matched_string = match.group(0) # 原始匹配到的日期字符串，用于检查空格
        year = match.group(1)
        month = match.group(2)
        day = match.group(3)
        
        # 格式化为标准YYYY-MM-DD 形式
        # 使用 f-string 的格式化功能，例如 {int(month):02d} 会将 3 格式化为 03
        standardized_date = f"{year}-{int(month):02d}-{int(day):02d}"
        
        msg = f"extract_timestamp_from_filing_decision: Extracted original '{original_matched_string}', standardized to '{standardized_date}'."
        logger.info(msg)
        print(msg) # Added print statement
        return original_matched_string, standardized_date
    else:
        msg = f"extract_timestamp_from_filing_decision: No timestamp found in filing decision document: {decision_text[:100]}..."
        logger.warning(msg)
        print(msg) # Added print statement
        return None, None # 返回两个None
