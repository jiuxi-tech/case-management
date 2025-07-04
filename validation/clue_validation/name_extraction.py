import logging
import re
import pandas as pd
from config import Config

logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)

def extract_name_from_report(report_text):
    if not report_text or pd.isna(report_text):
        msg = f"report_text 为空或无效: {report_text}"
        logger.info(msg)
        print(msg)
        return None
    start_marker = "（一）被反映人基本情况"
    start_idx = report_text.find(start_marker)
    if start_idx == -1:
        msg = f"未找到 '（一）被反映人基本情况' 标记: {report_text}"
        logger.warning(msg)
        print(msg)
        return None
    start_idx += len(start_marker)
    msg = f"原始报告文本: {report_text}"
    logger.debug(msg)
    print(msg)
    next_newline_idx = report_text.find("\n", start_idx)
    if next_newline_idx == -1:
        next_newline_idx = len(report_text)
    paragraph = report_text[start_idx:next_newline_idx].strip()
    if not paragraph:
        next_newline_idx = report_text.find("\n", start_idx)
        if next_newline_idx == -1:
            next_newline_idx = len(report_text)
        next_paragraph_start = report_text.find("\n", next_newline_idx + 1)
        if next_paragraph_start == -1:
            next_paragraph_start = len(report_text)
        paragraph = report_text[next_newline_idx + 1:next_paragraph_start].strip()
    msg = f"提取的段落: {paragraph}"
    logger.debug(msg)
    print(msg)
    end_idx = -1
    for char in [",", "，"]:
        temp_idx = paragraph.find(char)
        if temp_idx != -1 and (end_idx == -1 or temp_idx < end_idx):
            end_idx = temp_idx
    if end_idx == -1:
        msg = f"未找到逗号: {paragraph}"
        logger.warning(msg)
        print(msg)
        return None
    name = paragraph[:end_idx].strip()
    msg = f"提取姓名: {name} from paragraph: {paragraph}"
    logger.info(msg)
    print(msg)
    return name if name else None