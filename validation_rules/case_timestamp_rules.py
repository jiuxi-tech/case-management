import logging
import re
import pandas as pd # 新增导入 pandas
from datetime import datetime
from validation_rules.case_extractors_timestamp import extract_timestamp_from_filing_decision

logger = logging.getLogger(__name__)

def validate_filing_time(df, issues_list, filing_time_mismatch_indices):
    """
    验证“立案时间”字段的规则。
    规则1: 从“立案决定书”字段内容中提取落款时间，如果落款时间存在空格，
           在副本表中将“立案时间”字段标红，并在立案编号表中新增“BG立案决定书落款时间存在空格”。
    规则2: “立案时间”字段内容与“立案决定书”字段内容落款时间进行对比，精确匹配。
           不一致就在副本表中将“立案时间”字段标红，并在立案编号表中新增“AR立案时间与BG立案决定书落款时间不一致”。

    参数:
    df (pd.DataFrame): 原始Excel数据的DataFrame。
    issues_list (list): 包含所有问题的列表，每个问题是一个(索引, 案件编码, 涉案人员编码, 问题描述)元组。
    filing_time_mismatch_indices (set): 收集所有立案时间相关不一致的行索引，用于标红。

    返回:
    None (issues_list 和 filing_time_mismatch_indices 会在函数内部被修改)。
    """
    logger.info("开始验证立案时间规则...")
    print("开始验证立案时间规则...")

    # 获取列名，确保存在
    try:
        col_filing_time = "立案时间"
        col_filing_decision_doc = "立案决定书"
        
        if col_filing_time not in df.columns or col_filing_decision_doc not in df.columns:
            msg = f"缺少必要的列 '{col_filing_time}' 或 '{col_filing_decision_doc}'，跳过立案时间验证。"
            logger.warning(msg)
            print(msg)
            return

    except Exception as e:
        logger.error(f"获取立案时间相关列失败: {e}")
        print(f"获取立案时间相关列失败: {e}")
        return

    for index, row in df.iterrows():
        excel_filing_time_raw = str(row[col_filing_time]).strip() if pd.notna(row[col_filing_time]) else ''
        filing_decision_doc_text = str(row[col_filing_decision_doc]).strip() if pd.notna(row[col_filing_decision_doc]) else ''

        case_code = str(row.get("案件编码", "")).strip()
        person_code = str(row.get("涉案人员编码", "")).strip()

        # 规则1 和 规则2 的提取和比对
        original_filing_decision_timestamp_str, standardized_filing_decision_timestamp = \
            extract_timestamp_from_filing_decision(filing_decision_doc_text)
        
        is_filing_time_mismatch = False # 标记当前行是否有立案时间相关的错误

        # --- 规则1: 检查立案决定书落款时间是否存在空格 ---
        # 这里的检查是基于原始提取到的字符串，而不是标准化后的
        if original_filing_decision_timestamp_str and \
           (re.search(r'\s', original_filing_decision_timestamp_str) or '\xa0' in original_filing_decision_timestamp_str):
            issues_list.append((index, case_code, person_code, "BG立案决定书落款时间存在空格"))
            filing_time_mismatch_indices.add(index)
            is_filing_time_mismatch = True
            logger.info(f"行 {index + 1} - 规则1违规: 立案决定书落款时间存在空格，原始时间为: '{original_filing_decision_timestamp_str}'")
            print(f"行 {index + 1} - 规则1违规: 立案决定书落款时间存在空格，原始时间为: '{original_filing_decision_timestamp_str}'")

        # --- 规则2: “立案时间”字段内容与“立案决定书”字段内容落款时间进行对比 ---
        standardized_excel_filing_time = None
        if excel_filing_time_raw:
            try:
                # 尝试将 Excel 的日期字符串转换为 datetime 对象，然后标准化为YYYY-MM-DD
                # 使用 errors='coerce' 会将无法解析的日期转换为 NaT (Not a Time)
                standardized_excel_filing_time = pd.to_datetime(excel_filing_time_raw, errors='coerce').strftime('%Y-%m-%d')
                logger.debug(f"Excel立案时间 '{excel_filing_time_raw}' 标准化为 '{standardized_excel_filing_time}'")
                print(f"Excel立案时间 '{excel_filing_time_raw}' 标准化为 '{standardized_excel_filing_time}'")
            except Exception as e:
                logger.warning(f"行 {index + 1} - 无法解析Excel立案时间 '{excel_filing_time_raw}': {e}")
                print(f"行 {index + 1} - 无法解析Excel立案时间 '{excel_filing_time_raw}': {e}")
                # 如果解析失败，也视为不一致
                standardized_excel_filing_time = None

        # 进行比对
        if standardized_excel_filing_time and standardized_filing_decision_timestamp:
            if standardized_excel_filing_time != standardized_filing_decision_timestamp:
                issues_list.append((index, case_code, person_code, "AR立案时间与BG立案决定书落款时间不一致"))
                filing_time_mismatch_indices.add(index)
                is_filing_time_mismatch = True
                logger.info(f"行 {index + 1} - 规则2违规: Excel立案时间标准化后 ('{standardized_excel_filing_time}') 与决定书时间标准化后 ('{standardized_filing_decision_timestamp}') 不一致。")
                print(f"行 {index + 1} - 规则2违规: Excel立案时间标准化后 ('{standardized_excel_filing_time}') 与决定书时间标准化后 ('{standardized_filing_decision_timestamp}') 不一致。")
            else:
                logger.info(f"行 {index + 1} - 立案时间 ('{standardized_excel_filing_time}') 与决定书时间 ('{standardized_filing_decision_timestamp}') 一致。")
                print(f"行 {index + 1} - 立案时间 ('{standardized_excel_filing_time}') 与决定书时间 ('{standardized_filing_decision_timestamp}') 一致。")
        elif standardized_excel_filing_time is None and standardized_filing_decision_timestamp is None:
            logger.info(f"行 {index + 1} - Excel立案时间与决定书落款时间均为空或无法解析，跳过对比。")
            print(f"行 {index + 1} - Excel立案时间与决定书落款时间均为空或无法解析，跳过对比。")
        else: # 至少一个有值，但两者不一致或一方无法解析
            issues_list.append((index, case_code, person_code, "AR立案时间与BG立案决定书落款时间不一致（一方缺失或格式问题）"))
            filing_time_mismatch_indices.add(index)
            is_filing_time_mismatch = True
            logger.warning(f"行 {index + 1} - 规则2违规: Excel立案时间标准化后 ('{standardized_excel_filing_time}') 或决定书时间标准化后 ('{standardized_filing_decision_timestamp}') 缺失/格式错误，导致不一致。")
            print(f"行 {index + 1} - 规则2违规: Excel立案时间标准化后 ('{standardized_excel_filing_time}') 或决定书时间标准化后 ('{standardized_filing_decision_timestamp}') 缺失/格式错误，导致不一致。")

    logger.info("立案时间规则验证完成。")
    print("立案时间规则验证完成。")
