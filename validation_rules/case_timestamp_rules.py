import logging
import re
import pandas as pd
from datetime import datetime
from validation_rules.case_extractors_timestamp import extract_timestamp_from_filing_decision
import database # 导入 database 模块

logger = logging.getLogger(__name__)

def validate_filing_time(df, issues_list, filing_time_mismatch_indices, disciplinary_committee_filing_time_mismatch_indices, disciplinary_committee_filing_authority_mismatch_indices): # 新增参数
    """
    验证“立案时间”和“纪委立案时间”字段的规则。
    规则1: 从“立案决定书”字段内容中提取落款时间，如果落款时间存在空格，
           在副本表中将“立案时间”字段标红，并在立案编号表中新增“BG立案决定书落款时间存在空格”。
    规则2: “立案时间”字段内容与“立案决定书”字段内容落款时间进行对比，精确匹配。
           不一致就在副本表中将“立案时间”字段标红，并在立案编号表中新增“AR立案时间与BG立案决定书落款时间不一致”。
    规则3: 当立案决定书字段内容出现“纪立“和”审查“四个字时：
           a) 提取立案决定书字段的落款时间，转换后跟“纪委立案时间”字段值进行对比。
              不一致则标红“纪委立案时间”，并新增问题“AW纪委立案时间和BG立案决定书落款时间不一致“。
           b) 获取“纪委立案机关”和“填报单位名称“字段的值，从 `authority_agency_dict` 查询匹配，
              如果不能检索到记录，则标红“纪委立案机关”，并新增问题“A填报单位名称跟AV纪委立案机关不一致“。

    参数:
    df (pd.DataFrame): 原始Excel数据的DataFrame。
    issues_list (list): 包含所有问题的列表，每个问题是一个(索引, 案件编码, 涉案人员编码, 问题描述)元组。
    filing_time_mismatch_indices (set): 收集所有立案时间相关不一致的行索引，用于标红。
    disciplinary_committee_filing_time_mismatch_indices (set): 收集纪委立案时间不一致的行索引。
    disciplinary_committee_filing_authority_mismatch_indices (set): 收集纪委立案机关不一致的行索引。

    返回:
    None (issues_list 和 mismatch_indices 会在函数内部被修改)。
    """
    logger.info("开始验证立案时间规则...")
    print("开始验证立案时间规则...")

    # 获取列名，确保存在
    try:
        col_filing_time = "立案时间"
        col_filing_decision_doc = "立案决定书"
        col_disciplinary_committee_filing_time = "纪委立案时间" # 新增列
        col_disciplinary_committee_filing_authority = "纪委立案机关" # 新增列
        col_reporting_unit_name = "填报单位名称" # 新增列

        if not all(col in df.columns for col in [col_filing_time, col_filing_decision_doc, 
                                                 col_disciplinary_committee_filing_time, 
                                                 col_disciplinary_committee_filing_authority,
                                                 col_reporting_unit_name]):
            msg = f"缺少必要的列 '{col_filing_time}', '{col_filing_decision_doc}', '{col_disciplinary_committee_filing_time}', '{col_disciplinary_committee_filing_authority}' 或 '{col_reporting_unit_name}'，跳过立案时间相关验证。"
            logger.warning(msg)
            print(msg)
            return

    except Exception as e:
        logger.error(f"获取立案时间相关列失败: {e}")
        print(f"获取立案时间相关列失败: {e}")
        return

    # 从数据库获取机关单位对应表数据
    authority_agency_data = database.get_authority_agency_dict()
    # 将查询结果转换为更方便查找的格式，例如 {(authority, agency, category): True}
    authority_agency_lookup = set()
    for row_db in authority_agency_data:
        authority_agency_lookup.add((row_db['authority'], row_db['agency'], row_db['category']))


    for index, row in df.iterrows():
        excel_filing_time_raw = str(row[col_filing_time]).strip() if pd.notna(row[col_filing_time]) else ''
        filing_decision_doc_text = str(row[col_filing_decision_doc]).strip() if pd.notna(row[col_filing_decision_doc]) else ''
        excel_disciplinary_committee_filing_time_raw = str(row[col_disciplinary_committee_filing_time]).strip() if pd.notna(row[col_disciplinary_committee_filing_time]) else '' # 纪委立案时间
        excel_disciplinary_committee_filing_authority = str(row[col_disciplinary_committee_filing_authority]).strip() if pd.notna(row[col_disciplinary_committee_filing_authority]) else '' # 纪委立案机关
        excel_reporting_unit_name = str(row[col_reporting_unit_name]).strip() if pd.notna(row[col_reporting_unit_name]) else '' # 填报单位名称

        case_code = str(row.get("案件编码", "")).strip()
        person_code = str(row.get("涉案人员编码", "")).strip()

        # 规则1 和 规则2 的提取和比对
        original_filing_decision_timestamp_str, standardized_filing_decision_timestamp = \
            extract_timestamp_from_filing_decision(filing_decision_doc_text)
        
        # --- 规则1: 检查立案决定书落款时间是否存在空格 ---
        if original_filing_decision_timestamp_str and \
           (re.search(r'\s', original_filing_decision_timestamp_str) or '\xa0' in original_filing_decision_timestamp_str):
            issues_list.append((index, case_code, person_code, "BG立案决定书落款时间存在空格"))
            filing_time_mismatch_indices.add(index)
            logger.info(f"行 {index + 1} - 规则1违规: 立案决定书落款时间存在空格，原始时间为: '{original_filing_decision_timestamp_str}'")
            print(f"行 {index + 1} - 规则1违规: 立案决定书落款时间存在空格，原始时间为: '{original_filing_decision_timestamp_str}'")

        # --- 规则2: “立案时间”字段内容与“立案决定书”字段内容落款时间进行对比 ---
        standardized_excel_filing_time = None
        if excel_filing_time_raw:
            try:
                standardized_excel_filing_time = pd.to_datetime(excel_filing_time_raw, errors='coerce').strftime('%Y-%m-%d')
                logger.debug(f"Excel立案时间 '{excel_filing_time_raw}' 标准化为 '{standardized_excel_filing_time}'")
                print(f"Excel立案时间 '{excel_filing_time_raw}' 标准化为 '{standardized_excel_filing_time}'")
            except Exception as e:
                logger.warning(f"行 {index + 1} - 无法解析Excel立案时间 '{excel_filing_time_raw}': {e}")
                print(f"行 {index + 1} - 无法解析Excel立案时间 '{excel_filing_time_raw}': {e}")
                standardized_excel_filing_time = None

        if standardized_excel_filing_time and standardized_filing_decision_timestamp:
            if standardized_excel_filing_time != standardized_filing_decision_timestamp:
                issues_list.append((index, case_code, person_code, "AR立案时间与BG立案决定书落款时间不一致"))
                filing_time_mismatch_indices.add(index)
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
            logger.warning(f"行 {index + 1} - 规则2违规: Excel立案时间标准化后 ('{standardized_excel_filing_time}') 或决定书时间标准化后 ('{standardized_filing_decision_timestamp}') 缺失/格式错误，导致不一致。")
            print(f"行 {index + 1} - 规则2违规: Excel立案时间标准化后 ('{standardized_excel_filing_time}') 或决定书时间标准化后 ('{standardized_filing_decision_timestamp}') 缺失/格式错误，导致不一致。")

        # --- 规则3: 纪委立案时间和纪委立案机关规则 ---
        # 检查“纪立”和“审查”（或“审查调查”）是否存在于立案决定书内容
        if filing_decision_doc_text and \
           ("纪立" in filing_decision_doc_text and ("审查" in filing_decision_doc_text or "审查调查" in filing_decision_doc_text)):
            
            # 规则3a: 提取落款时间并与“纪委立案时间”对比
            # original_filing_decision_timestamp_str 和 standardized_filing_decision_timestamp 已经从上面获取
            standardized_excel_disciplinary_committee_filing_time = None
            if excel_disciplinary_committee_filing_time_raw:
                try:
                    standardized_excel_disciplinary_committee_filing_time = pd.to_datetime(excel_disciplinary_committee_filing_time_raw, errors='coerce').strftime('%Y-%m-%d')
                    logger.debug(f"Excel纪委立案时间 '{excel_disciplinary_committee_filing_time_raw}' 标准化为 '{standardized_excel_disciplinary_committee_filing_time}'")
                    print(f"Excel纪委立案时间 '{excel_disciplinary_committee_filing_time_raw}' 标准化为 '{standardized_excel_disciplinary_committee_filing_time}'")
                except Exception as e:
                    logger.warning(f"行 {index + 1} - 无法解析Excel纪委立案时间 '{excel_disciplinary_committee_filing_time_raw}': {e}")
                    print(f"行 {index + 1} - 无法解析Excel纪委立案时间 '{excel_disciplinary_committee_filing_time_raw}': {e}")
                    standardized_excel_disciplinary_committee_filing_time = None

            if standardized_excel_disciplinary_committee_filing_time and standardized_filing_decision_timestamp:
                if standardized_excel_disciplinary_committee_filing_time != standardized_filing_decision_timestamp:
                    issues_list.append((index, case_code, person_code, "AW纪委立案时间和BG立案决定书落款时间不一致"))
                    disciplinary_committee_filing_time_mismatch_indices.add(index)
                    logger.info(f"行 {index + 1} - 规则3a违规: Excel纪委立案时间标准化后 ('{standardized_excel_disciplinary_committee_filing_time}') 与决定书时间标准化后 ('{standardized_filing_decision_timestamp}') 不一致。")
                    print(f"行 {index + 1} - 规则3a违规: Excel纪委立案时间标准化后 ('{standardized_excel_disciplinary_committee_filing_time}') 与决定书时间标准化后 ('{standardized_filing_decision_timestamp}') 不一致。")
                else:
                    logger.info(f"行 {index + 1} - 纪委立案时间 ('{standardized_excel_disciplinary_committee_filing_time}') 与决定书时间 ('{standardized_filing_decision_timestamp}') 一致。")
                    print(f"行 {index + 1} - 纪委立案时间 ('{standardized_excel_disciplinary_committee_filing_time}') 与决定书时间 ('{standardized_filing_decision_timestamp}') 一致。")
            elif standardized_excel_disciplinary_committee_filing_time is None and standardized_filing_decision_timestamp is None:
                logger.info(f"行 {index + 1} - Excel纪委立案时间与决定书落款时间均为空或无法解析，跳过对比。")
                print(f"行 {index + 1} - Excel纪委立案时间与决定书落款时间均为空或无法解析，跳过对比。")
            else:
                issues_list.append((index, case_code, person_code, "AW纪委立案时间和BG立案决定书落款时间不一致（一方缺失或格式问题）"))
                disciplinary_committee_filing_time_mismatch_indices.add(index)
                logger.warning(f"行 {index + 1} - 规则3a违规: Excel纪委立案时间标准化后 ('{standardized_excel_disciplinary_committee_filing_time}') 或决定书时间标准化后 ('{standardized_filing_decision_timestamp}') 缺失/格式错误，导致不一致。")
                print(f"行 {index + 1} - 规则3a违规: Excel纪委立案时间标准化后 ('{standardized_excel_disciplinary_committee_filing_time}') 或决定书时间标准化后 ('{standardized_filing_decision_timestamp}') 缺失/格式错误，导致不一致。")

            # 规则3b: 查询机关单位对应表
            found_match = False
            # 查找 authority_agency_lookup 中是否存在 (纪委立案机关, 填报单位名称, "NSL")
            if (excel_disciplinary_committee_filing_authority, excel_reporting_unit_name, "NSL") in authority_agency_lookup:
                found_match = True
            
            if not found_match:
                issues_list.append((index, case_code, person_code, "A填报单位名称跟AV纪委立案机关不一致"))
                disciplinary_committee_filing_authority_mismatch_indices.add(index)
                logger.info(f"行 {index + 1} - 规则3b违规: 纪委立案机关 ('{excel_disciplinary_committee_filing_authority}') 和填报单位名称 ('{excel_reporting_unit_name}') 在对应表中未找到匹配记录，或category不是'NSL'。")
                print(f"行 {index + 1} - 规则3b违规: 纪委立案机关 ('{excel_disciplinary_committee_filing_authority}') 和填报单位名称 ('{excel_reporting_unit_name}') 在对应表中未找到匹配记录，或category不是'NSL'。")
            else:
                logger.info(f"行 {index + 1} - 纪委立案机关 ('{excel_disciplinary_committee_filing_authority}') 和填报单位名称 ('{excel_reporting_unit_name}') 在对应表中找到匹配记录。")
                print(f"行 {index + 1} - 纪委立案机关 ('{excel_disciplinary_committee_filing_authority}') 和填报单位名称 ('{excel_reporting_unit_name}') 在对应表中找到匹配记录。")
        else:
            logger.info(f"行 {index + 1} - 立案决定书内容不包含 '纪立' 和 ('审查' 或 '审查调查')，跳过纪委立案时间及机关验证。")
            print(f"行 {index + 1} - 立案决定书内容不包含 '纪立' 和 ('审查' 或 '审查调查')，跳过纪委立案时间及机关验证。")

    logger.info("立案时间规则验证完成。")
    print("立案时间规则验证完成。")
