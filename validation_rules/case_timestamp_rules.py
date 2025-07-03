import logging
import re
import pandas as pd
from datetime import datetime
from validation_rules.case_extractors_timestamp import extract_timestamp_from_filing_decision
import database # 导入 database 模块

logger = logging.getLogger(__name__)

def validate_filing_time(df, issues_list, filing_time_mismatch_indices, disciplinary_committee_filing_time_mismatch_indices, supervisory_committee_filing_time_mismatch_indices, disciplinary_committee_filing_authority_mismatch_indices, supervisory_committee_filing_authority_mismatch_indices):
    """
    验证“立案时间”、“纪委立案时间”和“监委立案时间”字段的规则。
    规则1: 从“立案决定书”字段内容中提取落款时间，如果落款时间存在空格，
            在副本表中将“立案时间”字段标红，并在立案编号表中新增“BG立案决定书落款时间存在空格”。
    规则2: “立案时间”字段内容与“立案决定书”字段内容落款时间进行对比，精准匹配。
            不一致就在副本表中将“立案时间”字段标红，并在立案编号表中新增“AR立案时间与BG立案决定书落款时间不一致”。
    规则3: 当立案决定书字段内容出现“纪立“和”审查“四个字时：
            a) 提取立案决定书字段的落款时间，转换后跟“纪委立案时间”字段值进行对比。
                不一致则标红“纪委立案时间”，并新增问题“AW纪委立案时间和BG立案决定书落款时间不一致“。
            b) 获取“纪委立案机关”和“填报单位名称“字段的值，从 `authority_agency_dict` 查询匹配，
                如果不能检索到记录，则标红“纪委立案机关”，并新增问题“A填报单位名称跟AV纪委立案机关不一致“。
    规则4: 当立案决定书字段内容出现“监立“和”调查“四个字时：
            a) 提取立案决定书字段的落款时间，转换后跟“监委立案时间”字段值进行对比。
                不一致则标红“监委立案时间”，并新增问题“AZ监委立案时间和BG立案决定书落款时间不一致“。
            b) 获取“监委立案机关”和“填报单位名称“字段的值，从 `authority_agency_dict` 查询匹配，
                如果不能检索到记录，则标红“监委立案机关”，并新增问题“A填报单位名称跟AZ监委立案机关不一致“。
    规则5: 当立案决定书字段内容出现“纪监立“和”审查调查“七个字时：
            a) 提取立案决定书字段的落款时间，转换后跟“监委立案时间”和“纪委立案时间”同时进行对比，精准匹配。
                如果不一致则同时标红“监委立案时间”和“纪委立案时间”，并新增问题“AZ监委立案时间、AW纪委立案时间和BG立案决定书落款时间不一致“。
            b) 获取“监委立案机关”和“填报单位名称“字段的值，从 `authority_agency_dict` 查询匹配，
                如果不能检索到记录，则标红“监委立案机关”，并新增问题“A填报单位名称跟AZ监委立案机关不一致“和“A填报单位名称跟AV纪委立案机关不一致“。
            c) 获取“纪委立案机关”和“填报单位名称“字段的值，从 `authority_agency_dict` 查询匹配，
                如果不能检索到记录，则标红“纪委立案机关”，并新增问题“A填报单位名称跟AV纪委立案机关不一致“。

    参数:
    df (pd.DataFrame): 原始Excel数据的DataFrame。
    issues_list (list): 包含所有问题的列表，每个问题是一个(索引, 案件编码, 涉案人员编码, 问题描述)元组。
    filing_time_mismatch_indices (set): 收集所有立案时间相关不一致的行索引，用于标红。
    disciplinary_committee_filing_time_mismatch_indices (set): 收集纪委立案时间不一致的行索引。
    disciplinary_committee_filing_authority_mismatch_indices (set): 收集纪委立案机关不一致的行索引。
    supervisory_committee_filing_time_mismatch_indices (set): 收集监委立案时间不一致的行索引。
    supervisory_committee_filing_authority_mismatch_indices (set): 收集监委立案机关不一致的行索引。

    返回:
    None (issues_list 和 mismatch_indices 会在函数内部被修改)。
    """
    logger.info("开始验证立案时间相关规则...")
    print("开始验证立案时间相关规则...")

    # 获取列名，确保存在
    try:
        col_filing_time = "立案时间"
        col_filing_decision_doc = "立案决定书"
        col_disciplinary_committee_filing_time = "纪委立案时间"
        col_disciplinary_committee_filing_authority = "纪委立案机关"
        col_supervisory_committee_filing_time = "监委立案时间"
        col_supervisory_committee_filing_authority = "监委立案机关"
        col_reporting_unit_name = "填报单位名称"

        required_cols = [
            col_filing_time, col_filing_decision_doc, 
            col_disciplinary_committee_filing_time, col_disciplinary_committee_filing_authority,
            col_supervisory_committee_filing_time, col_supervisory_committee_filing_authority,
            col_reporting_unit_name
        ]
        if not all(col in df.columns for col in required_cols):
            msg = f"缺少必要的列 {required_cols} 中至少一个，跳过立案时间相关验证。"
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
        excel_disciplinary_committee_filing_time_raw = str(row[col_disciplinary_committee_filing_time]).strip() if pd.notna(row[col_disciplinary_committee_filing_time]) else ''
        excel_disciplinary_committee_filing_authority = str(row[col_disciplinary_committee_filing_authority]).strip() if pd.notna(row[col_disciplinary_committee_filing_authority]) else ''
        excel_supervisory_committee_filing_time_raw = str(row[col_supervisory_committee_filing_time]).strip() if pd.notna(row[col_supervisory_committee_filing_time]) else ''
        excel_supervisory_committee_filing_authority = str(row[col_supervisory_committee_filing_authority]).strip() if pd.notna(row[col_supervisory_committee_filing_authority]) else ''
        excel_reporting_unit_name = str(row[col_reporting_unit_name]).strip() if pd.notna(row[col_reporting_unit_name]) else ''

        case_code = str(row.get("案件编码", "")).strip()
        person_code = str(row.get("涉案人员编码", "")).strip()

        # 提取立案决定书落款时间
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
        else:
            issues_list.append((index, case_code, person_code, "AR立案时间与BG立案决定书落款时间不一致（一方缺失或格式问题）"))
            filing_time_mismatch_indices.add(index)
            logger.warning(f"行 {index + 1} - 规则2违规: Excel立案时间标准化后 ('{standardized_excel_filing_time}') 或决定书时间标准化后 ('{standardized_filing_decision_timestamp}') 缺失/格式错误，导致不一致。")
            print(f"行 {index + 1} - 规则2违规: Excel立案时间标准化后 ('{standardized_excel_filing_time}') 或决定书时间标准化后 ('{standardized_filing_decision_timestamp}') 缺失/格式错误，导致不一致。")

        # --- 规则3: 纪委立案时间和纪委立案机关规则 ---
        # 检查“纪立”和“审查”（或“审查调查”）是否存在于立案决定书内容
        if filing_decision_doc_text and \
           ("纪立" in filing_decision_doc_text and ("审查" in filing_decision_doc_text or "审查调查" in filing_decision_doc_text)):
            logger.info(f"行 {index + 1} - 立案决定书内容包含 '纪立' 和 '审查' 或 '审查调查'。执行规则3。")
            print(f"行 {index + 1} - 立案决定书内容包含 '纪立' 和 '审查' 或 '审查调查'。执行规则3。")

            # 规则3a: 提取落款时间并与“纪委立案时间”对比
            standardized_excel_disciplinary_committee_filing_time = None
            if excel_disciplinary_committee_filing_time_raw:
                try:
                    standardized_excel_disciplinary_committee_filing_time = pd.to_datetime(excel_disciplinary_committee_filing_time_raw, errors='coerce').strftime('%Y-%m-%d')
                    logger.debug(f"Excel纪委立案时间 '{excel_disciplinary_committee_filing_time_raw}' 标准化为 '{standardized_excel_disciplinary_committee_filing_time}'")
                    print(f"Excel纪委立案时间 '{excel_disciplinary_committee_filing_time_raw}' 标准化为 '{standardized_excel_disciplinary_committee_filing_time}'")
                except Exception as e:
                    logger.warning(f"行 {index + 1} - 规则3a: 无法解析Excel纪委立案时间 '{excel_disciplinary_committee_filing_time_raw}': {e}")
                    print(f"行 {index + 1} - 规则3a: 无法解析Excel纪委立案时间 '{excel_disciplinary_committee_filing_time_raw}': {e}")
                    standardized_excel_disciplinary_committee_filing_time = None

            if standardized_excel_disciplinary_committee_filing_time and standardized_filing_decision_timestamp:
                if standardized_excel_disciplinary_committee_filing_time != standardized_filing_decision_timestamp:
                    issues_list.append((index, case_code, person_code, "AW纪委立案时间和BG立案决定书落款时间不一致"))
                    disciplinary_committee_filing_time_mismatch_indices.add(index)
                    logger.info(f"行 {index + 1} - 规则3a违规: Excel纪委立案时间标准化后 ('{standardized_excel_disciplinary_committee_filing_time}') 与决定书时间标准化后 ('{standardized_filing_decision_timestamp}') 不一致。")
                    print(f"行 {index + 1} - 规则3a违规: Excel纪委立案时间标准化后 ('{standardized_excel_disciplinary_committee_filing_time}') 与决定书时间标准化后 ('{standardized_filing_decision_timestamp}') 不一致。")
                else:
                    logger.info(f"行 {index + 1} - 规则3a: 纪委立案时间 ('{standardized_excel_disciplinary_committee_filing_time}') 与决定书时间 ('{standardized_filing_decision_timestamp}') 一致。")
                    print(f"行 {index + 1} - 规则3a: 纪委立案时间 ('{standardized_excel_disciplinary_committee_filing_time}') 与决定书时间 ('{standardized_filing_decision_timestamp}') 一致。")
            elif standardized_excel_disciplinary_committee_filing_time is None and standardized_filing_decision_timestamp is None:
                logger.info(f"行 {index + 1} - 规则3a: Excel纪委立案时间与决定书落款时间均为空或无法解析，跳过对比。")
                print(f"行 {index + 1} - 规则3a: Excel纪委立案时间与决定书落款时间均为空或无法解析，跳过对比。")
            else:
                issues_list.append((index, case_code, person_code, "AW纪委立案时间和BG立案决定书落款时间不一致（一方缺失或格式问题）"))
                disciplinary_committee_filing_time_mismatch_indices.add(index)
                logger.warning(f"行 {index + 1} - 规则3a违规: Excel纪委立案时间标准化后 ('{standardized_excel_disciplinary_committee_filing_time}') 或决定书时间标准化后 ('{standardized_filing_decision_timestamp}') 缺失/格式错误，导致不一致。")
                print(f"行 {index + 1} - 规则3a违规: Excel纪委立案时间标准化后 ('{standardized_excel_disciplinary_committee_filing_time}') 或决定书时间标准化后 ('{standardized_filing_decision_timestamp}') 缺失/格式错误，导致不一致。")

            # 规则3b: 查询机关单位对应表
            found_match_disciplinary = False
            if (excel_disciplinary_committee_filing_authority, excel_reporting_unit_name, "NSL") in authority_agency_lookup:
                found_match_disciplinary = True
            
            if not found_match_disciplinary:
                issues_list.append((index, case_code, person_code, "A填报单位名称跟AV纪委立案机关不一致"))
                disciplinary_committee_filing_authority_mismatch_indices.add(index)
                logger.info(f"行 {index + 1} - 规则3b违规: 纪委立案机关 ('{excel_disciplinary_committee_filing_authority}') 和填报单位名称 ('{excel_reporting_unit_name}') 在对应表中未找到匹配记录，或category不是'NSL'。")
                print(f"行 {index + 1} - 规则3b违规: 纪委立案机关 ('{excel_disciplinary_committee_filing_authority}') 和填报单位名称 ('{excel_reporting_unit_name}') 在对应表中未找到匹配记录，或category不是'NSL'。")
            else:
                logger.info(f"行 {index + 1} - 规则3b: 纪委立案机关 ('{excel_disciplinary_committee_filing_authority}') 和填报单位名称 ('{excel_reporting_unit_name}') 在对应表中找到匹配记录。")
                print(f"行 {index + 1} - 规则3b: 纪委立案机关 ('{excel_disciplinary_committee_filing_authority}') 和填报单位名称 ('{excel_reporting_unit_name}') 在对应表中找到匹配记录。")
        else:
            logger.info(f"行 {index + 1} - 立案决定书内容不包含 '纪立' 和 ('审查' 或 '审查调查')，跳过纪委立案时间及机关验证。")
            print(f"行 {index + 1} - 立案决定书内容不包含 '纪立' 和 ('审查' 或 '审查调查')，跳过纪委立案时间及机关验证。")


        # --- 规则4: 监委立案时间和监委立案机关规则 ---
        # 检查“监立”和“调查”是否存在于立案决定书内容
        if filing_decision_doc_text and \
           ("监立" in filing_decision_doc_text and "调查" in filing_decision_doc_text):
            logger.info(f"行 {index + 1} - 立案决定书内容包含 '监立' 和 '调查'。执行规则4。")
            print(f"行 {index + 1} - 立案决定书内容包含 '监立' 和 '调查'。执行规则4。")

            # 规则4a: 提取落款时间并与“监委立案时间”对比
            standardized_excel_supervisory_committee_filing_time = None
            if excel_supervisory_committee_filing_time_raw:
                try:
                    standardized_excel_supervisory_committee_filing_time = pd.to_datetime(excel_supervisory_committee_filing_time_raw, errors='coerce').strftime('%Y-%m-%d')
                    logger.debug(f"Excel监委立案时间 '{excel_supervisory_committee_filing_time_raw}' 标准化为 '{standardized_excel_supervisory_committee_filing_time}'")
                    print(f"Excel监委立案时间 '{excel_supervisory_committee_filing_time_raw}' 标准化为 '{standardized_excel_supervisory_committee_filing_time}'")
                except Exception as e:
                    logger.warning(f"行 {index + 1} - 规则4a: 无法解析Excel监委立案时间 '{excel_supervisory_committee_filing_time_raw}': {e}")
                    print(f"行 {index + 1} - 规则4a: 无法解析Excel监委立案时间 '{excel_supervisory_committee_filing_time_raw}': {e}")
                    standardized_excel_supervisory_committee_filing_time = None

            if standardized_excel_supervisory_committee_filing_time and standardized_filing_decision_timestamp:
                if standardized_excel_supervisory_committee_filing_time != standardized_filing_decision_timestamp:
                    issues_list.append((index, case_code, person_code, "AZ监委立案时间和BG立案决定书落款时间不一致"))
                    supervisory_committee_filing_time_mismatch_indices.add(index)
                    logger.info(f"行 {index + 1} - 规则4a违规: Excel监委立案时间标准化后 ('{standardized_excel_supervisory_committee_filing_time}') 与决定书时间标准化后 ('{standardized_filing_decision_timestamp}') 不一致。")
                    print(f"行 {index + 1} - 规则4a违规: Excel监委立案时间标准化后 ('{standardized_excel_supervisory_committee_filing_time}') 与决定书时间标准化后 ('{standardized_filing_decision_timestamp}') 不一致。")
                else:
                    logger.info(f"行 {index + 1} - 规则4a: 监委立案时间 ('{standardized_excel_supervisory_committee_filing_time}') 与决定书时间 ('{standardized_filing_decision_timestamp}') 一致。")
                    print(f"行 {index + 1} - 规则4a: 监委立案时间 ('{standardized_excel_supervisory_committee_filing_time}') 与决定书时间 ('{standardized_filing_decision_timestamp}') 一致。")
            elif standardized_excel_supervisory_committee_filing_time is None and standardized_filing_decision_timestamp is None:
                logger.info(f"行 {index + 1} - 规则4a: Excel监委立案时间与决定书落款时间均为空或无法解析，跳过对比。")
                print(f"行 {index + 1} - 规则4a: Excel监委立案时间与决定书落款时间均为空或无法解析，跳过对比。")
            else:
                issues_list.append((index, case_code, person_code, "AZ监委立案时间和BG立案决定书落款时间不一致（一方缺失或格式问题）"))
                supervisory_committee_filing_time_mismatch_indices.add(index)
                logger.warning(f"行 {index + 1} - 规则4a违规: Excel监委立案时间标准化后 ('{standardized_excel_supervisory_committee_filing_time}') 或决定书时间标准化后 ('{standardized_filing_decision_timestamp}') 缺失/格式错误，导致不一致。")
                print(f"行 {index + 1} - 规则4a违规: Excel监委立案时间标准化后 ('{standardized_excel_supervisory_committee_filing_time}') 或决定书时间标准化后 ('{standardized_filing_decision_timestamp}') 缺失/格式错误，导致不一致。")

            # 规则4b: 查询机关单位对应表
            found_match_supervisory = False
            if (excel_supervisory_committee_filing_authority, excel_reporting_unit_name, "NSL") in authority_agency_lookup:
                found_match_supervisory = True
            
            if not found_match_supervisory:
                issues_list.append((index, case_code, person_code, "A填报单位名称跟AZ监委立案机关不一致"))
                supervisory_committee_filing_authority_mismatch_indices.add(index)
                logger.info(f"行 {index + 1} - 规则4b违规: 监委立案机关 ('{excel_supervisory_committee_filing_authority}') 和填报单位名称 ('{excel_reporting_unit_name}') 在对应表中未找到匹配记录，或category不是'NSL'。")
                print(f"行 {index + 1} - 规则4b违规: 监委立案机关 ('{excel_supervisory_committee_filing_authority}') 和填报单位名称 ('{excel_reporting_unit_name}') 在对应表中未找到匹配记录，或category不是'NSL'。")
            else:
                logger.info(f"行 {index + 1} - 规则4b: 监委立案机关 ('{excel_supervisory_committee_filing_authority}') 和填报单位名称 ('{excel_reporting_unit_name}') 在对应表中找到匹配记录。")
                print(f"行 {index + 1} - 规则4b: 监委立案机关 ('{excel_supervisory_committee_filing_authority}') 和填报单位名称 ('{excel_reporting_unit_name}') 在对应表中找到匹配记录。")
        else:
            logger.info(f"行 {index + 1} - 立案决定书内容不包含 '监立' 和 '调查'，跳过监委立案时间及机关验证。")
            print(f"行 {index + 1} - 立案决定书内容不包含 '监立' 和 '调查'，跳过监委立案时间及机关验证。")

        # --- 规则5: “纪监立”和“审查调查”规则 (新增) ---
        if filing_decision_doc_text and \
           ("纪监立" in filing_decision_doc_text and "审查调查" in filing_decision_doc_text):
            logger.info(f"行 {index + 1} - 立案决定书内容包含 '纪监立' 和 '审查调查'。执行规则5。")
            print(f"行 {index + 1} - 立案决定书内容包含 '纪监立' 和 '审查调查'。执行规则5。")

            # 规则5a: 提取落款时间并与“监委立案时间”和“纪委立案时间”同时进行对比
            mismatch_found_rule_5a = False
            
            # 标准化纪委立案时间
            standardized_excel_disciplinary_committee_filing_time = None
            if excel_disciplinary_committee_filing_time_raw:
                try:
                    standardized_excel_disciplinary_committee_filing_time = pd.to_datetime(excel_disciplinary_committee_filing_time_raw, errors='coerce').strftime('%Y-%m-%d')
                    logger.debug(f"Excel纪委立案时间 '{excel_disciplinary_committee_filing_time_raw}' 标准化为 '{standardized_excel_disciplinary_committee_filing_time}' (规则5a)")
                    print(f"Excel纪委立案时间 '{excel_disciplinary_committee_filing_time_raw}' 标准化为 '{standardized_excel_disciplinary_committee_filing_time}' (规则5a)")
                except Exception as e:
                    logger.warning(f"行 {index + 1} - 规则5a: 无法解析Excel纪委立案时间 '{excel_disciplinary_committee_filing_time_raw}': {e}")
                    print(f"行 {index + 1} - 规则5a: 无法解析Excel纪委立案时间 '{excel_disciplinary_committee_filing_time_raw}': {e}")
                    standardized_excel_disciplinary_committee_filing_time = None

            # 标准化监委立案时间
            standardized_excel_supervisory_committee_filing_time = None
            if excel_supervisory_committee_filing_time_raw:
                try:
                    standardized_excel_supervisory_committee_filing_time = pd.to_datetime(excel_supervisory_committee_filing_time_raw, errors='coerce').strftime('%Y-%m-%d')
                    logger.debug(f"Excel监委立案时间 '{excel_supervisory_committee_filing_time_raw}' 标准化为 '{standardized_excel_supervisory_committee_filing_time}' (规则5a)")
                    print(f"Excel监委立案时间 '{excel_supervisory_committee_filing_time_raw}' 标准化为 '{standardized_excel_supervisory_committee_filing_time}' (规则5a)")
                except Exception as e:
                    logger.warning(f"行 {index + 1} - 规则5a: 无法解析Excel监委立案时间 '{excel_supervisory_committee_filing_time_raw}': {e}")
                    print(f"行 {index + 1} - 规则5a: 无法解析Excel监委立案时间 '{excel_supervisory_committee_filing_time_raw}': {e}")
                    standardized_excel_supervisory_committee_filing_time = None

            # 进行时间对比
            if standardized_filing_decision_timestamp:
                if (standardized_excel_disciplinary_committee_filing_time and standardized_excel_disciplinary_committee_filing_time != standardized_filing_decision_timestamp) or \
                   (standardized_excel_supervisory_committee_filing_time and standardized_excel_supervisory_committee_filing_time != standardized_filing_decision_timestamp):
                    mismatch_found_rule_5a = True
                    issues_list.append((index, case_code, person_code, "AZ监委立案时间、AW纪委立案时间和BG立案决定书落款时间不一致"))
                    disciplinary_committee_filing_time_mismatch_indices.add(index) # 标红纪委立案时间
                    supervisory_committee_filing_time_mismatch_indices.add(index) # 标红监委立案时间
                    logger.info(f"行 {index + 1} - 规则5a违规: 监委/纪委立案时间与决定书落款时间不一致。决定书时间: '{standardized_filing_decision_timestamp}', 纪委时间: '{standardized_excel_disciplinary_committee_filing_time}', 监委时间: '{standardized_excel_supervisory_committee_filing_time}'")
                    print(f"行 {index + 1} - 规则5a违规: 监委/纪委立案时间与决定书落款时间不一致。决定书时间: '{standardized_filing_decision_timestamp}', 纪委时间: '{standardized_excel_disciplinary_committee_filing_time}', 监委时间: '{standardized_excel_supervisory_committee_filing_time}'")
                else:
                    logger.info(f"行 {index + 1} - 规则5a: 监委/纪委立案时间与决定书落款时间一致。")
                    print(f"行 {index + 1} - 规则5a: 监委/纪委立案时间与决定书落款时间一致。")
            else:
                logger.info(f"行 {index + 1} - 规则5a: 立案决定书落款时间为空或无法解析，跳过对比。")
                print(f"行 {index + 1} - 规则5a: 立案决定书落款时间为空或无法解析，跳过对比。")


            # 规则5b: 获取“监委立案机关”字段值和“填报单位名称“字段的值，然后从机关单位对应表查询
            found_match_supervisory_authority_rule_5 = False
            if (excel_supervisory_committee_filing_authority, excel_reporting_unit_name, "NSL") in authority_agency_lookup:
                found_match_supervisory_authority_rule_5 = True
            
            if not found_match_supervisory_authority_rule_5:
                # 原始问题描述
                issues_list.append((index, case_code, person_code, "A填报单位名称跟AZ监委立案机关不一致"))
                supervisory_committee_filing_authority_mismatch_indices.add(index)
                logger.info(f"行 {index + 1} - 规则5b违规: 监委立案机关 ('{excel_supervisory_committee_filing_authority}') 和填报单位名称 ('{excel_reporting_unit_name}') 在对应表中未找到匹配记录，或category不是'NSL'。")
                print(f"行 {index + 1} - 规则5b违规: 监委立案机关 ('{excel_supervisory_committee_filing_authority}') 和填报单位名称 ('{excel_reporting_unit_name}') 在对应表中未找到匹配记录，或category不是'NSL'。")
                
                # 新增问题描述
                issues_list.append((index, case_code, person_code, "A填报单位名称跟AV纪委立案机关不一致"))
                disciplinary_committee_filing_authority_mismatch_indices.add(index) # 同样标红纪委立案机关
                logger.info(f"行 {index + 1} - 规则5b新增违规: 填报单位名称 ('{excel_reporting_unit_name}') 与纪委立案机关 ('{excel_disciplinary_committee_filing_authority}') 也在规则5b条件下被视为不一致。")
                print(f"行 {index + 1} - 规则5b新增违规: 填报单位名称 ('{excel_reporting_unit_name}') 与纪委立案机关 ('{excel_disciplinary_committee_filing_authority}') 也在规则5b条件下被视为不一致。")
            else:
                logger.info(f"行 {index + 1} - 规则5b: 监委立案机关 ('{excel_supervisory_committee_filing_authority}') 和填报单位名称 ('{excel_reporting_unit_name}') 在对应表中找到匹配记录。")
                print(f"行 {index + 1} - 规则5b: 监委立案机关 ('{excel_supervisory_committee_filing_authority}') 和填报单位名称 ('{excel_reporting_unit_name}') 在对应表中找到匹配记录。")

            # 规则5c: 获取“纪委立案机关”字段值和“填报单位名称“字段的值，然后从机关单位对应表查询
            found_match_disciplinary_authority_rule_5 = False
            if (excel_disciplinary_committee_filing_authority, excel_reporting_unit_name, "NSL") in authority_agency_lookup:
                found_match_disciplinary_authority_rule_5 = True
            
            if not found_match_disciplinary_authority_rule_5:
                issues_list.append((index, case_code, person_code, "A填报单位名称跟AV纪委立案机关不一致"))
                disciplinary_committee_filing_authority_mismatch_indices.add(index)
                logger.info(f"行 {index + 1} - 规则5c违规: 纪委立案机关 ('{excel_disciplinary_committee_filing_authority}') 和填报单位名称 ('{excel_reporting_unit_name}') 在对应表中未找到匹配记录，或category不是'NSL'。")
                print(f"行 {index + 1} - 规则5c违规: 纪委立案机关 ('{excel_disciplinary_committee_filing_authority}') 和填报单位名称 ('{excel_reporting_unit_name}') 在对应表中未找到匹配记录，或category不是'NSL'。")
            else:
                logger.info(f"行 {index + 1} - 规则5c: 纪委立案机关 ('{excel_disciplinary_committee_filing_authority}') 和填报单位名称 ('{excel_reporting_unit_name}') 在对应表中找到匹配记录。")
                print(f"行 {index + 1} - 规则5c: 纪委立案机关 ('{excel_disciplinary_committee_filing_authority}') 和填报单位名称 ('{excel_reporting_unit_name}') 在对应表中找到匹配记录。")
        else:
            logger.info(f"行 {index + 1} - 立案决定书内容不包含 '纪监立' 和 '审查调查'，跳过相关验证。")
            print(f"行 {index + 1} - 立案决定书内容不包含 '纪监立' 和 '审查调查'，跳过相关验证。")

    logger.info("立案时间相关规则验证完成。")
    print("立案时间相关规则验证完成。")


def validate_confiscation_amount(df, issues_list, confiscation_amount_indices):
    """
    新增规则：与“审理报告”字段内容进行对比，查找字符串“收缴”，
    若出现“收缴”二字，将副本文件“收缴金额（万元）”字段标黄，
    并在立案编号表中添加问题描述：“CY审理报告中含有收缴二字，请人工再次确认“。

    参数:
    df (pd.DataFrame): 原始Excel数据的DataFrame。
    issues_list (list): 包含所有问题的列表，每个问题是一个(索引, 案件编码, 涉案人员编码, 问题描述)元组。
    confiscation_amount_indices (set): 收集所有“收缴金额（万元）”需要标黄的行索引。

    返回:
    None (issues_list 和 confiscation_amount_indices 会在函数内部被修改)。
    """
    logger.info("开始验证收缴金额相关规则...")
    print("开始验证收缴金额相关规则...")

    col_trial_report = "审理报告"
    col_confiscation_amount = "收缴金额（万元）"

    if col_trial_report not in df.columns or col_confiscation_amount not in df.columns:
        msg = f"缺少必要的列 '{col_trial_report}' 或 '{col_confiscation_amount}'，跳过收缴金额相关验证。"
        logger.warning(msg)
        print(msg)
        return

    for index, row in df.iterrows():
        trial_report_text = str(row[col_trial_report]).strip() if pd.notna(row[col_trial_report]) else ''
        case_code = str(row.get("案件编码", "")).strip()
        person_code = str(row.get("涉案人员编码", "")).strip()

        if "收缴" in trial_report_text:
            issues_list.append((index, case_code, person_code, "CY审理报告中含有收缴二字，请人工再次确认"))
            confiscation_amount_indices.add(index)
            logger.info(f"行 {index + 1} - '审理报告' 中包含 '收缴'。'收缴金额（万元）' 字段将标黄。案件编码: {case_code}, 涉案人员编码: {person_code}")
            print(f"行 {index + 1} - '审理报告' 中包含 '收缴'。'收缴金额（万元）' 字段将标黄。案件编码: {case_code}, 涉案人员编码: {person_code}")

    logger.info("收缴金额相关规则验证完成。")
    print("收缴金额相关规则验证完成。")


def validate_confiscation_of_property_amount(df, issues_list, confiscation_of_property_amount_indices):
    """
    新增规则：与“审理报告”字段内容进行对比，查找字符串“没收金额”，
    若出现“没收金额”四字，将副本文件“没收金额”字段标黄。
    并在立案编号表中添加问题描述：“CY审理报告中含有没收金额四字，请人工再次确认CG没收金额“。

    参数:
    df (pd.DataFrame): 原始Excel数据的DataFrame。
    issues_list (list): 包含所有问题的列表，每个问题是一个(索引, 案件编码, 涉案人员编码, 问题描述)元组。
    confiscation_of_property_amount_indices (set): 收集所有“没收金额”需要标黄的行索引。

    返回:
    None (issues_list 和 confiscation_of_property_amount_indices 会在函数内部被修改)。
    """
    logger.info("开始验证没收金额相关规则...")
    print("开始验证没收金额相关规则...")

    col_trial_report = "审理报告"
    col_confiscation_of_property_amount = "没收金额"

    if col_trial_report not in df.columns or col_confiscation_of_property_amount not in df.columns:
        msg = f"缺少必要的列 '{col_trial_report}' 或 '{col_confiscation_of_property_amount}'，跳过没收金额相关验证。"
        logger.warning(msg)
        print(msg)
        return

    for index, row in df.iterrows():
        trial_report_text = str(row[col_trial_report]).strip() if pd.notna(row[col_trial_report]) else ''
        case_code = str(row.get("案件编码", "")).strip()
        person_code = str(row.get("涉案人员编码", "")).strip()

        if "没收金额" in trial_report_text:
            issues_list.append((index, case_code, person_code, "CY审理报告中含有没收金额四字，请人工再次确认CG没收金额"))
            confiscation_of_property_amount_indices.add(index)
            logger.info(f"行 {index + 1} - '审理报告' 中包含 '没收金额'。'没收金额' 字段将标黄。案件编码: {case_code}, 涉案人员编码: {person_code}")
            print(f"行 {index + 1} - '审理报告' 中包含 '没收金额'。'没收金额' 字段将标黄。案件编码: {case_code}, 涉案人员编码: {person_code}")

    logger.info("没收金额相关规则验证完成。")
    print("没收金额相关规则验证完成。")


def validate_order_for_reparations_amount(df, issues_list, order_for_reparations_amount_indices):
    """
    新增规则：与“审理报告”字段内容进行对比，查找字符串“责令退赔”，
    若出现“责令退赔”四字，将副本文件“责令退赔金额”字段标黄。
    并在立案编号表中添加问题描述：“CY审理报告中含有责令退赔四字，请人工再次确认CH责令退赔金额“。

    参数:
    df (pd.DataFrame): 原始Excel数据的DataFrame。
    issues_list (list): 包含所有问题的列表，每个问题是一个(索引, 案件编码, 涉案人员编码, 问题描述)元组。
    order_for_reparations_amount_indices (set): 收集所有“责令退赔金额”需要标黄的行索引。

    返回:
    None (issues_list 和 order_for_reparations_amount_indices 会在函数内部被修改)。
    """
    logger.info("开始验证责令退赔金额相关规则...")
    print("开始验证责令退赔金额相关规则...")

    col_trial_report = "审理报告"
    col_order_for_reparations_amount = "责令退赔金额"

    if col_trial_report not in df.columns or col_order_for_reparations_amount not in df.columns:
        msg = f"缺少必要的列 '{col_trial_report}' 或 '{col_order_for_reparations_amount}'，跳过责令退赔金额相关验证。"
        logger.warning(msg)
        print(msg)
        return

    for index, row in df.iterrows():
        trial_report_text = str(row[col_trial_report]).strip() if pd.notna(row[col_trial_report]) else ''
        case_code = str(row.get("案件编码", "")).strip()
        person_code = str(row.get("涉案人员编码", "")).strip()

        if "责令退赔" in trial_report_text:
            issues_list.append((index, case_code, person_code, "CY审理报告中含有责令退赔四字，请人工再次确认CH责令退赔金额"))
            order_for_reparations_amount_indices.add(index)
            logger.info(f"行 {index + 1} - '审理报告' 中包含 '责令退赔'。'责令退赔金额' 字段将标黄。案件编码: {case_code}, 涉案人员编码: {person_code}")
            print(f"行 {index + 1} - '审理报告' 中包含 '责令退赔'。'责令退赔金额' 字段将标黄。案件编码: {case_code}, 涉案人员编码: {person_code}")

    logger.info("责令退赔金额相关规则验证完成。")
    print("责令退赔金额相关规则验证完成。")


def validate_registered_handover_amount(df, issues_list, registered_handover_amount_indices):
    """
    新增规则：与“审理报告”字段内容进行对比，查找字符串“登记上交金额”，
    若出现“登记上交金额”这6个字，将副本文件“登记上交金额”字段标黄。
    并在立案编号表中添加问题描述：“CY审理报告中含有登记上交金额，请人工再次确认CI登记上交金额“。

    参数:
    df (pd.DataFrame): 原始Excel数据的DataFrame。
    issues_list (list): 包含所有问题的列表，每个问题是一个(索引, 案件编码, 涉案人员编码, 问题描述)元组。
    registered_handover_amount_indices (set): 收集所有“登记上交金额”需要标黄的行索引。

    返回:
    None (issues_list 和 registered_handover_amount_indices 会在函数内部被修改)。
    """
    logger.info("开始验证登记上交金额相关规则...")
    print("开始验证登记上交金额相关规则...")

    col_trial_report = "审理报告"
    col_registered_handover_amount = "登记上交金额"

    if col_trial_report not in df.columns or col_registered_handover_amount not in df.columns:
        msg = f"缺少必要的列 '{col_trial_report}' 或 '{col_registered_handover_amount}'，跳过登记上交金额相关验证。"
        logger.warning(msg)
        print(msg)
        return

    for index, row in df.iterrows():
        trial_report_text = str(row[col_trial_report]).strip() if pd.notna(row[col_trial_report]) else ''
        case_code = str(row.get("案件编码", "")).strip()
        person_code = str(row.get("涉案人员编码", "")).strip()

        if "登记上交金额" in trial_report_text:
            issues_list.append((index, case_code, person_code, "CY审理报告中含有登记上交金额，请人工再次确认CI登记上交金额"))
            registered_handover_amount_indices.add(index)
            logger.info(f"行 {index + 1} - '审理报告' 中包含 '登记上交金额'。'登记上交金额' 字段将标黄。案件编码: {case_code}, 涉案人员编码: {person_code}")
            print(f"行 {index + 1} - '审理报告' 中包含 '登记上交金额'。'登记上交金额' 字段将标黄。案件编码: {case_code}, 涉案人员编码: {person_code}")

    logger.info("登记上交金额相关规则验证完成。")
    print("登记上交金额相关规则验证完成。")