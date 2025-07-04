import logging
import pandas as pd
import re
from datetime import datetime # 新增导入

logger = logging.getLogger(__name__)

def validate_disposal_and_amount_rules(df, issues_list, disposal_spirit_mismatch_indices, closing_time_mismatch_indices, app_config):
    """
    验证与处分和金额相关的规则。
    新增 app_config 参数以匹配调用方传递的参数数量。

    规则1: 获取“是否违反中央八项规定精神”字段值，与“处分决定”字段内容进行对比，
             “处分决定”字段内容如果出现“违反中央八项规定精神”这十个文字，则结果为“是”，否则为“否”。
             比对结果如果不一致，副本文件“是否违反中央八项规定精神”字段标红，
             并且在立案编号表中添加问题描述：“BI是否违反中央八项规定精神与CU处分决定不一致“。

    规则2 (新增): 获取“结案时间”字段值，与“处分决定”字段内容进行对比，
                   查找字符串“本处分决定自xx年xx月xx日起生效”（其中xx表示变量），并提取日期。
                   比如从字符串“本处分决定自2025年3月20日起生效”中提取出“2025年3月20日”字符串并进行格式转换，
                   然后跟“结案时间”字段进行精准比对，若不一致则将副本文件“结案时间”字段标红，
                   并且在立案编号表中添加问题描述：“BN结案时间与CU处分决定不一致“。

    参数:
    df (pd.DataFrame): 原始Excel数据的DataFrame。
    issues_list (list): 包含所有问题的列表，每个问题是一个(索引, 案件编码, 涉案人员编码, 问题描述)元组。
    disposal_spirit_mismatch_indices (set): 收集“是否违反中央八项规定精神”不一致的行索引，用于标红。
    closing_time_mismatch_indices (set): 收集“结案时间”不一致的行索引，用于标红 (新增)。
    app_config (dict): Flask 应用的配置字典，包含Config类中的配置。

    返回:
    None (issues_list, disposal_spirit_mismatch_indices, 和 closing_time_mismatch_indices 会在函数内部被修改)。
    """
    logger.info("开始验证处分和金额相关规则...")

    try:
        col_spirit_violation = app_config['COLUMN_MAPPINGS']['central_eight_provisions']
        col_disposal_decision = app_config['COLUMN_MAPPINGS']['disciplinary_decision']
        col_closing_time = app_config['COLUMN_MAPPINGS']['closing_time']
        col_case_code = app_config['COLUMN_MAPPINGS']['case_code']
        col_person_code = app_config['COLUMN_MAPPINGS']['person_code']

        required_cols = [col_spirit_violation, col_disposal_decision, col_closing_time]
        if not all(col in df.columns for col in required_cols):
            msg = f"缺少必要的列 {required_cols} 中至少一个，跳过处分和金额相关验证。"
            logger.warning(msg)
            return

    except Exception as e:
        logger.error(f"获取处分和金额相关列失败: {e}")
        return

    for index, row in df.iterrows():
        case_code = str(row.get(col_case_code, "")).strip()
        person_code = str(row.get(col_person_code, "")).strip()

        # --- 规则1: “是否违反中央八项规定精神”与“处分决定”比对 ---
        excel_spirit_violation = str(row.get(col_spirit_violation, "")).strip() if pd.notna(row.get(col_spirit_violation)) else ''
        disposal_decision_text = str(row.get(col_disposal_decision, "")).strip() if pd.notna(row.get(col_disposal_decision)) else ''

        # 判断处分决定中是否包含“违反中央八项规定精神”
        decision_contains_violation_phrase = "违反中央八项规定精神" in disposal_decision_text

        # 预期结果：如果处分决定包含该短语，则预期为“是”，否则为“否”
        expected_spirit_violation = "是" if decision_contains_violation_phrase else "否"

        logger.info(f"行 {index + 1} - 处分决定内容: '{disposal_decision_text[:50]}...'")
        logger.info(f"行 {index + 1} - 处分决定是否包含 '违反中央八项规定精神': {decision_contains_violation_phrase}")
        logger.info(f"行 {index + 1} - Excel '是否违反中央八项规定精神' 字段值: '{excel_spirit_violation}'")
        logger.info(f"行 {index + 1} - 预期 '是否违反中央八项规定精神' 字段值: '{expected_spirit_violation}'")

        # 进行比对
        if excel_spirit_violation != expected_spirit_violation:
            issues_list.append((index, case_code, person_code, app_config['VALIDATION_RULES'].get("central_eight_provisions_mismatch", "是否违反中央八项规定精神与处分决定不一致"), "高")) # 增加风险等级
            disposal_spirit_mismatch_indices.add(index)
            logger.warning(f"行 {index + 1} - 规则违规: '{col_spirit_violation}' ('{excel_spirit_violation}') 与处分决定内容不一致，预期为 '{expected_spirit_violation}'。")
        else:
            logger.info(f"行 {index + 1} - '{col_spirit_violation}' 字段值与处分决定内容一致。")

        # --- 规则2 (新增): “结案时间”与“处分决定”中的生效日期比对 ---
        excel_closing_time_obj = None
        if pd.notna(row.get(col_closing_time)):
            try:
                # 尝试将Excel中的结案时间转换为日期对象，忽略时间部分
                if isinstance(row[col_closing_time], datetime):
                    excel_closing_time_obj = row[col_closing_time].date()
                else:
                    # pd.to_datetime 可以很好地处理多种日期格式
                    excel_closing_time_obj = pd.to_datetime(row[col_closing_time]).date()
            except Exception as e:
                logger.warning(f"行 {index + 1} - 无法解析 '{col_closing_time}' 字段 '{row[col_closing_time]}': {e}")


        # 使用正则表达式查找并提取日期
        # 正则表达式解释:
        # 本处分决定自 - 固定前缀
        # (\d{4}年\d{1,2}月\d{1,2}日) - 捕获组1: 匹配 'YYYY年MM月DD日' 格式的日期
        # 起生效 - 固定后缀
        match = re.search(r"本处分决定自(\d{4}年\d{1,2}月\d{1,2}日)起生效", disposal_decision_text)
        extracted_disposal_date = None

        if match:
            date_str = match.group(1)
            try:
                # 转换提取到的日期字符串为 datetime.date 对象
                # 替换“年”、“月”、“日”以便于datetime解析
                formatted_date_str = date_str.replace('年', '-').replace('月', '-').replace('日', '')
                extracted_disposal_date = datetime.strptime(formatted_date_str, "%Y-%m-%d").date()
                logger.info(f"行 {index + 1} - 从处分决定中提取的生效日期: '{date_str}' (格式化后: {extracted_disposal_date})")
            except ValueError as ve:
                logger.warning(f"行 {index + 1} - 无法解析处分决定中提取的日期 '{date_str}': {ve}")
        else:
            logger.info(f"行 {index + 1} - 处分决定中未找到匹配的生效日期模式。")

        # 进行“结案时间”与提取日期之间的比对
        if excel_closing_time_obj and extracted_disposal_date:
            if excel_closing_time_obj != extracted_disposal_date:
                issues_list.append((index, case_code, person_code, app_config['VALIDATION_RULES'].get("inconsistent_closing_time_with_decision", "结案时间与处分决定不一致"), "高")) # 增加风险等级
                closing_time_mismatch_indices.add(index)
                logger.warning(f"行 {index + 1} - 规则违规: '{col_closing_time}' ('{excel_closing_time_obj}') 与处分决定中提取的生效日期 ('{extracted_disposal_date}') 不一致。")
            else:
                logger.info(f"行 {index + 1} - '{col_closing_time}' 字段值与处分决定中提取的生效日期一致。")
        elif excel_closing_time_obj is None:
            logger.info(f"行 {index + 1} - '{col_closing_time}' 字段为空，跳过比对。")
        elif extracted_disposal_date is None:
            logger.info(f"行 {index + 1} - 未能从处分决定中提取到生效日期，跳过比对。")


    logger.info("处分和金额相关规则验证完成。")


def validate_disciplinary_sanction(df, issues_list, app_config):
    """
    Validate '党纪处分' against '处分决定' content.
    Highlights '党纪处分' in red if the corresponding keyword from '党纪处分' is not found in '处分决定'.
    同时，增加对党纪处分与党员身份不符的校验。

    Args:
        df (pd.DataFrame): The DataFrame containing the case data.
        issues_list (list): A list to append validation issues. Each issue is a dictionary:
                            {"案件编码": ..., "涉案人员编码": ..., "问题描述": ..., "风险等级": ..., "行号": ...}.
        app_config (dict): Flask 应用的配置字典，包含Config类中的配置。

    Returns:
        set: A set of row indices where '党纪处分' needs to be highlighted red.
    """
    disciplinary_sanction_mismatch_indices = set()
    
    sanction_keywords = app_config['DISCIPLINARY_SANCTION_KEYWORDS'] # 从 app_config 获取关键词

    logger.info("开始校验 '党纪处分' 字段与 '处分决定' 字段的一致性。")

    disciplinary_sanction_col = app_config['COLUMN_MAPPINGS'].get("disciplinary_sanction", "党纪处分")
    disposal_decision_col = app_config['COLUMN_MAPPINGS'].get("disciplinary_decision", "处分决定") 
    case_code_col = app_config['COLUMN_MAPPINGS'].get("case_code", "案件编码")
    person_code_col = app_config['COLUMN_MAPPINGS'].get("person_code", "涉案人员编码")
    party_member_col = app_config['COLUMN_MAPPINGS'].get("party_member", "是否中共党员") 

    missing_cols = []
    if disciplinary_sanction_col not in df.columns: missing_cols.append(disciplinary_sanction_col)
    if disposal_decision_col not in df.columns: missing_cols.append(disposal_decision_col)
    if party_member_col not in df.columns: missing_cols.append(party_member_col)
    
    if missing_cols:
        logger.warning(f"DataFrame 中缺少必需字段 {missing_cols}，跳过党纪处分校验。")
        return disciplinary_sanction_mismatch_indices

    for index, row in df.iterrows():
        disciplinary_sanction = str(row.get(disciplinary_sanction_col, "")).strip() if pd.notna(row.get(disciplinary_sanction_col)) else ""
        disposal_decision = str(row.get(disposal_decision_col, "")).strip() if pd.notna(row.get(disposal_decision_col)) else ""
        party_member_status = str(row.get(party_member_col, "")).strip() if pd.notna(row.get(party_member_col)) else ""
        
        case_code = str(row.get(case_code_col, "N/A")) 
        person_code = str(row.get(person_code_col, "N/A"))

        # --- 规则1: '党纪处分' 字段与 '处分决定' 内容的一致性校验 ---
        # 如果 '党纪处分' 有值，但 '处分决定' 中不包含任何 'sanction_keywords'，则标记
        if disciplinary_sanction and not any(kw in disposal_decision for kw in sanction_keywords):
            disciplinary_sanction_mismatch_indices.add(index)
            issue_description = app_config['VALIDATION_RULES'].get( # 从 app_config 获取问题描述
                "disciplinary_sanction_mismatch", 
                "党纪处分与处分决定不一致",
            )
            issues_list.append({
                "案件编码": case_code,
                "涉案人员编码": person_code,
                "问题描述": issue_description,
                "风险等级": "中", 
                "行号": index + 2
            })
            logger.debug(f"行 {index+2} (案件编码: {case_code}, 涉案人员编码: {person_code}): 党纪处分 '{disciplinary_sanction}' 与处分决定不匹配。")

        # --- 规则2: 党纪处分（处分决定）中出现开除党籍，但被调查人非中共党员 ---
        if "开除党籍" in disciplinary_sanction or "开除党籍" in disposal_decision:
            # 检查 party_member_status 是否明确表示非党员
            # 考虑到用户可能输入多种表示“非党员”的字符串，可以定义一个“非党员”的列表
            non_party_member_indicators = ["否", "非中共党员", "非党员"] 
            if party_member_status in non_party_member_indicators or (party_member_status and party_member_status not in ["是", "中共党员", "党员", "是（中共党员）"]):
                issue_description = app_config['VALIDATION_RULES'].get( # 从 app_config 获取问题描述
                    "disciplinary_sanction_party_member_mismatch", 
                    "党纪处分（开除党籍）与党员身份不符"
                )
                issues_list.append({
                    "案件编码": case_code,
                    "涉案人员编码": person_code,
                    "问题描述": issue_description,
                    "风险等级": "高", 
                    "行号": index + 2
                })
                logger.debug(f"行 {index+2} (案件编码: {case_code}, 涉案人员编码: {person_code}): 发现 '开除党籍' 但非中共党员。")

    logger.info("完成校验 '党纪处分' 字段与 '处分决定' 字段的一致性。")
    return disciplinary_sanction_mismatch_indices


def validate_administrative_sanction(df, issues_list, app_config):
    """
    Validate '政务处分' against '处分决定' content.
    Highlights '政务处分' in red if the corresponding keyword from '政务处分' is not found in '处分决定'.

    Args:
        df (pd.DataFrame): The DataFrame containing the case data.
        issues_list (list): A list to append validation issues. Each issue is a dictionary:
                            {"案件编码": ..., "涉案人员编码": ..., "问题描述": ..., "风险等级": ..., "行号": ...}.
        app_config (dict): Flask 应用的配置字典，包含Config类中的配置。

    Returns:
        set: A set of row indices where '政务处分' needs to be highlighted red.
    """
    administrative_sanction_mismatch_indices = set()
    
    # 新增：定义政务处分需要查找的关键字
    administrative_sanction_keywords = app_config['ADMINISTRATIVE_SANCTION_KEYWORDS'] # 从 app_config 获取关键词

    logger.info("开始校验 '政务处分' 字段与 '处分决定' 字段的一致性。")

    administrative_sanction_col = app_config['COLUMN_MAPPINGS'].get("administrative_sanction", "政务处分")
    disposal_decision_col = app_config['COLUMN_MAPPINGS'].get("disciplinary_decision", "处分决定") 
    case_code_col = app_config['COLUMN_MAPPINGS'].get("case_code", "案件编码")
    person_code_col = app_config['COLUMN_MAPPINGS'].get("person_code", "涉案人员编码")

    missing_cols = []
    if administrative_sanction_col not in df.columns: missing_cols.append(administrative_sanction_col)
    if disposal_decision_col not in df.columns: missing_cols.append(disposal_decision_col)
    
    if missing_cols:
        logger.warning(f"DataFrame 中缺少必需字段 {missing_cols}，跳过政务处分校验。")
        return administrative_sanction_mismatch_indices

    for index, row in df.iterrows():
        administrative_sanction = str(row.get(administrative_sanction_col, "")).strip() if pd.notna(row.get(administrative_sanction_col)) else ""
        disposal_decision = str(row.get(disposal_decision_col, "")).strip() if pd.notna(row.get(disposal_decision_col)) else ""
        
        case_code = str(row.get(case_code_col, "N/A")) 
        person_code = str(row.get(person_code_col, "N/A"))

        # --- 规则1: '政务处分' 字段与 '处分决定' 内容的一致性校验 ---
        # 只有当 '政务处分' 有值，但 '处分决定' 中不包含任何 'administrative_sanction_keywords'，则标记
        if administrative_sanction and not any(kw in disposal_decision for kw in administrative_sanction_keywords):
            administrative_sanction_mismatch_indices.add(index)
            issue_description = app_config['VALIDATION_RULES'].get( # 从 app_config 获取问题描述
                "administrative_sanction_mismatch", 
                "政务处分与处分决定不一致", 
            )
            issues_list.append({
                "案件编码": case_code,
                "涉案人员编码": person_code,
                "问题描述": issue_description,
                "风险等级": "中", 
                "行号": index + 2
            })
            logger.debug(f"行 {index+2} (案件编码: {case_code}, 涉案人员编码: {person_code}): 政务处分 '{administrative_sanction}' 与处分决定不匹配。")

    logger.info("完成校验 '政务处分' 字段与 '处分决定' 字段的一致性。")
    
    return administrative_sanction_mismatch_indices
