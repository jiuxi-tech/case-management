import pandas as pd
import logging
from config import Config

logger = logging.getLogger(__name__)

def validate_disciplinary_sanction(df, issues_list):
    """
    Validate '党纪处分' against '处分决定' content.
    Highlights '党纪处分' in red if the corresponding keyword from '党纪处分' is not found in '处分决定'.
    同时，增加对党纪处分与党员身份不符的校验。

    Args:
        df (pd.DataFrame): The DataFrame containing the case data.
        issues_list (list): A list to append validation issues. Each issue is a dictionary:
                            {"案件编码": ..., "涉案人员编码": ..., "问题描述": ..., "风险等级": ..., "行号": ...}.

    Returns:
        set: A set of row indices where '党纪处分' needs to be highlighted red.
    """
    disciplinary_sanction_mismatch_indices = set()
    
    sanction_keywords = Config.DISCIPLINARY_SANCTION_KEYWORDS

    logger.info("开始校验 '党纪处分' 字段与 '处分决定' 字段的一致性。")
    print("\n--- 党纪处分与处分决定匹配规则 ---")

    disciplinary_sanction_col = Config.COLUMN_MAPPINGS.get("disciplinary_sanction", "党纪处分")
    disposal_decision_col = Config.COLUMN_MAPPINGS.get("disciplinary_decision", "处分决定") 
    case_code_col = Config.COLUMN_MAPPINGS.get("case_code", "案件编码")
    person_code_col = Config.COLUMN_MAPPINGS.get("person_code", "涉案人员编码")
    party_member_col = Config.COLUMN_MAPPINGS.get("party_member", "是否中共党员") 

    missing_cols = []
    if disciplinary_sanction_col not in df.columns: missing_cols.append(disciplinary_sanction_col)
    if disposal_decision_col not in df.columns: missing_cols.append(disposal_decision_col)
    if party_member_col not in df.columns: missing_cols.append(party_member_col)
    
    if missing_cols:
        logger.warning(f"DataFrame 中缺少必需字段 {missing_cols}，跳过党纪处分校验。")
        print(f"警告：DataFrame 中缺少必需字段 {missing_cols}，跳过党纪处分校验。")
        return disciplinary_sanction_mismatch_indices

    for index, row in df.iterrows():
        disciplinary_sanction = str(row[disciplinary_sanction_col]).strip() if pd.notna(row[disciplinary_sanction_col]) else ""
        disposal_decision = str(row[disposal_decision_col]).strip() if pd.notna(row[disposal_decision_col]) else ""
        party_member_status = str(row[party_member_col]).strip() if pd.notna(row[party_member_col]) else ""
        
        case_code = str(row.get(case_code_col, "N/A")) 
        person_code = str(row.get(person_code_col, "N/A"))

        # --- 规则1: '党纪处分' 字段与 '处分决定' 内容的一致性校验 ---
        # 如果 '党纪处分' 有值，但 '处分决定' 中不包含任何 'sanction_keywords'，则标记
        if disciplinary_sanction and not any(kw in disposal_decision for kw in sanction_keywords):
            disciplinary_sanction_mismatch_indices.add(index)
            issue_description = Config.VALIDATION_RULES.get(
                "disciplinary_sanction_mismatch", 
                "BO党纪处分与CU处分决定不一致",
            )
            issues_list.append({
                "案件编码": case_code,
                "涉案人员编码": person_code,
                "问题描述": issue_description,
                "风险等级": "中", 
                "行号": index + 2
            })
            logger.debug(f"行 {index+2} (案件编码: {case_code}, 涉案人员编码: {person_code}): 党纪处分 '{disciplinary_sanction}' 与处分决定不匹配。")
            print(f"行 {index+2} (案件编码: {case_code}, 涉案人员编码: {person_code}): 党纪处分 '{disciplinary_sanction}' 与处分决定不一致。")

        # --- 规则2: 党纪处分（处分决定）中出现开除党籍，但被调查人非中共党员 ---
        if "开除党籍" in disciplinary_sanction or "开除党籍" in disposal_decision:
            if party_member_status not in ["是", "中共党员", "党员", "是（中共党员）"]: 
                issue_description = Config.VALIDATION_RULES.get(
                    "disciplinary_sanction_party_member_mismatch", 
                    "党纪处分（开除党籍）与党员身份不符（描述缺失）"
                )
                issues_list.append({
                    "案件编码": case_code,
                    "涉案人员编码": person_code,
                    "问题描述": issue_description,
                    "风险等级": "高", 
                    "行号": index + 2
                })
                logger.debug(f"行 {index+2} (案件编码: {case_code}, 涉案人员编码: {person_code}): 发现 '开除党籍' 但非中共党员。")
                print(f"行 {index+2} (案件编码: {case_code}, 涉案人员编码: {person_code}): 发现 '开除党籍' 但非中共党员。")

    print("--- 党纪处分与处分决定匹配规则校验完成 ---")
    logger.info("完成校验 '党纪处分' 字段与 '处分决定' 字段的一致性。")
    return disciplinary_sanction_mismatch_indices


def validate_administrative_sanction(df, issues_list):
    """
    Validate '政务处分' against '处分决定' content.
    Highlights '政务处分' in red if the corresponding keyword from '政务处分' is not found in '处分决定'.

    Args:
        df (pd.DataFrame): The DataFrame containing the case data.
        issues_list (list): A list to append validation issues. Each issue is a dictionary:
                            {"案件编码": ..., "涉案人员编码": ..., "问题描述": ..., "风险等级": ..., "行号": ...}.

    Returns:
        set: A set of row indices where '政务处分' needs to be highlighted red.
    """
    administrative_sanction_mismatch_indices = set()
    
    # 新增：定义政务处分需要查找的关键字
    # 建议将此列表也放到 Config 中，例如 Config.ADMINISTRATIVE_SANCTION_KEYWORDS
    administrative_sanction_keywords = Config.ADMINISTRATIVE_SANCTION_KEYWORDS # 假设已在Config中配置

    logger.info("开始校验 '政务处分' 字段与 '处分决定' 字段的一致性。")
    print("\n--- 政务处分与处分决定匹配规则 ---")

    administrative_sanction_col = Config.COLUMN_MAPPINGS.get("administrative_sanction", "政务处分")
    disposal_decision_col = Config.COLUMN_MAPPINGS.get("disciplinary_decision", "处分决定") 
    case_code_col = Config.COLUMN_MAPPINGS.get("case_code", "案件编码")
    person_code_col = Config.COLUMN_MAPPINGS.get("person_code", "涉案人员编码")

    missing_cols = []
    if administrative_sanction_col not in df.columns: missing_cols.append(administrative_sanction_col)
    if disposal_decision_col not in df.columns: missing_cols.append(disposal_decision_col)
    
    if missing_cols:
        logger.warning(f"DataFrame 中缺少必需字段 {missing_cols}，跳过政务处分校验。")
        print(f"警告：DataFrame 中缺少必需字段 {missing_cols}，跳过政务处分校验。")
        return administrative_sanction_mismatch_indices

    for index, row in df.iterrows():
        administrative_sanction = str(row[administrative_sanction_col]).strip() if pd.notna(row[administrative_sanction_col]) else ""
        disposal_decision = str(row[disposal_decision_col]).strip() if pd.notna(row[disposal_decision_col]) else ""
        
        case_code = str(row.get(case_code_col, "N/A")) 
        person_code = str(row.get(person_code_col, "N/A"))

        # --- 规则1: '政务处分' 字段与 '处分决定' 内容的一致性校验 ---
        # 只有当 '政务处分' 有值，但 '处分决定' 中不包含任何 'administrative_sanction_keywords'，则标记
        if administrative_sanction and not any(kw in disposal_decision for kw in administrative_sanction_keywords):
            administrative_sanction_mismatch_indices.add(index)
            issue_description = Config.VALIDATION_RULES.get(
                "administrative_sanction_mismatch", 
                "BR政务处分与CU处分决定不一致", 
            )
            issues_list.append({
                "案件编码": case_code,
                "涉案人员编码": person_code,
                "问题描述": issue_description,
                "风险等级": "中", 
                "行号": index + 2
            })
            logger.debug(f"行 {index+2} (案件编码: {case_code}, 涉案人员编码: {person_code}): 政务处分 '{administrative_sanction}' 与处分决定不匹配。")
            print(f"行 {index+2} (案件编码: {case_code}, 涉案人员编码: {person_code}): 政务处分 '{administrative_sanction}' 与处分决定不一致。")

    print("--- 政务处分与处分决定匹配规则校验完成 ---")
    logger.info("完成校验 '政务处分' 字段与 '处分决定' 字段的一致性。")
    
    return administrative_sanction_mismatch_indices