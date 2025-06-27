import logging
import pandas as pd
import re

logger = logging.getLogger(__name__)

def validate_disposal_and_amount_rules(df, issues_list, disposal_spirit_mismatch_indices):
    """
    验证与处分和金额相关的规则。
    规则: 获取“是否违反中央八项规定精神”字段值，与“处分决定”字段内容进行对比，
          “处分决定”字段内容如果出现“违反中央八项规定精神”这十个文字，则结果为“是”，否则为“否”。
          比对结果如果不一致，副本文件“是否违反中央八项规定精神”字段标红，
          并且在立案编号表中添加问题描述：“BI是否违反中央八项规定精神与CU处分决定不一致“。

    参数:
    df (pd.DataFrame): 原始Excel数据的DataFrame。
    issues_list (list): 包含所有问题的列表，每个问题是一个(索引, 案件编码, 涉案人员编码, 问题描述)元组。
    disposal_spirit_mismatch_indices (set): 收集“是否违反中央八项规定精神”不一致的行索引，用于标红。

    返回:
    None (issues_list 和 disposal_spirit_mismatch_indices 会在函数内部被修改)。
    """
    logger.info("开始验证处分和金额相关规则...")
    print("开始验证处分和金额相关规则...")

    try:
        col_spirit_violation = "是否违反中央八项规定精神"
        col_disposal_decision = "处分决定"

        required_cols = [col_spirit_violation, col_disposal_decision]
        if not all(col in df.columns for col in required_cols):
            msg = f"缺少必要的列 {required_cols} 中至少一个，跳过处分和金额相关验证。"
            logger.warning(msg)
            print(msg)
            return

    except Exception as e:
        logger.error(f"获取处分和金额相关列失败: {e}")
        print(f"获取处分和金额相关列失败: {e}")
        return

    for index, row in df.iterrows():
        excel_spirit_violation = str(row[col_spirit_violation]).strip() if pd.notna(row[col_spirit_violation]) else ''
        disposal_decision_text = str(row[col_disposal_decision]).strip() if pd.notna(row[col_disposal_decision]) else ''

        case_code = str(row.get("案件编码", "")).strip()
        person_code = str(row.get("涉案人员编码", "")).strip()

        # 判断处分决定中是否包含“违反中央八项规定精神”
        decision_contains_violation_phrase = "违反中央八项规定精神" in disposal_decision_text

        # 预期结果：如果处分决定包含该短语，则预期为“是”，否则为“否”
        expected_spirit_violation = "是" if decision_contains_violation_phrase else "否"

        logger.info(f"行 {index + 1} - 处分决定内容: '{disposal_decision_text[:50]}...'")
        logger.info(f"行 {index + 1} - 处分决定是否包含 '违反中央八项规定精神': {decision_contains_violation_phrase}")
        logger.info(f"行 {index + 1} - Excel '是否违反中央八项规定精神' 字段值: '{excel_spirit_violation}'")
        logger.info(f"行 {index + 1} - 预期 '是否违反中央八项规定精神' 字段值: '{expected_spirit_violation}'")

        print(f"行 {index + 1} - 处分决定内容: '{disposal_decision_text[:50]}...'")
        print(f"行 {index + 1} - 处分决定是否包含 '违反中央八项规定精神': {decision_contains_violation_phrase}")
        print(f"行 {index + 1} - Excel '是否违反中央八项规定精神' 字段值: '{excel_spirit_violation}'")
        print(f"行 {index + 1} - 预期 '是否违反中央八项规定精神' 字段值: '{expected_spirit_violation}'")

        # 进行比对
        if excel_spirit_violation != expected_spirit_violation:
            issues_list.append((index, case_code, person_code, "BI是否违反中央八项规定精神与CU处分决定不一致"))
            disposal_spirit_mismatch_indices.add(index)
            logger.warning(f"行 {index + 1} - 规则违规: '是否违反中央八项规定精神' ('{excel_spirit_violation}') 与处分决定内容不一致，预期为 '{expected_spirit_violation}'。")
            print(f"行 {index + 1} - 规则违规: '是否违反中央八项规定精神' ('{excel_spirit_violation}') 与处分决定内容不一致，预期为 '{expected_spirit_violation}'。")
        else:
            logger.info(f"行 {index + 1} - '是否违反中央八项规定精神' 字段值与处分决定内容一致。")
            print(f"行 {index + 1} - '是否违反中央八项规定精神' 字段值与处分决定内容一致。")

    logger.info("处分和金额相关规则验证完成。")
    print("处分和金额相关规则验证完成。")
