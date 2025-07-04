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
