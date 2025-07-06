import pandas as pd
import logging
from datetime import datetime

# 配置日志记录器
logger = logging.getLogger(__name__)

def validate_filing_time(df, issues_list, app_config):
    """
    验证立案时间相关规则。
    
    参数:
    df (pd.DataFrame): 原始Excel数据的DataFrame。
    issues_list (list): 包含所有问题的列表，每个问题是一个(索引, 案件编码, 涉案人员编码, 问题描述)元组。
    app_config (dict): Flask 应用的配置字典，包含Config类中的配置。
    
    返回:
    None (issues_list 会在函数内部被修改)。
    """
    logger.info("开始验证立案时间相关规则...")
    
    # 这里可以添加具体的立案时间验证逻辑
    # 目前为占位符实现
    
    logger.info("立案时间相关规则验证完成。")

def validate_order_for_reparations_amount(df, issues_list, order_for_reparations_amount_indices, app_config):
    """
    新增规则：与"审理报告"字段内容进行对比，查找字符串"责令退赔"，
    若出现"责令退赔"这4个字，将副本文件"责令退赔金额"字段标黄。
    并在立案编号表中添加问题描述："CY审理报告中含有责令退赔四字，请人工再次确认CH责令退赔金额"。

    参数:
    df (pd.DataFrame): 原始Excel数据的DataFrame。
    issues_list (list): 包含所有问题的列表，每个问题是一个(索引, 案件编码, 涉案人员编码, 问题描述)元组。
    order_for_reparations_amount_indices (set): 收集所有"责令退赔金额"需要标黄的行索引。
    app_config (dict): Flask 应用的配置字典，包含Config类中的配置。

    返回:
    None (issues_list 和 order_for_reparations_amount_indices 会在函数内部被修改)。
    """
    logger.info("开始验证责令退赔金额相关规则...")

    col_trial_report = app_config['COLUMN_MAPPINGS']['trial_report']
    col_order_for_reparations_amount = app_config['COLUMN_MAPPINGS']['compensation_amount']
    col_case_code = app_config['COLUMN_MAPPINGS']['case_code']
    col_person_code = app_config['COLUMN_MAPPINGS']['person_code']

    if col_trial_report not in df.columns or col_order_for_reparations_amount not in df.columns:
        msg = f"缺少必要的列 '{col_trial_report}' 或 '{col_order_for_reparations_amount}'，跳过责令退赔金额相关验证。"
        logger.warning(msg)
        return

    for index, row in df.iterrows():
        trial_report_text = str(row.get(col_trial_report, "")).strip() if pd.notna(row.get(col_trial_report)) else ''
        case_code = str(row.get(col_case_code, "")).strip()
        person_code = str(row.get(col_person_code, "")).strip()

        if "责令退赔" in trial_report_text:
            issues_list.append((index, case_code, person_code, app_config['VALIDATION_RULES'].get("highlight_compensation_amount", "审理报告中含有责令退赔四字，请人工再次确认责令退赔金额"), "中")) # 增加风险等级
            order_for_reparations_amount_indices.add(index)
            logger.info(f"行 {index + 1} - '{col_trial_report}' 中包含 '责令退赔'。'{col_order_for_reparations_amount}' 字段将标黄。案件编码: {case_code}, 涉案人员编码: {person_code}")

    logger.info("责令退赔金额相关规则验证完成。")


def validate_registered_handover_amount_single_row(row, index, excel_case_code, excel_person_code, issues_list, registered_handover_amount_indices, app_config):
    """
    CG.登记上交金额规则（单行验证）：与"审理报告"字段内容进行对比，查找字符串"登记上交金额"，
    若出现"登记上交金额"这6个字，将副本文件"登记上交金额"字段标黄。
    并在立案编号表中添加问题描述："CY审理报告中含有登记上交金额字样，请人工再次确认CG登记上交金额"。

    参数:
    row (pd.Series): DataFrame 的当前行数据。
    index (int): 当前行的索引。
    excel_case_code (str): Excel 中的案件编码。
    excel_person_code (str): Excel 中的涉案人员编码。
    issues_list (list): 包含所有问题的列表，每个问题是一个(索引, 案件编码, 涉案人员编码, 问题描述)元组。
    registered_handover_amount_indices (set): 收集所有"登记上交金额"需要标黄的行索引。
    app_config (dict): Flask 应用的配置字典，包含Config类中的配置。

    返回:
    None (issues_list 和 registered_handover_amount_indices 会在函数内部被修改)。
    """
    col_trial_report = app_config['COLUMN_MAPPINGS']['trial_report']
    col_registered_handover_amount = app_config['COLUMN_MAPPINGS']['registered_handover_amount']

    trial_report_text = str(row.get(col_trial_report, "")).strip() if pd.notna(row.get(col_trial_report)) else ''
    
    # 获取配置中的问题描述
    issue_description = app_config['VALIDATION_RULES'].get("highlight_case_registered_handover_amount", "CY审理报告中含有登记上交金额字样，请人工再次确认CG登记上交金额")

    if "登记上交金额" in trial_report_text:
        issues_list.append((index, excel_case_code, excel_person_code, issue_description, "中")) # 增加风险等级
        registered_handover_amount_indices.add(index)
        logger.warning(f"<立案 - (CG.登记上交金额)> - 行 {index + 2} - CY审理报告中含有登记上交金额字样，请人工再次确认CG登记上交金额")


def validate_registered_handover_amount(df, issues_list, registered_handover_amount_indices, app_config):
    """
    CG.登记上交金额规则：与"审理报告"字段内容进行对比，查找字符串"登记上交金额"，
    若出现"登记上交金额"这6个字，将副本文件"登记上交金额"字段标黄。
    并在立案编号表中添加问题描述："CY审理报告中含有登记上交金额，请人工再次确认CG登记上交金额"。

    参数:
    df (pd.DataFrame): 原始Excel数据的DataFrame。
    issues_list (list): 包含所有问题的列表，每个问题是一个(索引, 案件编码, 涉案人员编码, 问题描述)元组。
    registered_handover_amount_indices (set): 收集所有"登记上交金额"需要标黄的行索引。
    app_config (dict): Flask 应用的配置字典，包含Config类中的配置。

    返回:
    None (issues_list 和 registered_handover_amount_indices 会在函数内部被修改)。
    """
    col_trial_report = app_config['COLUMN_MAPPINGS']['trial_report']
    col_registered_handover_amount = app_config['COLUMN_MAPPINGS']['registered_handover_amount']
    col_case_code = app_config['COLUMN_MAPPINGS']['case_code']
    col_person_code = app_config['COLUMN_MAPPINGS']['person_code']

    if col_trial_report not in df.columns or col_registered_handover_amount not in df.columns:
        msg = f"<立案 - (CG.登记上交金额)> - 缺少必要的列 '{col_trial_report}' 或 '{col_registered_handover_amount}'，跳过登记上交金额相关验证。"
        logger.warning(msg)
        return

    for index, row in df.iterrows():
        case_code = row.get(app_config['COLUMN_MAPPINGS']['case_code'], '')
        person_code = row.get(app_config['COLUMN_MAPPINGS']['person_code'], '')
        trial_report_text = str(row.get(col_trial_report, "")).strip() if pd.notna(row.get(col_trial_report)) else ''

        if "登记上交金额" in trial_report_text:
            issues_list.append((index, case_code, person_code, app_config['VALIDATION_RULES'].get("highlight_case_registered_handover_amount", "CY审理报告中含有登记上交金额字样，请人工再次确认CG登记上交金额"), "中")) # 增加风险等级
            registered_handover_amount_indices.add(index)
            logger.warning(f"<立案 - (CG.登记上交金额)> - 行 {index + 2} - CY审理报告中含有登记上交金额字样，请人工再次确认CG登记上交金额")
