import logging
import pandas as pd

logger = logging.getLogger(__name__)

def validate_recovery_amount_rules(row, index, excel_case_code, excel_person_code, issues_list, recovery_amount_highlight_indices,
                                 excel_recovery_amount, app_config):
    """
    验证追缴失职渎职滥用职权造成的损失金额相关规则。
    当该字段有值时进行标黄处理。

    参数:
        row (pd.Series): DataFrame 的当前行数据。
        index (int): 当前行的索引。
        excel_case_code (str): Excel 中的案件编码。
        excel_person_code (str): Excel 中的涉案人员编码。
        issues_list (list): 用于收集所有发现问题的列表。
        recovery_amount_highlight_indices (set): 用于收集追缴金额有值的行索引。
        excel_recovery_amount (str): Excel 中的追缴失职渎职滥用职权造成的损失金额。
        app_config (dict): Flask 应用的配置字典。
    """
    
    # 规则: 追缴失职渎职滥用职权造成的损失金额字段有值时标黄
    if pd.notna(excel_recovery_amount) and str(excel_recovery_amount).strip() != '':
        recovery_amount_highlight_indices.add(index)
        issues_list.append({
            '案件编码': excel_case_code,
            '涉案人员编码': excel_person_code,
            '行号': index + 2,
            '比对字段': f"CJ{app_config['COLUMN_MAPPINGS']['recovery_amount']}",
            '被比对字段': f"CJ{app_config['COLUMN_MAPPINGS']['recovery_amount']}",
            '问题描述': f"CJ{index + 2}{app_config['COLUMN_MAPPINGS']['recovery_amount']}请再次确认",
            '列名': app_config['COLUMN_MAPPINGS']['recovery_amount']
        })
        logger.warning(f"<立案 - （1.追缴失职渎职滥用职权造成的损失金额）> - 行 {index + 2} - 追缴失职渎职滥用职权造成的损失金额有值，请再次确认")
    else:
        logger.info(f"<立案 - （1.追缴失职渎职滥用职权造成的损失金额）> - 行 {index + 2} - 追缴失职渎职滥用职权造成的损失金额为空")