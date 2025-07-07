import logging
import pandas as pd
import re
from datetime import datetime
from .case_document_validators import parse_chinese_date

logger = logging.getLogger(__name__)

def validate_trial_closing_time_rules(row, index, excel_case_code, excel_person_code, issues_list, trial_closing_time_mismatch_indices,
                                      excel_trial_closing_time, trial_text_raw, app_config):
    """
    验证审结时间字段规则。
    比较 Excel 中的审结时间与审理报告落款时间。
    使用 CS 作为字段标识，参考年龄规则的日志风格和编号表字段结构。

    参数:
        row (pd.Series): DataFrame 的当前行数据。
        index (int): 当前行的索引。
        excel_case_code (str): Excel 中的案件编码。
        excel_person_code (str): Excel 中的涉案人员编码。
        issues_list (list): 用于收集所有发现问题的列表。
        trial_closing_time_mismatch_indices (set): 用于收集审结时间不匹配的行索引。
        excel_trial_closing_time (str or datetime): Excel 中的审结时间字段值。
        trial_text_raw (str): 审理报告的原始文本。
        app_config (dict): Flask 应用的配置字典。
    """
    
    # 检查审结时间和审理报告是否都有值
    if pd.notna(excel_trial_closing_time) and pd.notna(trial_text_raw) and trial_text_raw.strip():
        excel_closing_date_obj = None
        
        # 解析Excel中的审结时间
        if isinstance(excel_trial_closing_time, datetime):
            excel_closing_date_obj = excel_trial_closing_time.date()
        elif isinstance(excel_trial_closing_time, str):
            try:
                excel_closing_date_obj = pd.to_datetime(excel_trial_closing_time).date()
            except ValueError:
                trial_closing_time_mismatch_indices.add(index)
                issues_list.append({
                    '案件编码': excel_case_code,
                    '涉案人员编码': excel_person_code,
                    '行号': index + 2,
                    '比对字段': f"CS{app_config['COLUMN_MAPPINGS']['trial_closing_time']}",
                    '被比对字段': f"CY{app_config['COLUMN_MAPPINGS']['trial_report']}",
                    '问题描述': f"CS{index + 2}{app_config['COLUMN_MAPPINGS']['trial_closing_time']}格式不正确",
                    '列名': app_config['COLUMN_MAPPINGS']['trial_closing_time']
                })
                logger.warning(f"<立案 - （审结时间格式）> - 行 {index + 2} - 审结时间 '{excel_trial_closing_time}' 格式不正确")
                return
        
        if excel_closing_date_obj:
            # 从审理报告最后一行提取落款时间
            lines = trial_text_raw.strip().split('\n')
            if lines:
                last_line = lines[-1].strip()
                date_match = re.search(r'(\d{4}年\d{1,2}月\d{1,2}日)', last_line)
                if date_match:
                    extracted_closing_date_str = date_match.group(1)
                    extracted_closing_date_obj = parse_chinese_date(extracted_closing_date_str)

                    if extracted_closing_date_obj:
                        if excel_closing_date_obj != extracted_closing_date_obj:
                            trial_closing_time_mismatch_indices.add(index)
                            issues_list.append({
                                '案件编码': excel_case_code,
                                '涉案人员编码': excel_person_code,
                                '行号': index + 2,
                                '比对字段': f"CS{app_config['COLUMN_MAPPINGS']['trial_closing_time']}",
                                '被比对字段': f"CY{app_config['COLUMN_MAPPINGS']['trial_report']}",
                                '问题描述': f"CS{index + 2}{app_config['COLUMN_MAPPINGS']['trial_closing_time']}与CY{index + 2}审理报告不一致",
                                '列名': app_config['COLUMN_MAPPINGS']['trial_closing_time']
                            })
                            logger.warning(f"<立案 - （审结时间与审理报告）> - 行 {index + 2} - 审结时间 '{excel_closing_date_obj}' 与审理报告落款时间 '{extracted_closing_date_obj}' 不一致")
                        else:
                            logger.info(f"行 {index + 2} - 审结时间一致：Excel: {excel_closing_date_obj}, 审理报告落款: {extracted_closing_date_obj}")
                    else:
                        trial_closing_time_mismatch_indices.add(index)
                        issues_list.append({
                            '案件编码': excel_case_code,
                            '涉案人员编码': excel_person_code,
                            '行号': index + 2,
                            '比对字段': f"CS{app_config['COLUMN_MAPPINGS']['trial_closing_time']}",
                            '被比对字段': f"CY{app_config['COLUMN_MAPPINGS']['trial_report']}",
                            '问题描述': f"CY{index + 2}{app_config['COLUMN_MAPPINGS']['trial_report']}落款时间格式不正确或未找到",
                            '列名': app_config['COLUMN_MAPPINGS']['trial_closing_time']
                        })
                        logger.warning(f"<立案 - （审理报告落款时间格式）> - 行 {index + 2} - 从审理报告最后一行 '{last_line}' 中提取的日期 '{extracted_closing_date_str}' 无法解析")
                else:
                    trial_closing_time_mismatch_indices.add(index)
                    issues_list.append({
                        '案件编码': excel_case_code,
                        '涉案人员编码': excel_person_code,
                        '行号': index + 2,
                        '比对字段': f"CS{app_config['COLUMN_MAPPINGS']['trial_closing_time']}",
                        '被比对字段': f"CY{app_config['COLUMN_MAPPINGS']['trial_report']}",
                        '问题描述': f"CY{index + 2}{app_config['COLUMN_MAPPINGS']['trial_report']}落款时间未找到",
                        '列名': app_config['COLUMN_MAPPINGS']['trial_closing_time']
                    })
                    logger.warning(f"<立案 - （审理报告落款时间未找到）> - 行 {index + 2} - 审理报告最后一行 '{last_line}' 未找到日期格式")
            else:
                trial_closing_time_mismatch_indices.add(index)
                issues_list.append({
                    '案件编码': excel_case_code,
                    '涉案人员编码': excel_person_code,
                    '行号': index + 2,
                    '比对字段': f"CS{app_config['COLUMN_MAPPINGS']['trial_closing_time']}",
                    '被比对字段': f"CY{app_config['COLUMN_MAPPINGS']['trial_report']}",
                    '问题描述': f"CY{index + 2}{app_config['COLUMN_MAPPINGS']['trial_report']}为空，无法比对审结时间",
                    '列名': app_config['COLUMN_MAPPINGS']['trial_closing_time']
                })
                logger.warning(f"<立案 - （审理报告为空）> - 行 {index + 2} - 审理报告为空，无法提取落款时间")
    
    elif pd.notna(excel_trial_closing_time) and (pd.isna(trial_text_raw) or not trial_text_raw.strip()):
        trial_closing_time_mismatch_indices.add(index)
        issues_list.append({
            '案件编码': excel_case_code,
            '涉案人员编码': excel_person_code,
            '行号': index + 2,
            '比对字段': f"CS{app_config['COLUMN_MAPPINGS']['trial_closing_time']}",
            '被比对字段': f"CY{app_config['COLUMN_MAPPINGS']['trial_report']}",
            '问题描述': f"CS{index + 2}{app_config['COLUMN_MAPPINGS']['trial_closing_time']}有值但CY{index + 2}审理报告为空，无法比对",
            '列名': app_config['COLUMN_MAPPINGS']['trial_closing_time']
        })
        logger.warning(f"<立案 - （审结时间有值但审理报告为空）> - 行 {index + 2} - 审结时间有值但审理报告为空，无法比对")
    
    elif (pd.isna(excel_trial_closing_time) or not str(excel_trial_closing_time).strip()) and pd.notna(trial_text_raw) and trial_text_raw.strip():
        trial_closing_time_mismatch_indices.add(index)
        issues_list.append({
            '案件编码': excel_case_code,
            '涉案人员编码': excel_person_code,
            '行号': index + 2,
            '比对字段': f"CS{app_config['COLUMN_MAPPINGS']['trial_closing_time']}",
            '被比对字段': f"CY{app_config['COLUMN_MAPPINGS']['trial_report']}",
            '问题描述': f"CS{index + 2}{app_config['COLUMN_MAPPINGS']['trial_closing_time']}为空但CY{index + 2}审理报告有值，无法比对",
            '列名': app_config['COLUMN_MAPPINGS']['trial_closing_time']
        })
        logger.warning(f"<立案 - （审结时间为空但审理报告有值）> - 行 {index + 2} - 审结时间为空但审理报告有值，无法比对")