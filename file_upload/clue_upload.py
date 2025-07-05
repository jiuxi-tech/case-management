# clue_file_processor.py
import os
import sys
import pandas as pd
import logging
from flask import flash, redirect, url_for

# 导入通用函数
from .upload_utils import handle_file_upload_and_initial_checks

# 导入验证规则模块和辅助函数
try:
    from validation.clue_validation.clue_validation import validate_clue_data
    from excel_formatter import format_clue_excel
    from db_utils import get_db, get_authority_agency_dict
except ImportError as e:
    # 打印到标准错误输出，确保能看到
    print(f"ERROR: 无法导入必要的模块或函数: {e}", file=sys.stderr)
    logging.getLogger(__name__).error(f"无法导入必要的模块或函数: {e}", exc_info=True)
    # 强制重新抛出异常
    raise

logger = logging.getLogger(__name__)

def process_clue_upload(request, app):
    """
    处理线索登记表文件的上传、保存和验证。
    如果文件不符合要求或验证失败，会闪现错误消息并重定向。

    参数:
        request (flask.request): Flask 请求对象，包含上传的文件。
        app (flask.Flask): Flask 应用实例，用于访问 app.config。

    返回:
        flask.redirect: 重定向到上传页面。
    """
    # 使用通用函数处理文件上传和初步检查
    file_path, original_filename, df, error_response = handle_file_upload_and_initial_checks(
        request, app, 'file', 'CLUE_FOLDER', app.config['REQUIRED_FILENAME_PATTERN'], '线索登记表'
    )
    if error_response:
        return error_response

    try:
        # 检查必要的表头，使用 app.config
        required_headers = list(app.config['CLUE_REQUIRED_HEADERS']) + [
            app.config['COLUMN_MAPPINGS']["organization_measure"],
            app.config['COLUMN_MAPPINGS']["acceptance_time"]
        ]

        if not all(header in df.columns for header in required_headers):
            missing_headers = [header for header in required_headers if header not in df.columns]
            logger.error(f"缺少必要表头: {missing_headers}")
            flash(f'Excel文件缺少必要的表头: {", ".join(missing_headers)}', 'error')
            return redirect(request.url)

        disposal_report_column = app.config['COLUMN_MAPPINGS'].get("disposal_report", "处置情况报告")
        if disposal_report_column not in df.columns:
            logger.error(f"Excel文件缺少必要表头: {disposal_report_column}")
            flash(f'Excel文件缺少必要的表头“{disposal_report_column}”', 'error')
            return redirect(request.url)

        if df[disposal_report_column].isnull().all():
            logger.error(f"线索登记表“{disposal_report_column}”字段为空")
            flash(f'线索登记表“{disposal_report_column}”字段为空', 'error')
            return redirect(request.url)

        # 获取机构映射数据
        agency_mapping_db = get_authority_agency_dict(category='NSL')

        # 调用线索数据验证函数，并传入 agency_mapping_db
        issues_list, error_count = validate_clue_data(df, app.config, agency_mapping_db)
        logger.info(f"validate_clue_data 返回了 {len(issues_list)} 个问题和 {error_count} 个错误。")

        # 处理并生成问题报告文件
        if issues_list:
            data_for_issues_df = []
            seen_issues = set() # 用于去重
            for issue_item in issues_list:
                # 确保 issue_item 是字典，并处理可能的元组格式
                if isinstance(issue_item, tuple):
                    # 假设元组格式为 (行号, 受理线索编码, 问题描述, 风险等级)
                    issue_dict = {
                        "行号": issue_item[0] + 2, # 调整行号，因为Excel从1开始，且有表头
                        "受理线索编码": issue_item[1],
                        "问题描述": issue_item[2],
                        "风险等级": issue_item[3] if len(issue_item) > 3 else "中"
                    }
                elif isinstance(issue_item, dict):
                    issue_dict = issue_item
                else:
                    logger.warning(f"issues_list 中发现未知类型项: {type(issue_item)}. 跳过处理。")
                    continue

                # 使用 frozenset 来判断字典是否重复
                issue_hashable = frozenset(issue_dict.items())
                if issue_hashable not in seen_issues:
                    data_for_issues_df.append({
                        '序号': len(data_for_issues_df) + 1, # 动态生成序号
                        '受理线索编码': issue_dict.get('受理线索编码', 'N/A'),
                        '受理人员编码': issue_dict.get('受理人员编码', ''),
                        '行号': issue_dict.get('行号', ''),
                        '比对字段': issue_dict.get('比对字段', ''),
                        '被比对字段': issue_dict.get('被比对字段', ''),
                        '问题': issue_dict.get('问题描述', '无描述'),
                        '风险等级': issue_dict.get('风险等级', '中')
                    })
                    seen_issues.add(issue_hashable)

            issues_df = pd.DataFrame(data_for_issues_df)

            issue_filename = f"线索编号{app.config['TODAY_DATE']}.xlsx" # 使用 app.config['TODAY_DATE']
            issue_path = os.path.join(app.config['CLUE_FOLDER'], issue_filename) # 直接使用 app.config
            issues_df.to_excel(issue_path, index=False)
            logger.info(f"生成线索编号文件: {issue_path}")

        # 生成带有高亮和问题的副本文件
        original_filename_copy = original_filename.replace('.xlsx', '_副本.xlsx').replace('.xls', '_副本.xlsx')
        original_path_copy = os.path.join(app.config['CLUE_FOLDER'], original_filename_copy) # 直接使用 app.config

        # 由于 clue_file_processor 中的 format_excel 不使用 case_file_processor 中的大量高亮参数
        format_clue_excel(df,
                          issues_list=issues_list,
                          output_path=original_path_copy
                     )

        logger.info("线索登记表处理成功")
        flash('文件上传处理成功！', 'success')
        return redirect(request.url)
    except Exception as e:
        logger.error(f"线索登记表处理失败: {str(e)}", exc_info=True)
        flash(f'文件处理失败: {str(e)}', 'error')
        return redirect(request.url)