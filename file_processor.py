import os
import pandas as pd
import logging
from datetime import datetime
from flask import flash, redirect, url_for
from config import Config

logger = logging.getLogger(__name__)

from validation_rules.validation_core import get_validation_issues
from excel_formatter import format_excel

def process_upload(request, app):
    logger.info("开始处理文件上传请求")
    if 'file' not in request.files:
        logger.error("未选择文件")
        flash('未选择文件', 'error')
        return redirect(request.url)
    file = request.files['file']
    if file.filename == '':
        logger.error("文件名为空")
        flash('未选择文件', 'error')
        return redirect(request.url)
    if not any(file.filename.lower().endswith(ext) for ext in Config.ALLOWED_EXTENSIONS):
        logger.error(f"文件格式错误: {file.filename}")
        flash('请上传Excel文件（.xlsx 或 .xls）', 'error')
        return redirect(request.url)
    if Config.REQUIRED_FILENAME_PATTERN not in file.filename:
        logger.error(f"文件名不符合要求: {file.filename}")
        flash('文件名必须包含“线索登记表”', 'error')
        return redirect(request.url)

    # 使用 CLUE_FOLDER 作为保存路径
    file_path = os.path.join(Config.CLUE_FOLDER, file.filename)
    logger.info(f"文件保存路径: {file_path}")  # 记录保存路径
    os.makedirs(os.path.dirname(file_path), exist_ok=True)  # 确保目录存在
    file.save(file_path)
    if not os.path.exists(file_path):
        logger.error(f"文件保存失败: {file_path} 不存在")
        flash(f'文件保存失败: {file_path} 不存在', 'error')
        return redirect(request.url)
    logger.info(f"文件保存成功: {file_path}")
    try:
        df = pd.read_excel(file_path)
        required_headers = Config.REQUIRED_HEADERS + [Config.COLUMN_MAPPINGS["organization_measure"], Config.COLUMN_MAPPINGS["acceptance_time"]]
        if not all(header in df.columns for header in required_headers):
            logger.error(f"缺少必要表头: {required_headers}")
            flash('Excel文件缺少必要的表头“填报单位名称”、“办理机关”、“被反映人”、“处置情况报告”、“受理时间”或“组织措施”', 'error')
            return redirect(request.url)

        mismatch_indices, issues_list = get_validation_issues(df)

        if issues_list:
            issues_df = pd.DataFrame(columns=['序号', '问题'])
            current_index = 1
            for i, (index, issue) in enumerate(issues_list, 1):
                issues_df = pd.concat([issues_df, pd.DataFrame({'序号': [current_index], '问题': [issue]})], ignore_index=True)
                current_index += 1
            issue_filename = f"线索编号{datetime.now().strftime('%Y%m%d')}.xlsx"
            issue_path = os.path.join(Config.CLUE_FOLDER, issue_filename)
            issues_df.to_excel(issue_path, index=False)
            logger.info(f"生成线索编号文件: {issue_path}")

        original_filename = file.filename.replace('.xlsx', '_副本.xlsx').replace('.xls', '_副本.xlsx')
        original_path = os.path.join(Config.CLUE_FOLDER, original_filename)
        format_excel(df, mismatch_indices, original_path, issues_list)

        if not all(header in df.columns for header in required_headers) or \
           not any(file.filename.lower().endswith(ext) for ext in Config.ALLOWED_EXTENSIONS) or \
           Config.REQUIRED_FILENAME_PATTERN not in file.filename or \
           file.filename == '':
            pass  # 错误已在校验阶段显示
        else:
            flash('程序执行成功！', 'success')
            logger.info("程序执行成功")

        return redirect(request.url)
    except Exception as e:
        logger.error(f"文件处理失败: {str(e)}")
        flash(f'文件处理失败: {str(e)}', 'error')
        return redirect(request.url)