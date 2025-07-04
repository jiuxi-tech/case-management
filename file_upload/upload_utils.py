# upload_utils.py
import os
import pandas as pd
import logging
from flask import flash, redirect, url_for
from werkzeug.utils import secure_filename

logger = logging.getLogger(__name__)

def allowed_file(filename, allowed_extensions):
    """
    检查文件扩展名是否在允许的列表中。

    参数:
        filename (str): 文件名。
        allowed_extensions (set): 允许的扩展名集合（不带前导点）。

    返回:
        bool: 如果文件扩展名在允许列表中则为 True，否则为 False。
    """
    if '.' not in filename:
        return False

    file_extension = filename.rsplit('.', 1)[1].lower()
    return file_extension in allowed_extensions

def handle_file_upload_and_initial_checks(request, app, file_key, folder_config_key, filename_pattern, file_type_chinese):
    """
    处理文件上传、保存和初步检查（扩展名、文件名模式）。

    参数:
        request (flask.request): Flask 请求对象，包含上传的文件。
        app (flask.Flask): Flask 应用实例，用于访问 app.config。
        file_key (str): 请求中文件字段的键名，例如 'case_file' 或 'file'。
        folder_config_key (str): app.config 中存储文件保存目录键名，例如 'CASE_FOLDER' 或 'CLUE_FOLDER'。
        filename_pattern (str): 文件名中必须包含的模式，例如 '立案登记表' 或 '线索登记表'。
        file_type_chinese (str): 文件类型的中文描述，用于错误消息，例如 '立案登记表' 或 '线索登记表'。

    返回:
        tuple: (file_path, original_filename, df, error_response)
               其中 error_response 是一个 flask.redirect 对象，如果发生错误，则非 None。
               成功时，df 是一个 pandas DataFrame。
    """
    logger.info(f"开始处理 {file_type_chinese} 上传请求")

    if file_key not in request.files:
        logger.error(f"未选择 {file_type_chinese} 文件")
        flash('未选择文件', 'error')
        return None, None, None, redirect(request.url)

    file = request.files[file_key]
    if file.filename == '':
        logger.error(f"{file_type_chinese} 文件名为空")
        flash('未选择文件', 'info')
        return None, None, None, redirect(request.url)

    allowed_extensions = app.config['ALLOWED_EXTENSIONS']
    if not allowed_file(file.filename, allowed_extensions):
        logger.error(f"{file_type_chinese} 文件格式错误: {file.filename}")
        flash(f'{file_type_chinese} 上传文件格式不对，请上传Excel文件（.xlsx 或 .xls）', 'error')
        return None, None, None, redirect(request.url)

    if filename_pattern not in file.filename:
        logger.error(f"{file_type_chinese} 文件名不符合要求: {file.filename}")
        flash(f'{file_type_chinese} 文件名必须包含“{filename_pattern}”', 'error')
        return None, None, None, redirect(request.url)

    target_folder = app.config[folder_config_key]
    original_filename = secure_filename(file.filename)
    file_path = os.path.join(target_folder, original_filename)

    logger.info(f"{file_type_chinese} 文件保存路径: {file_path}")
    os.makedirs(os.path.dirname(file_path), exist_ok=True)

    try:
        file.save(file_path)
    except Exception as e:
        logger.error(f"文件保存失败: {file_path} - {e}", exc_info=True)
        flash(f'文件保存失败: {e}', 'error')
        return None, None, None, redirect(request.url)

    if not os.path.exists(file_path):
        logger.error(f"{file_type_chinese} 文件保存失败: {file_path} 不存在")
        flash(f'文件保存失败: {file_path} 不存在', 'error')
        return None, None, None, redirect(request.url)
    logger.info(f"{file_type_chinese} 文件保存成功: {file_path}")

    try:
        df = pd.read_excel(file_path)
        return file_path, original_filename, df, None
    except Exception as e:
        logger.error(f"读取 {file_type_chinese} 文件失败: {str(e)}", exc_info=True)
        flash(f'读取文件内容失败，请确保它是有效的Excel文件: {str(e)}', 'error')
        return None, None, None, redirect(request.url)