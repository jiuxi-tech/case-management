import os
import pandas as pd
import logging
from datetime import datetime
from flask import flash, redirect, url_for
from config import Config

logger = logging.getLogger(__name__)

# 从你的项目结构中导入，如果不存在，请确保提供这些文件的内容
from validation_rules.validation_core import get_validation_issues
from excel_formatter import format_excel
from validation_rules.case_validation_core import validate_case_relationships, generate_case_files


def process_upload(request, app):
    logger.info("开始处理文件上传请求")
    if 'file' not in request.files:
        logger.error("未选择文件")
        flash('未选择文件', 'error')
        return redirect(request.url)
    file = request.files['file']
    if file.filename == '':
        logger.error("文件名为空")
        flash('未选择文件', 'info') # 更改为info，表示操作完成但无文件
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
        required_headers = Config.CLUE_REQUIRED_HEADERS + [Config.COLUMN_MAPPINGS["organization_measure"], Config.COLUMN_MAPPINGS["acceptance_time"]]
        if not all(header in df.columns for header in required_headers):
            logger.error(f"缺少必要表头: {required_headers}")
            flash('Excel文件缺少必要的表头“填报单位名称”、“办理机关”、“被反映人”、“处置情况报告”、“受理时间”或“组织措施”', 'error')
            return redirect(request.url)

        # 这一部分是处理线索登记表的逻辑，与本次修改的立案登记表无关，但保留原样
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
            logger.info("程序执行成功")

        flash('文件上传处理成功！', 'success')  # 添加成功提示

        return redirect(request.url)
    except Exception as e:
        logger.error(f"文件处理失败: {str(e)}")
        flash(f'文件处理失败: {str(e)}', 'error')
        return redirect(request.url)

def process_case_upload(request, app):
    logger.info("开始处理立案登记表上传请求")
    if 'case_file' not in request.files:
        logger.error("未选择立案登记表文件")
        flash('未选择文件', 'error')
        return redirect(request.url)
    file = request.files['case_file']
    if file.filename == '':
        logger.error("立案登记表文件名为空")
        flash('任务已处理完毕', 'info')
        return redirect(request.url)
    # 校验文件格式和文件名
    if not any(file.filename.lower().endswith(ext) for ext in Config.ALLOWED_EXTENSIONS):
        logger.error(f"立案登记表文件格式错误: {file.filename}")
        flash('上传文件格式不对', 'error')
        return redirect(request.url)
    if "立案登记表" not in file.filename:
        logger.error(f"立案登记表文件名不符合要求: {file.filename}")
        flash('文件名必须包含立案登记', 'error')
        return redirect(request.url)

    # 使用 CASE_FOLDER 作为保存路径
    file_path = os.path.join(Config.CASE_FOLDER, file.filename)
    logger.info(f"立案登记表文件保存路径: {file_path}")
    os.makedirs(os.path.dirname(file_path), exist_ok=True)
    file.save(file_path)
    if not os.path.exists(file_path):
        logger.error(f"立案登记表文件保存失败: {file_path} 不存在")
        flash(f'文件保存失败: {file_path} 不存在', 'error')
        return redirect(request.url)
    logger.info(f"立案登记表文件保存成功: {file_path}")

    try:
        df = pd.read_excel(file_path)
        required_headers = Config.CASE_REQUIRED_HEADERS  # 使用立案登记表的表头
        if not all(header in df.columns for header in required_headers):
            logger.error(f"立案登记表缺少必要表头: {required_headers}")
            flash('Excel文件缺少必要的表头', 'error')
            return redirect(request.url)

        # 验证字段关系
        # *** 关键修改：现在接收所有七个返回值 ***
        mismatch_indices, gender_mismatch_indices, age_mismatch_indices, issues_list, birth_date_mismatch_indices, education_mismatch_indices, ethnicity_mismatch_indices = validate_case_relationships(df)
        
        # *** 关键修改：现在传递所有七个索引列表给 generate_case_files ***
        copy_path, case_num_path = generate_case_files(
            df, 
            file.filename, 
            Config.BASE_UPLOAD_FOLDER, 
            mismatch_indices, 
            gender_mismatch_indices, 
            issues_list,
            age_mismatch_indices,
            birth_date_mismatch_indices,
            education_mismatch_indices,
            ethnicity_mismatch_indices # 新增的参数
        )

        flash('文件上传处理成功！', 'success')  # 添加成功提示
        logger.info("立案登记表处理成功")

        return redirect(request.url)
    except Exception as e:
        logger.error(f"立案登记表处理失败: {str(e)}")
        flash(f'文件处理失败: {str(e)}', 'error')
        return redirect(request.url)
