import os
import pandas as pd
from datetime import datetime
import xlsxwriter
import logging
from config import Config
from db_utils import get_db
from flask import flash, redirect, url_for

# 配置日志，输出到控制台
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[logging.StreamHandler()]
)
logger = logging.getLogger(__name__)

def extract_name_from_report(report_text):
    """从处置情况报告中提取被反映人姓名"""
    if not report_text or pd.isna(report_text):
        logger.info(f"report_text 为空或无效: {report_text}")
        return None
    # 找到“（一）被反映人基本情况”字符串
    start_marker = "（一）被反映人基本情况"
    start_idx = report_text.find(start_marker)
    if start_idx == -1:
        logger.warning(f"未找到 '（一）被反映人基本情况' 标记: {report_text}")
        return None
    start_idx += len(start_marker)
    logger.debug(f"原始报告文本: {report_text}")
    # 定位下一个自然段落（以换行符或双换行符为分隔）
    next_newline_idx = report_text.find("\n", start_idx)
    if next_newline_idx == -1:
        next_newline_idx = len(report_text)
    paragraph = report_text[start_idx:next_newline_idx].strip()
    if not paragraph:
        # 查找下一个段落
        next_newline_idx = report_text.find("\n", start_idx)
        if next_newline_idx == -1:
            next_newline_idx = len(report_text)
        next_paragraph_start = report_text.find("\n", next_newline_idx + 1)
        if next_paragraph_start == -1:
            next_paragraph_start = len(report_text)
        paragraph = report_text[next_newline_idx + 1:next_paragraph_start].strip()
    logger.debug(f"提取的段落: {paragraph}")
    # 提取第一个逗号（支持中文逗号）前的内容
    end_idx = -1
    for char in [",", "，"]:
        temp_idx = paragraph.find(char)
        if temp_idx != -1 and (end_idx == -1 or temp_idx < end_idx):
            end_idx = temp_idx
    if end_idx == -1:
        logger.warning(f"未找到逗号: {paragraph}")
        return None
    name = paragraph[:end_idx].strip()
    logger.info(f"提取姓名: {name} from paragraph: {paragraph}")
    return name if name else None

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

    file_path = os.path.join(Config.UPLOAD_FOLDER, file.filename)
    file.save(file_path)
    logger.info(f"文件保存成功: {file_path}")
    try:
        df = pd.read_excel(file_path)
        # Check required headers
        required_headers = Config.REQUIRED_HEADERS + ["被反映人", "处置情况报告"]
        if not all(header in df.columns for header in required_headers):
            logger.error(f"缺少必要表头: {required_headers}")
            flash('Excel文件缺少必要的表头“填报单位名称”、“办理机关”、“被反映人”或“处置情况报告”', 'error')
            return redirect(request.url)

        # Compare with database (original rule)
        with get_db() as conn:
            cursor = conn.cursor()
            cursor.execute('SELECT authority, agency FROM authority_agency_dict WHERE category = ?', ('NSL',))
            db_records = cursor.fetchall()
            db_dict = {(str(row['authority']).strip().lower(), str(row['agency']).strip().lower()) for row in db_records}

        mismatch_indices = set()  # 去重索引
        issues_list = []  # 存储每种不匹配类型的记录
        for index, row in df.iterrows():
            logger.debug(f"处理行 {index + 1}")
            # Original rule: 填报单位名称 vs 办理机关
            agency = str(row["填报单位名称"]).strip().lower() if pd.notna(row["填报单位名称"]) else ''
            authority = str(row["办理机关"]).strip().lower() if pd.notna(row["办理机关"]) else ''
            if not authority or not agency or (authority, agency) not in db_dict:
                mismatch_indices.add(index)
                issues_list.append((index, "C2填报单位名称与H2办理机关不一致"))
                logger.info(f"行 {index + 1} - 填报单位名称与办理机关不一致")

            # New rule: 被反映人 vs 处置情况报告
            reported_person = str(row["被反映人"]).strip() if pd.notna(row["被反映人"]) else ''
            report_text = row["处置情况报告"]
            report_name = extract_name_from_report(report_text)

            logger.info(f"行 {index + 1} - 被反映人: {reported_person}, 提取姓名: {report_name}")
            if pd.isna(report_text):
                mismatch_indices.add(index)
                issues_list.append((index, "E2被反映人与AB2处置情况报告姓名不一致 (报告为空)"))
                logger.info(f"行 {index + 1} - 处置情况报告为空")
            elif reported_person and report_name and reported_person != report_name:
                mismatch_indices.add(index)
                issues_list.append((index, "E2被反映人与AB2处置情况报告姓名不一致"))
                logger.info(f"行 {index + 1} - 被反映人与处置情况报告姓名不一致")

        if issues_list:
            issues_df = pd.DataFrame(columns=['序号', '问题'])
            for i, (index, issue) in enumerate(issues_list, 1):
                issues_df = pd.concat([issues_df, pd.DataFrame({'序号': [i], '问题': [issue]})], ignore_index=True)
            issue_filename = f"线索编号{datetime.now().strftime('%Y%m%d')}.xlsx"
            issue_path = os.path.join(Config.UPLOAD_FOLDER, issue_filename)
            issues_df.to_excel(issue_path, index=False)
            logger.info(f"生成线索编号文件: {issue_path}")

        original_filename = file.filename.replace('.xlsx', '_副本.xlsx').replace('.xls', '_副本.xlsx')
        original_path = os.path.join(Config.UPLOAD_FOLDER, original_filename)
        with pd.ExcelWriter(original_path, engine='xlsxwriter') as writer:
            df.to_excel(writer, sheet_name='Sheet1', index=False)
            workbook = writer.book
            worksheet = writer.sheets['Sheet1']
            red_format = workbook.add_format({'bg_color': '#FF0000'})  # Red for mismatches
            yellow_format = workbook.add_format({'bg_color': '#FFFF00'})  # Yellow for empty report
            for idx in mismatch_indices:
                row = df.iloc[idx]
                # Original rule marking
                worksheet.write(f'C{idx + 2}', row["填报单位名称"], red_format)
                worksheet.write(f'H{idx + 2}', row["办理机关"], red_format)
                # New rule marking
                if "处置情况报告" in df.columns:
                    report_text = row["处置情况报告"]
                    if pd.isna(report_text):
                        worksheet.write(f'AB{idx + 2}', row["处置情况报告"], yellow_format)
                    if "被反映人" in df.columns:
                        reported_person = str(row["被反映人"]).strip() if pd.notna(row["被反映人"]) else ''
                        report_name = extract_name_from_report(report_text)
                        if reported_person and report_name and reported_person != report_name:
                            worksheet.write(f'E{idx + 2}', row["被反映人"], red_format)
                            # 仅标红被反映人字段，处置情况报告不标红
            logger.info(f"生成副本文件: {original_path}")

        # 仅在校验失败或程序异常时显示错误，成功生成文件时显示成功
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