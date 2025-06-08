import os
import pandas as pd
from datetime import datetime
import xlsxwriter
from config import Config
from db_utils import get_db
from flask import flash, redirect, url_for  # 添加必要的 Flask 导入

def process_upload(request, app):
    if 'file' not in request.files:
        flash('未选择文件', 'error')
        return redirect(request.url)
    file = request.files['file']
    if file.filename == '':
        flash('未选择文件', 'error')
        return redirect(request.url)
    if not any(file.filename.lower().endswith(ext) for ext in Config.ALLOWED_EXTENSIONS):
        flash('请上传Excel文件（.xlsx 或 .xls）', 'error')
        return redirect(request.url)
    if Config.REQUIRED_FILENAME_PATTERN not in file.filename:
        flash('文件名必须包含“线索登记表”', 'error')
        return redirect(request.url)

    file_path = os.path.join(Config.UPLOAD_FOLDER, file.filename)
    file.save(file_path)
    try:
        df = pd.read_excel(file_path)
        if not all(header in df.columns for header in Config.REQUIRED_HEADERS):
            flash('Excel文件缺少必要的表头“填报单位名称”或“办理机关”', 'error')
            return redirect(request.url)

        with get_db() as conn:
            cursor = conn.cursor()
            cursor.execute('SELECT authority, agency FROM authority_agency_dict WHERE category = ?', ('NSL',))
            db_records = cursor.fetchall()
            db_dict = {(str(row['authority']).strip().lower(), str(row['agency']).strip().lower()) for row in db_records}

        mismatches = []
        mismatch_indices = []
        for index, row in df.iterrows():
            agency = str(row["填报单位名称"]).strip().lower() if pd.notna(row["填报单位名称"]) else ''
            authority = str(row["办理机关"]).strip().lower() if pd.notna(row["办理机关"]) else ''
            if not authority or not agency or (authority, agency) not in db_dict:
                mismatches.append(f"行 {index + 1}")
                mismatch_indices.append(index)
                if app.debug:
                    print(f"Debug - Row {index + 1}: authority='{authority}', agency='{agency}', db_match={((authority, agency) in db_dict)}")

        if mismatches:
            issues_df = pd.DataFrame(columns=['序号', '问题'])
            for i, _ in enumerate(mismatches, 1):
                issues_df = pd.concat([issues_df, pd.DataFrame({'序号': [i], '问题': ['C2填报单位名称与H2办理机关不一致']})], ignore_index=True)
            issue_filename = f"线索编号{datetime.now().strftime('%Y%m%d')}.xlsx"
            issue_path = os.path.join(Config.UPLOAD_FOLDER, issue_filename)
            issues_df.to_excel(issue_path, index=False)

        original_filename = file.filename.replace('.xlsx', '_副本.xlsx').replace('.xls', '_副本.xlsx')
        original_path = os.path.join(Config.UPLOAD_FOLDER, original_filename)
        with pd.ExcelWriter(original_path, engine='xlsxwriter') as writer:
            df.to_excel(writer, sheet_name='Sheet1', index=False)
            workbook = writer.book
            worksheet = writer.sheets['Sheet1']
            cell_format = workbook.add_format({'bg_color': '#FF0000'})
            for idx in mismatch_indices:
                worksheet.write(f'C{idx + 2}', df.at[idx, '填报单位名称'], cell_format)
                worksheet.write(f'H{idx + 2}', df.at[idx, '办理机关'], cell_format)

        if mismatches:
            flash('线索对比有异常，请查看生成的问题表', 'error')
        else:
            flash('文件上传并验证成功！', 'success')

        return redirect(request.url)
    except Exception as e:
        flash(f'文件处理失败: {str(e)}', 'error')
        return redirect(request.url)