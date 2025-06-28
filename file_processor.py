import os
import pandas as pd
import logging
from datetime import datetime
from flask import flash, redirect, url_for
from config import Config

logger = logging.getLogger(__name__)

# 确保这些导入路径是正确的，并且文件存在
from validation_rules.validation_core import get_validation_issues
from excel_formatter import format_excel
from validation_rules.case_validators import validate_case_relationships
from validation_rules.case_generators import generate_case_files


def process_upload(request, app):
    logger.info("开始处理线索登记表上传请求")
    if 'file' not in request.files:
        logger.error("未选择文件")
        flash('未选择文件', 'error')
        return redirect(request.url)
    file = request.files['file']
    if file.filename == '':
        logger.error("文件名为空")
        flash('未选择文件', 'info')
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
    logger.info(f"文件保存路径: {file_path}")
    os.makedirs(os.path.dirname(file_path), exist_ok=True)
    try:
        file.save(file_path)
    except Exception as e:
        logger.error(f"文件保存失败: {file_path} - {e}")
        flash(f'文件保存失败: {e}', 'error')
        return redirect(request.url)

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

        disposal_report_column = Config.COLUMN_MAPPINGS.get("disposal_report", "处置情况报告")
        if disposal_report_column not in df.columns:
            logger.error(f"Excel文件缺少必要表头: {disposal_report_column}")
            flash(f'Excel文件缺少必要的表头“{disposal_report_column}”', 'error')
            return redirect(request.url)
            
        if df[disposal_report_column].isnull().all():
            logger.error(f"线索登记表“{disposal_report_column}”字段为空")
            flash(f'线索登记表“{disposal_report_column}”字段为空', 'error')
            return redirect(request.url)

        # *** 重点：这里的 get_validation_issues 必须返回2个值 (mismatch_indices, issues_list) ***
        # issues_list 中的每个元素应该是一个3元组 (original_df_index, clue_code_value, issue_description)
        mismatch_indices, issues_list = get_validation_issues(df)
        logger.info(f"get_validation_issues 返回了 {len(mismatch_indices)} 个不匹配索引和 {len(issues_list)} 个问题。")

        if issues_list:
            data_for_issues_df = []
            for i, (original_df_index, clue_code_value, issue_description) in enumerate(issues_list):
                current_serial_number = i + 1
                data_for_issues_df.append({
                    '序号': current_serial_number,
                    '受理线索编码': clue_code_value,
                    '问题': issue_description
                })
            issues_df = pd.DataFrame(data_for_issues_df)
            
            issue_filename = f"线索编号{datetime.now().strftime('%Y%m%d')}.xlsx"
            issue_path = os.path.join(Config.CLUE_FOLDER, issue_filename)
            issues_df.to_excel(issue_path, index=False)
            logger.info(f"生成线索编号文件: {issue_path}")

        original_filename_copy = file.filename.replace('.xlsx', '_副本.xlsx').replace('.xls', '_副本.xlsx')
        original_path = os.path.join(Config.CLUE_FOLDER, original_filename_copy)
        # format_excel 函数也需要能够正确处理 issues_list 中每个元素是3元组的情况
        format_excel(df, mismatch_indices, original_path, issues_list)

        logger.info("程序执行成功")
        flash('文件上传处理成功！', 'success')
        return redirect(request.url)
    except Exception as e:
        logger.error(f"文件处理失败: {str(e)}", exc_info=True) # 打印详细 traceback
        flash(f'文件处理失败: {str(e)}', 'error')
        return redirect(request.url)

#-------------------------------------------------------------------------------------------------------
# 注意：以下是 `process_case_upload` 函数的修正，修复了之前由于我引入的`---`语法错误和参数不匹配问题。
#-------------------------------------------------------------------------------------------------------

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
    
    if not any(file.filename.lower().endswith(ext) for ext in Config.ALLOWED_EXTENSIONS):
        logger.error(f"立案登记表文件格式错误: {file.filename}")
        flash('上传文件格式不对', 'error')
        return redirect(request.url)
    if "立案登记表" not in file.filename:
        logger.error(f"立案登记表文件名不符合要求: {file.filename}")
        flash('文件名必须包含立案登记', 'error')
        return redirect(request.url)

    file_path = os.path.join(Config.CASE_FOLDER, file.filename)
    logger.info(f"立案登记表文件保存路径: {file_path}")
    os.makedirs(os.path.dirname(file_path), exist_ok=True)
    try:
        file.save(file_path)
    except Exception as e:
        logger.error(f"文件保存失败: {file_path} - {e}")
        flash(f'文件保存失败: {e}', 'error')
        return redirect(request.url)

    if not os.path.exists(file_path):
        logger.error(f"立案登记表文件保存失败: {file_path} 不存在")
        flash(f'文件保存失败: {file_path} 不存在', 'error')
        return redirect(request.url)
    logger.info(f"立案登记表文件保存成功: {file_path}")

    try:
        df = pd.read_excel(file_path)
        # 【最终修正】彻底移除“填报单位”，只保留“填报单位名称”
        required_headers = [
            "被调查人", "立案报告", "处分决定", "审查调查报告", "审理报告",
            "案件编码", "涉案人员编码", "立案时间", "立案决定书", 
            "纪委立案时间", "纪委立案机关", "监委立案时间", "监委立案机关", 
            "填报单位名称", # 确保这里是正确的
            "是否违反中央八项规定精神", "是否主动交代问题", 
            "结案时间", 
            "是否属于本应撤销党内职务，但本人没有党内职务而给予严重警告处分", 
            "追缴失职渎职滥用职权造成的损失金额", "审理受理时间", "审结时间",
            "审理机关" # 移除“填报单位”
        ]

        if not all(header in df.columns for header in required_headers):
            # 将缺少的表头找出并报告
            missing_headers = [header for header in required_headers if header not in df.columns]
            logger.error(f"立案登记表缺少必要表头: {missing_headers}")
            flash(f'Excel文件缺少必要的表头: {", ".join(missing_headers)}', 'error')
            return redirect(request.url)

        # 验证字段关系 - 接收所有返回值
        # 核心修改：在解包时添加 disposal_decision_keyword_mismatch_indices
        (mismatch_indices, gender_mismatch_indices, age_mismatch_indices, brief_case_details_mismatch_indices, issues_list, 
         birth_date_mismatch_indices, education_mismatch_indices, ethnicity_mismatch_indices, 
         party_member_mismatch_indices, party_joining_date_mismatch_indices, filing_time_mismatch_indices, 
         disciplinary_committee_filing_time_mismatch_indices, disciplinary_committee_filing_authority_mismatch_indices, 
         supervisory_committee_filing_time_mismatch_indices, supervisory_committee_filing_authority_mismatch_indices, 
         case_report_keyword_mismatch_indices, disposal_spirit_mismatch_indices, voluntary_confession_highlight_indices, 
         closing_time_mismatch_indices, no_party_position_warning_mismatch_indices,
         recovery_amount_highlight_indices, trial_acceptance_time_mismatch_indices, 
         trial_closing_time_mismatch_indices, trial_authority_agency_mismatch_indices,
         disposal_decision_keyword_mismatch_indices) = validate_case_relationships(df) # 增加了新的返回值

        # 生成副本和立案编号文件 - 传递所有参数
        # 核心修改：在调用 generate_case_files 时添加 disposal_decision_keyword_mismatch_indices
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
            ethnicity_mismatch_indices,
            party_member_mismatch_indices,
            party_joining_date_mismatch_indices,
            brief_case_details_mismatch_indices,
            filing_time_mismatch_indices,
            disciplinary_committee_filing_time_mismatch_indices,
            disciplinary_committee_filing_authority_mismatch_indices,
            supervisory_committee_filing_time_mismatch_indices,
            supervisory_committee_filing_authority_mismatch_indices,
            case_report_keyword_mismatch_indices,
            disposal_spirit_mismatch_indices,
            voluntary_confession_highlight_indices,
            closing_time_mismatch_indices,
            no_party_position_warning_mismatch_indices,
            recovery_amount_highlight_indices,
            trial_acceptance_time_mismatch_indices,
            trial_closing_time_mismatch_indices,
            trial_authority_agency_mismatch_indices,
            disposal_decision_keyword_mismatch_indices # 增加了新的参数
        )

        flash('文件上传处理成功！', 'success')
        logger.info("立案登记表处理成功")

        return redirect(request.url)
    except Exception as e:
        logger.error(f"立案登记表处理失败: {str(e)}", exc_info=True) # 打印详细 traceback
        flash(f'文件处理失败: {str(e)}', 'error')
        return redirect(request.url)