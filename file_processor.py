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
# 【党纪处分功能融合】导入正确的党纪处分验证函数
# 【新增政务处分功能】导入政务处分验证函数
from validation_rules.case_validation_sanctions import validate_disciplinary_sanction, validate_administrative_sanction


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
        # issues_list 中的每个元素应该是一个字典，例如 {"案件编码": ..., "问题": ..., "风险等级": ..., "行号": ...}
        mismatch_indices, issues_list = get_validation_issues(df)
        logger.info(f"get_validation_issues 返回了 {len(mismatch_indices)} 个不匹配索引和 {len(issues_list)} 个问题。")

        if issues_list:
            data_for_issues_df = []
            for i, issue_item in enumerate(issues_list): # 遍历字典
                current_serial_number = i + 1
                data_for_issues_df.append({
                    '序号': current_serial_number,
                    '受理线索编码': issue_item.get('受理线索编码', 'N/A'), # 从字典中获取
                    '问题': issue_item.get('问题描述', '无描述') # 从字典中获取
                })
            issues_df = pd.DataFrame(data_for_issues_df)
            
            issue_filename = f"线索编号{datetime.now().strftime('%Y%m%d')}.xlsx"
            issue_path = os.path.join(Config.CLUE_FOLDER, issue_filename)
            issues_df.to_excel(issue_path, index=False)
            logger.info(f"生成线索编号文件: {issue_path}")

        original_filename_copy = file.filename.replace('.xlsx', '_副本.xlsx').replace('.xls', '_副本.xlsx')
        original_path = os.path.join(Config.CLUE_FOLDER, original_filename_copy)
        # format_excel 函数也需要能够正确处理 issues_list 中每个元素是字典的情况
        format_excel(df, mismatch_indices, original_path, issues_list)

        logger.info("程序执行成功")
        flash('文件上传处理成功！', 'success')
        return redirect(request.url)
    except Exception as e:
        logger.error(f"文件处理失败: {str(e)}", exc_info=True) # 打印详细 traceback
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
        required_headers = [
            "被调查人", "性别", "年龄", "出生年月", "学历", "民族", 
            "是否中共党员", "入党时间", 
            "立案报告", "处分决定", "审查调查报告", "审理报告", "简要案情", 
            "案件编码", "涉案人员编码", "立案时间", "立案决定书", 
            "纪委立案时间", "纪委立案机关", "监委立案时间", "监委立案机关", 
            "填报单位名称", 
            "是否违反中央八项规定精神", "是否主动交代问题", 
            "结案时间", 
            "是否属于本应撤销党内职务，但本人没有党内职务而给予严重警告处分", 
            "追缴失职渎职滥用职权造成的损失金额", "审理受理时间", "审结时间",
            "审理机关", 
            "收缴金额（万元）",
            "没收金额",
            "责令退赔金额",
            "登记上交金额"
        ]

        # 确保党纪处分和政务处分字段也在必填头中
        if Config.COLUMN_MAPPINGS.get("disciplinary_sanction") not in required_headers:
            required_headers.append(Config.COLUMN_MAPPINGS.get("disciplinary_sanction"))
        if Config.COLUMN_MAPPINGS.get("administrative_sanction") not in required_headers: # 【新增】政务处分添加到必填头
            required_headers.append(Config.COLUMN_MAPPINGS.get("administrative_sanction"))


        if not all(header in df.columns for header in required_headers):
            missing_headers = [header for header in required_headers if header not in df.columns]
            logger.error(f"立案登记表缺少必要表头: {missing_headers}")
            flash(f'Excel文件缺少必要的表头: {", ".join(missing_headers)}', 'error')
            return redirect(request.url)

        # 调用主要的校验函数
        # ！！！关键修改在这里：重新包含 disciplinary_sanction_mismatch_indices
        (mismatch_indices, gender_mismatch_indices, age_mismatch_indices, brief_case_details_mismatch_indices, issues_list, 
         birth_date_mismatch_indices, education_mismatch_indices, ethnicity_mismatch_indices, 
         party_member_mismatch_indices, party_joining_date_mismatch_indices, filing_time_mismatch_indices, 
         disciplinary_committee_filing_time_mismatch_indices, disciplinary_committee_filing_authority_mismatch_indices, 
         supervisory_committee_filing_time_mismatch_indices, supervisory_committee_filing_authority_mismatch_indices, 
         case_report_keyword_mismatch_indices, disposal_spirit_mismatch_indices, voluntary_confession_highlight_indices, 
         closing_time_mismatch_indices, no_party_position_warning_mismatch_indices,
         recovery_amount_highlight_indices, trial_acceptance_time_mismatch_indices, 
         trial_closing_time_mismatch_indices, trial_authority_agency_mismatch_indices,
         disposal_decision_keyword_mismatch_indices,
         trial_report_non_representative_mismatch_indices, 
         trial_report_detention_mismatch_indices,
         confiscation_amount_indices,
         confiscation_of_property_amount_indices,
         compensation_amount_highlight_indices,
         registered_handover_amount_indices,
         disciplinary_sanction_mismatch_indices # <-- 重新添加，与 case_validators.py 的 32 个返回值匹配
         ) = validate_case_relationships(df) 
        
        # 调用党纪处分规则 (这个调用是独立的，并且正确)
        # 它的返回值应该合并到 issues_list 中，如果 validate_disciplinary_sanction 已经做了
        # 如果 validate_disciplinary_sanction 返回的是它自己的不匹配索引，我们需要用一个新变量接收
        # 再次检查 case_validation_sanctions.py 中的 validate_disciplinary_sanction 函数签名
        # 我之前的假设是它返回不匹配索引，并修改 issues_list
        # 如果它只返回不匹配索引，这里应该像 administrative_sanction_mismatch_indices 一样接收
        # 根据 case_validators.py，它是直接调用并更新了 issues_list
        # 所以，我们不再在这里重复调用 validate_disciplinary_sanction
        # 因为它已经在 case_validators.py 内部被 validate_case_relationships 调用了
        # 移除重复调用：
        # disciplinary_sanction_mismatch_indices = validate_disciplinary_sanction(df, issues_list) # 这一行应该被移除
        
        # 【新增】调用政务处分规则 (这个调用是新的，且需要保留)
        administrative_sanction_mismatch_indices = validate_administrative_sanction(df, issues_list)

        mismatch_indices = list(set(mismatch_indices))
        
        # Ensure issues_list contains dictionaries for proper de-duplication
        issues_list_unique = []
        seen_issues = set()
        for issue_dict_or_tuple in issues_list:
            if isinstance(issue_dict_or_tuple, tuple):
                issue_dict_converted = {
                    "行号": issue_dict_or_tuple[0] + 2, # Adjust row number
                    "案件编码": issue_dict_or_tuple[1],
                    "涉案人员编码": issue_dict_or_tuple[2],
                    "问题描述": issue_dict_or_tuple[3],
                    "风险等级": issue_dict_or_tuple[4] if len(issue_dict_or_tuple) > 4 else "中" # 确保风险等级也被传递
                }
                issue_hashable = frozenset(issue_dict_converted.items())
                if issue_hashable not in seen_issues:
                    issues_list_unique.append(issue_dict_converted)
                    seen_issues.add(issue_hashable)
            elif isinstance(issue_dict_or_tuple, dict):
                issue_hashable = frozenset(issue_dict_or_tuple.items()) 
                if issue_hashable not in seen_issues:
                    issues_list_unique.append(issue_dict_or_tuple)
                    seen_issues.add(issue_hashable)
            else:
                logger.warning(f"Unexpected item type in issues_list: {type(issue_dict_or_tuple)}. Skipping de-duplication for this item.")
        issues_list = issues_list_unique

        copy_path, case_num_path = generate_case_files(
            df, 
            file.filename, 
            Config.BASE_UPLOAD_FOLDER, 
            mismatch_indices, 
            gender_mismatch_indices, 
            issues_list, # issues_list should now contain dictionaries
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
            disposal_decision_keyword_mismatch_indices,
            trial_report_non_representative_mismatch_indices, 
            trial_report_detention_mismatch_indices,
            confiscation_amount_indices,
            confiscation_of_property_amount_indices,
            compensation_amount_highlight_indices,
            registered_handover_amount_indices,
            disciplinary_sanction_mismatch_indices,
            administrative_sanction_mismatch_indices # 【新增】将政务处分不匹配索引传递给 generate_case_files
        )

        flash('文件上传处理成功！', 'success')
        logger.info("立案登记表处理成功")

        return redirect(request.url)
    except Exception as e:
        logger.error(f"立案登记表处理失败: {str(e)}", exc_info=True)
        flash(f'文件处理失败: {str(e)}', 'error')
        return redirect(request.url)