# case_file_processor.py
import os
import pandas as pd
import logging
from flask import flash, redirect, url_for

# 导入通用函数
from .upload_utils import handle_file_upload_and_initial_checks

# 导入验证规则模块和辅助函数
try:
    from validation_rules.case_validators import validate_case_relationships
    from validation_rules.case_generators import generate_case_files
    from excel_formatter import format_excel
except ImportError as e:
    logging.getLogger(__name__).error(f"无法导入必要的模块或函数: {e}", exc_info=True)

logger = logging.getLogger(__name__)

def process_case_upload(request, app):
    """
    处理立案登记表文件的上传、保存和验证。
    执行多项复杂验证规则，并生成带有高亮和问题列表的副本文件。

    参数:
        request (flask.request): Flask 请求对象，包含上传的文件。
        app (flask.Flask): Flask 应用实例，用于访问 app.config。

    返回:
        flask.redirect: 重定向到上传页面。
    """
    # 使用通用函数处理文件上传和初步检查
    file_path, original_filename, df, error_response = handle_file_upload_and_initial_checks(
        request, app, 'case_file', 'CASE_FOLDER', '立案登记表', '立案登记表'
    )
    if error_response:
        return error_response

    try:
        # 检查必要的表头，使用 app.config
        required_headers = list(app.config['CASE_REQUIRED_HEADERS'])

        # 确保党纪处分和政务处分字段也在必填头中
        disciplinary_sanction_col = app.config['COLUMN_MAPPINGS'].get("disciplinary_sanction")
        if disciplinary_sanction_col and disciplinary_sanction_col not in required_headers:
            required_headers.append(disciplinary_sanction_col)

        administrative_sanction_col = app.config['COLUMN_MAPPINGS'].get("administrative_sanction")
        if administrative_sanction_col and administrative_sanction_col not in required_headers:
            required_headers.append(administrative_sanction_col)

        if not all(header in df.columns for header in required_headers):
            missing_headers = [header for header in required_headers if header not in df.columns]
            logger.error(f"立案登记表缺少必要表头: {missing_headers}")
            flash(f'Excel文件缺少必要的表头: {", ".join(missing_headers)}', 'error')
            return redirect(request.url)

        # 初始化 issues_list
        issues_list = []

        # 调用主要的校验函数，接收所有返回的索引和问题列表
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
         disciplinary_sanction_mismatch_indices,
         administrative_sanction_mismatch_indices
         ) = validate_case_relationships(df, app.config, issues_list) # 传递 app.config 和 issues_list

        # 将所有不匹配索引合并并去重
        all_mismatch_indices = list(set(list(mismatch_indices) +
                                         list(gender_mismatch_indices) +
                                         list(age_mismatch_indices) +
                                         list(brief_case_details_mismatch_indices) +
                                         list(birth_date_mismatch_indices) +
                                         list(education_mismatch_indices) +
                                         list(ethnicity_mismatch_indices) +
                                         list(party_member_mismatch_indices) +
                                         list(party_joining_date_mismatch_indices) +
                                         list(filing_time_mismatch_indices) +
                                         list(disciplinary_committee_filing_time_mismatch_indices) +
                                         list(disciplinary_committee_filing_authority_mismatch_indices) +
                                         list(supervisory_committee_filing_time_mismatch_indices) +
                                         list(supervisory_committee_filing_authority_mismatch_indices) +
                                         list(case_report_keyword_mismatch_indices) +
                                         list(disposal_spirit_mismatch_indices) +
                                         list(voluntary_confession_highlight_indices) +
                                         list(closing_time_mismatch_indices) +
                                         list(no_party_position_warning_mismatch_indices) +
                                         list(recovery_amount_highlight_indices) +
                                         list(trial_acceptance_time_mismatch_indices) +
                                         list(trial_closing_time_mismatch_indices) +
                                         list(trial_authority_agency_mismatch_indices) +
                                         list(disposal_decision_keyword_mismatch_indices) +
                                         list(trial_report_non_representative_mismatch_indices) +
                                         list(trial_report_detention_mismatch_indices) +
                                         list(confiscation_amount_indices) +
                                         list(confiscation_of_property_amount_indices) +
                                         list(compensation_amount_highlight_indices) +
                                         list(registered_handover_amount_indices) +
                                         list(disciplinary_sanction_mismatch_indices) +
                                         list(administrative_sanction_mismatch_indices)
                                         ))

        # 确保 issues_list 包含字典，并进行去重
        issues_list_unique = []
        seen_issues = set()
        for issue_dict_or_tuple in issues_list:
            if isinstance(issue_dict_or_tuple, tuple):
                # 假设元组格式为 (行号, 案件编码, 涉案人员编码, 问题描述, 风险等级)
                issue_dict_converted = {
                    "行号": issue_dict_or_tuple[0] + 2, # 调整行号，因为Excel从1开始，且有表头
                    "案件编码": issue_dict_or_tuple[1],
                    "涉案人员编码": issue_dict_or_tuple[2],
                    "问题描述": issue_dict_or_tuple[3],
                    "风险等级": issue_dict_or_tuple[4] if len(issue_dict_or_tuple) > 4 else "中"
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
                logger.warning(f"issues_list 中发现未知类型项: {type(issue_dict_or_tuple)}. 跳过去重。")
        issues_list = issues_list_unique

        # 生成案件文件和问题报告
        copy_path, case_num_path = generate_case_files(
            df,
            original_filename,
            app.config['CASE_FOLDER'], # 直接使用 app.config
            all_mismatch_indices,
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
            disposal_decision_keyword_mismatch_indices,
            trial_report_non_representative_mismatch_indices,
            trial_report_detention_mismatch_indices,
            confiscation_amount_indices,
            confiscation_of_property_amount_indices,
            compensation_amount_highlight_indices,
            registered_handover_amount_indices,
            disciplinary_sanction_mismatch_indices,
            administrative_sanction_mismatch_indices
        )

        flash('文件上传处理成功！', 'success')
        logger.info("立案登记表处理成功")

        return redirect(request.url)
    except Exception as e:
        logger.error(f"立案登记表处理失败: {str(e)}", exc_info=True)
        flash(f'文件处理失败: {str(e)}', 'error')
        return redirect(request.url)