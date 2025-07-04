import os
import pandas as pd
import logging
from datetime import datetime # 导入 datetime 用于获取当前日期
from flask import flash, redirect, url_for
from config import Config
from werkzeug.utils import secure_filename

# 导入验证规则模块和辅助函数
# 假设 validation_rules 目录已添加到 sys.path
try:
    from validation_rules.clue_validation import validate_clue_data
    # 导入 case_validators 中的 validate_case_relationships
    from validation_rules.case_validators import validate_case_relationships 
    from validation_rules.case_generators import generate_case_files # 确保导入 generate_case_files
    from excel_formatter import format_excel # 确保导入 format_excel
    from validation_rules.case_validation_sanctions import validate_administrative_sanction # 确保导入政务处分验证
except ImportError as e:
    logging.getLogger(__name__).error(f"无法导入必要的模块或函数: {e}", exc_info=True)
    # 在生产环境中，这里可能需要更健壮的错误处理，例如退出应用或禁用相关功能

logger = logging.getLogger(__name__)

def allowed_file(filename):
    """
    检查文件扩展名是否在允许的列表中。
    """
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in Config.ALLOWED_EXTENSIONS

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
        flash(f'文件名必须包含“{Config.REQUIRED_FILENAME_PATTERN}”', 'error')
        return redirect(request.url)

    # 从 app.config 中获取 CLUE_FOLDER
    clue_folder = app.config['CLUE_FOLDER']
    original_filename = secure_filename(file.filename)
    file_path = os.path.join(clue_folder, original_filename)
    
    logger.info(f"文件保存路径: {file_path}")
    os.makedirs(os.path.dirname(file_path), exist_ok=True)

    try:
        file.save(file_path)
    except Exception as e:
        logger.error(f"文件保存失败: {file_path} - {e}", exc_info=True)
        flash(f'文件保存失败: {e}', 'error')
        return redirect(request.url)

    if not os.path.exists(file_path):
        logger.error(f"文件保存失败: {file_path} 不存在")
        flash(f'文件保存失败: {file_path} 不存在', 'error')
        return redirect(request.url)
    logger.info(f"文件保存成功: {file_path}")

    try:
        df = pd.read_excel(file_path)
        
        # 检查必要的表头
        required_headers = Config.CLUE_REQUIRED_HEADERS + [
            Config.COLUMN_MAPPINGS["organization_measure"], 
            Config.COLUMN_MAPPINGS["acceptance_time"]
        ]
        if not all(header in df.columns for header in required_headers):
            missing_headers = [header for header in required_headers if header not in df.columns]
            logger.error(f"缺少必要表头: {missing_headers}")
            flash(f'Excel文件缺少必要的表头: {", ".join(missing_headers)}', 'error')
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

        # 调用线索数据验证函数
        issues_list, error_count = validate_clue_data(df, app.config)
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
                        '问题': issue_dict.get('问题描述', '无描述'),
                        '风险等级': issue_dict.get('风险等级', '中')
                    })
                    seen_issues.add(issue_hashable)
            
            issues_df = pd.DataFrame(data_for_issues_df)
            
            issue_filename = f"线索编号{Config().TODAY_DATE}.xlsx" # 使用 Config().TODAY_DATE 获取当前日期
            issue_path = os.path.join(clue_folder, issue_filename)
            issues_df.to_excel(issue_path, index=False)
            logger.info(f"生成线索编号文件: {issue_path}")

        # 生成带有高亮和问题的副本文件
        original_filename_copy = original_filename.replace('.xlsx', '_副本.xlsx').replace('.xls', '_副本.xlsx')
        original_path_copy = os.path.join(clue_folder, original_filename_copy)
        
        # 对于线索表，mismatch_indices可以是一个空集合，因为高亮主要基于issues_list
        # 并且 format_excel 的 apply_clue_table_formats 会处理 issues_list。
        # 其他所有 mismatch_indices 集合都为空，因为它们是针对案件表的。
        format_excel(df, set(), original_path_copy, issues_list,
                     gender_mismatch_indices=set(), age_mismatch_indices=set(),
                     birth_date_mismatch_indices=set(), education_mismatch_indices=set(), ethnicity_mismatch_indices=set(),
                     party_member_mismatch_indices=set(), party_joining_date_mismatch_indices=set(),
                     brief_case_details_mismatch_indices=set(), filing_time_mismatch_indices=set(),
                     disciplinary_committee_filing_time_mismatch_indices=set(),
                     disciplinary_committee_filing_authority_mismatch_indices=set(),
                     supervisory_committee_filing_time_mismatch_indices=set(),
                     supervisory_committee_filing_authority_mismatch_indices=set(),
                     case_report_keyword_mismatch_indices=set(), disposal_spirit_mismatch_indices=set(),
                     voluntary_confession_highlight_indices=set(), closing_time_mismatch_indices=set(),
                     no_party_position_warning_mismatch_indices=set(),
                     recovery_amount_highlight_indices=set(),
                     trial_acceptance_time_mismatch_indices=set(),
                     trial_closing_time_mismatch_indices=set(),
                     trial_authority_agency_mismatch_indices=set(),
                     disposal_decision_keyword_mismatch_indices=set(),
                     trial_report_non_representative_mismatch_indices=set(),
                     trial_report_detention_mismatch_indices=set(),
                     confiscation_amount_indices=set(),
                     confiscation_of_property_amount_indices=set(),
                     compensation_amount_highlight_indices=set(),
                     registered_handover_amount_indices=set(),
                     disciplinary_sanction_mismatch_indices=set(),
                     administrative_sanction_mismatch_indices=set()
                    )

        logger.info("线索登记表处理成功")
        flash('文件上传处理成功！', 'success')
        return redirect(request.url)
    except Exception as e:
        logger.error(f"线索登记表处理失败: {str(e)}", exc_info=True)
        flash(f'文件处理失败: {str(e)}', 'error')
        return redirect(request.url)


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
    logger.info("开始处理立案登记表上传请求")

    # 将 'file' 改为 'case_file' 以匹配 routes.py 中的表单字段名
    if 'case_file' not in request.files:
        logger.error("未选择立案登记表文件")
        flash('未选择文件', 'error')
        return redirect(request.url)

    file = request.files['case_file'] # 同样修改这里
    if file.filename == '':
        logger.error("立案登记表文件名为空")
        flash('未选择文件', 'info')
        return redirect(request.url)
    
    if not any(file.filename.lower().endswith(ext) for ext in Config.ALLOWED_EXTENSIONS):
        logger.error(f"立案登记表文件格式错误: {file.filename}")
        flash('上传文件格式不对', 'error')
        return redirect(request.url)

    if "立案登记表" not in file.filename: # 确保这里检查的是“立案登记表”
        logger.error(f"立案登记表文件名不符合要求: {file.filename}")
        flash('文件名必须包含“立案登记表”', 'error')
        return redirect(request.url)

    # 从 app.config 中获取 CASE_FOLDER
    case_folder = app.config['CASE_FOLDER']
    original_filename = secure_filename(file.filename)
    file_path = os.path.join(case_folder, original_filename)
    
    logger.info(f"立案登记表文件保存路径: {file_path}")
    os.makedirs(os.path.dirname(file_path), exist_ok=True)

    try:
        file.save(file_path)
    except Exception as e:
        logger.error(f"文件保存失败: {file_path} - {e}", exc_info=True)
        flash(f'文件保存失败: {e}', 'error')
        return redirect(request.url)

    if not os.path.exists(file_path):
        logger.error(f"立案登记表文件保存失败: {file_path} 不存在")
        flash(f'文件保存失败: {file_path} 不存在', 'error')
        return redirect(request.url)
    logger.info(f"立案登记表文件保存成功: {file_path}")

    try:
        df = pd.read_excel(file_path)
        
        # 检查必要的表头
        required_headers = [
            app.config['COLUMN_MAPPINGS']["investigated_person"], app.config['COLUMN_MAPPINGS']["gender"], app.config['COLUMN_MAPPINGS']["age"], 
            app.config['COLUMN_MAPPINGS']["birth_date"], app.config['COLUMN_MAPPINGS']["education"], app.config['COLUMN_MAPPINGS']["ethnicity"], 
            app.config['COLUMN_MAPPINGS']["party_member"], app.config['COLUMN_MAPPINGS']["party_joining_date"], 
            app.config['COLUMN_MAPPINGS']["case_report"], app.config['COLUMN_MAPPINGS']["disciplinary_decision"], 
            app.config['COLUMN_MAPPINGS']["investigation_report"], app.config['COLUMN_MAPPINGS']["trial_report"], 
            app.config['COLUMN_MAPPINGS']["brief_case_details"], app.config['COLUMN_MAPPINGS']["case_code"], 
            app.config['COLUMN_MAPPINGS']["person_code"], app.config['COLUMN_MAPPINGS']["filing_time"], 
            app.config['COLUMN_MAPPINGS']["filing_decision_doc"], app.config['COLUMN_MAPPINGS']["disciplinary_committee_filing_time"], 
            app.config['COLUMN_MAPPINGS']["disciplinary_committee_filing_authority"], app.config['COLUMN_MAPPINGS']["supervisory_committee_filing_time"], 
            app.config['COLUMN_MAPPINGS']["supervisory_committee_filing_authority"], app.config['COLUMN_MAPPINGS']["reporting_agency"], 
            app.config['COLUMN_MAPPINGS']["central_eight_provisions"], app.config['COLUMN_MAPPINGS']["voluntary_confession"], 
            app.config['COLUMN_MAPPINGS']["closing_time"], app.config['COLUMN_MAPPINGS']["no_party_position_warning"], 
            app.config['COLUMN_MAPPINGS']["recovery_amount"], app.config['COLUMN_MAPPINGS']["trial_acceptance_time"], 
            app.config['COLUMN_MAPPINGS']["trial_closing_time"], app.config['COLUMN_MAPPINGS']["trial_authority"], 
            app.config['COLUMN_MAPPINGS']["confiscation_amount"], app.config['COLUMN_MAPPINGS']["confiscation_of_property_amount"], 
            app.config['COLUMN_MAPPINGS']["compensation_amount"], app.config['COLUMN_MAPPINGS']["registered_handover_amount"]
        ]

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
         administrative_sanction_mismatch_indices # 添加此行以匹配 validate_case_relationships 的 33 个返回值
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
                                         list(administrative_sanction_mismatch_indices) # 添加政务处分不匹配索引
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
        # 传递 case_folder 而不是 UPLOAD_FOLDER
        copy_path, case_num_path = generate_case_files(
            df, 
            original_filename, # 传递原始文件名
            case_folder, # 使用 case_folder 作为输出目录
            all_mismatch_indices, # 使用合并后的所有不匹配索引
            gender_mismatch_indices, 
            issues_list, # issues_list 应该包含字典
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
