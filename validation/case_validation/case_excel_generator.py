import os
import pandas as pd
from datetime import datetime
import logging
import os
from db_utils import get_authority_agency_dict
from .case_validation_additional import validate_name_rules, validate_gender_rules, validate_age_rules, validate_birth_date_rules, validate_education_rules, validate_ethnicity_rules, validate_party_member_rules, validate_party_joining_date_rules, validate_brief_case_details_rules, validate_filing_time_rules, validate_disciplinary_committee_filing_time_rules, validate_supervisory_committee_filing_time_rules, validate_disciplinary_committee_filing_authority_rules, validate_supervisory_committee_filing_authority_rules, validate_case_report_rules, validate_central_eight_provisions_rules, validate_voluntary_confession_rules, validate_disciplinary_sanction_rules, validate_no_party_position_warning_rules
from .case_validation_administrative_sanction import validate_administrative_sanction_rules
from .case_validation_confiscation_amount import validate_confiscation_amount_rules
from .case_validation_confiscation_of_property_amount import validate_confiscation_of_property_amount_rules
from .case_validation_compensation_amount import validate_compensation_amount_rules
from .case_validation_trial_acceptance_time import validate_trial_acceptance_time_rules
from .case_validation_recovery_amount import validate_recovery_amount_rules
from .case_validation_trial_closing_time import validate_trial_closing_time_rules
from .case_validation_trial_authority import validate_trial_authority_rules
from .case_validation_trial_report import validate_trial_report_rules
from .case_validation_disciplinary_decision import validate_disciplinary_decision_rules
from .case_timestamp_rules import validate_registered_handover_amount_single_row

logger = logging.getLogger(__name__)

def generate_investigatee_number_file(df, original_filename, upload_dir, app_config):
    """
    生成独立的被调查人验证编号表Excel文件。
    只包含被调查人相关的验证规则结果。
    
    参数:
    df (pd.DataFrame): 原始Excel数据的DataFrame。
    original_filename (str): 原始上传的文件名。
    upload_dir (str): 上传文件的根目录。
    app_config: 应用配置对象。
    
    返回:
    str: 生成的立案编号表文件路径，如果生成失败返回None。
    """
    try:
        # 创建输出目录
        case_dir = upload_dir
        os.makedirs(case_dir, exist_ok=True)
        
        # 初始化问题列表和不匹配索引
        issues_list = []
        mismatch_indices = set()
        
        # 初始化机关单位对应表查询集合
        authority_agency_data = get_authority_agency_dict()
        authority_agency_lookup = set()
        for row_db in authority_agency_data:
            authority_agency_lookup.add((row_db['authority'], row_db['agency'], row_db['category']))
        
        # 遍历每一行数据，执行被调查人验证规则
        for index, row in df.iterrows():
            try:
                # 提取必要的字段（使用动态列映射）
                excel_case_code = str(row.get(app_config['COLUMN_MAPPINGS']['case_code'], '')).strip()
                excel_person_code = str(row.get(app_config['COLUMN_MAPPINGS']['person_code'], '')).strip()
                investigated_person = str(row.get(app_config['COLUMN_MAPPINGS']['investigated_person'], '')).strip()
                report_text_raw = str(row.get(app_config['COLUMN_MAPPINGS']['case_report'], '')).strip()
                decision_text_raw = str(row.get(app_config['COLUMN_MAPPINGS']['disciplinary_decision'], '')).strip()
                investigation_text_raw = str(row.get(app_config['COLUMN_MAPPINGS']['investigation_report'], '')).strip()
                trial_text_raw = str(row.get(app_config['COLUMN_MAPPINGS']['trial_report'], '')).strip()
                
                # 跳过空行或关键字段为空的行
                if not investigated_person or not excel_case_code:
                    continue
                    
                # 执行被调查人验证规则
                validate_name_rules(
                    row, index, excel_case_code, excel_person_code, issues_list, mismatch_indices,
                    investigated_person, report_text_raw, decision_text_raw, 
                    investigation_text_raw, trial_text_raw, app_config
                )
                
                # 执行性别验证规则
                excel_gender = str(row.get(app_config['COLUMN_MAPPINGS']['gender'], '')).strip()
                gender_mismatch_indices = set()
                validate_gender_rules(
                    row, index, excel_case_code, excel_person_code, issues_list, gender_mismatch_indices,
                    excel_gender, report_text_raw, decision_text_raw,
                    investigation_text_raw, trial_text_raw, app_config
                )
                
                # 执行年龄验证规则
                excel_age = None
                if pd.notna(row.get(app_config['COLUMN_MAPPINGS']['age'])):
                    try:
                        excel_age = int(row.get(app_config['COLUMN_MAPPINGS']['age']))
                    except ValueError:
                        logger.warning(f"行 {index + 1} - Excel '{app_config['COLUMN_MAPPINGS']['age']}' 字段 '{row.get(app_config['COLUMN_MAPPINGS']['age'])}' 不是有效数字。")
                age_mismatch_indices = set()
                current_year = datetime.now().year
                validate_age_rules(
                    row, index, excel_case_code, excel_person_code, issues_list, age_mismatch_indices,
                    excel_age, current_year, report_text_raw, decision_text_raw,
                    investigation_text_raw, trial_text_raw, app_config
                )
                
                # 执行出生年月验证规则
                excel_birth_date = str(row.get(app_config['COLUMN_MAPPINGS']['birth_date'], '')).strip()
                birth_date_mismatch_indices = set()
                validate_birth_date_rules(
                    row, index, excel_case_code, excel_person_code, issues_list, birth_date_mismatch_indices,
                    excel_birth_date, report_text_raw, decision_text_raw,
                    investigation_text_raw, trial_text_raw, app_config
                )
                
                # 执行学历验证规则
                excel_education = str(row.get(app_config['COLUMN_MAPPINGS']['education'], '')).strip()
                education_mismatch_indices = set()
                validate_education_rules(
                    row, index, excel_case_code, excel_person_code, issues_list, education_mismatch_indices,
                    excel_education, report_text_raw, decision_text_raw,
                    investigation_text_raw, trial_text_raw, app_config
                )
                
                # 执行民族验证规则
                excel_ethnicity = str(row.get(app_config['COLUMN_MAPPINGS']['ethnicity'], '')).strip()
                ethnicity_mismatch_indices = set()
                validate_ethnicity_rules(
                    row, index, excel_case_code, excel_person_code, issues_list, ethnicity_mismatch_indices,
                    excel_ethnicity, report_text_raw, decision_text_raw,
                    investigation_text_raw, trial_text_raw, app_config
                )
                
                # 执行是否中共党员验证规则
                excel_party_member = str(row.get(app_config['COLUMN_MAPPINGS']['party_member'], '')).strip()
                party_member_mismatch_indices = set()
                validate_party_member_rules(
                    row, index, excel_case_code, excel_person_code, issues_list, party_member_mismatch_indices,
                    excel_party_member, report_text_raw, decision_text_raw, app_config
                )
                
                # 执行入党时间验证规则
                excel_party_joining_date = str(row.get(app_config['COLUMN_MAPPINGS']['party_joining_date'], '')).strip()
                party_joining_date_mismatch_indices = set()
                validate_party_joining_date_rules(
                    row, index, excel_case_code, excel_person_code, issues_list, party_joining_date_mismatch_indices,
                    excel_party_member, excel_party_joining_date, report_text_raw, app_config
                )
                
                # 执行简要案情验证规则
                excel_brief_case_details = str(row.get(app_config['COLUMN_MAPPINGS']['brief_case_details'], '')).strip()
                brief_case_details_mismatch_indices = set()
                validate_brief_case_details_rules(
                    row, index, excel_case_code, excel_person_code, issues_list, brief_case_details_mismatch_indices,
                    excel_brief_case_details, investigated_person, report_text_raw, decision_text_raw, app_config
                )
                
                # 执行立案时间验证规则
                excel_filing_time = str(row.get(app_config['COLUMN_MAPPINGS']['filing_time'], '')).strip()
                excel_filing_decision_doc = str(row.get(app_config['COLUMN_MAPPINGS']['filing_decision_doc'], '')).strip()
                filing_time_mismatch_indices = set()
                validate_filing_time_rules(
                    row, index, excel_case_code, excel_person_code, issues_list, filing_time_mismatch_indices,
                    excel_filing_time, excel_filing_decision_doc, app_config
                )
                
                # 执行纪委立案时间验证规则
                excel_disciplinary_committee_filing_time = str(row.get(app_config['COLUMN_MAPPINGS']['disciplinary_committee_filing_time'], '')).strip()
                disciplinary_committee_filing_time_mismatch_indices = set()
                validate_disciplinary_committee_filing_time_rules(
                    row, index, excel_case_code, excel_person_code, issues_list, disciplinary_committee_filing_time_mismatch_indices,
                    excel_disciplinary_committee_filing_time, excel_filing_decision_doc, app_config
                )
                
                # 执行监委立案时间验证规则
                excel_supervisory_committee_filing_time = str(row.get(app_config['COLUMN_MAPPINGS']['supervisory_committee_filing_time'], '')).strip()
                supervisory_committee_filing_time_mismatch_indices = set()
                validate_supervisory_committee_filing_time_rules(
                    row, index, excel_case_code, excel_person_code, issues_list, supervisory_committee_filing_time_mismatch_indices,
                    excel_supervisory_committee_filing_time, excel_filing_decision_doc, app_config
                )
                
                # 执行纪委立案机关验证规则
                excel_disciplinary_committee_filing_authority = str(row.get(app_config['COLUMN_MAPPINGS']['disciplinary_committee_filing_authority'], '')).strip()
                excel_reporting_unit_name = str(row.get(app_config['COLUMN_MAPPINGS']['reporting_agency'], '')).strip()
                disciplinary_committee_filing_authority_mismatch_indices = set()
                validate_disciplinary_committee_filing_authority_rules(
                    row, index, excel_case_code, excel_person_code, issues_list, disciplinary_committee_filing_authority_mismatch_indices,
                    excel_disciplinary_committee_filing_authority, excel_reporting_unit_name, authority_agency_lookup, app_config
                )
                
                # 执行监委立案机关验证规则
                excel_supervisory_committee_filing_authority = str(row.get(app_config['COLUMN_MAPPINGS']['supervisory_committee_filing_authority'], '')).strip()
                supervisory_committee_filing_authority_mismatch_indices = set()
                validate_supervisory_committee_filing_authority_rules(
                    row, index, excel_case_code, excel_person_code, issues_list, supervisory_committee_filing_authority_mismatch_indices,
                    excel_supervisory_committee_filing_authority, excel_reporting_unit_name, authority_agency_lookup, app_config
                )
                
                # 立案报告规则验证
                excel_case_report = row.get(app_config['COLUMN_MAPPINGS']['case_report'], '')
                case_report_mismatch_indices = set()
                
                if pd.notna(excel_case_report) and excel_case_report.strip():
                    # 获取其他报告字段
                    excel_disciplinary_decision = row.get(app_config['COLUMN_MAPPINGS']['disciplinary_decision'], '')
                    excel_trial_report = row.get(app_config['COLUMN_MAPPINGS']['trial_report'], '')
                    excel_investigation_report = row.get(app_config['COLUMN_MAPPINGS']['investigation_report'], '')
                    
                    # 定义需要检查的关键字
                    case_report_keywords_to_check = ['贪污', '受贿', '挪用', '滥用职权', '玩忽职守']
                    
                    validate_case_report_rules(
                        row, index, excel_case_code, excel_person_code, issues_list, case_report_mismatch_indices,
                        case_report_keywords_to_check, excel_case_report, excel_disciplinary_decision, 
                        excel_investigation_report, excel_trial_report, app_config
                    )
                
                # 是否违反中央八项规定精神规则验证
                excel_central_eight_provisions = row.get(app_config['COLUMN_MAPPINGS']['central_eight_provisions'], '')
                central_eight_provisions_mismatch_indices = set()
                
                if pd.notna(excel_central_eight_provisions):
                    excel_central_eight_provisions = str(excel_central_eight_provisions).strip()
                    
                    validate_central_eight_provisions_rules(
                        row, index, excel_case_code, excel_person_code, issues_list, central_eight_provisions_mismatch_indices,
                        excel_central_eight_provisions, excel_disciplinary_decision, app_config
                    )
                
                # 是否主动交代问题规则
                excel_voluntary_confession = row.get(app_config['COLUMN_MAPPINGS']['voluntary_confession'], "")
                excel_trial_report = row.get(app_config['COLUMN_MAPPINGS']['trial_report'], "")
                voluntary_confession_highlight_indices = set()
                
                if pd.notna(excel_trial_report):
                    excel_trial_report = str(excel_trial_report).strip()
                    
                    validate_voluntary_confession_rules(
                        row, index, excel_case_code, excel_person_code, issues_list, voluntary_confession_highlight_indices,
                        excel_voluntary_confession, excel_trial_report, app_config
                    )
                
                # 执行党纪处分验证规则
                excel_disciplinary_sanction = str(row.get(app_config['COLUMN_MAPPINGS']['disciplinary_sanction'], '')).strip()
                disciplinary_sanction_mismatch_indices = set()
                validate_disciplinary_sanction_rules(
                    row, index, excel_case_code, excel_person_code, issues_list, disciplinary_sanction_mismatch_indices,
                    excel_disciplinary_sanction, decision_text_raw, app_config
                )
                
                # 执行是否属于本应撤销党内职务验证规则
                excel_no_party_position_warning = str(row.get(app_config['COLUMN_MAPPINGS']['no_party_position_warning'], '')).strip()
                no_party_position_warning_mismatch_indices = set()
                validate_no_party_position_warning_rules(
                    row, index, excel_case_code, excel_person_code, issues_list, no_party_position_warning_mismatch_indices,
                    excel_no_party_position_warning, decision_text_raw, app_config
                )
                
                # 执行政务处分验证规则
                excel_administrative_sanction = str(row.get(app_config['COLUMN_MAPPINGS']['administrative_sanction'], '')).strip()
                administrative_sanction_mismatch_indices = set()
                validate_administrative_sanction_rules(
                    row, index, excel_case_code, excel_person_code, issues_list, administrative_sanction_mismatch_indices,
                    excel_administrative_sanction, decision_text_raw, app_config
                )
                
                # 执行收缴金额验证规则
                excel_confiscation_amount = str(row.get(app_config['COLUMN_MAPPINGS']['confiscation_amount'], '')).strip()
                confiscation_amount_indices = set()
                validate_confiscation_amount_rules(
                    row, index, excel_case_code, excel_person_code, issues_list, confiscation_amount_indices,
                    excel_confiscation_amount, excel_trial_report, app_config
                )
                
                # 执行没收金额验证规则
                excel_confiscation_of_property_amount = str(row.get(app_config['COLUMN_MAPPINGS']['confiscation_of_property_amount'], '')).strip()
                confiscation_of_property_amount_indices = set()
                validate_confiscation_of_property_amount_rules(
                    row, index, excel_case_code, excel_person_code, issues_list, confiscation_of_property_amount_indices,
                    excel_confiscation_of_property_amount, excel_trial_report, app_config
                )
                
                # 执行责令退赔金额验证规则
                excel_compensation_amount = str(row.get(app_config['COLUMN_MAPPINGS']['compensation_amount'], '')).strip()
                compensation_amount_highlight_indices = set()
                validate_compensation_amount_rules(
                    row, index, excel_case_code, excel_person_code, issues_list, compensation_amount_highlight_indices,
                    excel_compensation_amount, excel_trial_report, app_config
                )
                
                # 执行追缴失职渎职滥用职权造成的损失金额验证规则
                excel_recovery_amount = row.get(app_config['COLUMN_MAPPINGS']['recovery_amount'])
                recovery_amount_highlight_indices = set()
                validate_recovery_amount_rules(
                    row, index, excel_case_code, excel_person_code, issues_list, recovery_amount_highlight_indices,
                    excel_recovery_amount, app_config
                )
                
                # 执行登记上交金额验证规则
                registered_handover_amount_indices = set()
                validate_registered_handover_amount_single_row(
                    row, index, excel_case_code, excel_person_code, issues_list, registered_handover_amount_indices, app_config
                )
                
                # 执行审理受理时间验证规则
                excel_trial_acceptance_time = row.get(app_config['COLUMN_MAPPINGS']['trial_acceptance_time'])
                trial_acceptance_time_mismatch_indices = set()
                validate_trial_acceptance_time_rules(
                    row, index, excel_case_code, excel_person_code, issues_list, trial_acceptance_time_mismatch_indices,
                    excel_trial_acceptance_time, excel_trial_report, app_config
                )
                
                # 执行审理机关验证规则
                excel_trial_authority = str(row.get(app_config['COLUMN_MAPPINGS']['trial_authority'], '')).strip()
                excel_reporting_agency = str(row.get(app_config['COLUMN_MAPPINGS']['reporting_agency'], '')).strip()
                trial_authority_mismatch_indices = set()
                sl_authority_agency_mappings = get_authority_agency_dict('SL')
                validate_trial_authority_rules(
                    row, index, excel_case_code, excel_person_code, issues_list, trial_authority_mismatch_indices,
                    excel_trial_authority, excel_reporting_agency, sl_authority_agency_mappings, app_config
                )
                
                # 执行审结时间验证规则
                excel_trial_closing_time = row.get(app_config['COLUMN_MAPPINGS']['trial_closing_time'])
                trial_closing_time_mismatch_indices = set()
                validate_trial_closing_time_rules(
                    row, index, excel_case_code, excel_person_code, issues_list, trial_closing_time_mismatch_indices,
                    excel_trial_closing_time, excel_trial_report, app_config
                )
                
                # 执行处分决定验证规则
                disciplinary_decision_mismatch_indices = set()
                validate_disciplinary_decision_rules(
                    row, index, excel_case_code, excel_person_code, issues_list, disciplinary_decision_mismatch_indices,
                    excel_disciplinary_decision, app_config
                )
                
                # 执行审理报告验证规则
                trial_report_mismatch_indices = set()
                validate_trial_report_rules(
                    row, index, excel_case_code, excel_person_code, issues_list, trial_report_mismatch_indices,
                    excel_trial_report, app_config
                )
                
            except Exception as e:
                logger.error(f"处理第 {index + 2} 行时发生错误: {str(e)}")
                continue
        
        # 生成立案编号表文件
        case_num_filename = f"立案编号表_{datetime.now().strftime('%Y%m%d')}.xlsx"
        case_num_path = os.path.join(case_dir, case_num_filename)
        
        # 准备数据
        data = []
        for i, issue_item in enumerate(issues_list):
            case_code = issue_item.get('案件编码', '')
            person_code = issue_item.get('涉案人员编码', '')
            row_number = issue_item.get('行号', '')
            compare_field = issue_item.get('比对字段', '')
            compared_field = issue_item.get('被比对字段', '')
            issue_description = issue_item.get('问题描述', '')
            column_name = issue_item.get('列名', '')
            
            data.append({
                '序号': i + 1,
                '案件编码': case_code,
                '涉案人员编码': person_code,
                '行号': row_number,
                '比对字段': compare_field,
                '被比对字段': compared_field,
                '问题': issue_description
            })
        
        # 如果没有发现问题，创建一个提示行
        if not data:
            data.append({
                '序号': 1,
                '案件编码': '',
                '涉案人员编码': '',
                '行号': '',
                '比对字段': '',
                '被比对字段': '',
                '问题': '未发现被调查人相关问题'
            })
        
        # 创建DataFrame
        issues_df = pd.DataFrame(data)
        
        # 写入Excel文件
        with pd.ExcelWriter(case_num_path, engine='xlsxwriter') as writer:
            issues_df.to_excel(writer, sheet_name='被调查人问题列表', index=False)
            workbook = writer.book
            worksheet = writer.sheets['被调查人问题列表']
            
            # 定义格式
            left_align_text_format = workbook.add_format({
                'align': 'left',
                'valign': 'vcenter',
                'num_format': '@'  # 强制文本格式
            })
            
            # 设置列格式
            columns_to_format = ['序号', '案件编码', '涉案人员编码', '行号', '比对字段', '被比对字段', '问题']
            for col_name in columns_to_format:
                if col_name in issues_df.columns:
                    col_idx = issues_df.columns.get_loc(col_name)
                    worksheet.set_column(col_idx, col_idx, None, left_align_text_format)
            
            # 自动调整列宽
            for i, col in enumerate(issues_df.columns):
                max_len = max(issues_df[col].astype(str).map(len).max(), len(col)) + 2
                worksheet.set_column(i, i, max_len)
        
        logger.info(f"成功生成被调查人立案编号表: {case_num_path}")
        return case_num_path
        
    except Exception as e:
        logger.error(f"生成被调查人立案编号表失败: {str(e)}", exc_info=True)
        return None