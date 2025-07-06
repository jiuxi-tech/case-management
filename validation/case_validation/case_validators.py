import logging
import pandas as pd
from datetime import datetime
from config import Config # 导入Config
import re
from db_utils import get_authority_agency_dict

# 从 case_validation_helpers 导入核心验证函数
from .case_validation_helpers import (
    validate_brief_case_details_rules
)

# 从 case_validation_extended 导入扩展验证函数
from .case_validation_extended import (
    validate_ethnicity_rules,
    validate_party_member_rules,
    validate_party_joining_date_rules
)

# 从 case_validation_additional 导入额外验证函数
from .case_validation_additional import (
    validate_education_rules
)

# 从 case_validation_additional 导入其他验证函数
from .case_validation_additional import (
    validate_name_rules,
    validate_gender_rules,
    validate_age_rules,
    validate_birth_date_rules,
    validate_case_report_keywords_rules,
    validate_voluntary_confession_rules,
    validate_no_party_position_warning_rules
)

# 导入立案时间规则
from .case_timestamp_rules import (
    validate_filing_time,
    validate_confiscation_amount,
    validate_confiscation_of_property_amount,
    validate_registered_handover_amount
)
# 导入处分和金额相关规则
from .case_disposal_amount_rules import validate_disposal_and_amount_rules

# 【党纪处分功能新增】: 导入党纪处分验证函数
from .case_validation_sanctions import validate_disciplinary_sanction, validate_administrative_sanction # 确保导入了 validate_administrative_sanction

# 导入新拆分的文件中的验证函数
from .case_document_validators import ( 
    validate_trial_acceptance_time_vs_report,
    validate_trial_closing_time_vs_report,
    validate_trial_authority_vs_reporting_agency,
    validate_disposal_decision_keywords,
    validate_trial_report_keywords,
    highlight_recovery_amount
)

logger = logging.getLogger(__name__)

def validate_case_relationships(df, app_config, issues_list):
    """
    验证立案登记表Excel中各字段之间的关系和数据有效性。

    参数:
        df (pd.DataFrame): 包含立案登记表数据的DataFrame。
        app_config (dict): Flask 应用的配置字典，包含Config类中的配置。
        issues_list (list): 用于收集所有发现问题的列表，每个问题是一个字典或元组。

    返回:
        tuple: 包含多个不一致索引集合和更新后的 issues_list。
               返回顺序与 file_processor.py 中接收的顺序一致。
    """
    mismatch_indices = set()
    gender_mismatch_indices = set()
    age_mismatch_indices = set()
    brief_case_details_mismatch_indices = set()
    birth_date_mismatch_indices = set()
    education_mismatch_indices = set()
    ethnicity_mismatch_indices = set()
    party_member_mismatch_indices = set()
    party_joining_date_mismatch_indices = set()
    filing_time_mismatch_indices = set()
    disciplinary_committee_filing_time_mismatch_indices = set()
    disciplinary_committee_filing_authority_mismatch_indices = set()
    supervisory_committee_filing_time_mismatch_indices = set()
    supervisory_committee_filing_authority_mismatch_indices = set()
    case_report_keyword_mismatch_indices = set()
    disposal_spirit_mismatch_indices = set()
    voluntary_confession_highlight_indices = set()
    closing_time_mismatch_indices = set()
    no_party_position_warning_mismatch_indices = set()
    recovery_amount_highlight_indices = set()
    trial_acceptance_time_mismatch_indices = set()
    trial_closing_time_mismatch_indices = set()
    trial_authority_agency_mismatch_indices = set()
    disposal_decision_keyword_mismatch_indices = set()
    
    # --- START OF NEW RULE ADDITION ---
    # New sets for trial report validation
    trial_report_non_representative_mismatch_indices = set()
    trial_report_detention_mismatch_indices = set()
    confiscation_amount_indices = set()
    confiscation_of_property_amount_indices = set()
    compensation_amount_highlight_indices = set()
    registered_handover_amount_indices = set()
    # 【党纪处分功能新增】: 初始化党纪处分相关的索引集合
    disciplinary_sanction_mismatch_indices = set()
    administrative_sanction_mismatch_indices = set() # 初始化政务处分不匹配索引
    # --- END OF NEW RULE ADDITION ---

    # issues_list 不再在这里初始化，而是作为参数传入并直接修改

    required_headers = [
        "被调查人", "性别", "年龄", "出生年月", "学历", "民族",
        "是否中共党员", "入党时间", "立案报告", "处分决定",
        "审查调查报告", "审理报告", "简要案情",
        "案件编码", "涉案人员编码",
        "立案时间", "立案决定书",
        "纪委立案时间", "纪委立案机关", "监委立案时间", "监委立案机关", "填报单位名称",
        "是否违反中央八项规定精神",
        "是否主动交代问题",
        "结案时间",
        "是否属于本应撤销党内职务，但本人没有党内职务而给予严重警告处分",
        "追缴失职渎职滥用职权造成的损失金额",
        "审理受理时间",
        "审结时间",
        "审理机关",
        "收缴金额（万元）",
        "没收金额",
        "责令退赔金额",
        "登记上交金额",
        # 【党纪处分功能新增】: 确保包含党纪处分列
        "党纪处分",
        # 【新增】政务处分也应该在这里检查
        app_config['COLUMN_MAPPINGS'].get("administrative_sanction") # 从 app_config 获取
    ]
    
    # 过滤掉 None 值，因为 get() 可能返回 None
    required_headers = [h for h in required_headers if h is not None]

    if not all(header in df.columns for header in required_headers):
        missing_headers = [header for header in required_headers if header not in df.columns]
        msg = f"缺少必要的表头: {missing_headers}"
        logger.error(msg)
        # Ensure all new sets are returned even if headers are missing
        return (mismatch_indices, gender_mismatch_indices, age_mismatch_indices, brief_case_details_mismatch_indices, issues_list,
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
                        administrative_sanction_mismatch_indices) # 确保这里也返回 administrative_sanction_mismatch_indices

    current_year = datetime.now().year

    # 使用 app_config 获取配置，而不是直接使用 Config
    # 原始代码中 case_report_keywords_to_check 是硬编码的，但其内容与 Config.DISPOSAL_DECISION_KEYWORDS 相同。
    # 因此，这里改为从 app_config 获取，以保持一致性。
    case_report_keywords_to_check = app_config['DISPOSAL_DECISION_KEYWORDS']
    
    # 从数据库获取机关单位字典数据
    authority_agency_db_data = get_authority_agency_dict()
    # 将数据库查询结果转换为更易于查找的列表，只包含SL类别的
    sl_authority_agency_mappings = []
    for record_raw in authority_agency_db_data:
        record = record_raw # record_raw 已经是字典
        if record['category'] == 'SL':
            sl_authority_agency_mappings.append({
                'authority': record['authority'],
                'agency': record['agency']
            })

    # 遍历DataFrame的每一行
    for index, row in df.iterrows():
        logger.debug(f"Processing row {index + 1}")

        investigated_person = str(row.get(app_config['COLUMN_MAPPINGS']["investigated_person"], "")).strip()
        if not investigated_person:
            logger.info(f"行 {index + 1} - '{app_config['COLUMN_MAPPINGS']['investigated_person']}' 字段为空，跳过此行验证。")
            continue

        excel_voluntary_confession = str(row.get(app_config['COLUMN_MAPPINGS']["voluntary_confession"], "")).strip()
        excel_gender = str(row.get(app_config['COLUMN_MAPPINGS']["gender"], "")).strip()
        
        excel_age = None
        if pd.notna(row.get(app_config['COLUMN_MAPPINGS']["age"])):
            try:
                excel_age = int(row.get(app_config['COLUMN_MAPPINGS']["age"]))
            except ValueError:
                logger.warning(f"行 {index + 1} - Excel '{app_config['COLUMN_MAPPINGS']['age']}' 字段 '{row.get(app_config['COLUMN_MAPPINGS']['age'])}' 不是有效数字。")
                age_mismatch_indices.add(index)
                issues_list.append((index, row.get(app_config['COLUMN_MAPPINGS']["case_code"], ""), row.get(app_config['COLUMN_MAPPINGS']["person_code"], ""), "N2年龄字段格式不正确", "高")) # 增加风险等级

        excel_brief_case_details = str(row.get(app_config['COLUMN_MAPPINGS']["brief_case_details"], "")).strip()
        excel_birth_date = str(row.get(app_config['COLUMN_MAPPINGS']["birth_date"], "")).strip()
        excel_education = str(row.get(app_config['COLUMN_MAPPINGS']["education"], "")).strip()
        excel_ethnicity = str(row.get(app_config['COLUMN_MAPPINGS']["ethnicity"], "")).strip()
        excel_party_member = str(row.get(app_config['COLUMN_MAPPINGS']["party_member"], "")).strip()
        excel_party_joining_date = str(row.get(app_config['COLUMN_MAPPINGS']["party_joining_date"], "")).strip()
        excel_no_party_position_warning = str(row.get(app_config['COLUMN_MAPPINGS']["no_party_position_warning"], "")).strip()
        
        # 将原始数据传递给新的验证函数，而不是在这里处理逻辑
        excel_trial_acceptance_time = row.get(app_config['COLUMN_MAPPINGS']["trial_acceptance_time"])
        excel_trial_closing_time = row.get(app_config['COLUMN_MAPPINGS']["trial_closing_time"])
        excel_trial_authority = str(row.get(app_config['COLUMN_MAPPINGS']["trial_authority"], "")).strip()
        excel_reporting_agency = str(row.get(app_config['COLUMN_MAPPINGS']["reporting_agency"], "")).strip()

        excel_case_code = str(row.get(app_config['COLUMN_MAPPINGS']["case_code"], "")).strip()
        excel_person_code = str(row.get(app_config['COLUMN_MAPPINGS']["person_code"], "")).strip()

        report_text_raw = row.get(app_config['COLUMN_MAPPINGS']["case_report"], "") if pd.notna(row.get(app_config['COLUMN_MAPPINGS']["case_report"])) else ''
        decision_text_raw = row.get(app_config['COLUMN_MAPPINGS']["disciplinary_decision"], "") if pd.notna(row.get(app_config['COLUMN_MAPPINGS']["disciplinary_decision"])) else ''
        investigation_text_raw = row.get(app_config['COLUMN_MAPPINGS']["investigation_report"], "") if pd.notna(row.get(app_config['COLUMN_MAPPINGS']["investigation_report"])) else ''
        trial_text_raw = row.get(app_config['COLUMN_MAPPINGS']["trial_report"], "") if pd.notna(row.get(app_config['COLUMN_MAPPINGS']["trial_report"])) else ''
        filing_decision_doc_raw = row.get(app_config['COLUMN_MAPPINGS']["filing_decision_doc"], "") if pd.notna(row.get(app_config['COLUMN_MAPPINGS']["filing_decision_doc"])) else ''
        
        # --- 调用辅助函数进行验证 ---
        # 传递 app_config 给可能需要它的辅助函数
        validate_gender_rules(row, index, excel_case_code, excel_person_code, issues_list, gender_mismatch_indices,
                              excel_gender, report_text_raw, decision_text_raw, investigation_text_raw, trial_text_raw, app_config)

        validate_age_rules(row, index, excel_case_code, excel_person_code, issues_list, age_mismatch_indices,
                           excel_age, current_year, report_text_raw, decision_text_raw, investigation_text_raw, trial_text_raw, app_config)

        validate_brief_case_details_rules(row, index, excel_case_code, excel_person_code, issues_list, brief_case_details_mismatch_indices,
                                          excel_brief_case_details, investigated_person, report_text_raw, decision_text_raw, app_config)

        validate_birth_date_rules(row, index, excel_case_code, excel_person_code, issues_list, birth_date_mismatch_indices,
                                  excel_birth_date, report_text_raw, decision_text_raw, investigation_text_raw, trial_text_raw, app_config)

        validate_education_rules(row, index, excel_case_code, excel_person_code, issues_list, education_mismatch_indices,
                                 excel_education, report_text_raw, decision_text_raw, investigation_text_raw, trial_text_raw, app_config)

        validate_ethnicity_rules(row, index, excel_case_code, excel_person_code, issues_list, ethnicity_mismatch_indices,
                                 excel_ethnicity, report_text_raw, decision_text_raw, investigation_text_raw, trial_text_raw, app_config)

        validate_party_member_rules(row, index, excel_case_code, excel_person_code, issues_list, party_member_mismatch_indices,
                                    excel_party_member, report_text_raw, decision_text_raw, app_config)

        validate_party_joining_date_rules(row, index, excel_case_code, excel_person_code, issues_list, party_joining_date_mismatch_indices,
                                          excel_party_member, excel_party_joining_date, report_text_raw, app_config)

        validate_name_rules(row, index, excel_case_code, excel_person_code, issues_list, mismatch_indices,
                            investigated_person, report_text_raw, decision_text_raw, investigation_text_raw, trial_text_raw, app_config)

        validate_case_report_keywords_rules(row, index, excel_case_code, excel_person_code, issues_list, case_report_keyword_mismatch_indices,
                                            case_report_keywords_to_check, report_text_raw, decision_text_raw, investigation_text_raw, trial_text_raw, app_config)
        
        validate_voluntary_confession_rules(row, index, excel_case_code, excel_person_code, issues_list, voluntary_confession_highlight_indices,
                                            excel_voluntary_confession, trial_text_raw, app_config)

        validate_no_party_position_warning_rules(row, index, excel_case_code, excel_person_code, issues_list, no_party_position_warning_mismatch_indices,
                                                 excel_no_party_position_warning, decision_text_raw, app_config)

        # 调用新拆分的函数来处理这些特定验证
        highlight_recovery_amount(row, index, excel_case_code, excel_person_code, issues_list, recovery_amount_highlight_indices, app_config)
        validate_trial_acceptance_time_vs_report(row, index, excel_case_code, excel_person_code, issues_list, trial_acceptance_time_mismatch_indices, app_config)
        validate_trial_closing_time_vs_report(row, index, excel_case_code, excel_person_code, issues_list, trial_closing_time_mismatch_indices, app_config)
        validate_trial_authority_vs_reporting_agency(row, index, excel_case_code, excel_person_code, issues_list, trial_authority_agency_mismatch_indices, sl_authority_agency_mappings, app_config)
        validate_disposal_decision_keywords(row, index, excel_case_code, excel_person_code, issues_list, disposal_decision_keyword_mismatch_indices, app_config)
        validate_trial_report_keywords(row, index, excel_case_code, excel_person_code, issues_list, 
                                       trial_report_non_representative_mismatch_indices, 
                                       trial_report_detention_mismatch_indices, 
                                       compensation_amount_highlight_indices, app_config)

    # 调用立案时间规则验证函数
    validate_filing_time(df, issues_list, filing_time_mismatch_indices,
                         disciplinary_committee_filing_time_mismatch_indices,
                         disciplinary_committee_filing_authority_mismatch_indices,
                         supervisory_committee_filing_time_mismatch_indices,
                         supervisory_committee_filing_authority_mismatch_indices, app_config)

    # 调用处分和金额相关规则验证函数
    validate_disposal_and_amount_rules(df, issues_list, disposal_spirit_mismatch_indices, closing_time_mismatch_indices, app_config)

    # 调用没收金额验证函数
    validate_confiscation_of_property_amount(df, issues_list, confiscation_of_property_amount_indices, app_config)

    # 调用收缴金额验证函数
    validate_confiscation_amount(df, issues_list, confiscation_amount_indices, app_config)

    # 调用登记上交金额验证函数
    validate_registered_handover_amount(df, issues_list, registered_handover_amount_indices, app_config)

    # 【党纪处分功能新增】: 调用党纪处分相关规则验证函数
    # validate_disciplinary_sanction 应该返回不匹配的行索引，并直接修改 issues_list
    new_disciplinary_mismatches = validate_disciplinary_sanction(df, issues_list, app_config)
    disciplinary_sanction_mismatch_indices.update(new_disciplinary_mismatches)

    # 【新增】调用政务处分规则，并将其不匹配索引合并到总的索引和问题列表中
    new_administrative_mismatches = validate_administrative_sanction(df, issues_list, app_config)
    administrative_sanction_mismatch_indices.update(new_administrative_mismatches)

    # 返回所有可能的不一致索引集以及更新后的 issues_list
    return (mismatch_indices, gender_mismatch_indices, age_mismatch_indices, brief_case_details_mismatch_indices, issues_list,
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
            administrative_sanction_mismatch_indices) # 确保这里返回 administrative_sanction_mismatch_indices
