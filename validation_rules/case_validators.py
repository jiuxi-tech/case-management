import logging
import pandas as pd
from datetime import datetime
from config import Config
import re

# 从 case_validation_helpers 导入核心验证函数
from validation_rules.case_validation_helpers import (
    validate_gender_rules,
    validate_age_rules,
    validate_brief_case_details_rules
)

# 从 case_validation_extended 导入扩展验证函数
from validation_rules.case_validation_extended import (
    validate_birth_date_rules,
    validate_education_rules,
    validate_ethnicity_rules,
    validate_party_member_rules,
    validate_party_joining_date_rules
)

# 从 case_validation_additional 导入其他验证函数
from validation_rules.case_validation_additional import (
    validate_name_rules,
    validate_case_report_keywords_rules,
    validate_voluntary_confession_rules
)

# 导入立案时间规则
from validation_rules.case_timestamp_rules import validate_filing_time
# 导入处分和金额相关规则
from validation_rules.case_disposal_amount_rules import validate_disposal_and_amount_rules

logger = logging.getLogger(__name__)

def validate_case_relationships(df):
    """Validate relationships between fields in the case registration Excel."""
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

    issues_list = [] 
    
    required_headers = [
        "被调查人", "性别", "年龄", "出生年月", "学历", "民族", 
        "是否中共党员", "入党时间", "立案报告", "处分决定", 
        "审查调查报告", "审理报告", "简要案情",
        "案件编码", "涉案人员编码",
        "立案时间", "立案决定书",
        "纪委立案时间", "纪委立案机关", "监委立案时间", "监委立案机关", "填报单位名称",
        "是否违反中央八项规定精神",
        "是否主动交代问题",
        "结案时间"
    ]
    if not all(header in df.columns for header in required_headers):
        msg = f"缺少必要的表头: {required_headers}"
        logger.error(msg)
        print(msg)
        return mismatch_indices, gender_mismatch_indices, age_mismatch_indices, brief_case_details_mismatch_indices, issues_list, \
               birth_date_mismatch_indices, education_mismatch_indices, ethnicity_mismatch_indices, \
               party_member_mismatch_indices, party_joining_date_mismatch_indices, filing_time_mismatch_indices, \
               disciplinary_committee_filing_time_mismatch_indices, disciplinary_committee_filing_authority_mismatch_indices, \
               supervisory_committee_filing_time_mismatch_indices, supervisory_committee_filing_authority_mismatch_indices, \
               case_report_keyword_mismatch_indices, disposal_spirit_mismatch_indices, voluntary_confession_highlight_indices, \
               closing_time_mismatch_indices

    current_year = datetime.now().year

    case_report_keywords_to_check = ["人大代表", "政协委员", "党委委员", "中共党代表", "纪委委员"]

    for index, row in df.iterrows():
        logger.debug(f"Processing row {index + 1}")
        print(f"处理行 {index + 1}")

        investigated_person = str(row.get("被调查人", "")).strip() # 使用 .get()
        if not investigated_person:
            logger.info(f"行 {index + 1} - '被调查人' 字段为空，跳过此行验证。")
            print(f"行 {index + 1} - '被调查人' 字段为空，跳过此行验证。")
            continue

        # 将 excel_voluntary_confession 的定义提前，确保其在调用辅助函数前被定义
        excel_voluntary_confession = str(row.get("是否主动交代问题", "")).strip()
        
        excel_gender = str(row.get("性别", "")).strip() # 使用 .get()
        
        excel_age = None
        if pd.notna(row.get("年龄")): # 使用 .get()
            try:
                excel_age = int(row.get("年龄")) # 使用 .get()
            except ValueError:
                logger.warning(f"行 {index + 1} - Excel '年龄' 字段 '{row.get('年龄')}' 不是有效数字。")
                print(f"行 {index + 1} - Excel '年龄' 字段 '{row.get('年龄')}' 不是有效数字。")
                age_mismatch_indices.add(index)
                issues_list.append((index, row.get("案件编码", ""), row.get("涉案人员编码", ""), "N2年龄字段格式不正确"))

        excel_brief_case_details = str(row.get("简要案情", "")).strip() # 使用 .get()
        excel_birth_date = str(row.get("出生年月", "")).strip() # 使用 .get()
        excel_education = str(row.get("学历", "")).strip() # 使用 .get()
        excel_ethnicity = str(row.get("民族", "")).strip() # 使用 .get()
        excel_party_member = str(row.get("是否中共党员", "")).strip() # 使用 .get()
        excel_party_joining_date = str(row.get("入党时间", "")).strip() # 使用 .get()

        excel_case_code = str(row.get("案件编码", "")).strip() # 使用 .get()
        excel_person_code = str(row.get("涉案人员编码", "")).strip() # 使用 .get()

        report_text_raw = row.get("立案报告", "") if pd.notna(row.get("立案报告")) else '' # 使用 .get()
        decision_text_raw = row.get("处分决定", "") if pd.notna(row.get("处分决定")) else '' # 使用 .get()
        investigation_text_raw = row.get("审查调查报告", "") if pd.notna(row.get("审查调查报告")) else '' # 使用 .get()
        trial_text_raw = row.get("审理报告", "") if pd.notna(row.get("审理报告")) else '' # 使用 .get()
        filing_decision_doc_raw = row.get("立案决定书", "") if pd.notna(row.get("立案决定书")) else '' # 使用 .get()

        # --- 调用辅助函数进行验证 ---
        validate_gender_rules(row, index, excel_case_code, excel_person_code, issues_list, gender_mismatch_indices,
                             excel_gender, report_text_raw, decision_text_raw, investigation_text_raw, trial_text_raw)

        validate_age_rules(row, index, excel_case_code, excel_person_code, issues_list, age_mismatch_indices,
                          excel_age, current_year, report_text_raw, decision_text_raw, investigation_text_raw, trial_text_raw)

        validate_brief_case_details_rules(row, index, excel_case_code, excel_person_code, issues_list, brief_case_details_mismatch_indices,
                                         excel_brief_case_details, investigated_person, report_text_raw, decision_text_raw)

        validate_birth_date_rules(row, index, excel_case_code, excel_person_code, issues_list, birth_date_mismatch_indices,
                                 excel_birth_date, report_text_raw, decision_text_raw, investigation_text_raw, trial_text_raw)

        validate_education_rules(row, index, excel_case_code, excel_person_code, issues_list, education_mismatch_indices,
                                excel_education, report_text_raw)

        validate_ethnicity_rules(row, index, excel_case_code, excel_person_code, issues_list, ethnicity_mismatch_indices,
                                excel_ethnicity, report_text_raw, decision_text_raw, investigation_text_raw, trial_text_raw)

        validate_party_member_rules(row, index, excel_case_code, excel_person_code, issues_list, party_member_mismatch_indices,
                                   excel_party_member, report_text_raw, decision_text_raw)

        validate_party_joining_date_rules(row, index, excel_case_code, excel_person_code, issues_list, party_joining_date_mismatch_indices,
                                        excel_party_member, excel_party_joining_date, report_text_raw)

        validate_name_rules(row, index, excel_case_code, excel_person_code, issues_list, mismatch_indices,
                           investigated_person, report_text_raw, decision_text_raw, investigation_text_raw, trial_text_raw)

        validate_case_report_keywords_rules(row, index, excel_case_code, excel_person_code, issues_list, case_report_keyword_mismatch_indices,
                                           case_report_keywords_to_check, report_text_raw, decision_text_raw, investigation_text_raw, trial_text_raw)
        
        validate_voluntary_confession_rules(row, index, excel_case_code, excel_person_code, issues_list, voluntary_confession_highlight_indices,
                                           excel_voluntary_confession, trial_text_raw)

    # 调用立案时间规则验证函数
    # 修正：validate_filing_time 函数目前接受 7 个参数，因此移除末尾的两个额外参数。
    validate_filing_time(df, issues_list, filing_time_mismatch_indices,
                        disciplinary_committee_filing_time_mismatch_indices,
                        disciplinary_committee_filing_authority_mismatch_indices,
                        supervisory_committee_filing_time_mismatch_indices,
                        supervisory_committee_filing_authority_mismatch_indices)

    # 调用处分和金额相关规则验证函数，现在也传递 closing_time_mismatch_indices
    validate_disposal_and_amount_rules(df, issues_list, disposal_spirit_mismatch_indices, closing_time_mismatch_indices)

    # 返回所有可能的不一致索引集以及更新后的 issues_list
    return mismatch_indices, gender_mismatch_indices, age_mismatch_indices, brief_case_details_mismatch_indices, issues_list, \
           birth_date_mismatch_indices, education_mismatch_indices, ethnicity_mismatch_indices, \
           party_member_mismatch_indices, party_joining_date_mismatch_indices, filing_time_mismatch_indices, \
           disciplinary_committee_filing_time_mismatch_indices, disciplinary_committee_filing_authority_mismatch_indices, \
           supervisory_committee_filing_time_mismatch_indices, supervisory_committee_filing_authority_mismatch_indices, \
           case_report_keyword_mismatch_indices, disposal_spirit_mismatch_indices, voluntary_confession_highlight_indices, \
           closing_time_mismatch_indices