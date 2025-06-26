import logging
import pandas as pd
from datetime import datetime
from config import Config
import re # 导入 re 模块用于去除空白符

# 从新的拆分文件中导入所有提取器函数
from validation_rules.case_extractors_names import (
    extract_name_from_case_report,
    extract_name_from_decision,
    extract_name_from_trial_report
)
from validation_rules.case_extractors_gender import (
    extract_gender_from_case_report,
    extract_gender_from_decision_report,
    extract_gender_from_investigation_report,
    extract_gender_from_trial_report
)
from validation_rules.case_extractors_birth_info import (
    extract_birth_year_from_case_report,
    extract_birth_year_from_decision_report,
    extract_birth_year_from_investigation_report,
    extract_birth_year_from_trial_report,
    extract_birth_date_from_case_report,
    extract_birth_date_from_decision_report,
    extract_birth_date_from_investigation_report,
    extract_birth_date_from_trial_report
)
from validation_rules.case_extractors_demographics import (
    extract_education_from_case_report,
    extract_ethnicity_from_case_report,
    extract_ethnicity_from_decision_report,
    extract_ethnicity_from_investigation_report,
    extract_ethnicity_from_trial_report,
    extract_suspected_violation_from_case_report, # 新增导入
    extract_suspected_violation_from_decision # 新增导入
)
from validation_rules.case_extractors_party_info import (
    extract_party_member_from_case_report,
    extract_party_member_from_decision_report,
    extract_party_joining_date_from_case_report
)
# 新增导入立案时间规则
from validation_rules.case_timestamp_rules import validate_filing_time


logger = logging.getLogger(__name__)

def validate_case_relationships(df):
    """Validate relationships between fields in the case registration Excel."""
    mismatch_indices = set()
    gender_mismatch_indices = set()
    age_mismatch_indices = set()
    brief_case_details_mismatch_indices = set() # 新增简要案情不一致索引集合
    birth_date_mismatch_indices = set()
    education_mismatch_indices = set()
    ethnicity_mismatch_indices = set()
    party_member_mismatch_indices = set()
    party_joining_date_mismatch_indices = set()
    filing_time_mismatch_indices = set() # 新增立案时间不一致索引集合
    disciplinary_committee_filing_time_mismatch_indices = set() # 新增纪委立案时间不一致索引集合
    disciplinary_committee_filing_authority_mismatch_indices = set() # 新增纪委立案机关不一致索引集合
    supervisory_committee_filing_time_mismatch_indices = set() # 新增监委立案时间不一致索引集合
    supervisory_committee_filing_authority_mismatch_indices = set() # 新增监委立案机关不一致索引集合

    # issues_list 现在将包含 (index, case_code, person_code, issue_description)
    issues_list = [] 
    
    # 定义立案登记表所需的所有表头，包括新增的编码
    required_headers = [
        "被调查人", "性别", "年龄", "出生年月", "学历", "民族", 
        "是否中共党员", "入党时间", "立案报告", "处分决定", 
        "审查调查报告", "审理报告", "简要案情", # 现有字段
        "案件编码", "涉案人员编码", # 新增字段
        "立案时间", "立案决定书", # 现有立案时间相关字段
        "纪委立案时间", "纪委立案机关", "监委立案时间", "监委立案机关", "填报单位名称" # 新增立案时间/机关相关字段
    ]
    if not all(header in df.columns for header in required_headers):
        logger.error(f"Missing required headers for case registration: {required_headers}")
        print(f"缺少必要的表头: {required_headers}")
        # 返回所有可能的不一致索引集，确保与调用处的解包数量一致
        return mismatch_indices, gender_mismatch_indices, age_mismatch_indices, brief_case_details_mismatch_indices, issues_list, \
               birth_date_mismatch_indices, education_mismatch_indices, ethnicity_mismatch_indices, \
               party_member_mismatch_indices, party_joining_date_mismatch_indices, filing_time_mismatch_indices, \
               disciplinary_committee_filing_time_mismatch_indices, disciplinary_committee_filing_authority_mismatch_indices, \
               supervisory_committee_filing_time_mismatch_indices, supervisory_committee_filing_authority_mismatch_indices # 添加新的返回值

    current_year = datetime.now().year

    for index, row in df.iterrows():
        logger.debug(f"Processing row {index + 1}")
        print(f"处理行 {index + 1}")

        investigated_person = str(row["被调查人"]).strip() if pd.notna(row["被调查人"]) else ''
        if not investigated_person:
            logger.info(f"行 {index + 1} - '被调查人' 字段为空，跳过此行验证。")
            print(f"行 {index + 1} - '被调查人' 字段为空，跳过此行验证。")
            continue

        excel_gender = str(row["性别"]).strip() if pd.notna(row["性别"]) else ''
        
        excel_age = None
        if pd.notna(row["年龄"]):
            try:
                excel_age = int(row["年龄"])
            except ValueError:
                logger.warning(f"行 {index + 1} - Excel '年龄' 字段 '{row['年龄']}' 不是有效数字。")
                print(f"行 {index + 1} - Excel '年龄' 字段 '{row['年龄']}' 不是有效数字。")
                age_mismatch_indices.add(index)
                # 修改 issues_list 的添加方式，包含编码字段
                issues_list.append((index, row.get("案件编码", ""), row.get("涉案人员编码", ""), "N2年龄字段格式不正确"))

        excel_brief_case_details = str(row["简要案情"]).strip() if pd.notna(row["简要案情"]) else '' # 获取简要案情
        excel_birth_date = str(row["出生年月"]).strip() if pd.notna(row["出生年月"]) else ''
        excel_education = str(row["学历"]).strip() if pd.notna(row["学历"]) else '' 
        excel_ethnicity = str(row["民族"]).strip() if pd.notna(row["民族"]) else ''
        excel_party_member = str(row["是否中共党员"]).strip() if pd.notna(row["是否中共党员"]) else ''
        excel_party_joining_date = str(row["入党时间"]).strip() if pd.notna(row["入党时间"]) else ''

        # 新增：获取案件编码和涉案人员编码
        excel_case_code = str(row["案件编码"]).strip() if pd.notna(row["案件编码"]) else ''
        excel_person_code = str(row["涉案人员编码"]).strip() if pd.notna(row["涉案人员编码"]) else ''

        report_text_raw = row["立案报告"] if pd.notna(row["立案报告"]) else ''
        decision_text_raw = row["处分决定"] if pd.notna(row["处分决定"]) else ''
        investigation_text_raw = row["审查调查报告"] if pd.notna(row["审查调查报告"]) else ''
        trial_text_raw = row["审理报告"] if pd.notna(row["审理报告"]) else ''
        # 获取立案决定书内容
        filing_decision_doc_raw = row["立案决定书"] if pd.notna(row["立案决定书"]) else ''


        # --- Gender matching rules ---
        extracted_gender_from_report = extract_gender_from_case_report(report_text_raw)
        if extracted_gender_from_report is None or (excel_gender and excel_gender != extracted_gender_from_report):
            gender_mismatch_indices.add(index)
            issues_list.append((index, excel_case_code, excel_person_code, "M2性别与BF2立案报告不一致"))
            logger.info(f"行 {index + 1} - 性别不匹配: Excel性别 ('{excel_gender}') vs 立案报告提取性别 ('{extracted_gender_from_report}')")
            print(f"行 {index + 1} - 性别不匹配: Excel性别 ('{excel_gender}') vs 立案报告提取性别 ('{extracted_gender_from_report}')")

        extracted_gender_from_decision = extract_gender_from_decision_report(decision_text_raw)
        if extracted_gender_from_decision is None or (excel_gender and excel_gender != extracted_gender_from_decision):
            gender_mismatch_indices.add(index) 
            issues_list.append((index, excel_case_code, excel_person_code, "M2性别与CU2处分决定不一致"))
            logger.info(f"行 {index + 1} - 性别不匹配: Excel性别 ('{excel_gender}') vs 处分决定提取性别 ('{extracted_gender_from_decision}')")
            print(f"行 {index + 1} - 性别不匹配: Excel性别 ('{excel_gender}') vs 处分决定提取性别 ('{extracted_gender_from_decision}')")

        extracted_gender_from_investigation = extract_gender_from_investigation_report(investigation_text_raw)
        if extracted_gender_from_investigation is None or (excel_gender and excel_gender != extracted_gender_from_investigation):
            gender_mismatch_indices.add(index)
            issues_list.append((index, excel_case_code, excel_person_code, "M2性别与CX2审查调查报告不一致"))
            logger.info(f"行 {index + 1} - 性别不匹配: Excel性别 ('{excel_gender}') vs 审查调查报告提取性别 ('{extracted_gender_from_investigation}')")
            print(f"行 {index + 1} - 性别不匹配: Excel性别 ('{excel_gender}') vs 审查调查报告提取性别 ('{extracted_gender_from_investigation}')")

        extracted_gender_from_trial = extract_gender_from_trial_report(trial_text_raw)
        if extracted_gender_from_trial is None or (excel_gender and excel_gender != extracted_gender_from_trial):
            gender_mismatch_indices.add(index)
            issues_list.append((index, excel_case_code, excel_person_code, "M2性别与CY2审理报告不一致"))
            logger.info(f"行 {index + 1} - 性别不匹配: Excel性别 ('{excel_gender}') vs 审理报告提取性别 ('{extracted_gender_from_trial}')")
            print(f"行 {index + 1} - 性别不匹配: Excel性别 ('{excel_gender}') vs 审理报告提取性别 ('{extracted_gender_from_trial}')")

        # --- Age matching rules ---
        extracted_birth_year_from_report = extract_birth_year_from_case_report(report_text_raw)
        calculated_age_from_report = None
        if extracted_birth_year_from_report is not None:
            calculated_age_from_report = current_year - extracted_birth_year_from_report
        if (calculated_age_from_report is None) or \
           (excel_age is not None and calculated_age_from_report is not None and excel_age != calculated_age_from_report):
            age_mismatch_indices.add(index)
            issues_list.append((index, excel_case_code, excel_person_code, "N2年龄与BF2立案报告不一致"))
            logger.info(f"行 {index + 1} - 年龄不匹配: Excel年龄 ('{excel_age}') vs 立案报告计算年龄 ('{calculated_age_from_report}')")
            print(f"行 {index + 1} - 年龄不匹配: Excel年龄 ('{excel_age}') vs 立案报告计算年龄 ('{calculated_age_from_report}')")

        extracted_birth_year_from_decision = extract_birth_year_from_decision_report(decision_text_raw)
        calculated_age_from_decision = None
        if extracted_birth_year_from_decision is not None:
            calculated_age_from_decision = current_year - extracted_birth_year_from_decision
        if (calculated_age_from_decision is None) or \
           (excel_age is not None and calculated_age_from_decision is not None and excel_age != calculated_age_from_decision):
            age_mismatch_indices.add(index)
            issues_list.append((index, excel_case_code, excel_person_code, "N2年龄与CU2处分决定不一致"))
            logger.info(f"行 {index + 1} - 年龄不匹配: Excel年龄 ('{excel_age}') vs 处分决定计算年龄 ('{calculated_age_from_decision}')")
            print(f"行 {index + 1} - 年龄不匹配: Excel年龄 ('{excel_age}') vs 处分决定计算年龄 ('{calculated_age_from_decision}')")

        extracted_birth_year_from_investigation = extract_birth_year_from_investigation_report(investigation_text_raw)
        calculated_age_from_investigation = None
        if extracted_birth_year_from_investigation is not None:
            calculated_age_from_investigation = current_year - extracted_birth_year_from_investigation
        if (calculated_age_from_investigation is None) or \
           (excel_age is not None and calculated_age_from_investigation is not None and excel_age != calculated_age_from_investigation):
            age_mismatch_indices.add(index)
            issues_list.append((index, excel_case_code, excel_person_code, "N2年龄与CX2审查调查报告不一致"))
            logger.info(f"行 {index + 1} - 年龄不匹配: Excel年龄 ('{excel_age}') vs 审查调查报告计算年龄 ('{calculated_age_from_investigation}')")
            print(f"行 {index + 1} - 年龄不匹配: Excel年龄 ('{excel_age}') vs 审查调查报告计算年龄 ('{calculated_age_from_investigation}')")

        extracted_birth_year_from_trial = extract_birth_year_from_trial_report(trial_text_raw)
        calculated_age_from_trial = None
        if extracted_birth_year_from_trial is not None:
            calculated_age_from_trial = current_year - extracted_birth_year_from_trial
        if (calculated_age_from_trial is None) or \
           (excel_age is not None and calculated_age_from_trial is not None and excel_age != calculated_age_from_trial):
            age_mismatch_indices.add(index)
            issues_list.append((index, excel_case_code, excel_person_code, "N2年龄与CY2审理报告不一致"))
            logger.info(f"行 {index + 1} - 年龄不匹配: Excel年龄 ('{excel_age}') vs 审理报告计算年龄 ('{calculated_age_from_trial}')")
            print(f"行 {index + 1} - 年龄不匹配: Excel年龄 ('{excel_age}') vs 审理报告计算年龄 ('{calculated_age_from_trial}')")

        # --- Brief Case Details matching rules ---
        is_brief_case_details_mismatch = False
        extracted_brief_case_details = None

        if pd.isna(row["处分决定"]) or decision_text_raw == '':
            # 当处分决定字段为空时，与立案报告的涉嫌违纪问题段落进行对比
            extracted_brief_case_details = extract_suspected_violation_from_case_report(report_text_raw)
            logger.info(f"行 {index + 1} - 处分决定为空，从立案报告提取简要案情：'{extracted_brief_case_details}'")
            print(f"行 {index + 1} - 处分决定为空，从立案报告提取简要案情：'{extracted_brief_case_details}'")
        else:
            # 当处分决定不为空时，与处分决定的涉嫌违纪问题段落进行对比
            # 需要提供被调查人姓名，这里使用excel_investigated_person
            extracted_brief_case_details = extract_suspected_violation_from_decision(decision_text_raw, investigated_person)
            logger.info(f"行 {index + 1} - 处分决定不为空，从处分决定提取简要案情：'{extracted_brief_case_details}'")
            print(f"行 {index + 1} - 处分决定不为空，从处分决定提取简要案情：'{extracted_brief_case_details}'")

        # 对比逻辑
        if extracted_brief_case_details is None:
            # 如果未能从任何报告中提取到内容，但Excel有值，则认为不一致
            if excel_brief_case_details:
                is_brief_case_details_mismatch = True
                issues_list.append((index, excel_case_code, excel_person_code, "BE简要案情与相关报告不一致（未能提取到内容）"))
                logger.info(f"行 {index + 1} - 简要案情不匹配: Excel有值 ('{excel_brief_case_details}') 但未能从报告中提取。")
                print(f"行 {index + 1} - 简要案情不匹配: Excel有值 ('{excel_brief_case_details}') 但未能从报告中提取。")
        else:
            # 清理Excel中的简要案情，移除所有空白符进行精准匹配
            cleaned_excel_brief_case_details = re.sub(r'\s+', '', excel_brief_case_details)
            
            if cleaned_excel_brief_case_details != extracted_brief_case_details:
                is_brief_case_details_mismatch = True
                issues_list.append((index, excel_case_code, excel_person_code, "BE简要案情与CU处分决定不一致")) # 统一问题描述
                logger.info(f"行 {index + 1} - 简要案情不匹配: Excel简要案情 ('{cleaned_excel_brief_case_details}') vs 提取简要案情 ('{extracted_brief_case_details}')")
                print(f"行 {index + 1} - 简要案情不匹配: Excel简要案情 ('{cleaned_excel_brief_case_details}') vs 提取简要案情 ('{extracted_brief_case_details}')")

        if is_brief_case_details_mismatch:
            brief_case_details_mismatch_indices.add(index)


        # --- Birth Date matching rules ---
        extracted_birth_date_from_report = extract_birth_date_from_case_report(report_text_raw)
        is_birth_date_mismatch_report = False
        if pd.isna(row["出生年月"]) or excel_birth_date == '':
            if extracted_birth_date_from_report is not None:
                is_birth_date_mismatch_report = True
        elif extracted_birth_date_from_report is None:
            is_birth_date_mismatch_report = True
        elif excel_birth_date != extracted_birth_date_from_report:
            is_birth_date_mismatch_report = True
        if is_birth_date_mismatch_report:
            birth_date_mismatch_indices.add(index)
            issues_list.append((index, excel_case_code, excel_person_code, "O2出生年月与BF2立案报告不一致"))
            logger.info(f"行 {index + 1} - 出生年月不匹配: Excel出生年月 ('{excel_birth_date}') vs 立案报告提取出生年月 ('{extracted_birth_date_from_report}')")
            print(f"行 {index + 1} - 出生年月不匹配: Excel出生年月 ('{excel_birth_date}') vs 立案报告提取出生年月 ('{extracted_birth_date_from_report}')")

        extracted_birth_date_from_decision = extract_birth_date_from_decision_report(decision_text_raw)
        is_birth_date_mismatch_decision = False
        if pd.isna(row["出生年月"]) or excel_birth_date == '':
            if extracted_birth_date_from_decision is not None:
                is_birth_date_mismatch_decision = True
        elif extracted_birth_date_from_decision is None:
            is_birth_date_mismatch_decision = True
        elif excel_birth_date != extracted_birth_date_from_decision:
            is_birth_date_mismatch_decision = True
        if is_birth_date_mismatch_decision:
            birth_date_mismatch_indices.add(index)
            issues_list.append((index, excel_case_code, excel_person_code, "O2出生年月与CU2处分决定不一致"))
            logger.info(f"行 {index + 1} - 出生年月不匹配: Excel出生年月 ('{excel_birth_date}') vs 处分决定提取出生年月 ('{extracted_birth_date_from_decision}')")
            print(f"行 {index + 1} - 出生年月不匹配: Excel出生年月 ('{excel_birth_date}') vs 处分决定提取出生年月 ('{extracted_birth_date_from_decision}')")

        extracted_birth_date_from_investigation = extract_birth_date_from_investigation_report(investigation_text_raw)
        is_birth_date_mismatch_investigation = False
        if pd.isna(row["出生年月"]) or excel_birth_date == '':
            if extracted_birth_date_from_investigation is not None:
                is_birth_date_mismatch_investigation = True
        elif extracted_birth_date_from_investigation is None:
            is_birth_date_mismatch_investigation = True
        elif excel_birth_date != extracted_birth_date_from_investigation:
            is_birth_date_mismatch_investigation = True
        if is_birth_date_mismatch_investigation:
            birth_date_mismatch_indices.add(index)
            issues_list.append((index, excel_case_code, excel_person_code, "O2出生年月与CX2审查调查报告不一致"))
            logger.info(f"行 {index + 1} - 出生年月不匹配: Excel出生年月 ('{excel_birth_date}') vs 审查调查报告提取出生年月 ('{extracted_birth_date_from_investigation}')")
            print(f"行 {index + 1} - 出生年月不匹配: Excel出生年月 ('{excel_birth_date}') vs 审查调查报告提取出生年月 ('{extracted_birth_date_from_investigation}')")

        extracted_birth_date_from_trial = extract_birth_date_from_trial_report(trial_text_raw)
        is_birth_date_mismatch_trial = False
        if pd.isna(row["出生年月"]) or excel_birth_date == '':
            if extracted_birth_date_from_trial is not None:
                is_birth_date_mismatch_trial = True
        elif extracted_birth_date_from_trial is None:
            is_birth_date_mismatch_trial = True
        elif excel_birth_date != extracted_birth_date_from_trial:
            is_birth_date_mismatch_trial = True
        if is_birth_date_mismatch_trial:
            birth_date_mismatch_indices.add(index)
            issues_list.append((index, excel_case_code, excel_person_code, "O2出生年月与CY2审理报告不一致"))
            logger.info(f"行 {index + 1} - 出生年月不匹配: Excel出生年月 ('{excel_birth_date}') vs 审理报告提取出生年月 ('{extracted_birth_date_from_trial}')")
            print(f"行 {index + 1} - 出生年月不匹配: Excel出生年月 ('{excel_birth_date}') vs 审理报告提取出生年月 ('{extracted_birth_date_from_trial}')")

        # --- Education matching rules ---
        extracted_education_from_report = extract_education_from_case_report(report_text_raw)
        is_education_mismatch_report = False
        excel_education_normalized = excel_education
        if excel_education == "大学本科":
            excel_education_normalized = "本科"
        extracted_education_normalized = extracted_education_from_report
        if extracted_education_from_report == "大学本科":
            extracted_education_normalized = "本科"

        if not excel_education:
            if extracted_education_from_report is not None:
                is_education_mismatch_report = True
                issues_list.append((index, excel_case_code, excel_person_code, "P2学历与BF2立案报告不一致"))
                logger.info(f"行 {index + 1} - 学历不匹配: Excel学历为空，但立案报告中提取到学历 ('{extracted_education_from_report}')。")
                print(f"行 {index + 1} - 学历不匹配: Excel学历为空，但立案报告中提取到学历 ('{extracted_education_from_report}')。")
        else:
            if extracted_education_from_report is None:
                is_education_mismatch_report = True
                issues_list.append((index, excel_case_code, excel_person_code, "P2学历与BF2立案报告不一致"))
                logger.info(f"行 {index + 1} - 学历不匹配: Excel学历 ('{excel_education}') 有值，但立案报告中未提取到学历。")
                print(f"行 {index + 1} - 学历不匹配: Excel学历 ('{excel_education}') 有值，但立案报告中未提取到学历。")
            elif excel_education_normalized != extracted_education_normalized:
                is_education_mismatch_report = True
                issues_list.append((index, excel_case_code, excel_person_code, "P2学历与BF2立案报告不一致"))
                logger.info(f"行 {index + 1} - 学历不匹配: Excel学历 ('{excel_education}') vs 立案报告提取学历 ('{extracted_education_from_report}')。")
                print(f"行 {index + 1} - 学历不匹配: Excel学历 ('{excel_education}') vs 立案报告提取学历 ('{extracted_education_from_report}')")
        if is_education_mismatch_report:
            education_mismatch_indices.add(index)
        
        # --- Ethnicity matching rules ---
        extracted_ethnicity_from_report = extract_ethnicity_from_case_report(report_text_raw)
        is_ethnicity_mismatch_report = False
        if not excel_ethnicity:
            if extracted_ethnicity_from_report is not None:
                is_ethnicity_mismatch_report = True
                issues_list.append((index, excel_case_code, excel_person_code, "Q2民族与BF2立案报告不一致"))
                logger.info(f"行 {index + 1} - 民族不匹配: Excel民族为空，但立案报告中提取到民族 ('{extracted_ethnicity_from_report}')。")
                print(f"行 {index + 1} - 民族不匹配: Excel民族为空，但立案报告中提取到民族 ('{extracted_ethnicity_from_report}'))。")
        elif extracted_ethnicity_from_report is None:
            is_ethnicity_mismatch_report = True
            issues_list.append((index, excel_case_code, excel_person_code, "Q2民族与BF2立案报告不一致"))
            logger.info(f"行 {index + 1} - 民族不匹配: Excel民族 ('{excel_ethnicity}') 有值，但立案报告中未提取到民族。")
            print(f"行 {index + 1} - 民族不匹配: Excel民族 ('{excel_ethnicity}') 有值，但立案报告中未提取到民族。")
        elif excel_ethnicity != extracted_ethnicity_from_report:
            is_ethnicity_mismatch_report = True
            issues_list.append((index, excel_case_code, excel_person_code, "Q2民族与BF2立案报告不一致"))
            logger.info(f"行 {index + 1} - 民族不匹配: Excel民族 ('{excel_ethnicity}') vs 立案报告提取民族 ('{extracted_ethnicity_from_report}')。")
            print(f"行 {index + 1} - 民族不匹配: Excel民族 ('{excel_ethnicity}') vs 立案报告提取民族 ('{extracted_ethnicity_from_report}')")
        if is_ethnicity_mismatch_report:
            ethnicity_mismatch_indices.add(index)
        
        extracted_ethnicity_from_decision = extract_ethnicity_from_decision_report(decision_text_raw)
        is_ethnicity_mismatch_decision = False
        if not excel_ethnicity:
            if extracted_ethnicity_from_decision is not None:
                is_ethnicity_mismatch_decision = True
                issues_list.append((index, excel_case_code, excel_person_code, "Q2民族与CU2处分决定不一致"))
                logger.info(f"行 {index + 1} - 民族不匹配: Excel民族为空，但处分决定中提取到民族 ('{extracted_ethnicity_from_decision}')。")
                print(f"行 {index + 1} - 民族不匹配: Excel民族为空，但处分决定中提取到民族 ('{extracted_ethnicity_from_decision}')。")
        elif extracted_ethnicity_from_decision is None:
            is_ethnicity_mismatch_decision = True
            issues_list.append((index, excel_case_code, excel_person_code, "Q2民族与CU2处分决定不一致"))
            logger.info(f"行 {index + 1} - 民族不匹配: Excel民族 ('{excel_ethnicity}') 有值，但处分决定中未提取到民族。")
            print(f"行 {index + 1} - 民族不匹配: Excel民族 ('{excel_ethnicity}') 有值，但处分决定中未提取到民族。")
        elif excel_ethnicity != extracted_ethnicity_from_decision:
            is_ethnicity_mismatch_decision = True
            issues_list.append((index, excel_case_code, excel_person_code, "Q2民族与CU2处分决定不一致"))
            logger.info(f"行 {index + 1} - 民族不匹配: Excel民族 ('{excel_ethnicity}') vs 处分决定提取民族 ('{extracted_ethnicity_from_decision}')。")
            print(f"行 {index + 1} - 民族不匹配: Excel民族 ('{excel_ethnicity}') vs 处分决定提取民族 ('{extracted_ethnicity_from_decision}')")
        if is_ethnicity_mismatch_decision:
            ethnicity_mismatch_indices.add(index)

        extracted_ethnicity_from_investigation = extract_ethnicity_from_investigation_report(investigation_text_raw)
        is_ethnicity_mismatch_investigation = False
        if not excel_ethnicity:
            if extracted_ethnicity_from_investigation is not None:
                is_ethnicity_mismatch_investigation = True
                issues_list.append((index, excel_case_code, excel_person_code, "Q2民族与CX2审查调查报告不一致"))
                logger.info(f"行 {index + 1} - 民族不匹配: Excel民族为空，但审查调查报告中提取到民族 ('{extracted_ethnicity_from_investigation}')。")
                print(f"行 {index + 1} - 民族不匹配: Excel民族为空，但审查调查报告中提取到民族 ('{extracted_ethnicity_from_investigation}')。")
        elif extracted_ethnicity_from_investigation is None:
            is_ethnicity_mismatch_investigation = True
            issues_list.append((index, excel_case_code, excel_person_code, "Q2民族与CX2审查调查报告不一致"))
            logger.info(f"行 {index + 1} - 民族不匹配: Excel民族 ('{excel_ethnicity}') 有值，但审查调查报告中未提取到民族。")
            print(f"行 {index + 1} - 民族不匹配: Excel民族 ('{excel_ethnicity}') 有值，但审查调查报告中未提取到民族。")
        elif excel_ethnicity != extracted_ethnicity_from_investigation:
            is_ethnicity_mismatch_investigation = True
            issues_list.append((index, excel_case_code, excel_person_code, "Q2民族与CX2审查调查报告不一致"))
            logger.info(f"行 {index + 1} - 民族不匹配: Excel民族 ('{excel_ethnicity}') vs 审查调查报告提取民族 ('{extracted_ethnicity_from_investigation}')。")
            print(f"行 {index + 1} - 民族不匹配: Excel民族 ('{excel_ethnicity}') vs 审查调查报告提取民族 ('{extracted_ethnicity_from_investigation}')")
        if is_ethnicity_mismatch_investigation:
            ethnicity_mismatch_indices.add(index)

        extracted_ethnicity_from_trial = extract_ethnicity_from_trial_report(trial_text_raw)
        is_ethnicity_mismatch_trial = False
        if not excel_ethnicity:
            if extracted_ethnicity_from_trial is not None:
                is_ethnicity_mismatch_trial = True
                issues_list.append((index, excel_case_code, excel_person_code, "Q2民族与CY2审理报告不一致"))
                logger.info(f"行 {index + 1} - 民族不匹配: Excel民族为空，但审理报告中提取到民族 ('{extracted_ethnicity_from_trial}')。")
                print(f"行 {index + 1} - 民族不匹配: Excel民族为空，但审理报告中提取到民族 ('{extracted_ethnicity_from_trial}')。")
        elif extracted_ethnicity_from_trial is None:
            is_ethnicity_mismatch_trial = True
            issues_list.append((index, excel_case_code, excel_person_code, "Q2民族与CY2审理报告不一致"))
            logger.info(f"行 {index + 1} - 民族不匹配: Excel民族 ('{excel_ethnicity}') 有值，但审理报告中未提取到民族。")
            print(f"行 {index + 1} - 民族不匹配: Excel民族 ('{excel_ethnicity}') 有值，但审理报告中未提取到民族。")
        elif excel_ethnicity != extracted_ethnicity_from_trial:
            is_ethnicity_mismatch_trial = True
            issues_list.append((index, excel_case_code, excel_person_code, "Q2民族与CY2审理报告不一致"))
            logger.info(f"行 {index + 1} - 民族不匹配: Excel民族 ('{excel_ethnicity}') vs 审理报告提取民族 ('{extracted_ethnicity_from_trial}')。")
            print(f"行 {index + 1} - 民族不匹配: Excel民族 ('{excel_ethnicity}') vs 审理报告提取民族 ('{extracted_ethnicity_from_trial}')")
        if is_ethnicity_mismatch_trial:
            ethnicity_mismatch_indices.add(index)

        # --- Party Member matching rules ---
        extracted_party_member_from_report = extract_party_member_from_case_report(report_text_raw)
        is_party_member_mismatch_report = False
        if not excel_party_member:
            if extracted_party_member_from_report == "是":
                is_party_member_mismatch_report = True
                issues_list.append((index, excel_case_code, excel_person_code, "T2是否中共党员与BF2立案报告不一致"))
                logger.info(f"行 {index + 1} - 是否中共党员不匹配: Excel字段为空，但立案报告中提取到“是”。")
                print(f"行 {index + 1} - 是否中共党员不匹配: Excel字段为空，但立案报告中提取到“是”。")
        elif extracted_party_member_from_report is None:
            is_party_member_mismatch_report = True
            issues_list.append((index, excel_case_code, excel_person_code, "T2是否中共党员与BF2立案报告不一致"))
            logger.info(f"行 {index + 1} - 是否中共党员不匹配: Excel字段 ('{excel_party_member}') 有值，但立案报告中未明确提取到党员信息。")
            print(f"行 {index + 1} - 是否中共党员不匹配: Excel字段 ('{excel_party_member}') 有值，但立案报告中未明确提取到党员信息。")
        elif excel_party_member != extracted_party_member_from_report:
            is_party_member_mismatch_report = True
            issues_list.append((index, excel_case_code, excel_person_code, "T2是否中共党员与BF2立案报告不一致"))
            logger.info(f"行 {index + 1} - 是否中共党员不匹配: Excel字段 ('{excel_party_member}') vs 立案报告提取 ('{extracted_party_member_from_report}')。")
            print(f"行 {index + 1} - 是否中共党员不匹配: Excel字段 ('{excel_party_member}') vs 立案报告提取 ('{extracted_party_member_from_report}')。")
        if is_party_member_mismatch_report:
            party_member_mismatch_indices.add(index)

        extracted_party_member_from_decision = extract_party_member_from_decision_report(decision_text_raw)
        is_party_member_mismatch_decision = False
        if not excel_party_member:
            if extracted_party_member_from_decision == "是":
                is_party_member_mismatch_decision = True
                issues_list.append((index, excel_case_code, excel_person_code, "T2是否中共党员与CU2处分决定不一致"))
                logger.info(f"行 {index + 1} - 是否中共党员不匹配: Excel字段为空，但处分决定中提取到“是”。")
                print(f"行 {index + 1} - 是否中共党员不匹配: Excel字段为空，但处分决定中提取到“是”。")
            elif extracted_party_member_from_decision == "否":
                pass
        elif extracted_party_member_from_decision is None:
            is_party_member_mismatch_decision = True
            issues_list.append((index, excel_case_code, excel_person_code, "T2是否中共党员与CU2处分决定不一致"))
            logger.info(f"行 {index + 1} - 是否中共党员不匹配: Excel字段 ('{excel_party_member}') 有值，但处分决定中未明确提取到党员信息。")
            print(f"行 {index + 1} - 是否中共党员不匹配: Excel字段 ('{excel_party_member}') 有值，但处分决定中未明确提取到党员信息。")
        elif excel_party_member != extracted_party_member_from_decision:
            is_party_member_mismatch_decision = True
            issues_list.append((index, excel_case_code, excel_person_code, "T2是否中共党员与CU2处分决定不一致"))
            logger.info(f"行 {index + 1} - 是否中共党员不匹配: Excel字段 ('{excel_party_member}') vs 处分决定提取 ('{extracted_party_member_from_decision}')。")
            print(f"行 {index + 1} - 是否中共党员不匹配: Excel字段 ('{excel_party_member}') vs 处分决定提取 ('{extracted_party_member_from_decision}')。")
        if is_party_member_mismatch_decision:
            party_member_mismatch_indices.add(index)

        # --- Party Joining Date matching rule ---
        extracted_party_joining_date_from_report = extract_party_joining_date_from_case_report(report_text_raw)
        is_party_joining_date_mismatch = False

        if excel_party_member == "是":
            if not excel_party_joining_date:
                if extracted_party_joining_date_from_report is not None:
                    is_party_joining_date_mismatch = True
                    issues_list.append((index, excel_case_code, excel_person_code, "V2入党时间与BF2立案报告不一致"))
                    logger.info(f"行 {index + 1} - 入党时间不匹配: Excel入党时间为空，但立案报告中提取到 ('{extracted_party_joining_date_from_report}')。")
                    print(f"行 {index + 1} - 入党时间不匹配: Excel入党时间为空，但立案报告中提取到 ('{extracted_party_joining_date_from_report}')。")
            elif extracted_party_joining_date_from_report is None:
                is_party_joining_date_mismatch = True
                issues_list.append((index, excel_case_code, excel_person_code, "V2入党时间与BF2立案报告不一致"))
                logger.info(f"行 {index + 1} - 入党时间不匹配: Excel入党时间 ('{excel_party_joining_date}') 有值，但立案报告中未提取到。")
                print(f"行 {index + 1} - 入党时间不匹配: Excel入党时间 ('{excel_party_joining_date}') 有值，但立案报告中未提取到。")
            elif excel_party_joining_date != extracted_party_joining_date_from_report:
                is_party_joining_date_mismatch = True
                issues_list.append((index, excel_case_code, excel_person_code, "V2入党时间与BF2立案报告不一致"))
                logger.info(f"行 {index + 1} - 入党时间不匹配: Excel入党时间 ('{excel_party_joining_date}') vs 立案报告提取 ('{extracted_party_joining_date_from_report}')。")
                print(f"行 {index + 1} - 入党时间不匹配: Excel入党时间 ('{excel_party_joining_date}') vs 立案报告提取 ('{extracted_party_joining_date_from_report}')。")
        elif excel_party_member == "否":
            if excel_party_joining_date:
                is_party_joining_date_mismatch = True
                issues_list.append((index, excel_case_code, excel_person_code, "V2入党时间与BF2立案报告不一致"))
                logger.info(f"行 {index + 1} - 入党时间不匹配: Excel是否中共党员为“否”，但入党时间字段不为空 ('{excel_party_joining_date}')。")
                print(f"行 {index + 1} - 入党时间不匹配: Excel是否中共党员为“否”，但入党时间字段不为空 ('{excel_party_joining_date}')。")

        if is_party_joining_date_mismatch:
            party_joining_date_mismatch_indices.add(index)


        # --- Name matching rules ---
        report_name = extract_name_from_case_report(report_text_raw)
        if report_name and investigated_person != report_name:
            mismatch_indices.add(index)
            issues_list.append((index, excel_case_code, excel_person_code, "C2被调查人与BF2立案报告不一致"))
            logger.info(f"行 {index + 1} - 姓名不匹配: C2被调查人 ('{investigated_person}') vs BF2立案报告 ('{report_name}')")
            print(f"行 {index + 1} - 姓名不匹配: C2被调查人 ('{investigated_person}') vs BF2立案报告 ('{report_name}')")

        decision_name = extract_name_from_decision(decision_text_raw)
        if not decision_name or (decision_name and investigated_person != decision_name):
            mismatch_indices.add(index)
            issues_list.append((index, excel_case_code, excel_person_code, "C2被调查人与CU2处分决定不一致"))
            logger.info(f"行 {index + 1} - 姓名不匹配: C2被调查人 ('{investigated_person}') vs CU2处分决定 ('{decision_name}')")
            print(f"行 {index + 1} - 姓名不匹配: C2被调查人 ('{investigated_person}') vs CU2处分决定 ('{decision_name}')")

        investigation_name = extract_name_from_case_report(investigation_text_raw)
        if investigation_name and investigated_person != investigation_name:
            mismatch_indices.add(index)
            issues_list.append((index, excel_case_code, excel_person_code, "C2被调查人与CX2审查调查报告不一致"))
            logger.info(f"行 {index + 1} - 姓名不匹配: C2被调查人 ('{investigated_person}') vs CX2审查调查报告 ('{investigation_name}')")
            print(f"行 {index + 1} - 姓名不匹配: C2被调查人 ('{investigated_person}') vs CX2审查调查报告 ('{investigation_name}')")

        trial_name = extract_name_from_trial_report(trial_text_raw)
        if not trial_name or (trial_name and investigated_person != trial_name):
            mismatch_indices.add(index)
            issues_list.append((index, excel_case_code, excel_person_code, "C2被调查人与CY2审理报告不一致"))
            logger.info(f"行 {index + 1} - 姓名不匹配: C2被调查人 ('{investigated_person}') vs CY2审理报告 ('{trial_name}')")
            print(f"行 {index + 1} - 姓名不匹配: C2被调查人 ('{investigated_person}') vs CY2审理报告 ('{trial_name}')")

    # 调用立案时间规则验证函数
    validate_filing_time(df, issues_list, filing_time_mismatch_indices,
                           disciplinary_committee_filing_time_mismatch_indices,
                           disciplinary_committee_filing_authority_mismatch_indices,
                           supervisory_committee_filing_time_mismatch_indices, # 传递新增的集合
                           supervisory_committee_filing_authority_mismatch_indices) # 传递新增的集合

    # 返回所有可能的不一致索引集以及更新后的 issues_list
    return mismatch_indices, gender_mismatch_indices, age_mismatch_indices, brief_case_details_mismatch_indices, issues_list, \
           birth_date_mismatch_indices, education_mismatch_indices, ethnicity_mismatch_indices, \
           party_member_mismatch_indices, party_joining_date_mismatch_indices, filing_time_mismatch_indices, \
           disciplinary_committee_filing_time_mismatch_indices, disciplinary_committee_filing_authority_mismatch_indices, \
           supervisory_committee_filing_time_mismatch_indices, supervisory_committee_filing_authority_mismatch_indices # 添加新的返回值
