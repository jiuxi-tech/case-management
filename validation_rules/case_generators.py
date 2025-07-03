import logging
import pandas as pd
import os
from datetime import datetime
import xlsxwriter 
from config import Config # 确保 Config 被正确导入
from excel_formatter import format_excel # 确保这个导入是正确的

logger = logging.getLogger(__name__)

def generate_case_files(df, original_filename, upload_dir, mismatch_indices, gender_mismatch_indices, issues_list, 
                        age_mismatch_indices, birth_date_mismatch_indices, education_mismatch_indices, 
                        ethnicity_mismatch_indices, party_member_mismatch_indices, party_joining_date_mismatch_indices, 
                        brief_case_details_mismatch_indices, filing_time_mismatch_indices, 
                        disciplinary_committee_filing_time_mismatch_indices, 
                        disciplinary_committee_filing_authority_mismatch_indices, 
                        supervisory_committee_filing_time_mismatch_indices, 
                        supervisory_committee_filing_authority_mismatch_indices, 
                        case_report_keyword_mismatch_indices, disposal_spirit_mismatch_indices, 
                        voluntary_confession_highlight_indices, closing_time_mismatch_indices,
                        no_party_position_warning_mismatch_indices,
                        recovery_amount_highlight_indices,
                        trial_acceptance_time_mismatch_indices,
                        trial_closing_time_mismatch_indices,
                        trial_authority_agency_mismatch_indices,
                        disposal_decision_keyword_mismatch_indices,
                        trial_report_non_representative_mismatch_indices, 
                        trial_report_detention_mismatch_indices,
                        confiscation_amount_indices 
                        ):
    """
    根据分析结果生成副本和立案编号Excel文件。
    该函数将原始DataFrame写入一个副本文件，对不匹配的单元格进行标红。
    同时，它会生成一个立案编号文件，其中包含所有发现的问题列表。

    参数:
    df (pd.DataFrame): 原始Excel数据的DataFrame。
    original_filename (str): 原始上传的文件名。
    upload_dir (str): 上传文件的根目录 (此参数实际上未被使用，我们将直接使用 Config.CASE_FOLDER)。
    mismatch_indices (set): 姓名不匹配的行索引集合。
    gender_mismatch_indices (set): 性别不匹配的行索引集合。
    issues_list (list): 包含所有问题的列表，每个问题是一个(索引, 案件编码, 涉案人员编码, 问题描述)元组。
    age_mismatch_indices (set): 年龄不匹配的行索引集合。
    birth_date_mismatch_indices (set): 出生年月不匹配的行索引集合。
    education_mismatch_indices (set): 学历不匹配的行索引集合。
    ethnicity_mismatch_indices (set): 民族不匹配的行索引集合。
    party_member_mismatch_indices (set): 是否中共党员不匹配的行索引集合。
    party_joining_date_mismatch_indices (set): 入党时间不匹配的行索引集合。
    brief_case_details_mismatch_indices (set): 简要案情不匹配的行索引集合。
    filing_time_mismatch_indices (set): 立案时间不匹配的行索引集合。
    disciplinary_committee_filing_time_mismatch_indices (set): 纪委立案时间不匹配的行索引集合。
    disciplinary_committee_filing_authority_mismatch_indices (set): 纪委立案机关不匹配的行索引集合。
    supervisory_committee_filing_time_mismatch_indices (set): 监委立案时间不一致的行索引集合。
    supervisory_committee_filing_authority_mismatch_indices (set): 监委立案机关不一致的行索引集合。
    case_report_keyword_mismatch_indices (set): 立案报告关键字不一致的行索引集合。
    disposal_spirit_mismatch_indices (set): 是否违反中央八项规定精神不一致的行索引集合。
    voluntary_confession_highlight_indices (set): 是否主动交代问题标黄索引集合。
    closing_time_mismatch_indices (set): 结案时间不一致的行索引集合。
    no_party_position_warning_mismatch_indices (set): 是否属于本应撤销党内职务，但本人没有党内职务而给予严重警告处分不一致的行索引集合。
    recovery_amount_highlight_indices (set): 追缴失职渎职滥用职权造成的损失金额标黄索引集合。
    trial_acceptance_time_mismatch_indices (set): 审理受理时间不一致的行索引集合。
    trial_closing_time_mismatch_indices (set): 审结时间与审理报告落款时间不一致的行索引集合。
    trial_authority_agency_mismatch_indices (set): 审理机关与填报单位不一致的行索引集合。
    disposal_decision_keyword_mismatch_indices (set): 处分决定关键词不一致的行索引集合。
    trial_report_non_representative_mismatch_indices (set): 审理报告中非人大代表/政协委员等关键词的行索引集合。
    trial_report_detention_mismatch_indices (set): 审理报告中出现“扣押”关键词的行索引集合。
    confiscation_amount_indices (set): 收缴金额（万元）需要高亮的行索引集合。 

    返回:
    tuple: (copy_path, case_num_path) 生成的副本文件路径和立案编号文件路径。
            如果生成失败，返回 (None, None)。
    """
    # 修正点：直接使用 Config.CASE_FOLDER 作为文件保存的根目录
    # 这样可以确保与 file_processor.py 中保存上传文件的逻辑一致
    case_dir = Config.CASE_FOLDER 
    os.makedirs(case_dir, exist_ok=True) # 确保目录存在

    copy_filename = original_filename.replace('.xlsx', '_副本.xlsx').replace('.xls', '_副本.xlsx')
    copy_path = os.path.join(case_dir, copy_filename)
    
    try:
        format_excel(
            df, 
            mismatch_indices, 
            copy_path, 
            issues_list, 
            gender_mismatch_indices, 
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
            confiscation_amount_indices 
        )
        logger.info(f"Generated copy file with highlights: {copy_path}")
        print(f"生成高亮后的副本文件: {copy_path}")
    except Exception as e:
        logger.error(f"生成高亮副本文件失败: {e}", exc_info=True)
        print(f"生成高亮副本文件失败: {e}")
        return None, None

    # 立案编号文件的命名方式和路径也调整，确保在同一个目录下
    case_num_filename = f"立案编号{datetime.now().strftime('%Y%m%d')}.xlsx" # 使用当前日期生成文件名
    case_num_path = os.path.join(case_dir, case_num_filename)
    
    issues_df = pd.DataFrame(columns=['序号', '案件编码', '涉案人员编码', '问题'])
    
    if not issues_list:
        issues_df = pd.DataFrame({'序号': [1], '案件编码': [''], '涉案人员编码': [''], '问题': ['无问题']})
    else:
        data = []
        for i, issue_item in enumerate(issues_list):
            # 根据您提供的 issues_list 结构，它应该是 (original_df_index, clue_code_value, issue_description) 3个元素
            # 或者如果是案件，可能是 (original_df_index, case_code_value, person_code_value, issue_description) 4个元素
            # 这里的逻辑需要根据您实际传入 issues_list 的结构来确定。
            # 暂时按照立案登记表可能包含人员编码的情况处理，如果不对，请根据实际传入的 issues_list 结构调整
            if len(issue_item) == 4:
                original_index, case_code, person_code, issue_description = issue_item
            elif len(issue_item) == 3: # 如果 issues_list 只有三个元素 (index, code, description)
                original_index, case_code, issue_description = issue_item
                person_code = '' # 如果没有涉案人员编码，可以留空或给定默认值
            else:
                logger.warning(f"issues_list 中存在未知格式的元素，跳过: {issue_item}")
                continue # 跳过无法解析的行

            data.append({'序号': i + 1, '案件编码': case_code, '涉案人员编码': person_code, '问题': issue_description})
        issues_df = pd.DataFrame(data)

    try:
        issues_df.to_excel(case_num_path, index=False)
        logger.info(f"Generated case number file: {case_num_path}")
        print(f"生成立案编号表: {case_num_path}")
    except Exception as e:
        logger.error(f"生成立案编号文件失败: {e}", exc_info=True)
        print(f"生成立案编号文件失败: {e}")
        return copy_path, None 

    return copy_path, case_num_path