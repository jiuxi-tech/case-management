import logging
import pandas as pd
import os
from datetime import datetime
import xlsxwriter # 即使不用直接写入，但pd.ExcelWriter可能间接使用，保留
from config import Config
# 【关键修改】导入 excel_formatter 模块
from excel_formatter import format_excel

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
                        trial_authority_agency_mismatch_indices): # 接收 trial_authority_agency_mismatch_indices
    """
    根据分析结果生成副本和立案编号Excel文件。
    该函数将原始DataFrame写入一个副本文件，对不匹配的单元格进行标红。
    同时，它会生成一个立案编号文件，其中包含所有发现的问题列表。

    参数:
    df (pd.DataFrame): 原始Excel数据的DataFrame。
    original_filename (str): 原始上传的文件名。
    upload_dir (str): 上传文件的根目录。
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

    返回:
    tuple: (copy_path, case_num_path) 生成的副本文件路径和立案编号文件路径。
            如果生成失败，返回 (None, None)。
    """
    today = datetime.now().strftime('%Y%m%d')
    case_dir = os.path.join(upload_dir, today, 'case')
    os.makedirs(case_dir, exist_ok=True)

    copy_filename = original_filename.replace('.xlsx', '_副本.xlsx').replace('.xls', '_副本.xlsx')
    copy_path = os.path.join(case_dir, copy_filename)
    
    # 【关键修改】直接调用 excel_formatter.py 中的 format_excel 函数
    # format_excel 将负责写入 DataFrame 并应用所有格式
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
            no_party_position_warning_mismatch_indices, # 传递新参数
            recovery_amount_highlight_indices, # 传递新参数
            trial_acceptance_time_mismatch_indices, # 传递新参数
            trial_closing_time_mismatch_indices, # 传递新参数
            trial_authority_agency_mismatch_indices # 传递新参数
        )
        logger.info(f"Generated copy file with highlights: {copy_path}")
        print(f"生成高亮后的副本文件: {copy_path}")
    except Exception as e:
        logger.error(f"生成高亮副本文件失败: {e}", exc_info=True)
        print(f"生成高亮副本文件失败: {e}")
        return None, None

    case_num_filename = f"立案编号{today}.xlsx"
    case_num_path = os.path.join(case_dir, case_num_filename)
    
    issues_df = pd.DataFrame(columns=['序号', '案件编码', '涉案人员编码', '问题'])
    
    if not issues_list:
        issues_df = pd.DataFrame({'序号': [1], '案件编码': [''], '涉案人员编码': [''], '问题': ['无问题']})
    else:
        data = []
        for i, (original_index, case_code, person_code, issue_description) in enumerate(issues_list):
            data.append({'序号': i + 1, '案件编码': case_code, '涉案人员编码': person_code, '问题': issue_description})
        issues_df = pd.DataFrame(data)

    try:
        issues_df.to_excel(case_num_path, index=False)
        logger.info(f"Generated case number file: {case_num_path}")
        print(f"生成立案编号表: {case_num_path}")
    except Exception as e:
        logger.error(f"生成立案编号文件失败: {e}", exc_info=True)
        print(f"生成立案编号文件失败: {e}")
        return copy_path, None # 返回副本路径，但编号文件生成失败

    return copy_path, case_num_path