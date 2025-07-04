import logging
import pandas as pd
import os
from datetime import datetime
import xlsxwriter 
from config import Config 
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
                        trial_authority_agency_mismatch_indices,
                        disposal_decision_keyword_mismatch_indices,
                        trial_report_non_representative_mismatch_indices, 
                        trial_report_detention_mismatch_indices,
                        confiscation_amount_indices,
                        confiscation_of_property_amount_indices,
                        compensation_amount_highlight_indices,
                        registered_handover_amount_indices, 
                        disciplinary_sanction_mismatch_indices 
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
    issues_list (list): 包含所有问题的列表，每个问题是一个字典。
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
    confiscation_of_property_amount_indices (set): 没收金额需要高亮的行索引集合。
    compensation_amount_highlight_indices (set): 责令退赔金额需要高亮的行索引集合。
    registered_handover_amount_indices (set): 登记上交金额需要高亮的行索引集合。
    disciplinary_sanction_mismatch_indices (set): 党纪处分不匹配的行索引集合。

    返回:
    tuple: (copy_path, case_num_path) 生成的副本文件路径和立案编号文件路径。
            如果生成失败，返回 (None, None)。
    """
    case_dir = Config.CASE_FOLDER 
    os.makedirs(case_dir, exist_ok=True) 

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
            confiscation_amount_indices,
            confiscation_of_property_amount_indices,
            compensation_amount_highlight_indices,
            registered_handover_amount_indices,
            disciplinary_sanction_mismatch_indices 
        )
        logger.info(f"Generated copy file with highlights: {copy_path}")
        print(f"生成高亮后的副本文件: {copy_path}")
    except Exception as e:
        logger.error(f"生成高亮副本文件失败: {e}", exc_info=True)
        print(f"生成高亮副本文件失败: {e}")
        return None, None

    case_num_filename = f"立案编号{datetime.now().strftime('%Y%m%d')}.xlsx" 
    case_num_path = os.path.join(case_dir, case_num_filename)
    
    data = []
    for i, issue_item in enumerate(issues_list):
        # 确保字典中存在这些键，即使值为空，也用空字符串填充
        case_code = issue_item.get('案件编码', '')
        person_code = issue_item.get('涉案人员编码', '')
        # 如果是线索表，可能没有案件编码和涉案人员编码，而是受理线索编码
        if not case_code and not person_code:
            case_code = issue_item.get('受理线索编码', '') # 尝试获取线索编码

        issue_description = issue_item.get('问题描述', '')
        data.append({'序号': i + 1, '案件编码': case_code, '涉案人员编码': person_code, '问题': issue_description})
    
    issues_df = pd.DataFrame(data)

    try:
        with pd.ExcelWriter(case_num_path, engine='xlsxwriter') as writer:
            issues_df.to_excel(writer, sheet_name='Sheet1', index=False)
            workbook = writer.book
            worksheet = writer.sheets['Sheet1']

            # 定义通用的左对齐和文本格式
            # 这里的关键是 'num_format': '@'，它强制将单元格内容视为文本
            left_align_text_format = workbook.add_format({
                'align': 'left', 
                'valign': 'vcenter',
                'num_format': '@' # 强制文本格式
            })

            # 设置所有相关列的对齐方式和文本格式
            # 对于“序号”列，通常希望居中或右对齐，但为了统一，这里也设置为左对齐文本
            columns_to_format = ['序号', '案件编码', '涉案人员编码', '问题'] 
            
            for col_name in columns_to_format:
                if col_name in issues_df.columns:
                    col_idx = issues_df.columns.get_loc(col_name)
                    worksheet.set_column(col_idx, col_idx, None, left_align_text_format)
            
            # 自动调整列宽以适应内容
            for i, col in enumerate(issues_df.columns):
                # 考虑列名和实际数据中最大长度来设置列宽
                # astype(str) 确保所有内容都被视为字符串来计算长度
                max_len = max(issues_df[col].astype(str).map(len).max(), len(col)) + 2 # 加2是为了留出一点边距
                worksheet.set_column(i, i, max_len)

        logger.info(f"Generated case number file: {case_num_path}")
        print(f"生成立案编号表: {case_num_path}")
    except Exception as e:
        logger.error(f"生成立案编号文件失败: {e}", exc_info=True)
        print(f"生成立案编号文件失败: {e}")
        return copy_path, None 

    return copy_path, case_num_path