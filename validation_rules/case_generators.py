import logging
import pandas as pd
import os
from datetime import datetime
import xlsxwriter
from config import Config

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
                       no_party_position_warning_mismatch_indices):
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

    返回:
    tuple: (copy_path, case_num_path) 生成的副本文件路径和立案编号文件路径。
           如果生成失败，返回 (None, None)。
    """
    today = datetime.now().strftime('%Y%m%d')
    case_dir = os.path.join(upload_dir, today, 'case')
    os.makedirs(case_dir, exist_ok=True)

    copy_filename = original_filename.replace('.xlsx', '_副本.xlsx').replace('.xls', '_副本.xlsx')
    copy_path = os.path.join(case_dir, copy_filename)
    
    with pd.ExcelWriter(copy_path, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='Sheet1', index=False)
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']
        
        red_format = workbook.add_format({'bg_color': Config.FORMATS["red"]})
        yellow_format = workbook.add_format({'bg_color': '#FFFF00'})

        try:
            col_index_investigated_person = df.columns.get_loc("被调查人")
            col_index_gender = df.columns.get_loc("性别")
            col_index_age = df.columns.get_loc("年龄")
            col_index_brief_case_details = df.columns.get_loc("简要案情")
            col_index_birth_date = df.columns.get_loc("出生年月")
            col_index_education = df.columns.get_loc("学历")
            col_index_ethnicity = df.columns.get_loc("民族")
            col_index_party_member = df.columns.get_loc("是否中共党员")
            col_index_party_joining_date = df.columns.get_loc("入党时间")
            col_index_filing_time = df.columns.get_loc("立案时间")
            col_index_disciplinary_committee_filing_time = df.columns.get_loc("纪委立案时间")
            col_index_disciplinary_committee_filing_authority = df.columns.get_loc("纪委立案机关")
            col_index_supervisory_committee_filing_time = df.columns.get_loc("监委立案时间")
            col_index_supervisory_committee_filing_authority = df.columns.get_loc("监委立案机关")
            col_index_case_report = df.columns.get_loc("立案报告")
            col_index_disposal_spirit = df.columns.get_loc("是否违反中央八项规定精神")
            col_index_voluntary_confession = df.columns.get_loc("是否主动交代问题")
            col_index_closing_time = df.columns.get_loc("结案时间")
            col_index_no_party_position_warning = df.columns.get_loc("是否属于本应撤销党内职务，但本人没有党内职务而给予严重警告处分")

        except KeyError as e:
            logger.error(f"Excel 文件缺少必要的列: {e}")
            print(f"Excel 文件缺少必要的列: {e}")
            return None, None

        for idx in range(len(df)):
            if idx in mismatch_indices:
                worksheet.write(idx + 1, col_index_investigated_person, 
                                df.iloc[idx]["被调查人"] if pd.notna(df.iloc[idx]["被调查人"]) else '', red_format)
            
            if idx in gender_mismatch_indices:
                worksheet.write(idx + 1, col_index_gender,
                                df.iloc[idx]["性别"] if pd.notna(df.iloc[idx]["性别"]) else '', red_format)

            if idx in age_mismatch_indices:
                worksheet.write(idx + 1, col_index_age,
                                df.iloc[idx]["年龄"] if pd.notna(df.iloc[idx]["年龄"]) else '', red_format)

            if idx in brief_case_details_mismatch_indices:
                worksheet.write(idx + 1, col_index_brief_case_details,
                                df.iloc[idx]["简要案情"] if pd.notna(df.iloc[idx]["简要案情"]) else '', red_format)

            if idx in birth_date_mismatch_indices:
                worksheet.write(idx + 1, col_index_birth_date,
                                df.iloc[idx]["出生年月"] if pd.notna(df.iloc[idx]["出生年月"]) else '', red_format)

            if idx in education_mismatch_indices:
                worksheet.write(idx + 1, col_index_education,
                                df.iloc[idx]["学历"] if pd.notna(df.iloc[idx]["学历"]) else '', red_format)

            if idx in ethnicity_mismatch_indices:
                worksheet.write(idx + 1, col_index_ethnicity,
                                df.iloc[idx]["民族"] if pd.notna(df.iloc[idx]["民族"]) else '', red_format)

            if idx in party_member_mismatch_indices:
                worksheet.write(idx + 1, col_index_party_member,
                                df.iloc[idx]["是否中共党员"] if pd.notna(df.iloc[idx]["是否中共党员"]) else '', red_format)

            if idx in party_joining_date_mismatch_indices:
                worksheet.write(idx + 1, col_index_party_joining_date,
                                df.iloc[idx]["入党时间"] if pd.notna(df.iloc[idx]["入党时间"]) else '', red_format)

            if idx in filing_time_mismatch_indices:
                worksheet.write(idx + 1, col_index_filing_time,
                                df.iloc[idx]["立案时间"] if pd.notna(df.iloc[idx]["立案时间"]) else '', red_format)
            
            if idx in disciplinary_committee_filing_time_mismatch_indices:
                worksheet.write(idx + 1, col_index_disciplinary_committee_filing_time,
                                df.iloc[idx]["纪委立案时间"] if pd.notna(df.iloc[idx]["纪委立案时间"]) else '', red_format)

            if idx in disciplinary_committee_filing_authority_mismatch_indices:
                worksheet.write(idx + 1, col_index_disciplinary_committee_filing_authority,
                                df.iloc[idx]["纪委立案机关"] if pd.notna(df.iloc[idx]["纪委立案机关"]) else '', red_format)

            if idx in supervisory_committee_filing_time_mismatch_indices:
                worksheet.write(idx + 1, col_index_supervisory_committee_filing_time,
                                df.iloc[idx]["监委立案时间"] if pd.notna(df.iloc[idx]["监委立案时间"]) else '', red_format)

            if idx in supervisory_committee_filing_authority_mismatch_indices:
                worksheet.write(idx + 1, col_index_supervisory_committee_filing_authority,
                                df.iloc[idx]["监委立案机关"] if pd.notna(df.iloc[idx]["监委立案机关"]) else '', red_format)
            
            if idx in case_report_keyword_mismatch_indices:
                worksheet.write(idx + 1, col_index_case_report,
                                df.iloc[idx]["立案报告"] if pd.notna(df.iloc[idx]["立案报告"]) else '', red_format)
            
            if idx in disposal_spirit_mismatch_indices:
                worksheet.write(idx + 1, col_index_disposal_spirit,
                                df.iloc[idx]["是否违反中央八项规定精神"] if pd.notna(df.iloc[idx]["是否违反中央八项规定精神"]) else '', red_format)

            if idx in voluntary_confession_highlight_indices:
                worksheet.write(idx + 1, col_index_voluntary_confession,
                                df.iloc[idx]["是否主动交代问题"] if pd.notna(df.iloc[idx]["是否主动交代问题"]) else '', yellow_format)

            if idx in closing_time_mismatch_indices:
                worksheet.write(idx + 1, col_index_closing_time,
                                df.iloc[idx]["结案时间"] if pd.notna(df.iloc[idx]["结案时间"]) else '', red_format)

            if idx in no_party_position_warning_mismatch_indices:
                worksheet.write(idx + 1, col_index_no_party_position_warning,
                                df.iloc[idx]["是否属于本应撤销党内职务，但本人没有党内职务而给予严重警告处分"] if pd.notna(df.iloc[idx]["是否属于本应撤销党内职务，但本人没有党内职务而给予严重警告处分"]) else '', red_format)

    logger.info(f"Generated copy file with highlights: {copy_path}")
    print(f"生成高亮后的副本文件: {copy_path}")

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

    issues_df.to_excel(case_num_path, index=False)
    logger.info(f"Generated case number file: {case_num_path}")
    print(f"生成立案编号表: {case_num_path}")

    return copy_path, case_num_path