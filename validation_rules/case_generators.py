import logging
import pandas as pd
import os
from datetime import datetime
import xlsxwriter
from config import Config # Assuming Config class exists in config.py

logger = logging.getLogger(__name__)

def generate_case_files(df, original_filename, upload_dir, mismatch_indices, gender_mismatch_indices, issues_list, age_mismatch_indices, birth_date_mismatch_indices, education_mismatch_indices, ethnicity_mismatch_indices, party_member_mismatch_indices, party_joining_date_mismatch_indices):
    """Generate copy and case number Excel files based on analysis."""
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

        try:
            col_index_investigated_person = df.columns.get_loc("被调查人")
            col_index_gender = df.columns.get_loc("性别")
            col_index_age = df.columns.get_loc("年龄")
            col_index_birth_date = df.columns.get_loc("出生年月")
            col_index_education = df.columns.get_loc("学历") 
            col_index_ethnicity = df.columns.get_loc("民族") 
            col_index_party_member = df.columns.get_loc("是否中共党员") 
            col_index_party_joining_date = df.columns.get_loc("入党时间")
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


    logger.info(f"Generated copy file with highlights: {copy_path}")
    print(f"生成高亮后的副本文件: {copy_path}")


    case_num_filename = f"立案编号{today}.xlsx"
    case_num_path = os.path.join(case_dir, case_num_filename)
    issues_df = pd.DataFrame(columns=['序号', '问题'])
    if not issues_list:
        issues_df = pd.DataFrame({'序号': [1], '问题': ['无问题']})
    else:
        data = [{'序号': i + 1, '问题': issue} for i, (index, issue) in enumerate(issues_list)]
        issues_df = pd.DataFrame(data)

    issues_df.to_excel(case_num_path, index=False)
    logger.info(f"Generated case number file: {case_num_path}")
    print(f"生成立案编号表: {case_num_path}")

    return copy_path, case_num_path
