import os
import pandas as pd
from datetime import datetime
import logging
from .case_validation_additional import validate_name_rules, validate_gender_rules

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
        
        # 遍历每一行数据，执行被调查人验证规则
        for index, row in df.iterrows():
            try:
                # 提取必要的字段
                excel_case_code = str(row.get('案件编码', '')).strip()
                excel_person_code = str(row.get('涉案人员编码', '')).strip()
                investigated_person = str(row.get('被调查人', '')).strip()
                report_text_raw = str(row.get('立案报告', '')).strip()
                decision_text_raw = str(row.get('处分决定', '')).strip()
                investigation_text_raw = str(row.get('审查调查报告', '')).strip()
                trial_text_raw = str(row.get('审理报告', '')).strip()
                
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
                excel_gender = str(row.get('性别', '')).strip()
                gender_mismatch_indices = set()
                validate_gender_rules(
                    row, index, excel_case_code, excel_person_code, issues_list, gender_mismatch_indices,
                    excel_gender, report_text_raw, decision_text_raw,
                    investigation_text_raw, trial_text_raw, app_config
                )
                
            except Exception as e:
                logger.error(f"处理第 {index + 2} 行时发生错误: {str(e)}")
                continue
        
        # 生成立案编号表文件
        case_num_filename = f"立案编号表_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
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