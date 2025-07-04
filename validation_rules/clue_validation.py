import logging
import pandas as pd
import re
from datetime import datetime

logger = logging.getLogger(__name__)

def extract_name_from_report(report_content, investigated_person_excel):
    """
    从报告文本中提取姓名，并与Excel中的姓名进行比对。
    优先匹配“关于XX同志”模式，如果未找到，则尝试匹配报告开头的人名。
    """
    # 尝试匹配“关于XX同志”模式
    match_comrade = re.search(r'关于(.{1,10})同志', report_content)
    if match_comrade:
        extracted_name = match_comrade.group(1).strip()
        logger.debug(f"从报告中提取姓名 (同志模式): {extracted_name}")
        return extracted_name
    
    # 如果没有匹配到“同志”模式，尝试从报告开头提取人名
    # 匹配报告开头的姓名，通常是“姓名，性别，民族，出生年月”这种格式
    match_start = re.match(r'^\s*([A-Za-z\u4e00-\u9fa5]{2,5})\s*[男女，，\s]', report_content)
    if match_start:
        extracted_name = match_start.group(1).strip()
        logger.debug(f"从报告中提取姓名 (开头模式): {extracted_name}")
        return extracted_name

    logger.debug("未能从报告中提取到姓名。")
    return None

def extract_gender_from_report(report_content):
    """从报告文本中提取性别。"""
    if "，男，" in report_content or " 男，" in report_content:
        return "男"
    if "，女，" in report_content or " 女，" in report_content:
        return "女"
    return None

def extract_birth_date_from_report(report_content):
    """从报告文本中提取出生年月。"""
    match = re.search(r'(\d{4}年\d{1,2}月(?!日))', report_content) # 匹配“XXXX年X月”，排除“X日”
    if match:
        return match.group(1).replace('年', '/').replace('月', '')
    return None

def extract_ethnicity_from_report(report_content):
    """从报告文本中提取民族。"""
    match = re.search(r'，(汉族|壮族|满族|回族|苗族|维吾尔族|土家族|彝族|蒙古族|藏族|布依族|侗族|瑶族|朝鲜族|白族|哈尼族|哈萨克族|黎族|傣族|畲族|傈僳族|仡佬族|东乡族|拉祜族|景颇族|佤族|水族|纳西族|羌族|土族|仫佬族|锡伯族|柯尔克孜族|达斡尔族|京族|布朗族|撒拉族|毛南族|阿昌族|普米族|鄂温克族|怒族|京族|基诺族|德昂族|保安族|俄罗斯族|裕固族|乌孜别克族|门巴族|鄂伦春族|独龙族|塔塔尔族|赫哲族|珞巴族高山族)，', report_content)
    if match:
        return match.group(1)
    return None

def extract_education_from_report(report_content):
    """从报告文本中提取学历。"""
    education_keywords = [
        "博士研究生", "硕士研究生", "研究生", "大学本科", "本科", "大学专科", "大专",
        "中专", "高中", "初中", "小学"
    ]
    for keyword in education_keywords:
        if keyword in report_content:
            return keyword
    return None

def extract_party_member_from_report(report_content):
    """从报告文本中提取是否中共党员信息。"""
    if "加入中国共产党" in report_content or "中共党员" in report_content or "共产党员" in report_content:
        return "是"
    if "非中共党员" in report_content:
        return "否"
    return None

def extract_party_joining_date_from_report(report_content):
    """从报告文本中提取入党时间。"""
    match = re.search(r'(\d{4}年\d{1,2}月)加入中国共产党', report_content)
    if match:
        return match.group(1).replace('年', '/').replace('月', '')
    return None

def validate_clue_data(df, app_config):
    """
    验证线索登记表中的数据一致性。
    """
    issues_list = []
    error_count = 0

    # 确保所有需要的列都存在
    required_columns = [
        app_config['COLUMN_MAPPINGS']['investigated_person'],
        app_config['COLUMN_MAPPINGS']['disposal_report'],
        app_config['COLUMN_MAPPINGS']['gender'],
        app_config['COLUMN_MAPPINGS']['age'],
        app_config['COLUMN_MAPPINGS']['birth_date'],
        app_config['COLUMN_MAPPINGS']['ethnicity'],
        app_config['COLUMN_MAPPINGS']['education'],
        app_config['COLUMN_MAPPINGS']['party_member'],
        app_config['COLUMN_MAPPINGS']['party_joining_date'],
        app_config['COLUMN_MAPPINGS']['organization_measure'],
        app_config['COLUMN_MAPPINGS']['acceptance_time'],
        app_config['COLUMN_MAPPINGS']['completion_time'],
        app_config['COLUMN_MAPPINGS']['disposal_method_1'],
        app_config['COLUMN_MAPPINGS']['accepted_clue_code']
    ]
    
    for col in required_columns:
        if col not in df.columns:
            issues_list.append({
                "受理线索编码": "N/A",
                "问题描述": f"缺少必要列: {col}",
                "风险等级": "高"
            })
            error_count += 1
            logger.error(f"缺少必要列: {col}")
            return issues_list, error_count # 如果缺少关键列，直接返回

    for index, row in df.iterrows():
        original_df_index = index # 记录原始DataFrame的索引
        
        investigated_person_excel = str(row[app_config['COLUMN_MAPPINGS']['investigated_person']]).strip()
        disposal_report_content = str(row[app_config['COLUMN_MAPPINGS']['disposal_report']]).strip()
        
        accepted_clue_code = str(row.get(app_config['COLUMN_MAPPINGS']['accepted_clue_code'], 'N/A')).strip()

        # 规则1: E2被反映人与AB2处置情况报告姓名不一致
        extracted_name = extract_name_from_report(disposal_report_content, investigated_person_excel)
        if investigated_person_excel and extracted_name and investigated_person_excel != extracted_name:
            issues_list.append({
                "受理线索编码": accepted_clue_code,
                "问题描述": app_config['VALIDATION_RULES']["inconsistent_name"],
                "风险等级": "高"
            })
            error_count += 1
            logger.warning(f"行 {original_df_index + 2} - 姓名不匹配: Excel '{investigated_person_excel}' vs 报告 '{extracted_name}'")
        elif investigated_person_excel and not extracted_name and disposal_report_content: # 报告有内容但未提取到姓名
             issues_list.append({
                "受理线索编码": accepted_clue_code,
                "问题描述": app_config['VALIDATION_RULES']["empty_report"],
                "风险等级": "高"
            })
             error_count += 1
             logger.warning(f"行 {original_df_index + 2} - 姓名不匹配: Excel '{investigated_person_excel}' vs 报告为空或未提取到姓名")


        # 规则2: 性别比对
        excel_gender = str(row.get(app_config['COLUMN_MAPPINGS']['gender'], '')).strip()
        extracted_gender = extract_gender_from_report(disposal_report_content)
        if excel_gender and extracted_gender and excel_gender != extracted_gender:
            issues_list.append({
                "受理线索编码": accepted_clue_code,
                "问题描述": f"行 {original_df_index + 2} - 性别不匹配: Excel '{excel_gender}' vs 报告 '{extracted_gender}'",
                "风险等级": "中"
            })
            error_count += 1
            logger.warning(f"行 {original_df_index + 2} - 性别不匹配: Excel '{excel_gender}' vs 报告 '{extracted_gender}'")
        elif excel_gender and not extracted_gender and disposal_report_content:
            issues_list.append({
                "受理线索编码": accepted_clue_code,
                "问题描述": f"行 {original_df_index + 2} - 性别不匹配: Excel '{excel_gender}' 有值，但报告中未提取到性别",
                "风险等级": "中"
            })
            error_count += 1
            logger.warning(f"行 {original_df_index + 2} - 性别不匹配: Excel '{excel_gender}' 有值，但报告中未提取到性别")


        # 规则3: 年龄比对 (根据出生年月和当前年份推算)
        excel_age = None
        if pd.notna(row.get(app_config['COLUMN_MAPPINGS']['age'])):
            try:
                excel_age = int(row.get(app_config['COLUMN_MAPPINGS']['age']))
            except ValueError:
                issues_list.append({
                    "受理线索编码": accepted_clue_code,
                    "问题描述": f"行 {original_df_index + 2} - 年龄字段格式不正确",
                    "风险等级": "高"
                })
                error_count += 1
                logger.warning(f"行 {original_df_index + 2} - Excel 年龄字段 '{row.get(app_config['COLUMN_MAPPINGS']['age'])}' 不是有效数字。")

        extracted_birth_date_str = extract_birth_date_from_report(disposal_report_content)
        if excel_age is not None and extracted_birth_date_str:
            try:
                extracted_birth_year = int(extracted_birth_date_str.split('/')[0])
                current_year = datetime.now().year
                calculated_age = current_year - extracted_birth_year
                # 允许1岁的误差
                if not (excel_age == calculated_age or excel_age == calculated_age - 1 or excel_age == calculated_age + 1):
                    issues_list.append({
                        "受理线索编码": accepted_clue_code,
                        "问题描述": f"行 {original_df_index + 2} - 年龄不匹配: Excel '{excel_age}' vs 报告推算 '{calculated_age}'",
                        "风险等级": "中"
                    })
                    error_count += 1
                    logger.warning(f"行 {original_df_index + 2} - 年龄不匹配: Excel '{excel_age}' vs 报告推算 '{calculated_age}'")
            except ValueError:
                logger.warning(f"行 {original_df_index + 2} - 从报告中提取的出生年份 '{extracted_birth_date_str}' 无法解析。")
        elif excel_age is not None and not extracted_birth_date_str and disposal_report_content:
            issues_list.append({
                "受理线索编码": accepted_clue_code,
                "问题描述": f"行 {original_df_index + 2} - 年龄有值但报告中未提取到出生年月，无法比对",
                "风险等级": "中"
            })
            error_count += 1
            logger.warning(f"行 {original_df_index + 2} - 年龄有值但报告中未提取到出生年月，无法比对")


        # 规则4: 出生年月比对
        excel_birth_date = str(row.get(app_config['COLUMN_MAPPINGS']['birth_date'], '')).strip()
        if excel_birth_date and extracted_birth_date_str and excel_birth_date != extracted_birth_date_str:
            issues_list.append({
                "受理线索编码": accepted_clue_code,
                "问题描述": f"行 {original_df_index + 2} - 出生年月不匹配: Excel '{excel_birth_date}' vs 报告 '{extracted_birth_date_str}'",
                "风险等级": "中"
            })
            error_count += 1
            logger.warning(f"行 {original_df_index + 2} - 出生年月不匹配: Excel '{excel_birth_date}' vs 报告 '{extracted_birth_date_str}'")
        elif excel_birth_date and not extracted_birth_date_str and disposal_report_content:
            issues_list.append({
                "受理线索编码": accepted_clue_code,
                "问题描述": f"行 {original_df_index + 2} - 出生年月有值但报告中未提取到出生年月，无法比对",
                "风险等级": "中"
            })
            error_count += 1
            logger.warning(f"行 {original_df_index + 2} - 出生年月有值但报告中未提取到出生年月，无法比对")


        # 规则5: 学历比对
        excel_education = str(row.get(app_config['COLUMN_MAPPINGS']['education'], '')).strip()
        extracted_education = extract_education_from_report(disposal_report_content)
        if excel_education and extracted_education and excel_education not in extracted_education and extracted_education not in excel_education:
            issues_list.append({
                "受理线索编码": accepted_clue_code,
                "问题描述": f"行 {original_df_index + 2} - 学历不匹配: Excel '{excel_education}' vs 报告 '{extracted_education}'",
                "风险等级": "中"
            })
            error_count += 1
            logger.warning(f"行 {original_df_index + 2} - 学历不匹配: Excel '{excel_education}' vs 报告 '{extracted_education}'")
        elif excel_education and not extracted_education and disposal_report_content:
            issues_list.append({
                "受理线索编码": accepted_clue_code,
                "问题描述": f"行 {original_df_index + 2} - 学历有值但报告中未提取到学历，无法比对",
                "风险等级": "中"
            })
            error_count += 1
            logger.warning(f"行 {original_df_index + 2} - 学历有值但报告中未提取到学历，无法比对")


        # 规则6: 民族比对
        excel_ethnicity = str(row.get(app_config['COLUMN_MAPPINGS']['ethnicity'], '')).strip()
        extracted_ethnicity = extract_ethnicity_from_report(disposal_report_content)
        if excel_ethnicity and extracted_ethnicity and excel_ethnicity != extracted_ethnicity:
            issues_list.append({
                "受理线索编码": accepted_clue_code,
                "问题描述": f"行 {original_df_index + 2} - 民族不匹配: Excel '{excel_ethnicity}' vs 报告 '{extracted_ethnicity}'",
                "风险等级": "中"
            })
            error_count += 1
            logger.warning(f"行 {original_df_index + 2} - 民族不匹配: Excel '{excel_ethnicity}' vs 报告 '{extracted_ethnicity}'")
        elif excel_ethnicity and not extracted_ethnicity and disposal_report_content:
            issues_list.append({
                "受理线索编码": accepted_clue_code,
                "问题描述": f"行 {original_df_index + 2} - 民族有值但报告中未提取到民族，无法比对",
                "风险等级": "中"
            })
            error_count += 1
            logger.warning(f"行 {original_df_index + 2} - 民族有值但报告中未提取到民族，无法比对")


        # 规则7: 是否中共党员比对
        excel_party_member = str(row.get(app_config['COLUMN_MAPPINGS']['party_member'], '')).strip()
        extracted_party_member = extract_party_member_from_report(disposal_report_content)
        if excel_party_member and extracted_party_member and excel_party_member != extracted_party_member:
            issues_list.append({
                "受理线索编码": accepted_clue_code,
                "问题描述": f"行 {original_df_index + 2} - 是否中共党员不匹配: Excel '{excel_party_member}' vs 报告 '{extracted_party_member}'",
                "风险等级": "中"
            })
            error_count += 1
            logger.warning(f"行 {original_df_index + 2} - 是否中共党员不匹配: Excel '{excel_party_member}' vs 报告 '{extracted_party_member}'")
        elif excel_party_member and not extracted_party_member and disposal_report_content:
            issues_list.append({
                "受理线索编码": accepted_clue_code,
                "问题描述": f"行 {original_df_index + 2} - 是否中共党员有值但报告中未提取到，无法比对",
                "风险等级": "中"
            })
            error_count += 1
            logger.warning(f"行 {original_df_index + 2} - 是否中共党员有值但报告中未提取到，无法比对")


        # 规则8: 入党时间比对 (仅当Excel中“是否中共党员”为“是”时才比对)
        excel_party_joining_date = str(row.get(app_config['COLUMN_MAPPINGS']['party_joining_date'], '')).strip()
        extracted_party_joining_date = extract_party_joining_date_from_report(disposal_report_content)
        
        if excel_party_member == '是':
            if excel_party_joining_date and extracted_party_joining_date and excel_party_joining_date != extracted_party_joining_date:
                issues_list.append({
                    "受理线索编码": accepted_clue_code,
                    "问题描述": f"行 {original_df_index + 2} - 入党时间不匹配: Excel '{excel_party_joining_date}' vs 报告 '{extracted_party_joining_date}'",
                    "风险等级": "中"
                })
                error_count += 1
                logger.warning(f"行 {original_df_index + 2} - 入党时间不匹配: Excel '{excel_party_joining_date}' vs 报告 '{extracted_party_joining_date}'")
            elif excel_party_joining_date and not extracted_party_joining_date and disposal_report_content:
                issues_list.append({
                    "受理线索编码": accepted_clue_code,
                    "问题描述": f"行 {original_df_index + 2} - 入党时间有值但报告中未提取到，无法比对",
                    "风险等级": "中"
                })
                error_count += 1
                logger.warning(f"行 {original_df_index + 2} - 入党时间有值但报告中未提取到，无法比对")
        elif excel_party_member == '否' and excel_party_joining_date:
            issues_list.append({
                "受理线索编码": accepted_clue_code,
                "问题描述": f"行 {original_df_index + 2} - Excel是否中共党员为“否”，但入党时间字段不为空 ('{excel_party_joining_date}')。",
                "风险等级": "低"
            })
            error_count += 1
            logger.warning(f"行 {original_df_index + 2} - Excel是否中共党员为“否”，但入党时间字段不为空 ('{excel_party_joining_date}')。")


        # 规则9: 组织措施与处置情况报告比对
        excel_organization_measure = str(row.get(app_config['COLUMN_MAPPINGS']['organization_measure'], '')).strip()
        
        # 使用Config中的关键词列表
        organization_measure_keywords = app_config['ORGANIZATION_MEASURE_KEYWORDS']
        
        found_match = False
        if excel_organization_measure:
            for keyword in organization_measure_keywords:
                if keyword in disposal_report_content:
                    found_match = True
                    break
            if not found_match:
                issues_list.append({
                    "受理线索编码": accepted_clue_code,
                    "问题描述": f"行 {original_df_index + 2} - 组织措施 ('{excel_organization_measure}') 与处置情况报告不一致，报告中未提及相关关键词。",
                    "风险等级": "中"
                })
                error_count += 1
                logger.warning(f"行 {original_df_index + 2} - 组织措施 ('{excel_organization_measure}') 与处置情况报告不一致。")
        
        # 规则10: 受理时间与处置情况报告落款时间比对
        excel_acceptance_time = row.get(app_config['COLUMN_MAPPINGS']['acceptance_time'])
        
        if pd.notna(excel_acceptance_time) and disposal_report_content:
            lines = disposal_report_content.strip().split('\n')
            report_date = None
            if lines:
                # 尝试从最后一行提取日期
                last_line = lines[-1].strip()
                match = re.search(r'(\d{4}年\d{1,2}月\d{1,2}日)', last_line)
                if match:
                    try:
                        report_date_str = match.group(1)
                        report_date = datetime.strptime(report_date_str, '%Y年%m月%d日').date()
                    except ValueError:
                        pass # 继续尝试其他行

            if report_date:
                excel_date_obj = None
                if isinstance(excel_acceptance_time, datetime):
                    excel_date_obj = excel_acceptance_time.date()
                elif isinstance(excel_acceptance_time, str):
                    try:
                        excel_date_obj = pd.to_datetime(excel_acceptance_time).date()
                    except ValueError:
                        pass

                if excel_date_obj and excel_date_obj > report_date:
                    issues_list.append({
                        "受理线索编码": accepted_clue_code,
                        "问题描述": f"行 {original_df_index + 2} - 受理时间 ('{excel_acceptance_time}') 晚于处置情况报告落款时间 ('{report_date}')。",
                        "风险等级": "高"
                    })
                    error_count += 1
                    logger.warning(f"行 {original_df_index + 2} - 受理时间晚于处置情况报告落款时间。")
            else:
                issues_list.append({
                    "受理线索编码": accepted_clue_code,
                    "问题描述": f"行 {original_df_index + 2} - 处置情况报告中未能提取到有效的落款时间。",
                    "风险等级": "中"
                })
                error_count += 1
                logger.warning(f"行 {original_df_index + 2} - 处置情况报告中未能提取到有效的落款时间。")
        elif pd.notna(excel_acceptance_time) and not disposal_report_content:
            issues_list.append({
                "受理线索编码": accepted_clue_code,
                "问题描述": f"行 {original_df_index + 2} - 受理时间有值但处置情况报告为空，无法比对。",
                "风险等级": "中"
            })
            error_count += 1
            logger.warning(f"行 {original_df_index + 2} - 受理时间有值但处置情况报告为空，无法比对。")
        
        # 规则11: 办结时间与处置情况报告落款时间比对
        excel_completion_time = row.get(app_config['COLUMN_MAPPINGS']['completion_time'])
        
        if pd.notna(excel_completion_time) and disposal_report_content:
            lines = disposal_report_content.strip().split('\n')
            report_date = None
            if lines:
                last_line = lines[-1].strip()
                match = re.search(r'(\d{4}年\d{1,2}月\d{1,2}日)', last_line)
                if match:
                    try:
                        report_date_str = match.group(1)
                        report_date = datetime.strptime(report_date_str, '%Y年%m月%d日').date()
                    except ValueError:
                        pass

            if report_date:
                excel_date_obj = None
                if isinstance(excel_completion_time, datetime):
                    excel_date_obj = excel_completion_time.date()
                elif isinstance(excel_completion_time, str):
                    try:
                        excel_date_obj = pd.to_datetime(excel_completion_time).date()
                    except ValueError:
                        pass

                if excel_date_obj and excel_date_obj != report_date:
                    issues_list.append({
                        "受理线索编码": accepted_clue_code,
                        "问题描述": f"行 {original_df_index + 2} - 办结时间 ('{excel_completion_time}') 与处置情况报告落款时间 ('{report_date}') 不一致。",
                        "风险等级": "高"
                    })
                    error_count += 1
                    logger.warning(f"行 {original_df_index + 2} - 办结时间与处置情况报告落款时间不一致。")
            else:
                issues_list.append({
                    "受理线索编码": accepted_clue_code,
                    "问题描述": f"行 {original_df_index + 2} - 处置情况报告中未能提取到有效的落款时间。",
                    "风险等级": "中"
                })
                error_count += 1
                logger.warning(f"行 {original_df_index + 2} - 处置情况报告中未能提取到有效的落款时间。")
        elif pd.notna(excel_completion_time) and not disposal_report_content:
            issues_list.append({
                "受理线索编码": accepted_clue_code,
                "问题描述": f"行 {original_df_index + 2} - 办结时间有值但处置情况报告为空，无法比对。",
                "风险等级": "中"
            })
            error_count += 1
            logger.warning(f"行 {original_df_index + 2} - 办结时间有值但处置情况报告为空，无法比对。")

        # 规则12: 处置方式1二级与处置情况报告比对
        excel_disposal_method_1 = str(row.get(app_config['COLUMN_MAPPINGS']['disposal_method_1'], '')).strip()
        
        # 假设处置方式1二级的值会出现在处置情况报告中
        if excel_disposal_method_1 and excel_disposal_method_1 not in disposal_report_content:
            issues_list.append({
                "受理线索编码": accepted_clue_code,
                "问题描述": f"行 {original_df_index + 2} - 处置方式1二级 ('{excel_disposal_method_1}') 未在处置情况报告中提及。",
                "风险等级": "中"
            })
            error_count += 1
            logger.warning(f"行 {original_df_index + 2} - 处置方式1二级 ('{excel_disposal_method_1}') 未在处置情况报告中提及。")
        
    return issues_list, error_count
