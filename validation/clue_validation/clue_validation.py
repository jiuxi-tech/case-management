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
    """从报告文本中提取出生年月。
    
    从"（一）被反映人基本情况"段落中提取类似"1966年12月生"的出生年月信息。
    """
    # 首先查找"（一）被反映人基本情况"段落
    basic_info_match = re.search(r'（一）被反映人基本情况[\s\S]*?(?=（二）|$)', report_content)
    if basic_info_match:
        basic_info_section = basic_info_match.group(0)
        # 在基本情况段落中查找"XXXX年XX月生"格式的出生年月
        birth_match = re.search(r'(\d{4}年\d{1,2}月)生', basic_info_section)
        if birth_match:
            # 转换为标准日期格式 YYYY/MM
            birth_date = birth_match.group(1).replace('年', '/').replace('月', '')
            return birth_date
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

def extract_party_joining_date_from_report(report_content):
    """从报告文本中提取入党时间。"""
    match = re.search(r'(\d{4}年\d{1,2}月)加入中国共产党', report_content)
    if match:
        return match.group(1).replace('年', '/').replace('月', '')
    return None

def validate_clue_data(df, app_config, agency_mapping_db):
    """
    验证线索登记表中的数据一致性。
    """
    issues_list = []
    error_count = 0

    # 确保所有需要的列都存在
    required_columns = [
        app_config['COLUMN_MAPPINGS']['mentioned_person'],
        app_config['COLUMN_MAPPINGS']['disposal_report'],
        app_config['COLUMN_MAPPINGS']['birth_date'],
        app_config['COLUMN_MAPPINGS']['ethnicity'],

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

            })
            error_count += 1
            logger.error(f"缺少必要列: {col}")
            return issues_list, error_count # 如果缺少关键列，直接返回

    for index, row in df.iterrows():
        original_df_index = index # 记录原始DataFrame的索引
        
        investigated_person_excel = str(row.get(app_config['COLUMN_MAPPINGS']['mentioned_person'], '')).strip()
        disposal_report_content = str(row.get(app_config['COLUMN_MAPPINGS']['disposal_report'], '')).strip()
        
        # 处理nan值
        if disposal_report_content.lower() == 'nan':
            disposal_report_content = ''
        if investigated_person_excel.lower() == 'nan':
            investigated_person_excel = ''
        
        accepted_clue_code = str(row.get(app_config['COLUMN_MAPPINGS']['accepted_clue_code'], 'N/A')).strip()
        accepted_personnel_code = str(row.get(app_config['COLUMN_MAPPINGS']['accepted_personnel_code'], 'N/A')).strip()

        # 规则1: 填报单位名称与办理机关不一致 (统一处理)
        reporting_agency_excel = str(row.get(app_config['COLUMN_MAPPINGS']['reporting_agency'], '')).strip()
        authority_excel = str(row.get(app_config['COLUMN_MAPPINGS']['authority'], '')).strip()

        if reporting_agency_excel and authority_excel:
            # 检查是否存在匹配的记录
            match_found = False
            for mapping in agency_mapping_db:
                if mapping['authority'] == authority_excel and mapping['agency'] == reporting_agency_excel:
                    match_found = True
                    break
                    
            if not match_found:
                # 构建比对字段和被比对字段的描述
                compared_field = f"C{original_df_index + 2}填报单位名称"
                being_compared_field = f"H{original_df_index + 2}办理机关"
                issues_list.append({
                    "受理线索编码": accepted_clue_code,
                    "受理人员编码": accepted_personnel_code,
                    "行号": original_df_index + 2,
                    "比对字段": compared_field,
                    "被比对字段": being_compared_field,
                    "问题描述": f"C{original_df_index + 2}填报单位名称与H{original_df_index + 2}办理机关不一致",
                    "列名": app_config['COLUMN_MAPPINGS']['reporting_agency'] # 添加列名用于标红
                })
                error_count += 1
                logger.warning(
                    f"<线索 - （1.填报单位名称）> - 行 {original_df_index + 2} - 填报单位名称 '{reporting_agency_excel}' (len: {len(reporting_agency_excel)}) 与办理机关 '{authority_excel}' (len: {len(authority_excel)}) 不一致，且不在数据库映射中。数据库查询语句为：SELECT authority, agency FROM authority_agency_dict WHERE category = 'NSL' AND authority = '{authority_excel}' AND agency = '{reporting_agency_excel}'")

        # 规则2: E2被反映人与AB2处置情况报告姓名不一致
        extracted_name = extract_name_from_report(disposal_report_content, investigated_person_excel)
        if investigated_person_excel and extracted_name and investigated_person_excel != extracted_name:
            # 构建比对字段和被比对字段的描述
            compared_field = f"E{original_df_index + 2}被反映人"
            being_compared_field = f"AB{original_df_index + 2}处置情况报告"
            issues_list.append({
                "受理线索编码": accepted_clue_code,
                "受理人员编码": accepted_personnel_code,
                "行号": original_df_index + 2,
                "比对字段": compared_field,
                "被比对字段": being_compared_field,
                "问题描述": f"E{original_df_index + 2}被反映人与AB{original_df_index + 2}处置情况报告姓名不一致",
                "列名": app_config['COLUMN_MAPPINGS']['mentioned_person'] # 添加列名用于标红
            })
            error_count += 1
            logger.warning(f"<线索 - （2.被反映人）> - 行 {original_df_index + 2} - 被反映人 '{investigated_person_excel}' 与 处置情况报告的姓名（{extracted_name}）不一致。")
        elif investigated_person_excel and not extracted_name and disposal_report_content: # 报告有内容但未提取到姓名
            # 构建比对字段和被比对字段的描述
            compared_field = f"E{original_df_index + 2}被反映人"
            being_compared_field = f"AB{original_df_index + 2}处置情况报告"
            issues_list.append({
                "受理线索编码": accepted_clue_code,
                "受理人员编码": accepted_personnel_code,
                "行号": original_df_index + 2,
                "比对字段": compared_field,
                "被比对字段": being_compared_field,
                "问题描述": f"E{original_df_index + 2}被反映人与AB{original_df_index + 2}处置情况报告姓名不一致 (报告为空)",
                "列名": app_config['COLUMN_MAPPINGS']['mentioned_person'] # 添加列名用于标红
            })
            error_count += 1
            logger.warning(f"<线索 - （2.被反映人）> - 行 {original_df_index + 2} - 被反映人 '{investigated_person_excel}' 与 处置情况报告的姓名为空或未提取到。")

        # 规则3: 收缴金额（万元）检查
        if "收缴金额（万元）" in df.columns and disposal_report_content and "收缴" in disposal_report_content:
            # 构建比对字段和被比对字段的描述
            compared_field = f"Q{original_df_index + 2}收缴金额（万元）"
            being_compared_field = f"AB{original_df_index + 2}处置情况报告"
            issues_list.append({
                "受理线索编码": accepted_clue_code,
                "受理人员编码": accepted_personnel_code,
                "行号": original_df_index + 2,
                "比对字段": compared_field,
                "被比对字段": being_compared_field,
                "问题描述": f"Q{original_df_index + 2}收缴金额（万元）与AB{original_df_index + 2}处置情况报告对比结果是AB{original_df_index + 2}处置情况报告出现收缴二字",
                "列名": "收缴金额（万元）" # 添加列名用于标黄
            })
            error_count += 1
            logger.warning(f"<线索 - （3.收缴金额（万元））> - 行 {original_df_index + 2} - 处置情况报告出现【收缴】二字。")

        # 规则4: 没收金额检查
        if "没收金额" in df.columns and disposal_report_content and "没收" in disposal_report_content:
            # 构建比对字段和被比对字段的描述
            compared_field = f"R{original_df_index + 2}没收金额"
            being_compared_field = f"AB{original_df_index + 2}处置情况报告"
            issues_list.append({
                "受理线索编码": accepted_clue_code,
                "受理人员编码": accepted_personnel_code,
                "行号": original_df_index + 2,
                "比对字段": compared_field,
                "被比对字段": being_compared_field,
                "问题描述": f"R{original_df_index + 2}没收金额与AB{original_df_index + 2}处置情况报告对比结果是AB{original_df_index + 2}处置情况报告出现没收二字",
                "列名": "没收金额" # 添加列名用于标黄
            })
            error_count += 1
            logger.warning(f"<线索 - （4.没收金额）> - 行 {original_df_index + 2} - 处置情况报告出现【没收】二字。")

        # 规则5: 责令退赔金额检查
        if "责令退赔金额" in df.columns and disposal_report_content and "责令退赔" in disposal_report_content:
            # 构建比对字段和被比对字段的描述
            compared_field = f"S{original_df_index + 2}责令退赔金额"
            being_compared_field = f"AB{original_df_index + 2}处置情况报告"
            issues_list.append({
                "受理线索编码": accepted_clue_code,
                "受理人员编码": accepted_personnel_code,
                "行号": original_df_index + 2,
                "比对字段": compared_field,
                "被比对字段": being_compared_field,
                "问题描述": f"S{original_df_index + 2}责令退赔金额与AB{original_df_index + 2}处置情况报告对比结果是AB{original_df_index + 2}处置情况报告出现责令退赔字样",
                "列名": "责令退赔金额" # 添加列名用于标黄
            })
            error_count += 1
            logger.warning(f"<线索 - （5.责令退赔金额）> - 行 {original_df_index + 2} - 处置情况报告出现【责令退赔】字样。")

        # 规则6: 登记上交金额检查
        if "登记上交金额" in df.columns and disposal_report_content and "登记上交金额" in disposal_report_content:
            # 构建比对字段和被比对字段的描述
            compared_field = f"T{original_df_index + 2}登记上交金额"
            being_compared_field = f"AB{original_df_index + 2}处置情况报告"
            issues_list.append({
                "受理线索编码": accepted_clue_code,
                "受理人员编码": accepted_personnel_code,
                "行号": original_df_index + 2,
                "比对字段": compared_field,
                "被比对字段": being_compared_field,
                "问题描述": f"T{original_df_index + 2}登记上交金额与AB{original_df_index + 2}处置情况报告对比结果是AB{original_df_index + 2}处置情况报告出现登记上交金额字样",
                "列名": "登记上交金额" # 添加列名用于标黄
            })
            error_count += 1
            logger.warning(f"<线索 - （6.登记上交金额）> - 行 {original_df_index + 2} - 处置情况报告出现【登记上交金额】字样。")

        # 规则7: 追缴失职渎职滥用职权造成的损失金额检查
        if "追缴失职渎职滥用职权造成的损失金额" in df.columns and disposal_report_content and "追缴" in disposal_report_content:
            # 构建比对字段和被比对字段的描述
            compared_field = f"U{original_df_index + 2}追缴失职渎职滥用职权造成的损失金额"
            being_compared_field = f"AB{original_df_index + 2}处置情况报告"
            issues_list.append({
                "受理线索编码": accepted_clue_code,
                "受理人员编码": accepted_personnel_code,
                "行号": original_df_index + 2,
                "比对字段": compared_field,
                "被比对字段": being_compared_field,
                "问题描述": f"U{original_df_index + 2}追缴失职渎职滥用职权造成的损失金额与AB{original_df_index + 2}处置情况报告对比结果是AB{original_df_index + 2}处置情况报告出现追缴字样",
                "列名": "追缴失职渎职滥用职权造成的损失金额" # 添加列名用于标黄
            })
            error_count += 1
            logger.warning(f"<线索 - （7.追缴失职渎职滥用职权造成的损失金额）> - 行 {original_df_index + 2} - 处置情况报告出现【追缴】字样。")

        # 规则8: 民族比对
        excel_ethnicity = str(row.get(app_config['COLUMN_MAPPINGS']['ethnicity'], '')).strip()
        extracted_ethnicity = extract_ethnicity_from_report(disposal_report_content)
        if excel_ethnicity and extracted_ethnicity and excel_ethnicity != extracted_ethnicity:
            # 构建比对字段和被比对字段的描述
            compared_field = f"W{original_df_index + 2}民族"
            being_compared_field = f"AB{original_df_index + 2}处置情况报告"
            issues_list.append({
                "受理线索编码": accepted_clue_code,
                "受理人员编码": accepted_personnel_code,
                "行号": original_df_index + 2,
                "比对字段": compared_field,
                "被比对字段": being_compared_field,
                "问题描述": f"W{original_df_index + 2}民族与AB{original_df_index + 2}处置情况报告民族不一致",
                "列名": app_config['COLUMN_MAPPINGS']['ethnicity']
            })
            error_count += 1
            logger.warning(f"<线索 - （8.民族）> - 行 {original_df_index + 2} - 民族不匹配: Excel '{excel_ethnicity}' vs 报告 '{extracted_ethnicity}'")
        elif excel_ethnicity and not extracted_ethnicity and disposal_report_content:
            # 构建比对字段和被比对字段的描述
            compared_field = f"W{original_df_index + 2}民族"
            being_compared_field = f"AB{original_df_index + 2}处置情况报告"
            issues_list.append({
                "受理线索编码": accepted_clue_code,
                "受理人员编码": accepted_personnel_code,
                "行号": original_df_index + 2,
                "比对字段": compared_field,
                "被比对字段": being_compared_field,
                "问题描述": f"W{original_df_index + 2}民族有值但AB{original_df_index + 2}处置情况报告中未提取到民族，无法比对",
                "列名": app_config['COLUMN_MAPPINGS']['ethnicity']
            })
            error_count += 1
            logger.warning(f"<线索 - （8.民族）> - 行 {original_df_index + 2} - 民族有值但报告中未提取到民族，无法比对")

        # 规则9: 出生年月比对
        excel_birth_date = str(row.get(app_config['COLUMN_MAPPINGS']['birth_date'], '')).strip()
        extracted_birth_date_str = extract_birth_date_from_report(disposal_report_content)
        if excel_birth_date and extracted_birth_date_str and excel_birth_date != extracted_birth_date_str:
            # 构建比对字段和被比对字段的描述
            compared_field = f"X{original_df_index + 2}出生年月"
            being_compared_field = f"AB{original_df_index + 2}处置情况报告"
            issues_list.append({
                "受理线索编码": accepted_clue_code,
                "受理人员编码": accepted_personnel_code,
                "行号": original_df_index + 2,
                "比对字段": compared_field,
                "被比对字段": being_compared_field,
                "问题描述": f"X{original_df_index + 2}出生年月与AB{original_df_index + 2}处置情况报告的出生年月不一致",
                "列名": app_config['COLUMN_MAPPINGS']['birth_date']
            })
            error_count += 1
            logger.warning(f"<线索 - （9.出生年月）> - 行 {original_df_index + 2} - 出生年月不匹配: Excel '{excel_birth_date}' vs 报告 '{extracted_birth_date_str}'")
        elif excel_birth_date and not extracted_birth_date_str and disposal_report_content:
            # 构建比对字段和被比对字段的描述
            compared_field = f"X{original_df_index + 2}出生年月"
            being_compared_field = f"AB{original_df_index + 2}处置情况报告"
            issues_list.append({
                "受理线索编码": accepted_clue_code,
                "受理人员编码": accepted_personnel_code,
                "行号": original_df_index + 2,
                "比对字段": compared_field,
                "被比对字段": being_compared_field,
                "问题描述": f"X{original_df_index + 2}出生年月有值但AB{original_df_index + 2}处置情况报告中未提取到出生年月，无法比对",
                "列名": app_config['COLUMN_MAPPINGS']['birth_date']
            })
            error_count += 1
            logger.warning(f"<线索 - （9.出生年月）> - 行 {original_df_index + 2} - 出生年月有值但报告中未提取到出生年月，无法比对")


        # 规则10: 入党时间比对
        excel_party_joining_date = str(row.get(app_config['COLUMN_MAPPINGS']['party_joining_date'], '')).strip()
        extracted_party_joining_date = extract_party_joining_date_from_report(disposal_report_content)
        
        # 构建字段信息
        compared_field = f"AC{original_df_index + 2}入党时间"
        being_compared_field = f"AB{original_df_index + 2}处置情况报告的入党时间"
        
        # 使用智能日期比较，避免格式差异导致的误判（如'1990/01' vs '1990/1'）
        def normalize_date_format(date_str):
            """标准化日期格式，将'1990/1'转换为'1990/01'"""
            if not date_str:
                return date_str
            # 匹配YYYY/M或YYYY/MM格式
            match = re.match(r'(\d{4})/(\d{1,2})$', str(date_str))
            if match:
                year, month = match.groups()
                return f"{year}/{month.zfill(2)}"
            return str(date_str)
        
        normalized_excel_date = normalize_date_format(excel_party_joining_date)
        normalized_extracted_date = normalize_date_format(extracted_party_joining_date)
        
        if excel_party_joining_date and extracted_party_joining_date and normalized_excel_date != normalized_extracted_date:
            issues_list.append({
                "受理线索编码": accepted_clue_code,
                "受理人员编码": accepted_personnel_code,
                "行号": original_df_index + 2,
                "比对字段": compared_field,
                "被比对字段": being_compared_field,
                "问题描述": f"AC{original_df_index + 2}入党时间与AB{original_df_index + 2}处置情况报告的入党时间不一致",
                "列名": app_config['COLUMN_MAPPINGS']['party_joining_date']
            })
            error_count += 1
            logger.warning(f"<线索 - （10.入党时间）> - 行 {original_df_index + 2} - 入党时间不匹配: Excel '{excel_party_joining_date}' vs 报告 '{extracted_party_joining_date}'")
        elif excel_party_joining_date and not extracted_party_joining_date and disposal_report_content:
            issues_list.append({
                "受理线索编码": accepted_clue_code,
                "受理人员编码": accepted_personnel_code,
                "行号": original_df_index + 2,
                "比对字段": compared_field,
                "被比对字段": being_compared_field,
                "问题描述": f"AC{original_df_index + 2}入党时间有值但AB{original_df_index + 2}处置情况报告中未提取到入党时间，无法比对",
                "列名": app_config['COLUMN_MAPPINGS']['party_joining_date']
            })
            error_count += 1
            logger.warning(f"<线索 - （10.入党时间）> - 行 {original_df_index + 2} - 入党时间有值但报告中未提取到入党时间，无法比对")



        # 规则11: 办结时间与处置情况报告落款时间比对
        excel_completion_time = row.get(app_config['COLUMN_MAPPINGS']['completion_time'])
        
        # 构建字段信息
        compared_field = f"BT{original_df_index + 2}办结时间"
        being_compared_field = f"AB{original_df_index + 2}处置情况报告落款时间"
        
        if pd.notna(excel_completion_time) and disposal_report_content:
            lines = disposal_report_content.strip().split('\n')
            report_date = None
            
            # 查找"核查组成员签字"段落，并提取其上一行的日期
            for i, line in enumerate(lines):
                if "核查组成员签字" in line and i > 0:
                    # 获取上一行内容
                    prev_line = lines[i-1].strip()
                    # 匹配日期格式：YYYY年M月D日 或 YYYY年MM月DD日
                    match = re.search(r'(\d{4})年(\d{1,2})月(\d{1,2})日', prev_line)
                    if match:
                        try:
                            year, month, day = match.groups()
                            report_date_str = f"{year}年{month}月{day}日"
                            report_date = datetime.strptime(report_date_str, '%Y年%m月%d日').date()
                            break
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
                        "受理人员编码": accepted_personnel_code,
                        "行号": original_df_index + 2,
                        "比对字段": compared_field,
                        "被比对字段": being_compared_field,
                        "问题描述": f"BT{original_df_index + 2}办结时间与AB{original_df_index + 2}处置情况报告落款时间不一致",
                        "列名": app_config['COLUMN_MAPPINGS']['completion_time']
                    })
                    error_count += 1
                    logger.warning(f"<线索 - （11.办结时间）> - 行 {original_df_index + 2} - 办结时间不匹配: Excel '{excel_completion_time}' vs 报告落款时间 '{report_date}'")
            else:
                issues_list.append({
                    "受理线索编码": accepted_clue_code,
                    "受理人员编码": accepted_personnel_code,
                    "行号": original_df_index + 2,
                    "比对字段": compared_field,
                    "被比对字段": being_compared_field,
                    "问题描述": f"BT{original_df_index + 2}办结时间有值但AB{original_df_index + 2}处置情况报告中未能提取到有效的落款时间，无法比对",
                    "列名": app_config['COLUMN_MAPPINGS']['completion_time']
                })
                error_count += 1
                logger.warning(f"<线索 - （11.办结时间）> - 行 {original_df_index + 2} - 处置情况报告中未能提取到有效的落款时间")
        elif pd.notna(excel_completion_time) and not disposal_report_content:
            issues_list.append({
                "受理线索编码": accepted_clue_code,
                "受理人员编码": accepted_personnel_code,
                "行号": original_df_index + 2,
                "比对字段": compared_field,
                "被比对字段": being_compared_field,
                "问题描述": f"BT{original_df_index + 2}办结时间有值但AB{original_df_index + 2}处置情况报告为空，无法比对",
                "列名": app_config['COLUMN_MAPPINGS']['completion_time']
            })
            error_count += 1
            logger.warning(f"<线索 - （11.办结时间）> - 行 {original_df_index + 2} - 办结时间有值但处置情况报告为空，无法比对")

        # 规则12: 组织措施与处置情况报告比对
        excel_organization_measure = str(row.get(app_config['COLUMN_MAPPINGS']['organization_measure'], '')).strip()
        
        # 处理nan值
        if excel_organization_measure.lower() == 'nan':
            excel_organization_measure = ''
        
        # 构建字段信息
        compared_field = f"CC{original_df_index + 2}组织措施"
        being_compared_field = f"AB{original_df_index + 2}处置情况报告的组织措施"
        
        # 使用Config中的关键词列表
        organization_measure_keywords = app_config['ORGANIZATION_MEASURE_KEYWORDS']
        
        if excel_organization_measure and disposal_report_content:
            # 检查处置报告中是否包含任何组织措施关键词
            report_contains_keyword = False
            matched_keyword = None
            
            for keyword in organization_measure_keywords:
                if keyword in disposal_report_content:
                    report_contains_keyword = True
                    matched_keyword = keyword
                    break
            
            # 如果处置报告中不包含任何组织措施关键词，或者与Excel中的组织措施不一致，则标红
            excel_contains_keyword = any(keyword in excel_organization_measure for keyword in organization_measure_keywords)
            
            if not report_contains_keyword or not excel_contains_keyword or (matched_keyword and matched_keyword not in excel_organization_measure):
                issues_list.append({
                    "受理线索编码": accepted_clue_code,
                    "受理人员编码": accepted_personnel_code,
                    "行号": original_df_index + 2,
                    "比对字段": compared_field,
                    "被比对字段": being_compared_field,
                    "问题描述": f"CC{original_df_index + 2}组织措施与AB{original_df_index + 2}处置情况报告的组织措施不一致",
                    "列名": app_config['COLUMN_MAPPINGS']['organization_measure']
                })
                error_count += 1
                if not report_contains_keyword:
                    logger.warning(f"<线索 - （12.组织措施）> - 行 {original_df_index + 2} - 处置情况报告中未找到组织措施关键词")
                elif not excel_contains_keyword:
                    logger.warning(f"<线索 - （12.组织措施）> - 行 {original_df_index + 2} - Excel组织措施'{excel_organization_measure}'不包含预定义关键词")
                else:
                    logger.warning(f"<线索 - （12.组织措施）> - 行 {original_df_index + 2} - 组织措施不一致: Excel '{excel_organization_measure}' vs 报告 '{matched_keyword}'")
        elif excel_organization_measure and not disposal_report_content:
            issues_list.append({
                "受理线索编码": accepted_clue_code,
                "受理人员编码": accepted_personnel_code,
                "行号": original_df_index + 2,
                "比对字段": compared_field,
                "被比对字段": being_compared_field,
                "问题描述": f"CC{original_df_index + 2}组织措施有值但AB{original_df_index + 2}处置情况报告为空，无法比对",
                "列名": app_config['COLUMN_MAPPINGS']['organization_measure']
            })
            error_count += 1
            logger.warning(f"<线索 - （12.组织措施）> - 行 {original_df_index + 2} - 组织措施有值但处置情况报告为空，无法比对")
        elif not excel_organization_measure and disposal_report_content:
            # 检查处置报告中是否包含组织措施关键词，但Excel组织措施字段为空
            report_contains_keyword = False
            matched_keyword = None
            
            for keyword in organization_measure_keywords:
                if keyword in disposal_report_content:
                    report_contains_keyword = True
                    matched_keyword = keyword
                    break
            
            if report_contains_keyword:
                issues_list.append({
                    "受理线索编码": accepted_clue_code,
                    "受理人员编码": accepted_personnel_code,
                    "行号": original_df_index + 2,
                    "比对字段": compared_field,
                    "被比对字段": being_compared_field,
                    "问题描述": f"CC{original_df_index + 2}组织措施与AB{original_df_index + 2}处置情况报告的组织措施不一致",
                    "列名": app_config['COLUMN_MAPPINGS']['organization_measure']
                })
                error_count += 1
                logger.warning(f"<线索 - （12.组织措施）> - 行 {original_df_index + 2} - 组织措施字段为空但处置情况报告包含关键词'{matched_keyword}'")

        # 规则13: 受理时间与处置情况报告落款时间比对
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
       
                    })
                    error_count += 1
                    logger.warning(f"行 {original_df_index + 2} - 受理时间晚于处置情况报告落款时间。")
            else:
                issues_list.append({
                    "受理线索编码": accepted_clue_code,
                    "问题描述": f"行 {original_df_index + 2} - 处置情况报告中未能提取到有效的落款时间。",
    

                })
                error_count += 1
                logger.warning(f"行 {original_df_index + 2} - 处置情况报告中未能提取到有效的落款时间。")
        elif pd.notna(excel_acceptance_time) and not disposal_report_content:
            issues_list.append({
                "受理线索编码": accepted_clue_code,
                "问题描述": f"行 {original_df_index + 2} - 受理时间有值但处置情况报告为空，无法比对。",
            })
            error_count += 1
            logger.warning(f"行 {original_df_index + 2} - 受理时间有值但处置情况报告为空，无法比对。")
        
        # 规则14: 处置方式1二级与处置情况报告比对
        excel_disposal_method_1 = str(row.get(app_config['COLUMN_MAPPINGS']['disposal_method_1'], '')).strip()
        
        # 假设处置方式1二级的值会出现在处置情况报告中
        if excel_disposal_method_1 and excel_disposal_method_1 not in disposal_report_content:
            issues_list.append({
                "受理线索编码": accepted_clue_code,
                "问题描述": f"行 {original_df_index + 2} - 处置方式1二级 ('{excel_disposal_method_1}') 未在处置情况报告中提及。"
            })
            error_count += 1
            logger.warning(f"行 {original_df_index + 2} - 处置方式1二级 ('{excel_disposal_method_1}') 未在处置情况报告中提及。")
        
    return issues_list, error_count
