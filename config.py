# config.py
import os
from datetime import datetime

class Config:
    """
    应用配置类。
    包含文件处理、Excel 映射、验证规则、关键词列表、数据库路径和安全密钥等配置。
    """

    ALLOWED_EXTENSIONS = {'xlsx', 'xls'}
    REQUIRED_FILENAME_PATTERN = '线索登记表'

    TODAY_DATE = datetime.now().strftime('%Y%m%d')
    """
    获取当前日期字符串，格式为YYYYMMDD。
    在应用启动时确定，保持不变。
    """

    # Excel 列名与内部字段名的映射。
    # 用于将 Excel 表格中的中文列名转换为代码中使用的英文键名。
    COLUMN_MAPPINGS = {
        "organization_measure": "组织措施",
        "acceptance_time": "受理时间",
        "joining_party_time": "入党时间",
        "ethnicity": "民族",
        "birth_date": "出生年月",
        "completion_time": "办结时间",
        "disposal_method_1": "处置方式1二级",
        "disposal_report": "处置情况报告",
        "accepted_clue_code": "受理线索编码",
        "accepted_personnel_code": "受理人员编码",
        "row_number": "行号",
        "compared_field": "比对字段",
        "being_compared_field": "被比对字段",
        "case_code": "案件编码",
        "person_code": "涉案人员编码",
        "disciplinary_sanction": "党纪处分",
        "administrative_sanction": "政务处分",
        "party_member": "是否中共党员",
        "investigated_person": "被调查人",
        "mentioned_person": "被反映人",
        "case_report": "立案报告",
        "disciplinary_decision": "处分决定",
        "investigation_report": "审查调查报告",
        "trial_report": "审理报告",
        "gender": "性别",
        "age": "年龄",
        "education": "学历",
        "party_joining_date": "入党时间",
        "brief_case_details": "简要案情",
        "filing_time": "立案时间",
        "filing_decision_doc": "立案决定书",
        "disciplinary_committee_filing_time": "纪委立案时间",
        "disciplinary_committee_filing_authority": "纪委立案机关",
        "supervisory_committee_filing_time": "监委立案时间",
        "supervisory_committee_filing_authority": "监委立案机关",
        "central_eight_provisions": "是否违反中央八项规定精神",
        "voluntary_confession": "是否主动交代问题",
        "closing_time": "结案时间",
        "no_party_position_warning": "是否属于本应撤销党内职务，但本人没有党内职务而给予严重警告处分",
        "recovery_amount": "追缴失职渎职滥用职权造成的损失金额",
        "trial_acceptance_time": "审理受理时间",
        "trial_closing_time": "审结时间",
        "trial_authority": "审理机关",
        "confiscation_amount": "收缴金额（万元）",
        "confiscation_of_property_amount": "没收金额",
        "compensation_amount": "责令退赔金额",
        "registered_handover_amount": "登记上交金额",
        "reporting_agency": "填报单位名称",
        "authority": "办理机关"
    }

    # Excel 单元格格式定义。
    # 用于在生成报告时对特定单元格进行高亮显示。
    FORMATS = {
        "red": '#FF0000',
        "yellow": '#FFFF00'
    }

    # 验证规则描述。
    # 键为规则的内部标识，值为用户友好的中文描述。
    VALIDATION_RULES = {
        "inconsistent_agency": "CR审理机关与A填报单位不一致",
        "inconsistent_name": "E2被反映人与AB2处置情况报告姓名不一致",
        "empty_report": "E2被反映人与AB2处置情况报告姓名不一致 (报告为空)",
        "confirm_acceptance_time": "AF2受理时间请再次确认",
        "inconsistent_organization_measure": "CC2组织措施与AB2处置情况报告不一致",
        "inconsistent_joining_party_time": "AC2入党时间与AB2处置情况报告不一致",
        "highlight_collection_amount": "收缴金额（万元）与.*处置情况报告对比结果是.*处置情况报告出现收缴二字",
        "highlight_confiscation_amount": "没收金额与.*处置情况报告对比结果是.*处置情况报告出现没收二字",
        "highlight_compensation_amount": "责令退赔金额与.*处置情况报告对比结果是.*处置情况报告出现责令退赔字样",
        "highlight_registration_amount": "登记上交金额与.*处置情况报告对比结果是.*处置情况报告出现登记上交金额字样",
        "highlight_recovery_amount": "追缴失职渎职滥用职权造成的损失金额与.*处置情况报告对比结果是.*处置情况报告出现追缴字样",
        "inconsistent_ethnicity": "W.*民族与AB.*处置情况报告民族不一致",
        "highlight_birth_date": "X2出生年月与AB2处置情况报告出生年月不一致",
        "highlight_completion_time": "BT2办结时间与AB2处置情况报告落款时间不一致",
        "highlight_disposal_method_1": "AK2处置方式1二级请再次确认",
        "inconsistent_case_name_report": "C2被调查人与BF2立案报告不一致",
        "inconsistent_case_name_decision": "C2被调查人与CU2处分决定不一致",
        "inconsistent_case_name_investigation": "C2被调查人与CX2审查调查报告不一致",
        "inconsistent_case_name_trial": "C2被调查人与CY2审理报告不一致",
        "disposal_decision_keyword_highlight": "CU处分决定中出现非人大代表、非政协委员、非committee、非中共党代表、非纪委委员等字样",
        "highlight_confiscation_of_property_amount": "CY审理报告中含有没收金额四字，请人工再次确认CG没收金额",
        "highlight_compensation_from_trial_report": "CY审理报告中含有责令退赔四字，请人工再次确认CH责令退赔金额",
        "disciplinary_sanction_party_member_mismatch": "党纪处分（处分决定）中出现开除党籍，但被调查人非中共党员，请核实！",
        "administrative_sanction_mismatch": "BR政务处分与CU处分决定不一致",
        "inconsistent_agency_clue": "C2填报单位名称与H2办理机关不一致"
    }

    # 组织措施相关的关键词列表。
    # 用于匹配和识别 Excel 内容中的组织措施。
    ORGANIZATION_MEASURE_KEYWORDS = [
        "谈话提醒", "提醒谈话", "批评教育", "责令检查", "责令其做出书面检查",
        "责令其做出检查", "诫勉", "警示谈话", "通报批评", "责令公开道歉（检查）",
        "责令具结悔过"
    ]

    # 处分决定中需要特别关注的关键词列表。
    # 用于识别处分决定中可能存在的特殊情况。
    DISPOSAL_DECISION_KEYWORDS = [
        "非人大代表", "非政协委员", "非党委委员", "非中共党代表", "非纪委委员"
    ]

    # 党纪处分相关的关键词列表。
    # 用于识别和验证党纪处分类型。
    DISCIPLINARY_SANCTION_KEYWORDS = [
        "开除党籍", "留党察看", "撤销党内职务", "严重警告", "警告"
    ]

    # 政务处分相关的关键词列表。
    # 用于识别和验证政务处分类型。
    ADMINISTRATIVE_SANCTION_KEYWORDS = [
        "警告", "记过", "记大过", "降级", "撤职", "开除"
    ]

    # 线索登记表必需的表头列表。
    # 用于验证上传的线索登记表是否包含所有必要的列。
    CLUE_REQUIRED_HEADERS = [
        "填报单位名称", "办理机关", "被反映人", "处置情况报告", "受理时间",
        "收缴金额（万元）", "没收金额", "责令退赔金额", "登记上交金额",
        "追缴失职渎职滥用职权造成的损失金额", "民族", "出生年月", "入党时间",
        "办结时间", "组织措施", "处置方式1二级"
    ]

    # 立案登记表必需的表头列表。
    # 用于验证上传的立案登记表是否包含所有必要的列。
    CASE_REQUIRED_HEADERS = [
        "被调查人", "立案报告", "处分决定", "审查调查报告", "审理报告"
    ]

    DATABASE_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'case_management.db')
    """
    SQLite 数据库文件的路径。
    文件将创建在 config.py 文件的同级目录。
    """

    SECRET_KEY = os.environ.get('SECRET_KEY', 'your_very_secret_key_here')
    """
    Flask 应用的安全密钥。
    用于会话管理和加密。建议从环境变量中获取，以提高安全性。
    如果环境变量未设置，则使用默认值 'your_very_secret_key_here'。
    在生产环境中，请务必将其替换为强随机字符串。
    """

    # 确保这些文件存储目录已定义且存在，如果你的项目结构不同，请修改路径
    UPLOAD_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'uploads')
    CASE_FOLDER = os.path.join(UPLOAD_FOLDER, 'cases')
    CLUE_FOLDER = os.path.join(UPLOAD_FOLDER, 'clues')

    # 在应用启动时创建这些目录 (或者确保你的 __init__.py 或 app.py 中有类似逻辑)
    os.makedirs(CASE_FOLDER, exist_ok=True)
    os.makedirs(CLUE_FOLDER, exist_ok=True)