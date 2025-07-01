# config.py
import os
from datetime import datetime

class Config:
    BASE_UPLOAD_FOLDER = 'uploads'
    TODAY_DATE = datetime.now().strftime('%Y%m%d')
    UPLOAD_FOLDER = os.path.join(BASE_UPLOAD_FOLDER, TODAY_DATE)
    CLUE_FOLDER = os.path.join(UPLOAD_FOLDER, 'clue')
    CASE_FOLDER = os.path.join(UPLOAD_FOLDER, 'case')
    ALLOWED_EXTENSIONS = {'.xlsx', '.xls'}
    REQUIRED_FILENAME_PATTERN = '线索登记表'
    
    # 列配置
    COLUMN_MAPPINGS = {
        "organization_measure": "组织措施",
        "acceptance_time": "受理时间",
        "joining_party_time": "入党时间",
        "ethnicity": "民族",
        "birth_date": "出生年月",
        "completion_time": "办结时间",
        "disposal_method_1": "处置方式1二级",
        "disposal_report": "处置情况报告", 
        "accepted_clue_code": "受理线索编码" 
    }
    
    # Excel 格式
    FORMATS = {
        "red": '#FF0000',
        "yellow": '#FFFF00'
    }
    
    # 验证规则描述
    VALIDATION_RULES = {
        "inconsistent_agency": "C2填报单位名称与H2办理机关不一致",
        "inconsistent_name": "E2被反映人与AB2处置情况报告姓名不一致",
        "empty_report": "E2被反映人与AB2处置情况报告姓名不一致 (报告为空)",
        "confirm_acceptance_time": "AF2受理时间请再次确认",
        "inconsistent_organization_measure": "CC2组织措施与AB2处置情况报告不一致",
        "inconsistent_joining_party_time": "AC2入党时间与AB2处置情况报告不一致",
        "highlight_collection_amount": "Q2收缴金额请再次确认",
        "highlight_confiscation_amount": "R2没收金额请再次确认",
        "highlight_compensation_amount": "S2责令退赔金额请再次确认",
        "highlight_registration_amount": "T2登记上交金额请再次确认",
        "highlight_recovery_amount": "CJ追缴失职渎职滥用职权造成的损失金额请再次确认",
        "inconsistent_ethnicity": "W2民族与AB2处置情况报告民族不一致",
        "highlight_birth_date": "X2出生年月与AB2处置情况报告出生年月不一致",
        "highlight_completion_time": "BT2办结时间与AB2处置情况报告落款时间不一致",
        "highlight_disposal_method_1": "AK2处置方式1二级请再次确认",
        "inconsistent_case_name_report": "C2被调查人与BF2立案报告不一致",
        "inconsistent_case_name_decision": "C2被调查人与CU2处分决定不一致",
        "inconsistent_case_name_investigation": "C2被调查人与CX2审查调查报告不一致",
        "inconsistent_case_name_trial": "C2被调查人与CY2审理报告不一致",
        # 【新增】处分决定关键词检查规则
        "disposal_decision_keyword_highlight": "CU处分决定中出现非人大代表、非政协委员、非党委委员、非中共党代表、非纪委委员等字样"
    }
    
    # 组织措施关键词
    ORGANIZATION_MEASURE_KEYWORDS = [
        "谈话提醒", "提醒谈话", "批评教育", "责令检查", "责令其做出书面检查", 
        "责令其做出检查", "诫勉", "警示谈话", "通报批评", "责令公开道歉（检查）", 
        "责令具结悔过"
    ]

    # 【新增】处分决定中的特殊关键词
    DISPOSAL_DECISION_KEYWORDS = [
        "非人大代表", "非政协委员", "非党委委员", "非中共党代表", "非纪委委员"
    ]
    
    # 线索登记表必需的表头
    CLUE_REQUIRED_HEADERS = ["填报单位名称", "办理机关", "被反映人", "处置情况报告", "受理时间"]
    
    # 立案登记表必需的表头
    CASE_REQUIRED_HEADERS = ["被调查人", "立案报告", "处分决定", "审查调查报告", "审理报告"]
    
    # 数据库路径
    DATABASE_PATH = os.path.join(os.path.dirname(__file__), 'case_management.db')
    
    # 安全密钥
    SECRET_KEY = os.urandom(24) 

    # 创建目录
    if not os.path.exists(CLUE_FOLDER):
        os.makedirs(CLUE_FOLDER)
    if not os.path.exists(CASE_FOLDER):
        os.makedirs(CASE_FOLDER)