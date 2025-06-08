# 配置文件
import os

class Config:
    UPLOAD_FOLDER = 'uploads'
    ALLOWED_EXTENSIONS = {'.xlsx', '.xls'}
    REQUIRED_FILENAME_PATTERN = '线索登记表'
    REQUIRED_HEADERS = ["填报单位名称", "办理机关", "被反映人", "处置情况报告", "受理时间"]
    
    # 列配置
    COLUMN_MAPPINGS = {
        "organization_measure": "组织措施",
        "acceptance_time": "受理时间"
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
        "confirm_organization_measure": "CC2组织措施内容请再次确认"
    }
    
    # 组织措施关键词
    ORGANIZATION_MEASURES = [
        "谈话提醒", "提醒谈话", "批评教育", "责令检查", "责令其做出书面检查", 
        "责令其做出检查", "诫勉", "警示谈话", "通报批评", "责令公开道歉（检查）", 
        "责令具结悔过"
    ]
    
    # 数据库路径
    DATABASE_PATH = os.path.join(os.path.dirname(__file__), 'case_management.db')
    
    # 安全密钥
    SECRET_KEY = os.urandom(24)  # 生成一个24字节的随机密钥