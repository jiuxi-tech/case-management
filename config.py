import os

class Config:
    SECRET_KEY = os.environ.get('SECRET_KEY', 'default-secret-key')  # 可通过环境变量设置
    UPLOAD_FOLDER = os.path.join(os.getcwd(), 'uploads')
    DATABASE_PATH = os.path.join(os.getcwd(), 'case_management.db')
    ALLOWED_EXTENSIONS = {'.xlsx', '.xls'}
    REQUIRED_HEADERS = ["填报单位名称", "办理机关"]
    REQUIRED_FILENAME_PATTERN = "线索登记表"