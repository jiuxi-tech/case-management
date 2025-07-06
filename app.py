from flask import Flask, flash, redirect, url_for
from config import Config
from routes import init_routes
import os
import sys
import logging
import webbrowser
import time
from threading import Timer
from datetime import datetime # 导入 datetime 模块

from db_utils import init_db

def _get_base_path():
    """
    获取应用的基础路径。
    兼容 PyInstaller 打包环境，返回 .exe 所在目录；
    否则返回当前文件所在目录。
    """
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

def _ensure_directories(app_config):
    """
    确保必要的文件夹存在并配置到 Flask app config 中。
    包括 uploads、uploads/YYYYMMDD/clue 和 uploads/YYYYMMDD/case 文件夹。
    同时动态添加 validation_rules 目录到 sys.path。
    """
    base_path = _get_base_path()
    
    # 获取当前的日期字符串，用于构建日期相关的文件夹路径
    today_date = Config().TODAY_DATE

    # 确保 uploads 文件夹存在
    upload_folder = os.path.join(base_path, 'uploads')
    os.makedirs(upload_folder, exist_ok=True)
    app_config['UPLOAD_FOLDER'] = upload_folder

    # 确保 clue 文件夹存在
    clue_folder = os.path.join(upload_folder, today_date, 'clue')
    os.makedirs(clue_folder, exist_ok=True)
    app_config['CLUE_FOLDER'] = clue_folder

    # 确保 case 文件夹存在
    case_folder = os.path.join(upload_folder, today_date, 'case')
    os.makedirs(case_folder, exist_ok=True)
    app_config['CASE_FOLDER'] = case_folder

    # 动态添加 validation_rules 目录到 sys.path，以便导入自定义验证规则
    validation_path = os.path.join(base_path, 'validation') # 注意这里是 'validation'
    if os.path.exists(validation_path) and validation_path not in sys.path:
        sys.path.append(validation_path)

    # 动态添加 file_upload 目录到 sys.path，以便导入其中的模块
    file_upload_path = os.path.join(base_path, 'file_upload')
    if os.path.exists(file_upload_path) and file_upload_path not in sys.path:
        sys.path.append(file_upload_path)


def create_app():
    """
    创建并配置 Flask 应用实例。
    加载配置、确保必要目录存在并绑定路由。
    """
    app = Flask(__name__)
    app.config.from_object(Config)
    _ensure_directories(app.config)

    # 确保所有路由正确绑定到应用实例
    with app.app_context():
        init_routes(app)
        
    return app

def _open_browser_if_not_opened():
    """
    延迟打开浏览器，确保服务器启动后再打开，且只打开一次。
    """
    time.sleep(1)
    webbrowser.open('http://localhost:5000')

def _configure_logging(app, base_path):
    """
    配置应用的日志系统，将日志输出到文件和控制台。
    日志文件路径会根据写入权限进行调整。
    日志文件以当前日期为目录名，日志文件以日期作为日志名。
    """
    # 获取当前日期字符串，用于日志目录和文件名
    today_date = datetime.now().strftime('%Y%m%d')

    # 构建日志目录路径
    log_folder = os.path.join(base_path, 'logs', today_date)
    os.makedirs(log_folder, exist_ok=True) # 确保日志目录存在

    # 构建日志文件路径
    log_file = os.path.join(log_folder, f'{today_date}.log')

    # 检查日志目录是否有写入权限，若无则回退到临时目录
    # 对于 PyInstaller 打包的应用，确保日志文件始终可写
    if not os.access(log_folder, os.W_OK) and not getattr(sys, 'frozen', False):
        temp_log_folder = os.path.join(os.getenv('TEMP', '/tmp'), 'logs', today_date)
        os.makedirs(temp_log_folder, exist_ok=True)
        log_file = os.path.join(temp_log_folder, f'{today_date}.log')
    elif getattr(sys, 'frozen', False):
        if not os.access(log_folder, os.W_OK):
            temp_log_folder = os.path.join(os.getenv('TEMP', '/tmp'), 'logs', today_date)
            os.makedirs(temp_log_folder, exist_ok=True)
            log_file = os.path.join(temp_log_folder, f'{today_date}.log')

    # 获取根日志器并设置级别
    root_logger = logging.getLogger()
    root_logger.setLevel(logging.DEBUG)
    
    # 获取当前模块的日志器
    logger = logging.getLogger(__name__)
    
    # 定义业务日志过滤器
    class BusinessLogFilter(logging.Filter):
        def filter(self, record):
            # 只允许以 <立案 开头的业务日志消息通过
            return record.getMessage().startswith('<立案')
    
    # 定义业务日志格式器
    class BusinessLogFormatter(logging.Formatter):
        def format(self, record):
            # 只显示消息内容，不显示时间戳等信息
            return record.getMessage()
    
    # 定义标准日志格式（用于文件记录）
    standard_formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
    business_formatter = BusinessLogFormatter()

    # 移除所有现有的处理器，防止重复添加日志处理器
    if root_logger.handlers:
        for handler in root_logger.handlers:
            root_logger.removeHandler(handler)

    # 文件处理器：将日志写入文件（使用标准格式记录所有日志）
    file_handler = logging.FileHandler(log_file, encoding='utf-8')
    file_handler.setLevel(logging.DEBUG)
    file_handler.setFormatter(standard_formatter)
    root_logger.addHandler(file_handler)

    # 控制台处理器：将日志输出到控制台（只显示业务日志）
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.DEBUG)
    console_handler.setFormatter(business_formatter)
    console_handler.addFilter(BusinessLogFilter())  # 添加过滤器
    root_logger.addHandler(console_handler)

    # 将日志处理器也添加到 Flask 应用的日志器中
    app.logger.addHandler(file_handler)
    app.logger.addHandler(console_handler)
    
    # 设置Flask应用日志器级别
    app.logger.setLevel(logging.DEBUG)

    # 记录应用启动信息和关键路径
    logger.info("Application started in %s", base_path)
    logger.info("Log file set to: %s", log_file)
    logger.info("Upload folder set to: %s", app.config['UPLOAD_FOLDER'])
    logger.info("Clue folder set to: %s", app.config['CLUE_FOLDER'])
    logger.info("Case folder set to: %s", app.config['CASE_FOLDER'])

def run_app():
    """
    运行 Flask 应用。
    包括应用创建、模板路径设置、日志配置、数据库初始化、错误处理和启动服务器。
    """
    base_path = _get_base_path()
    app = create_app()

    # 根据运行环境调整模板路径
    if getattr(sys, 'frozen', False):
        app.template_folder = os.path.join(sys._MEIPASS, 'templates')
    else:
        app.template_folder = os.path.join(base_path, 'templates')

    _configure_logging(app, base_path)
    
    # 数据库初始化，在应用上下文内执行
    with app.app_context():
        init_db()
    
    # 自定义全局错误处理
    @app.errorhandler(Exception)
    def handle_exception(e):
        """
        捕获应用中的所有未处理异常。
        记录错误日志、闪现错误消息并重定向到上传页面。
        """
        app.logger.error(f"发生异常: {str(e)}", exc_info=True)  # 记录完整堆栈信息
        flash(f'发生错误: {str(e)}', 'error')
        return redirect(url_for('upload_case')), 500  # 返回 500 状态码

    # 自动打开浏览器，仅在应用首次启动时执行一次
    if not hasattr(run_app, '_browser_opened'):
        Timer(1, _open_browser_if_not_opened).start()
        run_app._browser_opened = True
    
    # 启动 Flask 开发服务器
    app.run(debug=False, host='0.0.0.0', port=5000, use_reloader=False)

if __name__ == '__main__':
    run_app()
