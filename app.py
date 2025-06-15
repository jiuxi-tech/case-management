from flask import Flask, flash, redirect, url_for
from config import Config
from routes import init_routes
import os
import sys
import logging
import webbrowser
import time
from threading import Timer
from db_utils import init_db

def create_app():
    app = Flask(__name__)
    app.config.from_object(Config)

    # 确保上传文件夹存在，使用 .exe 所在目录的 uploads
    base_path = os.path.dirname(os.path.abspath(__file__)) if not getattr(sys, 'frozen', False) else os.path.dirname(sys.executable)
    upload_folder = os.path.join(base_path, 'uploads')
    if not os.path.exists(upload_folder):
        os.makedirs(upload_folder)
    app.config['UPLOAD_FOLDER'] = upload_folder

    # 确保 CLUE_FOLDER 存在
    clue_folder = os.path.join(upload_folder, Config.TODAY_DATE, 'clue')
    if not os.path.exists(clue_folder):
        os.makedirs(clue_folder)
    app.config['CLUE_FOLDER'] = clue_folder

    # 确保 CASE_FOLDER 存在
    case_folder = os.path.join(upload_folder, Config.TODAY_DATE, 'case')
    if not os.path.exists(case_folder):
        os.makedirs(case_folder)
    app.config['CASE_FOLDER'] = case_folder

    # 动态添加 validation_rules 目录到 sys.path
    validation_rules_path = os.path.join(base_path, 'validation_rules')
    if os.path.exists(validation_rules_path) and validation_rules_path not in sys.path:
        sys.path.append(validation_rules_path)

    # 确保所有路由正确绑定
    with app.app_context():
        init_routes(app)

    return app

def open_browser():
    # 延迟 2 秒确保服务器启动
    time.sleep(2)
    if not hasattr(open_browser, 'opened'):
        webbrowser.open('http://localhost:5000')
        open_browser.opened = True

def run_app():
    # 获取 .exe 所在目录
    base_path = os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else os.getcwd()
    template_folder = os.path.join(base_path, 'templates')
    log_folder = os.path.join(base_path, 'logs')
    
    if getattr(sys, 'frozen', False):
        # PyInstaller 打包环境，调整模板路径
        template_folder = os.path.join(sys._MEIPASS, 'templates')
        if not os.path.exists(log_folder):
            os.makedirs(log_folder)
    
    # 显式创建应用实例，使用 create_app 设定的 UPLOAD_FOLDER
    app = create_app()
    app.template_folder = template_folder
    
    # 设置日志，输出到文件和控制台
    if not os.path.exists(log_folder):
        os.makedirs(log_folder, exist_ok=True)
    log_file = os.path.join(log_folder, 'app.log')
    if not os.access(base_path, os.W_OK):
        log_file = os.path.join(os.getenv('TEMP'), 'app.log')  # 回退到临时目录
    
    # 配置日志处理器
    logger = logging.getLogger(__name__)
    logger.setLevel(logging.INFO)
    formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')  # 移除 encoding 参数
    
    # 文件处理器
    file_handler = logging.FileHandler(log_file, encoding='utf-8')  # 将 encoding 应用到 FileHandler
    file_handler.setLevel(logging.INFO)
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)
    
    # 控制台处理器
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)
    
    logger.info("Application started in %s", base_path)
    logger.info("Log file set to: %s", log_file)
    logger.info("Template folder set to: %s", template_folder)
    logger.info("Upload folder set to: %s", app.config['UPLOAD_FOLDER'])
    logger.info("Clue folder set to: %s", app.config['CLUE_FOLDER'])
    logger.info("Case folder set to: %s", app.config['CASE_FOLDER'])

    # 数据库初始化
    with app.app_context():
        init_db()
    
    # 自定义错误处理
    @app.errorhandler(Exception)
    def handle_exception(e):
        logger.error(f"发生异常: {str(e)}")  # 记录到日志和控制台
        flash(f'发生错误: {str(e)}', 'error')
        return redirect(url_for('upload_case'))

    # 自动打开浏览器，仅首次执行
    Timer(1, open_browser).start()
    
    # 启用 development mode，无重载
    app.run(debug=False, host='0.0.0.0', port=5000, use_reloader=False)

if __name__ == '__main__':
    run_app()