from flask import Flask
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

    # 确保上传文件夹存在
    if not os.path.exists(app.config['UPLOAD_FOLDER']):
        os.makedirs(app.config['UPLOAD_FOLDER'])

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
    upload_folder = os.path.join(base_path, 'uploads')
    log_folder = os.path.join(base_path, 'logs')
    
    if getattr(sys, 'frozen', False):
        # PyInstaller 打包环境，调整模板和上传路径
        template_folder = os.path.join(sys._MEIPASS, 'templates')
        upload_folder = os.path.join(sys._MEIPASS, 'uploads')
        if not os.path.exists(upload_folder):
            os.makedirs(upload_folder)
        if not os.path.exists(log_folder):
            os.makedirs(log_folder)
    
    # 显式创建应用实例
    app = create_app()
    app.config['UPLOAD_FOLDER'] = upload_folder
    app.template_folder = template_folder
    
    # 设置日志，指定 UTF-8 编码
    log_file = os.path.join(log_folder, 'app.log')
    if not os.access(base_path, os.W_OK):
        log_file = os.path.join(os.getenv('TEMP'), 'app.log')  # 回退到临时目录
    logging.basicConfig(filename=log_file, level=logging.INFO,
                        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
                        encoding='utf-8')
    logger = logging.getLogger(__name__)
    logger.info("Application started in %s", base_path)
    logger.info("Log file set to: %s", log_file)
    logger.info("Template folder set to: %s", template_folder)

    # 数据库初始化
    with app.app_context():
        init_db()
    
    # 自动打开浏览器，仅首次执行
    Timer(1, open_browser).start()
    
    # 启用 development mode，无重载
    app.run(debug=True, host='0.0.0.0', port=5000, use_reloader=False)

if __name__ == '__main__':
    run_app()