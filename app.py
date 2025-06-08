from flask import Flask
from config import Config
from routes import init_routes
import os

def create_app():
    app = Flask(__name__)
    app.config.from_object(Config)

    # 确保上传文件夹存在
    if not os.path.exists(app.config['UPLOAD_FOLDER']):
        os.makedirs(app.config['UPLOAD_FOLDER'])

    init_routes(app)

    return app

if __name__ == '__main__':
    app = create_app()
    with app.app_context():
        from db_utils import init_db
        init_db()
    app.run(debug=True)