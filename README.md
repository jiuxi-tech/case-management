# case-management
案管系统

## 1. 项目结构
app.py：主应用文件，包含Flask路由、用户认证和文件上传逻辑。
templates/：包含HTML模板，使用Jinja2渲染。
base.html：基础模板，包含导航栏和通用样式。
login.html：登录页面。
register.html：注册页面。
index.html：首页。
upload.html：文件上传页面，显示上传的Excel内容。
uploads/：存储上传的Excel文件。

## 2. 项目依赖
Python库：flask, werkzeug, pandas, xlsxwriter
安装命令：pip install flask werkzeug pandas openpyxl xlsxwriter pandas

## 3. 功能：
登录/注册：简单的用户认证系统，使用哈希存储密码（werkzeug.security），通过session管理登录状态。
文件上传：支持上传.xlsx文件，使用pandas解析并显示内容，文件名添加时间戳避免冲突。
消息提示：使用Flask的flash功能显示成功或错误提示，样式美观。

## 4. 运行项目：
创建项目目录，保存以上文件。
确保templates/目录包含所有HTML文件，创建uploads/目录。
安装依赖：pip install flask werkzeug pandas openpyxl
运行app.py：python app.py
访问http://localhost:5000

## 5. 注意事项
请将app.secret_key替换为安全的随机字符串（可用os.urandom(24).hex()生成）。
当前用户数据存储在内存中（users字典），生产环境建议使用数据库（如SQLite或MySQL）。
上传文件保存在uploads/目录，确保有写权限。
界面支持响应式设计，适配移动端和桌面端。

## 6. 创建打包程序
pyinstaller --onefile --add-data "templates;templates" --add-data "uploads;uploads" package_app.py
