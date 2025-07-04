from functools import wraps
from flask import render_template, request, redirect, url_for, flash, session, current_app
from db_utils import get_user, create_user, get_authority_agency_dict, add_authority_agency, \
                     update_authority_agency, delete_authority_agency, get_db
from file_processor import process_clue_upload, process_case_upload
from werkzeug.security import generate_password_hash, check_password_hash

def login_required(f):
    """
    一个简单的认证装饰器，用于保护需要用户登录才能访问的路由。
    如果用户未登录，则重定向到登录页面并闪现错误消息。
    """
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'username' not in session:
            flash('请先登录以访问此页面。', 'warning')
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated_function

def init_routes(app):
    """
    初始化 Flask 应用的所有路由。

    参数:
        app (flask.Flask): Flask 应用实例。
    """

    @app.route('/')
    @login_required
    def index():
        """
        首页路由。
        如果用户已登录，显示首页；否则，由 login_required 装饰器重定向到登录页。
        """
        return render_template('index.html', title='首页')

    @app.route('/login', methods=['GET', 'POST'])
    def login():
        """
        用户登录路由。
        GET 请求显示登录表单。
        POST 请求处理用户提交的登录信息，验证用户名和密码。
        登录成功则设置会话并重定向到首页；失败则闪现错误消息。
        """
        if request.method == 'POST':
            username = request.form['username']
            password = request.form['password']
            user = get_user(username)
            if user and check_password_hash(user['password'], password):
                session['username'] = username
                flash('登录成功！', 'success')
                return redirect(url_for('index'))
            else:
                flash('用户名或密码错误', 'error')
        return render_template('login.html', title='登录')

    @app.route('/register', methods=['GET', 'POST'])
    def register():
        """
        用户注册路由。
        GET 请求显示注册表单。
        POST 请求处理用户提交的注册信息。
        如果用户名已存在则闪现错误消息；否则创建新用户并重定向到登录页。
        """
        if request.method == 'POST':
            username = request.form['username']
            password = request.form['password']
            user = get_user(username)
            if user:
                flash('用户名已存在', 'error')
            else:
                hashed_password = generate_password_hash(password)
                create_user(username, hashed_password)
                flash('注册成功，请登录', 'success')
                return redirect(url_for('login'))
        return render_template('register.html', title='注册')

    @app.route('/logout')
    @login_required
    def logout():
        """
        用户登出路由。
        清除用户会话，闪现登出消息，并重定向到登录页面。
        """
        session.pop('username', None)
        flash('已登出', 'success')
        return redirect(url_for('login'))

    @app.route('/upload_clue', methods=['GET', 'POST'])
    @login_required
    def upload_clue():
        """
        线索登记表上传路由。
        GET 请求显示上传表单。
        POST 请求调用 file_processor.process_upload 处理文件上传和验证。
        """
        if request.method == 'POST':
            # 将 app 实例传递给 process_upload，以便其可以访问 app.config
            return process_upload(request, current_app._get_current_object())
        return render_template('upload_clue.html', title='上传')

    @app.route('/authority_agency')
    @login_required
    def authority_agency():
        """
        机关单位管理页面路由。
        显示所有已记录的机关单位信息。
        """
        records = get_authority_agency_dict()
        return render_template('authority_agency.html', records=records, title='机关单位')

    @app.route('/authority_agency/add', methods=['GET', 'POST'])
    @login_required
    def add_authority_agency_route():
        """
        添加机关单位路由。
        GET 请求显示添加表单。
        POST 请求处理新机关单位的添加。
        """
        if request.method == 'POST':
            authority = request.form['authority']
            category = request.form['category']
            agency = request.form['agency']
            add_authority_agency(authority, category, agency)
            flash('新增成功！', 'success')
            return redirect(url_for('authority_agency'))
        return render_template('add_authority_agency.html', title='机关单位')

    @app.route('/authority_agency/update/<int:id>', methods=['GET', 'POST'])
    @login_required
    def update_authority_agency_route(id):
        """
        更新机关单位路由。
        GET 请求显示指定 ID 机关单位的更新表单。
        POST 请求处理指定 ID 机关单位的更新。
        """
        record = None
        try:
            with get_db() as conn:
                cursor = conn.cursor()
                cursor.execute('SELECT * FROM authority_agency_dict WHERE id = ?', (id,))
                record = cursor.fetchone()
        except Exception as e:
            flash(f"获取记录失败: {e}", 'error')
            current_app.logger.error(f"获取机关单位记录失败 (ID: {id}): {e}", exc_info=True)
            return redirect(url_for('authority_agency'))

        if not record:
            flash(f"未找到 ID 为 {id} 的记录。", 'error')
            return redirect(url_for('authority_agency'))

        if request.method == 'POST':
            authority = request.form['authority']
            category = request.form['category']
            agency = request.form['agency']
            try:
                update_authority_agency(id, authority, category, agency)
                flash('更新成功！', 'success')
                return redirect(url_for('authority_agency'))
            except Exception as e:
                flash(f"更新失败: {e}", 'error')
                current_app.logger.error(f"更新机关单位记录失败 (ID: {id}): {e}", exc_info=True)
                return render_template('update_authority_agency.html', record=record, title='机关单位')
        return render_template('update_authority_agency.html', record=record, title='机关单位')

    @app.route('/authority_agency/delete/<int:id>')
    @login_required
    def delete_authority_agency_route(id):
        """
        删除机关单位路由。
        删除指定 ID 的机关单位记录。
        """
        try:
            delete_authority_agency(id)
            flash('删除成功！', 'success')
        except Exception as e:
            flash(f"删除失败: {e}", 'error')
            current_app.logger.error(f"删除机关单位记录失败 (ID: {id}): {e}", exc_info=True)
        return redirect(url_for('authority_agency'))

    @app.route('/upload_case', methods=['GET', 'POST'])
    @login_required
    def upload_case():
        """
        立案登记表上传路由。
        GET 请求显示上传表单。
        POST 请求调用 file_processor.process_case_upload 处理文件上传和验证。
        """
        if request.method == 'POST':
            # 将 app 实例传递给 process_case_upload，以便其可以访问 app.config
            return process_case_upload(request, current_app._get_current_object())
        return render_template('upload_case.html', title='上传立案登记表')

