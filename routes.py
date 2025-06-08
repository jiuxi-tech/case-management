from flask import render_template, request, redirect, url_for, flash, session
from db_utils import get_user, create_user, get_authority_agency_dict, add_authority_agency, update_authority_agency, delete_authority_agency, get_db
from file_processor import process_upload
from werkzeug.security import generate_password_hash, check_password_hash

def init_routes(app):
    @app.route('/')
    def index():
        if 'username' in session:
            return render_template('index.html', title='首页')
        return redirect(url_for('login'))

    @app.route('/login', methods=['GET', 'POST'])
    def login():
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
    def logout():
        session.pop('username', None)
        flash('已登出', 'success')
        return redirect(url_for('login'))

    @app.route('/upload', methods=['GET', 'POST'])
    def upload():
        if 'username' not in session:
            return redirect(url_for('login'))
        if request.method == 'POST':
            return process_upload(request, app)
        return render_template('upload.html', title='上传')

    @app.route('/authority_agency')
    def authority_agency():
        if 'username' not in session:
            return redirect(url_for('login'))
        records = get_authority_agency_dict()
        return render_template('authority_agency.html', records=records, title='机关单位')

    @app.route('/authority_agency/add', methods=['GET', 'POST'])
    def add_authority_agency_route():
        if 'username' not in session:
            return redirect(url_for('login'))
        if request.method == 'POST':
            authority = request.form['authority']
            category = request.form['category']
            agency = request.form['agency']
            add_authority_agency(authority, category, agency)
            flash('新增成功！', 'success')
            return redirect(url_for('authority_agency'))
        return render_template('add_authority_agency.html', title='机关单位')

    @app.route('/authority_agency/update/<int:id>', methods=['GET', 'POST'])
    def update_authority_agency_route(id):
        if 'username' not in session:
            return redirect(url_for('login'))
        with get_db() as conn:
            cursor = conn.cursor()
            cursor.execute('SELECT * FROM authority_agency_dict WHERE id = ?', (id,))
            record = cursor.fetchone()
        if request.method == 'POST':
            authority = request.form['authority']
            category = request.form['category']
            agency = request.form['agency']
            update_authority_agency(id, authority, category, agency)
            flash('更新成功！', 'success')
            return redirect(url_for('authority_agency'))
        return render_template('update_authority_agency.html', record=record, title='机关单位')

    @app.route('/authority_agency/delete/<int:id>')
    def delete_authority_agency_route(id):
        if 'username' not in session:
            return redirect(url_for('login'))
        delete_authority_agency(id)
        flash('删除成功！', 'success')
        return redirect(url_for('authority_agency'))