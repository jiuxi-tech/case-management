from flask import Flask, render_template, request, redirect, url_for, flash, session
from werkzeug.security import generate_password_hash, check_password_hash
import os
import pandas as pd
from datetime import datetime
from database import init_db, get_user, create_user, get_db, get_authority_agency_dict, add_authority_agency, update_authority_agency, delete_authority_agency
import xlsxwriter

app = Flask(__name__)
app.secret_key = 'your-secret-key'  # 请替换为安全的密钥
UPLOAD_FOLDER = 'uploads'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# 确保上传文件夹存在
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

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
        if 'file' not in request.files:
            flash('未选择文件', 'error')
            return redirect(request.url)
        file = request.files['file']
        if file.filename == '':
            flash('未选择文件', 'error')
            return redirect(request.url)
        if not (file.filename.endswith('.xlsx') or file.filename.endswith('.xls')):
            flash('请上传Excel文件（.xlsx 或 .xls）', 'error')
            return redirect(request.url)
        if '线索登记表' not in file.filename:
            flash('文件名必须包含“线索登记表”', 'error')
            return redirect(request.url)

        file_path = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
        file.save(file_path)
        try:
            df = pd.read_excel(file_path)
            # Check required headers
            required_headers = ["填报单位名称", "办理机关"]
            if not all(header in df.columns for header in required_headers):
                flash('Excel文件缺少必要的表头“填报单位名称”或“办理机关”', 'error')
                return redirect(request.url)

            # Compare with database
            with get_db() as conn:
                cursor = conn.cursor()
                cursor.execute('SELECT authority, agency FROM authority_agency_dict WHERE category = ?', ('NSL',))
                db_records = cursor.fetchall()
                db_dict = {(str(row['authority']).strip().lower(), str(row['agency']).strip().lower()) for row in db_records}

            mismatches = []
            mismatch_indices = []
            for index, row in df.iterrows():
                agency = str(row["填报单位名称"]).strip().lower() if pd.notna(row["填报单位名称"]) else ''
                authority = str(row["办理机关"]).strip().lower() if pd.notna(row["办理机关"]) else ''
                if not authority or not agency or (authority, agency) not in db_dict:
                    mismatches.append(f"行 {index + 1}")
                    mismatch_indices.append(index)
                    if app.debug:
                        print(f"Debug - Row {index + 1}: authority='{authority}', agency='{agency}', db_match={((authority, agency) in db_dict)}")

            # Generate first Excel file with mismatches
            if mismatches:
                issues_df = pd.DataFrame(columns=['序号', '问题'])
                for i, _ in enumerate(mismatches, 1):
                    issues_df = pd.concat([issues_df, pd.DataFrame({'序号': [i], '问题': ['C2填报单位名称与H2办理机关不一致']})], ignore_index=True)
                issue_filename = f"线索编号{datetime.now().strftime('%Y%m%d')}.xlsx"
                issue_path = os.path.join(app.config['UPLOAD_FOLDER'], issue_filename)
                issues_df.to_excel(issue_path, index=False)

            # Generate second Excel file with original data and color marking
            original_filename = file.filename.replace('.xlsx', '_副本.xlsx').replace('.xls', '_副本.xlsx')
            original_path = os.path.join(app.config['UPLOAD_FOLDER'], original_filename)
            with pd.ExcelWriter(original_path, engine='xlsxwriter') as writer:
                df.to_excel(writer, sheet_name='Sheet1', index=False)
                workbook = writer.book
                worksheet = writer.sheets['Sheet1']
                cell_format = workbook.add_format({'bg_color': '#FF0000'})  # Red background for mismatches
                for idx in mismatch_indices:
                    worksheet.write(f'C{idx + 2}', df.at[idx, '填报单位名称'], cell_format)
                    worksheet.write(f'H{idx + 2}', df.at[idx, '办理机关'], cell_format)

            if mismatches:
                flash('线索对比有异常，请查看生成的问题表', 'error')
            else:
                flash('文件上传并验证成功！', 'success')

            return redirect(request.url)
        except Exception as e:
            flash(f'文件处理失败: {str(e)}', 'error')
            return redirect(request.url)
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

if __name__ == '__main__':
    init_db()  # 初始化数据库
    app.run(debug=True)