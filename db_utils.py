import sqlite3
import logging

DATABASE = 'case_management.db'

def get_db():
    conn = sqlite3.connect(DATABASE)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    with get_db() as conn:
        cursor = conn.cursor()
        # 创建 users 表
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS users (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                username TEXT UNIQUE NOT NULL,
                password TEXT NOT NULL
            )
        ''')
        # 创建 authority_agency_dict 表，包含 category 字段
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS authority_agency_dict (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                authority TEXT NOT NULL,
                category TEXT NOT NULL,
                agency TEXT NOT NULL
            )
        ''')
        conn.commit()

        # 检查 authority_agency_dict 表是否已初始化
        cursor.execute('SELECT COUNT(*) FROM authority_agency_dict')
        count = cursor.fetchone()[0]
        if count == 0:
            # 初始化 authority_agency_dict 表数据
            dict_data = [
                ("县市区旗纪委", "NSL", "平度市纪委监委第一纪检监察室"),
                ("县市区旗纪委", "NSL", "平度市纪委监委第二纪检监察室"),
                ("县市区旗纪委", "NSL", "平度市纪委监委第三纪检监察室"),
                ("县市区旗纪委", "NSL", "平度市纪委监委第四纪检监察室"),
                ("县市区旗纪委", "NSL", "平度市纪委监委第五纪检监察室"),
                ("县市区旗纪委", "NSL", "平度市纪委监委第六纪检监察室"),
                ("县市区旗纪委", "NSL", "平度市纪委监委第七纪检监察室"),
                ("县市区旗纪委派驻纪检组", "NSL", "平度市纪委监委派驻第一纪检监察组"),
                ("县市区旗纪委派驻纪检组", "NSL", "平度市纪委监委派驻第二纪检监察组"),
                ("县市区旗纪委派驻纪检组", "NSL", "平度市纪委监委派驻第三纪检监察组"),
                ("县市区旗纪委派驻纪检组", "NSL", "平度市纪委监委派驻第四纪检监察组"),
                ("县市区旗纪委派驻纪检组", "NSL", "平度市纪委监委派驻第五纪检监察组"),
                ("县市区旗纪委派驻纪检组", "NSL", "平度市纪委监委派驻第六纪检监察组"),
                ("县市区旗纪委派驻纪检组", "NSL", "平度市纪委监委派驻第七纪检监察组"),
                ("县市区旗纪委派驻纪检组", "NSL", "平度市纪委监委派驻第八纪检监察组"),
                ("县市区旗纪委派驻纪检组", "NSL", "平度市纪委监委派驻第九纪检监察组"),
                ("县市区旗纪委派驻纪检组", "NSL", "平度市纪委监委派驻第十纪检监察组"),
                ("乡镇纪委监委机构", "NSL", "平度市凤台街道纪检监察工委"),
                ("乡镇纪委监委机构", "NSL", "平度市同和街道纪检监察工委"),
                ("乡镇纪委监委机构", "NSL", "平度市李园街道纪检监察工委"),
                ("乡镇纪委监委机构", "NSL", "平度市东阁街道纪检监察工委"),
                ("乡镇纪委监委机构", "NSL", "平度市白沙河街道纪检监察工委"),
                ("乡镇纪委监委机构", "NSL", "平度市南村镇纪委"),
                ("乡镇纪委监委机构", "NSL", "平度市蓼兰镇纪委"),
                ("乡镇纪委监委机构", "NSL", "平度市崔家集镇纪委"),
                ("乡镇纪委监委机构", "NSL", "平度市明村镇纪委"),
                ("乡镇纪委监委机构", "NSL", "平度市田庄镇纪委"),
                ("乡镇纪委监委机构", "NSL", "平度市新河镇纪委"),
                ("乡镇纪委监委机构", "NSL", "平度市云山镇纪委"),
                ("乡镇纪委监委机构", "NSL", "平度市古岘镇纪委"),
                ("乡镇纪委监委机构", "NSL", "平度市仁兆镇纪委"),
                ("乡镇纪委监委机构", "NSL", "平度市店子镇纪委"),
                ("乡镇纪委监委机构", "NSL", "平度市大泽山镇纪委"),
                ("乡镇纪委监委机构", "NSL", "平度市旧店镇纪委"),
                ("其他", "NSL", "平度市开发区纪检监察工委"),
                ("企业单位纪检监察机构", "NSL", "平度市控股集团纪委"),
                ("企业单位纪检监察机构", "NSL", "平度市开发集团纪委"),
                ("企业单位纪检监察机构", "NSL", "平度市农旅集团纪委"),
                ("企业单位纪检监察机构", "NSL", "平度市“食在平度”集团纪委"),
                ("县市区旗纪委", "SL", "平度市纪委监委第一纪检监察室"),
                ("县市区旗纪委", "SL", "平度市纪委监委第二纪检监察室"),
                ("县市区旗纪委", "SL", "平度市纪委监委第三纪检监察室"),
                ("县市区旗纪委", "SL", "平度市纪委监委第四纪检监察室"),
                ("县市区旗纪委", "SL", "平度市纪委监委第五纪检监察室"),
                ("县市区旗纪委", "SL", "平度市纪委监委第六纪检监察室"),
                ("县市区旗纪委", "SL", "平度市纪委监委第七纪检监察室"),
                ("县市区旗纪委", "SL", "平度市纪委监委派驻第一纪检监察组"),
                ("县市区旗纪委", "SL", "平度市纪委监委派驻第二纪检监察组"),
                ("县市区旗纪委", "SL", "平度市纪委监委派驻第三纪检监察组"),
                ("县市区旗纪委", "SL", "平度市纪委监委派驻第四纪检监察组"),
                ("县市区旗纪委", "SL", "平度市纪委监委派驻第五纪检监察组"),
                ("县市区旗纪委", "SL", "平度市纪委监委派驻第六纪检监察组"),
                ("县市区旗纪委", "SL", "平度市纪委监委派驻第七纪检监察组"),
                ("县市区旗纪委", "SL", "平度市纪委监委派驻第八纪检监察组"),
                ("县市区旗纪委", "SL", "平度市纪委监委派驻第九纪检监察组"),
                ("县市区旗纪委", "SL", "平度市纪委监委派驻第十纪检监察组"),
            ]
            cursor.executemany('INSERT INTO authority_agency_dict (authority, category, agency) VALUES (?, ?, ?)', dict_data)
            logging.info(f"Initialized {len(dict_data)} records in authority_agency_dict")
            conn.commit()

def get_user(username):
    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute('SELECT * FROM users WHERE username = ?', (username,))
        return cursor.fetchone()

def create_user(username, hashed_password):
    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute('INSERT INTO users (username, password) VALUES (?, ?)', 
                      (username, hashed_password))
        conn.commit()

def get_authority_agency_dict():
    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute('SELECT * FROM authority_agency_dict')
        return cursor.fetchall()

def add_authority_agency(authority, category, agency):
    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute('INSERT INTO authority_agency_dict (authority, category, agency) VALUES (?, ?, ?)',
                      (authority, category, agency))
        conn.commit()

def update_authority_agency(id, authority, category, agency):
    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute('UPDATE authority_agency_dict SET authority = ?, category = ?, agency = ? WHERE id = ?',
                      (authority, category, agency, id))
        conn.commit()

def delete_authority_agency(id):
    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute('DELETE FROM authority_agency_dict WHERE id = ?', (id,))
        conn.commit()