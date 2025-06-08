import sqlite3
from config import Config

def get_db():
    conn = sqlite3.connect(Config.DATABASE_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS users (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                username TEXT UNIQUE NOT NULL,
                password TEXT NOT NULL
            )
        ''')
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS authority_agency_dict (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                authority TEXT NOT NULL,
                category TEXT NOT NULL,
                agency TEXT NOT NULL
            )
        ''')
        conn.commit()

        cursor.execute('SELECT COUNT(*) FROM authority_agency_dict')
        count = cursor.fetchone()[0]
        if count == 0:
            dict_data = [
                ("县市区旗纪委", "NSL", "平度市纪委监委第一纪检监察室"),
                ("县市区旗纪委", "NSL", "平度市纪委监委第二纪检监察室"),
                # ... 其余初始化数据保持不变
                ("县市区旗纪委", "SL", "平度市纪委监委派驻第十纪检监察组"),
            ]
            cursor.executemany('INSERT INTO authority_agency_dict (authority, category, agency) VALUES (?, ?, ?)', dict_data)
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