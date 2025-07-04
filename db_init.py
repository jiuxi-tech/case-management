import sqlite3

# 定义数据库文件名
DATABASE = 'case_management.db'

def init_db():
    """
    初始化数据库，包括创建 'users' 表和 'authority_agency_dict' 表。
    如果 'authority_agency_dict' 表为空，则插入预设的初始化数据。
    """
    with sqlite3.connect(DATABASE) as conn:
        cursor = conn.cursor()

        # 创建 users 表
        # 该表用于存储用户信息，包括 id (主键，自增), username (用户名，唯一且非空), password (密码，非空)。
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS users (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                username TEXT UNIQUE NOT NULL,
                password TEXT NOT NULL
            )
        ''')

        # 创建 authority_agency_dict 表
        # 该表用于存储权限机构字典数据，包括 id (主键，自增), authority (权限), category (类别), agency (机构)。
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS authority_agency_dict (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                authority TEXT NOT NULL,
                category TEXT NOT NULL,
                agency TEXT NOT NULL
            )
        ''')
        conn.commit() # 提交事务，确保表创建成功

        # 检查 authority_agency_dict 表是否已初始化
        cursor.execute('SELECT COUNT(*) FROM authority_agency_dict')
        count = cursor.fetchone()[0]

        # 如果表为空，则插入初始化数据
        if count == 0:
            # 预设的权限机构字典数据
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
            # 批量插入数据到 authority_agency_dict 表
            cursor.executemany('INSERT INTO authority_agency_dict (authority, category, agency) VALUES (?, ?, ?)', dict_data)
            conn.commit() # 提交事务，确保数据插入成功

def get_db():
    """
    获取数据库连接。
    设置 row_factory 为 sqlite3.Row，以便可以通过字典方式访问查询结果的列。
    """
    conn = sqlite3.connect(DATABASE)
    conn.row_factory = sqlite3.Row # 将查询结果的行转换为可字典访问的对象
    return conn

def get_user(username):
    """
    根据用户名从 'users' 表中检索用户信息。
    """
    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute('SELECT * FROM users WHERE username = ?', (username,))
        return cursor.fetchone() # 返回匹配的第一条记录

def create_user(username, hashed_password):
    """
    在 'users' 表中创建新用户。
    参数:
        username (str): 用户名。
        hashed_password (str): 经过哈希处理的密码。
    """
    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute('INSERT INTO users (username, password) VALUES (?, ?)',
                       (username, hashed_password))
        conn.commit() # 提交事务，保存新用户

def get_authority_agency_dict():
    """
    从 'authority_agency_dict' 表中获取所有权限机构字典数据。
    """
    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute('SELECT * FROM authority_agency_dict')
        return cursor.fetchall() # 返回所有记录

def add_authority_agency(authority, category, agency):
    """
    向 'authority_agency_dict' 表中添加新的权限机构记录。
    参数:
        authority (str): 权限名称。
        category (str): 类别。
        agency (str): 机构名称。
    """
    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute('INSERT INTO authority_agency_dict (authority, category, agency) VALUES (?, ?, ?)',
                       (authority, category, agency))
        conn.commit() # 提交事务，保存新记录

def update_authority_agency(id, authority, category, agency):
    """
    更新 'authority_agency_dict' 表中指定 ID 的权限机构记录。
    参数:
        id (int): 要更新的记录的 ID。
        authority (str): 新的权限名称。
        category (str): 新的类别。
        agency (str): 新的机构名称。
    """
    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute('UPDATE authority_agency_dict SET authority = ?, category = ?, agency = ? WHERE id = ?',
                       (authority, category, agency, id))
        conn.commit() # 提交事务，保存更新

def delete_authority_agency(id):
    """
    从 'authority_agency_dict' 表中删除指定 ID 的权限机构记录。
    参数:
        id (int): 要删除的记录的 ID。
    """
    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute('DELETE FROM authority_agency_dict WHERE id = ?', (id,))
        conn.commit() # 提交事务，删除记录