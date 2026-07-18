"""天眼查数据库统一连接模块"""
import sys
from pathlib import Path

# 兼容直接 python3 scripts/xxx.py 和 python3 xxx.py 两种运行方式
# 目录结构: tianyancha/scripts/db.py
# db.py 在 scripts/ 中

if __name__ != '__main__':
    _SCRIPT_DIR = Path(__file__).parent.parent  # tianyancha/
else:
    _SCRIPT_DIR = Path.cwd()

DB_PATH = _SCRIPT_DIR / "tianyancha.db"

def get_connection():
    """获取数据库连接"""
    import sqlite3
    return sqlite3.connect(str(DB_PATH))

def get_db_path():
    """获取数据库路径"""
    return DB_PATH
