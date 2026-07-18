#!/usr/bin/env python3
"""
天眼查数据库自动备份脚本
用法：python3 scripts/backup_db.py [--keep N]
  --keep N: 保留最近 N 份备份（默认 14）
"""
import sys
import shutil
import subprocess
from pathlib import Path
from datetime import datetime

DB_PATH = Path(__file__).parent.parent / "tianyancha.db"
BACKUP_DIR = Path(__file__).parent.parent / "backups"
BACKUP_DIR.mkdir(exist_ok=True)

def get_db_size(path):
    return path.stat().st_size if path.exists() else 0

def main():
    keep = 14
    if len(sys.argv) > 1 and sys.argv[1] == '--keep':
        keep = int(sys.argv[2])

    if not DB_PATH.exists():
        print(f"❌ 数据库不存在: {DB_PATH}")
        sys.exit(1)

    size_mb = get_db_size(DB_PATH) / 1024 / 1024
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    backup_path = BACKUP_DIR / f"tianyancha_{timestamp}.db"

    # 使用 VACUUM INTO 确保备份一致性
    import sqlite3
    try:
        conn = sqlite3.connect(str(DB_PATH))
        conn.execute(f"VACUUM INTO '{backup_path}'")
        conn.close()
    except Exception as e:
        # 回退到 shutil copy
        print(f"  VACUUM INTO 失败 ({e})，回退到文件复制...")
        shutil.copy2(DB_PATH, backup_path)

    backup_size_mb = get_db_size(backup_path) / 1024 / 1024
    print(f"✅ 备份完成: {backup_path.name}")
    print(f"   大小: {backup_size_mb:.1f} MB | 来源: {size_mb:.1f} MB")

    # 清理旧备份
    backups = sorted(BACKUP_DIR.glob("tianyancha_*.db"))
    while len(backups) > keep:
        old = backups.pop(0)
        old.unlink()
        print(f"  清理旧备份: {old.name}")
    print(f"   当前备份数: {min(len(backups), keep)} / {keep}")

if __name__ == '__main__':
    main()
