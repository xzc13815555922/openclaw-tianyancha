#!/usr/bin/env python3
"""
备份恢复演练脚本
验证最新备份文件可读、表结构与主库一致
用法：
    python3 scripts/verify_backup.py
"""""
import sqlite3
from pathlib import Path

from db import get_connection
BACKUP_DIR = Path(__file__).parent.parent / "backups"

def verify():
    if not BACKUP_DIR.exists():
        print("❌ backups/ 目录不存在")
        return False
    
    backups = sorted(BACKUP_DIR.glob("*.db"))
    if not backups:
        print("❌ 无备份文件")
        return False
    
    latest = backups[-1]
    print(f"最新备份: {latest.name} ({latest.stat().st_size / 1024 / 1024:.1f} MB)")
    
    # 验证备份文件可读
    try:
        bk = sqlite3.connect(str(latest))
        bk_cnt = bk.execute("SELECT COUNT(*) FROM enterprise_detail").fetchone()[0]
        bk_cols = [r[1] for r in bk.execute("PRAGMA table_info(enterprise_detail)").fetchall()]
        bk.close()
        print(f"  ✅ 备份可读 | 数据量: {bk_cnt:,} | 列数: {len(bk_cols)}")
    except Exception as e:
        print(f"  ❌ 备份不可读: {e}")
        return False
    
    # 对比主库
    try:
        main = sqlite3.connect(str(DB_PATH))
        main_cnt = main.execute("SELECT COUNT(*) FROM enterprise_detail").fetchone()[0]
        main_cols = [r[1] for r in main.execute("PRAGMA table_info(enterprise_detail)").fetchall()]
        main.close()
        
        if bk_cnt != main_cnt:
            print(f"  ⚠️ 数据量不一致: 备份={bk_cnt}, 主库={main_cnt}")
        else:
            print(f"  ✅ 数据量一致: {main_cnt:,}")
        
        if set(bk_cols) != set(main_cols):
            extra = set(main_cols) - set(bk_cols)
            missing = set(bk_cols) - set(main_cols)
            if extra: print(f"  ⚠️ 主库多字段: {extra}")
            if missing: print(f"  ⚠️ 备份多字段: {missing}")
        else:
            print(f"  ✅ 列结构一致")
    except Exception as e:
        print(f"  ❌ 主库对比失败: {e}")
        return False
    
    print(f"\n✅ 备份验证通过 ({latest.name})")
    return True

if __name__ == '__main__':
    verify()
