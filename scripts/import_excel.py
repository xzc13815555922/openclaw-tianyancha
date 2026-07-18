#!/usr/bin/env python3
"""
天眼查企业数据导入模块
功能：将 Excel 企业数据导入数据库，按统一社会信用代码去重（重复则覆盖）

使用方式：
    python3 scripts/import_excel.py <excel文件路径>       # 追加导入（重复覆盖）
    python3 scripts/import_excel.py <excel文件路径> --force  # 无提示直接覆盖

导入流程：
    Step 1: 读取 Excel（跳过天眼查免责行，header=1）
    Step 2: 按 统一社会信用代码 UPSERT（重复覆盖）
    Step 3: 记录来源文件 + 采集日期
    Step 4: 自动执行 annotate_all.py 做数据治理标注

Excel 格式要求：
    - 第0行：天眼查免责说明（自动跳过）
    - 第1行：列头（公司名称、统一社会信用代码、注册地址 等）
    - 第2行起：数据
"""

import sys
import sqlite3
from pathlib import Path
import pandas as pd

DB_PATH = Path(__file__).parent.parent / "tianyancha.db"

# ── 列名映射（天眼查 Excel → DB字段）─────────────────────────────
TYC_COL_MAP = {
    '公司名称': 'enterprise_name',
    '企业名称': 'enterprise_name',
    '登记状态': 'registration_status',
    '法定代表人': 'legal_representative',
    '企业规模': 'enterprise_scale',
    '注册资本': 'registered_capital',
    '成立日期': 'establishment_date',
    '营业期限': 'business_term',
    '所属省份': 'province',
    '所属城市': 'city',
    '所属区县': 'county',
    '企业(机构)类型': 'enterprise_type',
    '国标行业门类': 'industry_sector',
    '国标行业大类': 'industry_major',
    '国标行业中类': 'industry_mid',
    '曾用名': 'former_name',
    '英文名': 'english_name',
    '统一社会信用代码': 'unified_social_credit_code',
    '纳税人识别号': 'taxpayer_id',
    '注册号': 'registration_number',
    '组织机构代码': 'organization_code',
    '参保人数': 'social_security_count',
    '参保人数所属年报': 'social_security_year',
    '有效手机号': 'mobile_phone',
    '更多电话': 'more_phones',
    '注册地址': 'registered_address',
    '通信地址': 'contact_address',
    '网址': 'website',
    '邮箱': 'email',
    '经营范围': 'business_scope',
}

# ── UPSERT 字段（按统一社会信用代码去重）────────────────────────
# 仅包含数据库中实际存在的字段，与 TYC_COL_MAP 严格对应
UPSERT_FIELDS = [
    'enterprise_name', 'registration_status', 'legal_representative', 'registered_capital',
    'establishment_date', 'business_term',
    'province', 'city', 'county', 'enterprise_type',
    'industry_sector', 'industry_major', 'industry_mid',
    'former_name', 'english_name',
    'unified_social_credit_code', 'taxpayer_id', 'registration_number', 'organization_code',
    'social_security_count', 'social_security_year',
    'mobile_phone', 'more_phones', 'registered_address',
    'contact_address', 'website', 'email', 'business_scope', 'enterprise_scale',
]

def clean_val(v):
    """清洗单元格值：- / nan / None → None"""
    if v is None:
        return None
    s = str(v).strip()
    if s in ('-', 'nan', '', 'None'):
        return None
    return s

def import_excel(excel_path: Path, force: bool = False):
    print(f"📖 读取Excel: {excel_path.name}")

    # Step 1: 读取（跳过第0行免责说明，header=1）
    try:
        df = pd.read_excel(excel_path, header=1)
    except Exception as e:
        print(f"❌ 读取失败: {e}")
        return False

    # 列名重命名（中文列名 → DB字段名）
    df = df.rename(columns=TYC_COL_MAP)

    # 过滤空行（按 DB 字段名）
    df = df.dropna(subset=['enterprise_name'])
    total = len(df)
    print(f"  数据行数: {total}")

    # 提取来源信息
    src_file = excel_path.name
    # 从文件名提取日期，如：高级搜索20260522 → 2026-05-22
    import re
    m = re.search(r'(\d{8})', excel_path.name)
    if m:
        date_str = m.group(1)
        collected_date = f"{date_str[:4]}-{date_str[4:6]}-{date_str[6:8]}"
    else:
        from datetime import date
        collected_date = str(date.today())

    conn = sqlite3.connect(str(DB_PATH))
    cursor = conn.cursor()

    # Step 2: UPSERT
    upsert_cols = UPSERT_FIELDS + ['collected_date', 'source_file']
    placeholders = ','.join([':' + f for f in upsert_cols])

    # 构建 INSERT OR REPLACE 语句
    # 注意：id 为自增主键，INSERT OR REPLACE 会导致 id 变化（可接受）
    insert_sql = f"""
        INSERT OR REPLACE INTO enterprise_detail ({','.join(upsert_cols)})
        VALUES ({placeholders})
    """

    skipped = 0
    imported = 0

    for _, row in df.iterrows():
        # 统一社会信用代码是去重 key
        credit_code = clean_val(row.get('unified_social_credit_code'))
        if not credit_code:
            skipped += 1
            continue

        record = {'collected_date': collected_date, 'source_file': src_file}
        for tyc_col, db_field in TYC_COL_MAP.items():
            if db_field in upsert_cols:
                record[db_field] = clean_val(row.get(db_field))

        # 转换参保人数为整数
        ss = record.get('social_security_count')
        if ss is not None:
            try:
                record['social_security_count'] = int(float(ss))
            except:
                record['social_security_count'] = None

        try:
            cursor.execute(insert_sql, record)
            imported += 1
        except Exception as e:
            print(f"  ⚠️ 插入失败: {record.get('enterprise_name', '未知')[:30]} - {e}")

        if imported % 500 == 0:
            print(f"  已导入: {imported}/{total}")

    conn.commit()

    # 统计
    db_total = conn.execute("SELECT COUNT(*) FROM enterprise_detail").fetchone()[0]
    with_credit = conn.execute("SELECT COUNT(*) FROM enterprise_detail WHERE unified_social_credit_code IS NOT NULL AND unified_social_credit_code != ''").fetchone()[0]

    print(f"\n✅ 导入完成")
    print(f"  本次导入: {imported} 条（跳过 {skipped} 条无信用代码）")
    print(f"  数据库总量: {db_total:,} 条 | 有信用代码: {with_credit:,} 条")
    print(f"  采集日期: {collected_date} | 来源: {src_file}")

    conn.close()
    return imported  # 返回导入条数


def run_annotation():
    """执行数据治理标注"""
    import subprocess
    script_path = Path(__file__).parent / "annotate_all.py"
    print("\n🔄 执行数据治理标注...")
    result = subprocess.run(['python3', str(script_path)], capture_output=True, text=True)
    print(result.stdout.strip())
    if result.returncode != 0:
        print(f"⚠️ 标注脚本异常: {result.stderr}")
        return False
    return True


def main():
    if len(sys.argv) < 2:
        print("用法: python3 import_excel.py <excel文件> [--force] [--no-annotate]")
        print("  --force: 无确认提示直接执行")
        print("  --no-annotate: 跳过数据治理标注")
        sys.exit(1)

    excel_path = Path(sys.argv[1])
    if not excel_path.exists():
        print(f"❌ 文件不存在: {excel_path}")
        sys.exit(1)

    force = '--force' in sys.argv
    skip_annotate = '--no-annotate' in sys.argv

    imported = import_excel(excel_path, force=force)
    if isinstance(imported, (int, float)) and imported > 0 and not skip_annotate:
        annotate_ok = run_annotation()
        if not annotate_ok:
            print("⚠️ 警告: 数据导入成功但标注执行异常，请稍后手动执行:")
            print(f"   python3 scripts/annotate_all.py")
            # 不退出，标注失败不应导致导入回滚
    elif isinstance(imported, (int, float)) and imported == 0:
        print("⚠️ 未导入任何数据（可能全部已存在或无效）")


if __name__ == '__main__':
    main()