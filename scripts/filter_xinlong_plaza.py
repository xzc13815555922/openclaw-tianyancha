#!/usr/bin/env python3
"""
新龙广场 - 近一个月新注册、注册资本 > 100 万、非个体工商户
复用 query.py 的列定义、Excel 导出格式，附加自定义 WHERE 条件
"""
import sys
import sqlite3
import re
from pathlib import Path
from itertools import groupby

# 复用 query.py 的全部定义
sys.path.insert(0, str(Path(__file__).parent))
from query import (
    EXCEL_COL_ORDER,
    REMAINING_COLS,
    COLUMN_DISPLAY_NAMES,
    export_excel,
    build_filename,
    OUTPUT_DIR,
    DB_PATH,
)

# ── 业务参数 ──────────────────────────────────────────────────────
BUILDING = "新龙广场"
DATE_FROM = "2026-05-03"   # 近一个月起点（今天 2026-06-03）
DATE_TO   = "2026-06-03"
CAPITAL_MIN_WAN = 100
CAPITAL_OP = ">"   # "大于"，严格语义；如需"100 万及以上"改为 ">="

def parse_capital_to_wan(s: str) -> float:
    """把 '100万元' / '1亿' / '1000.5万' / '50000' 转成万元数值"""
    if not s:
        return 0.0
    s = s.strip()
    m = re.search(r"([\d,.]+)\s*亿", s)
    if m:
        return float(m.group(1).replace(",", "")) * 10000
    m = re.search(r"([\d,.]+)\s*万", s)
    if m:
        return float(m.group(1).replace(",", ""))
    m = re.search(r"([\d,.]+)", s)
    if m:
        # 没标"万"也没标"亿"，按元处理 → 转万元
        return float(m.group(1).replace(",", "")) / 10000
    return 0.0

def main():
    conn = sqlite3.connect(str(DB_PATH))
    all_cols = EXCEL_COL_ORDER + REMAINING_COLS
    cols_sql = ", ".join(all_cols)

    sql = f"SELECT {cols_sql} FROM enterprise_detail WHERE building = ?"
    rows = conn.execute(sql, (BUILDING,)).fetchall()
    conn.close()

    # 过滤：日期范围 + 注册资本 + 排除个体
    filtered = []
    for row in rows:
        d = dict(zip(all_cols, row))
        ed = (d.get("establishment_date") or "")[:10]
        if not (DATE_FROM <= ed <= DATE_TO):
            continue
        cap = parse_capital_to_wan(d.get("registered_capital") or "")
        if CAPITAL_OP == ">" and cap <= CAPITAL_MIN_WAN:
            continue
        if CAPITAL_OP == ">=" and cap < CAPITAL_MIN_WAN:
            continue
        name = d.get("enterprise_name") or ""
        etype = d.get("enterprise_type") or ""
        if "个体" in name or "个体" in etype:
            continue
        filtered.append(d)

    # 排序：先按 region/street/building 升序分组；组内按注册时间倒序，企业名兜底
    def gkey(d):
        return (d.get("region") or "", d.get("street") or "", d.get("building") or "")
    def dkey(d):
        return ((d.get("establishment_date") or "")[:10], d.get("enterprise_name") or "")
    filtered.sort(key=gkey)
    # 组内按日期倒序
    out = []
    for _, group in groupby(filtered, key=gkey):
        chunk = sorted(list(group), key=dkey, reverse=True)
        out.extend(chunk)
    filtered = out

    # 还原成 row tuple
    rows_out = [tuple(d[c] for c in all_cols) for d in filtered]
    result = {
        "type": "list",
        "rows": rows_out,
        "total": len(rows_out),
        "columns": all_cols,
    }
    filters = [("building", "=", BUILDING)]
    export_excel(result, filters)

    # 摘要
    print(f"\n📋 筛选条件：")
    print(f"  大楼宇：{BUILDING}")
    print(f"  注册时间：{DATE_FROM} ~ {DATE_TO}（近一个月）")
    print(f"  注册资本：{CAPITAL_OP} {CAPITAL_MIN_WAN} 万")
    print(f"  排除：个体工商户")
    print(f"📊 结果：共 {len(rows_out)} 家")

    # 列序自检
    print(f"\n✅ Excel 列数：{len(all_cols)}（前7列：{' → '.join(EXCEL_COL_ORDER)}）")

if __name__ == "__main__":
    main()
