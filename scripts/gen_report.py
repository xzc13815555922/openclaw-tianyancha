#!/usr/bin/env python3
"""生成天眼查数据治理汇总报告"""
import sqlite3
from pathlib import Path
from datetime import datetime

DB = Path.home() / '.openclaw/plugin-skills/tianyancha/tianyancha.db'
conn = sqlite3.connect(str(DB))

# 动态获取数据截止时间（数据库最新注册时间）
cutoff_date = conn.execute("SELECT MAX(establishment_date) FROM enterprise_detail WHERE establishment_date != '' AND establishment_date IS NOT NULL").fetchone()[0] or '未知'

report = []
report.append("=" * 60)
report.append("  天眼查企业数据治理报告")
report.append(f"  生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M')}")
report.append(f"  数据截止时间: {cutoff_date}")
report.append("=" * 60)

# 一、数据概况
report.append("\n一、数据概况")
total = conn.execute("SELECT COUNT(*) FROM enterprise_detail").fetchone()[0]
report.append(f"  企业总数: {total:,}")

# 二、区县分布
report.append("\n二、区县分布")
region_stats = conn.execute("""
    SELECT region, COUNT(*) as cnt 
    FROM enterprise_detail 
    GROUP BY region 
    ORDER BY cnt DESC
""").fetchall()
for r, c in region_stats:
    r_display = r if r else "(未标注)"
    pct = c / total * 100
    report.append(f"  {r_display:<15} {c:>8,} ({pct:>5.1f}%)")

# 三、街道分布
report.append("\n三、街道分布")
street_stats = conn.execute("""
    SELECT street, region, COUNT(*) as cnt 
    FROM enterprise_detail 
    WHERE street != ''
    GROUP BY street, region 
    ORDER BY cnt DESC
""").fetchall()
for s, r, c in street_stats:
    report.append(f"  {s:<10} ({r:<8}) {c:>8,}")

# 四、大楼宇分布
report.append("\n四、大楼宇分布")
building_stats = conn.execute("""
    SELECT building, region, street, COUNT(*) as cnt 
    FROM enterprise_detail 
    WHERE building != ''
    GROUP BY building, region, street 
    ORDER BY cnt DESC
""").fetchall()
report.append(f"  {'大楼宇':<20} {'区县':<10} {'街道':<8} {'企业数':>8}")
report.append("  " + "-" * 52)
for b, r, s, c in building_stats:
    r_display = r if r else "(空)"
    s_display = s if s else "(空)"
    report.append(f"  {b:<20} {r_display:<10} {s_display:<8} {c:>8,}")
report.append(f"  大楼宇企业合计: {sum(x[3] for x in building_stats):,}")

# 五、小楼宇匹配情况
report.append("\n五、小楼宇匹配情况")
sb_stats = conn.execute("""
    SELECT building,
           COUNT(*) as total,
           SUM(CASE WHEN small_building != '' THEN 1 ELSE 0 END) as matched
    FROM enterprise_detail 
    WHERE building != ''
    GROUP BY building 
    ORDER BY total DESC
""").fetchall()
matched_total = sum(x[2] for x in sb_stats)
building_total = sum(x[1] for x in sb_stats)
report.append(f"  大楼宇企业总数: {building_total:,}")
report.append(f"  已匹配小楼宇: {matched_total:,}")
report.append(f"  小楼宇匹配率: {matched_total/building_total*100:.1f}%")

report.append(f"\n  {'大楼宇':<20} {'总数':>8} {'已配':>8} {'匹配率':>8}")
report.append("  " + "-" * 50)
for b, t, m in sb_stats:
    rate = m/t*100 if t > 0 else 0
    report.append(f"  {b:<20} {t:>8} {m:>8} {rate:>7.1f}%")

# 六、数据治理字段填充率
report.append("\n六、数据治理字段填充率")
fields = [
    ('region', '区县'),
    ('street', '街道'),
    ('building', '大楼宇'),
    ('small_building', '小楼宇'),
]
total = conn.execute("SELECT COUNT(*) FROM enterprise_detail").fetchone()[0]
for field, name in fields:
    filled = conn.execute(f"SELECT COUNT(*) FROM enterprise_detail WHERE {field} != '' AND {field} IS NOT NULL").fetchone()[0]
    pct = filled / total * 100
    report.append(f"  {name:<12} {filled:>8,} / {total:,} ({pct:>5.1f}%)")

# 七、大楼宇列表（盐南+经开）
report.append("\n七、大楼宇列表（盐南高新区 20栋 + 经开区 7栋）")
report.append("  盐南高新区:")
yannan = [b for b, r, s, c in building_stats if r == '盐南高新区']
for b, r, s, c in building_stats:
    if r == '盐南高新区':
        report.append(f"    {b}")

report.append("\n  经开区:")
for b, r, s, c in building_stats:
    if r == '经开区':
        report.append(f"    {b}")

report.append("\n" + "=" * 60)
report.append("  报告生成完成")
report.append("=" * 60)

conn.close()

# 输出
output = '\n'.join(report)
print(output)

# 保存到文件
output_dir = Path.home() / '.openclaw/plugin-skills/tianyancha/output'
output_dir.mkdir(exist_ok=True)
output_path = output_dir / '数据治理汇总报告.txt'
with open(output_path, 'w', encoding='utf-8') as f:
    f.write(output)

print(f"\n✅ 已保存: {output_path}")