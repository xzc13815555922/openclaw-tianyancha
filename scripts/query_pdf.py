#!/usr/bin/env python3
"""
天眼查企业数据查询报告模块
功能：根据多列规则输出 PDF 报告（仅盐南高新区+经开区）

使用方式：
    python3 scripts/query_pdf.py                          # 全量统计（盐南+经开）
    python3 scripts/query_pdf.py --region 盐南高新区        # 按区县筛选
    python3 scripts/query_pdf.py --building 金融城        # 按大楼宇筛选
    python3 scripts/query_pdf.py --street 黄海街道        # 按街道筛选
    python3 scripts/query_pdf.py --year 2026             # 按成立年份筛选
    python3 scripts/query_pdf.py --type 有限              # 按企业类型筛选
    python3 scripts/query_pdf.py --not 个体               # 排除包含某关键词的记录
    python3 scripts/query_pdf.py --output pdf            # 输出PDF（默认）
    python3 scripts/query_pdf.py --group region           # 按某列分组统计

示例：
    # 盐南高新区 2026年新建企业（非个体工商户）
    python3 scripts/query_pdf.py --region 盐南高新区 --year 2026 --not 个体

    # 金融城 大楼宇 按街道分布统计
    python3 scripts/query_pdf.py --building 金融城 --group street

    # 步凤镇 企业类型分布
    python3 scripts/query_pdf.py --street 步凤镇 --group enterprise_type
"""

import sys
import sqlite3
import os
from pathlib import Path
from datetime import datetime

from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT

DB_PATH = Path(__file__).parent.parent / "tianyancha.db"
OUTPUT_DIR = Path(__file__).parent.parent / "output"
OUTPUT_DIR.mkdir(exist_ok=True)

# 注册中文字体
FONT_PATH = '/System/Library/Fonts/PingFang.ttc'  # macOS 拼音字体
FONT_FALLBACK = '/System/Library/Fonts/STHeiti Light.ttc'  # 黑体

def get_font():
    """获取可用的中文字体"""
    for path in [FONT_PATH, FONT_FALLBACK, '/System/Library/Fonts/Hiragino Sans GB.ttc']:
        if Path(path).exists():
            try:
                pdfmetrics.registerFont(TTFont('Chinese', path))
                return 'Chinese'
            except:
                pass
    return 'Helvetica'  # 回退到默认字体

CHINESE_FONT = get_font()

# 样式定义
styles = {
    'title': ParagraphStyle(
        'title',
        fontName=CHINESE_FONT,
        fontSize=18,
        leading=24,
        alignment=TA_CENTER,
        spaceAfter=12,
    ),
    'subtitle': ParagraphStyle(
        'subtitle',
        fontName=CHINESE_FONT,
        fontSize=12,
        leading=16,
        alignment=TA_CENTER,
        spaceAfter=6,
    ),
    'section': ParagraphStyle(
        'section',
        fontName=CHINESE_FONT,
        fontSize=14,
        leading=20,
        alignment=TA_LEFT,
        spaceBefore=12,
        spaceAfter=6,
    ),
    'normal': ParagraphStyle(
        'normal',
        fontName=CHINESE_FONT,
        fontSize=10,
        leading=14,
        alignment=TA_LEFT,
    ),
    'table_header': ParagraphStyle(
        'table_header',
        fontName=CHINESE_FONT,
        fontSize=9,
        leading=12,
        alignment=TA_CENTER,
    ),
    'table_cell': ParagraphStyle(
        'table_cell',
        fontName=CHINESE_FONT,
        fontSize=8,
        leading=11,
        alignment=TA_LEFT,
    ),
}

def build_where_clause(filters):
    """构建 WHERE 子句（仅限盐南高新区和经开区）"""
    conditions = ["region IN ('盐南高新区', '经开区')"]  # 固定筛选条件
    params = []
    for field, op, value in filters:
        if op == "=":
            conditions.append(f"{field} = ?")
            params.append(value)
        elif op == "like":
            conditions.append(f"{field} LIKE ?")
            params.append(f"%{value}%")
        elif op == "not_like":
            conditions.append(f"({field} IS NULL OR {field} NOT LIKE ?)")
            params.append(f"%{value}%")
        elif op == "year_eq":
            conditions.append(f"SUBSTR({field}, 1, 4) = ?")
            params.append(value)
        elif op == "month_eq":
            month_str = str(value).zfill(2)
            conditions.append(f"SUBSTR({field}, 6, 2) = ?")
            params.append(month_str)
        elif op == "month_gte":
            year_val = None
            for f, o, v in filters:
                if o == "year_eq": year_val = v
            if year_val:
                ym = f"{year_val}-{str(value).zfill(2)}"
                conditions.append(f"SUBSTR({field}, 1, 7) >= ?")
                params.append(ym)
            else:
                conditions.append(f"SUBSTR({field}, 6, 2) >= ?")
                params.append(str(value).zfill(2))
        elif op == "!=":
            conditions.append(f"{field} != ?")
            params.append(value)
        elif op == "is_null":
            conditions.append(f"{field} IS NULL OR {field} = ''")
    return conditions, params

def query(filters=None, group_by=None, limit=None):
    """执行查询"""
    conn = sqlite3.connect(str(DB_PATH))
    
    if filters:
        conditions, params = build_where_clause(filters)
        where_sql = "WHERE " + " AND ".join(conditions) if conditions else ""
    else:
        where_sql = "WHERE region IN ('盐南高新区', '经开区')"
        params = []
    
    if group_by:
        cols = f"'{group_by}', {group_by}, COUNT(*) as cnt"
        group_sql = f"""
            SELECT {cols}
            FROM enterprise_detail
            {where_sql}
            GROUP BY {group_by}
            ORDER BY COUNT(*) DESC
        """
        rows = conn.execute(group_sql, params).fetchall()
        conn.close()
        return {"type": "group", "group_by": group_by, "rows": rows}
    else:
        select_cols = ["enterprise_name", "registered_address", "region", "street", "building", "small_building", "enterprise_type", "establishment_date"]
        cols_sql = ", ".join(select_cols)
        sql = f"""
            SELECT {cols_sql}
            FROM enterprise_detail
            {where_sql}
        """
        if limit:
            sql += f" ORDER BY enterprise_name LIMIT {limit}"
        
        rows = conn.execute(sql, params).fetchall()
        conn.close()
        return {"type": "list", "rows": rows, "columns": select_cols}

def generate_pdf(result, title, filters_desc=""):
    """生成 PDF 报告"""
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M')
    
    output_path = OUTPUT_DIR / f"天眼查报告_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
    
    doc = SimpleDocTemplate(
        str(output_path),
        pagesize=A4,
        leftMargin=15*mm,
        rightMargin=15*mm,
        topMargin=20*mm,
        bottomMargin=20*mm,
    )
    
    elements = []
    
    # 标题
    elements.append(Paragraph("天眼查企业数据报告", styles['title']))
    elements.append(Paragraph(f"盐南高新区 | 经开区", styles['subtitle']))
    elements.append(Paragraph(f"生成时间: {timestamp}", styles['normal']))
    elements.append(Spacer(1, 5*mm))
    
    if filters_desc:
        elements.append(Paragraph(f"筛选条件: {filters_desc}", styles['normal']))
        elements.append(Spacer(1, 5*mm))
    
    if result["type"] == "group":
        # 分组统计表格
        elements.append(Paragraph(f"{result['group_by']} 分布统计", styles['section']))
        
        col_name = {"region": "区县", "street": "街道", "building": "大楼宇", "enterprise_type": "企业类型", "establishment_date": "成立年份"}
        headers = [col_name.get(result['group_by'], result['group_by']), "数量", "占比"]
        data = [headers]
        
        total = sum(row[2] for row in result["rows"])
        for row in result["rows"]:
            name = row[1] if row[1] else "(未标注)"
            pct = row[2] / total * 100 if total > 0 else 0
            data.append([name, str(row[2]), f"{pct:.1f}%"])
        
        data.append(["合计", str(total), "100.0%"])
        
        # 创建表格
        table = Table(data, colWidths=[100*mm, 40*mm, 40*mm])
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1F4E79')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('FONTNAME', (0, 0), (-1, -1), CHINESE_FONT),
            ('FONTSIZE', (0, 0), (-1, 0), 10),
            ('FONTSIZE', (0, 1), (-1, -1), 9),
            ('ALIGN', (1, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
            ('ROWBACKGROUNDS', (0, 1), (-1, -2), [colors.white, colors.HexColor('#D6E4F0')]),
            ('BACKGROUND', (0, -1), (-1, -1), colors.HexColor('#E2EFDA')),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
            ('TOPPADDING', (0, 0), (-1, -1), 4),
        ]))
        elements.append(table)
        
    elif result["type"] == "list":
        # 详细列表
        total = len(result["rows"])
        elements.append(Paragraph(f"查询结果 (共 {total} 条)", styles['section']))
        
        if total == 0:
            elements.append(Paragraph("无符合条件的数据", styles['normal']))
        else:
            # 截断显示前100条
            display_rows = result["rows"][:100]
            
            headers = ["企业名称", "注册地址", "区县", "街道", "大楼宇", "小楼宇", "类型", "成立日期"]
            data = [headers]
            
            for row in display_rows:
                data.append([str(v) if v else "" for v in row])
            
            if len(result["rows"]) > 100:
                data.append([f"... 还有 {total - 100} 条记录 (已截断显示前100条)", "", "", "", "", "", "", ""])
            
            col_widths = [45*mm, 60*mm, 20*mm, 20*mm, 25*mm, 25*mm, 15*mm, 20*mm]
            table = Table(data, colWidths=col_widths)
            table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1F4E79')),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
                ('FONTNAME', (0, 0), (-1, -1), CHINESE_FONT),
                ('FONTSIZE', (0, 0), (-1, 0), 8),
                ('FONTSIZE', (0, 1), (-1, -1), 7),
                ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
                ('ROWBACKGROUNDS', (0, 1), (-1, -2), [colors.white, colors.HexColor('#D6E4F0')]),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
                ('TOPPADDING', (0, 0), (-1, -1), 3),
            ]))
            elements.append(table)
    
    # 页脚
    elements.append(Spacer(1, 10*mm))
    elements.append(Paragraph(f"共 {total if result['type'] == 'list' else sum(r[2] for r in result['rows'])} 条记录 | 数据来源：天眼查", styles['normal']))
    
    doc.build(elements)
    return output_path

def _parse_int(raw, name, lo=None, hi=None):
    """严格整数解析 + 范围检查"""
    try:
        v = int(raw)
    except (TypeError, ValueError):
        print(f"❌ {name} 参数必须是整数，收到: {raw!r}", file=sys.stderr)
        sys.exit(2)
    if lo is not None and v < lo:
        print(f"❌ {name} 参数必须 ≥ {lo}，收到: {v}", file=sys.stderr)
        sys.exit(2)
    if hi is not None and v > hi:
        print(f"❌ {name} 参数必须 ≤ {hi}，收到: {v}", file=sys.stderr)
        sys.exit(2)
    return v


def main():
    filters = []
    group_by = None
    limit = None
    filters_desc_parts = []
    
    args = sys.argv[1:]
    i = 0
    while i < len(args):
        arg = args[i]
        if arg == "--region":
            filters.append(("region", "=", args[i+1]))
            filters_desc_parts.append(f"区县={args[i+1]}")
            i += 2
        elif arg == "--street":
            filters.append(("street", "=", args[i+1]))
            filters_desc_parts.append(f"街道={args[i+1]}")
            i += 2
        elif arg == "--building":
            filters.append(("building", "=", args[i+1]))
            filters_desc_parts.append(f"大楼宇={args[i+1]}")
            i += 2
        elif arg == "--small_building":
            filters.append(("small_building", "=", args[i+1]))
            i += 2
        elif arg == "--year":
            year_v = _parse_int(args[i+1], "--year", lo=1900, hi=2100)
            filters.append(("establishment_date", "year_eq", year_v))
            filters_desc_parts.append(f"成立年份={args[i+1]}")
            i += 2
        elif arg == "--month":
            month_v = _parse_int(args[i+1], "--month", lo=1, hi=12)
            filters.append(("establishment_date", "month_eq", month_v))
            filters_desc_parts.append(f"月份={args[i+1]}")
            i += 2
        elif arg == "--from-month":
            from_month_v = _parse_int(args[i+1], "--from-month", lo=1, hi=12)
            filters.append(("establishment_date", "month_gte", from_month_v))
            filters_desc_parts.append(f"{args[i+1]}月起")
            i += 2
        elif arg == "--type":
            filters.append(("enterprise_type", "like", args[i+1]))
            filters_desc_parts.append(f"类型包含{args[i+1]}")
            i += 2
        elif arg == "--scale":
            filters.append(("enterprise_scale", "like", args[i+1]))
            i += 2
        elif arg == "--not":
            filters.append(("enterprise_name", "not_like", args[i+1]))
            filters_desc_parts.append(f"排除{args[i+1]}")
            i += 2
        elif arg == "--group":
            group_by = args[i+1]
            i += 2
        elif arg == "--limit":
            limit = _parse_int(args[i+1], "--limit", lo=1)
            i += 2
        else:
            print(f"❌ 未识别的参数: {arg}", file=sys.stderr)
            print(f"   已识别参数: --region --street --building --small_building", file=sys.stderr)
            print(f"                  --year --month --from-month --not --type --scale --group --limit", file=sys.stderr)
            sys.exit(2)
    
    filters_desc = " | ".join(filters_desc_parts) if filters_desc_parts else "全部数据"
    
    print(f"查询条件: {filters_desc}")
    
    result = query(filters=filters, group_by=group_by, limit=limit)
    
    if result["type"] == "group":
        print(f"分组统计: {result['group_by']}")
        for row in result["rows"]:
            print(f"  {row[1]}: {row[2]}")
    else:
        print(f"查询结果: {len(result['rows'])} 条")
    
    output_path = generate_pdf(result, "天眼查企业数据报告", filters_desc)
    print(f"\n✅ PDF 已生成: {output_path}")

if __name__ == "__main__":
    main()