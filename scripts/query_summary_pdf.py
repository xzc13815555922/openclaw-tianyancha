#!/usr/bin/env python3
"""
天眼查企业数据汇总报告（PDF格式）
Page 1: 按区县汇总
Page 2: 按街道汇总
Page 3: 按大楼宇汇总

筛选条件：仅盐南高新区 + 经开区

统计口径：
- 总企业数（不含个体工商户）
- 其中有联系方式数（mobile_phone 为11位手机号码）
- 总个体工商户
- 其中有联系方式数

联系方式判定：mobile_phone 符合11位纯数字格式（正则 ^[0-9]{11}$）

排版质量检查机制：
- 所有表格宽度自动对齐（总宽度不超过页面可用宽度 186mm）
- 禁止任何表格列宽超出页面边界
"""

import sqlite3
from pathlib import Path
from datetime import datetime

from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.enums import TA_CENTER, TA_LEFT

DB_PATH = Path(__file__).parent.parent / "tianyancha.db"
OUTPUT_DIR = Path(__file__).parent.parent / "output"
OUTPUT_DIR.mkdir(exist_ok=True)

# 页面可用宽度(mm),A4 左右各 12mm 边距
PAGE_USABLE_WIDTH = (210 - 12 * 2) * mm  # 186mm in points

def check_table_width(col_widths, label=""):
    """排版质量检查:确保表格不超过页面可用宽度"""
    total = sum(col_widths)
    if total > PAGE_USABLE_WIDTH:
        print(f"⚠️ 表格宽度警告 [{label}]: {total}mm > {PAGE_USABLE_WIDTH}mm(超出 {total - PAGE_USABLE_WIDTH}mm)")
        return False
    return True

# 注册中文字体
def get_font():
    for path in ['/System/Library/Fonts/PingFang.ttc', '/System/Library/Fonts/STHeiti Light.ttc', '/System/Library/Fonts/Hiragino Sans GB.ttc']:
        if Path(path).exists():
            try:
                pdfmetrics.registerFont(TTFont('Chinese', path))
                return 'Chinese'
            except:
                pass
    return 'Helvetica'

CHINESE_FONT = get_font()

styles = {
    'title': ParagraphStyle('title', fontName=CHINESE_FONT, fontSize=18, leading=24, alignment=TA_CENTER, spaceAfter=8),
    'subtitle': ParagraphStyle('subtitle', fontName=CHINESE_FONT, fontSize=12, leading=16, alignment=TA_CENTER, spaceAfter=4),
    'section': ParagraphStyle('section', fontName=CHINESE_FONT, fontSize=13, leading=18, alignment=TA_LEFT, spaceBefore=10, spaceAfter=4),
    'normal': ParagraphStyle('normal', fontName=CHINESE_FONT, fontSize=9, leading=12, alignment=TA_LEFT),
    'small': ParagraphStyle('small', fontName=CHINESE_FONT, fontSize=8, leading=10, alignment=TA_LEFT),
}

PHONE_CONDITION = "AND LENGTH(mobile_phone) = 11 AND mobile_phone GLOB '[0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9]'"

def get_stats(conn, region=None, street=None, building=None, is_company=None, year=None):
    """获取统计数据,返回 (总数, 有电话数)"""
    conditions = ["region IN ('盐南高新区', '经开区')"]
    params = []

    if region:
        conditions.append("region = ?")
        params.append(region)
    if street:
        conditions.append("street = ?")
        params.append(street)
    if building:
        conditions.append("building = ?")
        params.append(building)
    if is_company == True:
        conditions.append("(enterprise_name NOT LIKE '%个体%' AND enterprise_type NOT LIKE '%个体%')")
    elif is_company == False:
        conditions.append("(enterprise_name LIKE '%个体%' OR enterprise_type LIKE '%个体%')")
    if year:
        conditions.append("SUBSTR(establishment_date, 1, 4) = ?")
        params.append(str(year))

    where = "WHERE " + " AND ".join(conditions)

    total = conn.execute(f"SELECT COUNT(*) FROM enterprise_detail {where}", params).fetchone()[0]
    with_phone = conn.execute(f"SELECT COUNT(*) FROM enterprise_detail {where} {PHONE_CONDITION}", params).fetchone()[0]
    return total, with_phone

def query_region_stats(conn):
    """按区县统计"""
    regions = ['盐南高新区', '经开区']
    results = []
    for region in regions:
        total_company, phone_company = get_stats(conn, region=region, is_company=True)
        new_company_2026, _ = get_stats(conn, region=region, is_company=True, year=2026)
        total_individual, phone_individual = get_stats(conn, region=region, is_company=False)
        new_individual_2026, _ = get_stats(conn, region=region, is_company=False, year=2026)
        results.append({
            'region': region,
            'total_company': total_company,
            'phone_company': phone_company,
            'new_company_2026': new_company_2026,
            'total_individual': total_individual,
            'phone_individual': phone_individual,
            'new_individual_2026': new_individual_2026,
        })
    return results

def query_street_stats(conn):
    """按街道统计"""
    streets = [
        ('盐南高新区', '黄海街道'),
        ('盐南高新区', '新都街道'),
        ('盐南高新区', '科城街道'),
        ('盐南高新区', '新河街道'),
        ('盐南高新区', '伍佑街道'),
        ('经开区', '新城街道'),
        ('经开区', '步凤镇'),
    ]

    results = []
    for region, street in streets:
        total_company, phone_company = get_stats(conn, region=region, street=street, is_company=True)
        new_company_2026, _ = get_stats(conn, region=region, street=street, is_company=True, year=2026)
        total_individual, phone_individual = get_stats(conn, region=region, street=street, is_company=False)
        new_individual_2026, _ = get_stats(conn, region=region, street=street, is_company=False, year=2026)
        results.append({
            'region': region,
            'street': street,
            'total_company': total_company,
            'phone_company': phone_company,
            'new_company_2026': new_company_2026,
            'total_individual': total_individual,
            'phone_individual': phone_individual,
            'new_individual_2026': new_individual_2026,
        })
    return results

def query_building_stats(conn):
    """按大楼宇统计"""
    rows = conn.execute("""
        SELECT building, region, street, COUNT(*) as total
        FROM enterprise_detail
        WHERE region IN ('盐南高新区', '经开区') AND building != ''
        GROUP BY building, region, street
        ORDER BY region,
            CASE street
                WHEN '黄海街道' THEN 1
                WHEN '新都街道' THEN 2
                WHEN '科城街道' THEN 3
                WHEN '新河街道' THEN 4
                WHEN '伍佑街道' THEN 5
                WHEN '新城街道' THEN 6
                WHEN '步凤镇' THEN 7
            END,
            total DESC
    """).fetchall()

    results = []
    for row in rows:
        building, region, street = row[0], row[1], row[2]
        total_company, phone_company = get_stats(conn, building=building, is_company=True)
        new_company_2026, _ = get_stats(conn, building=building, is_company=True, year=2026)
        total_individual, phone_individual = get_stats(conn, building=building, is_company=False)
        new_individual_2026, _ = get_stats(conn, building=building, is_company=False, year=2026)
        results.append({
            'building': building,
            'region': region,
            'street': street,
            'total_company': total_company,
            'phone_company': phone_company,
            'new_company_2026': new_company_2026,
            'total_individual': total_individual,
            'phone_individual': phone_individual,
            'new_individual_2026': new_individual_2026,
        })
    return results

def auto_fit_table(data, col_widths, min_font=7, max_font=9):
    """自动调整字体大小确保表格不超出页面"""
    total = sum(col_widths)
    overshoot = total - PAGE_USABLE_WIDTH
    if overshoot > 0 and min_font < max_font:
        # 缩小字体来补偿
        scale = PAGE_USABLE_WIDTH / total
        new_widths = [int(w * scale) for w in col_widths]
        return new_widths
    return col_widths

def generate_pdf():
    """生成 PDF 报告"""
    conn = sqlite3.connect(str(DB_PATH))

    # 动态获取数据截止时间(数据库最新注册时间)
    cutoff_date = conn.execute("SELECT MAX(establishment_date) FROM enterprise_detail WHERE establishment_date != '' AND establishment_date IS NOT NULL").fetchone()[0] or '未知'

    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M')
    output_path = OUTPUT_DIR / f"天眼查汇总报告_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"

    doc = SimpleDocTemplate(
        str(output_path),
        pagesize=A4,
        leftMargin=12*mm,
        rightMargin=12*mm,
        topMargin=15*mm,
        bottomMargin=15*mm,
    )

    elements = []

    # ========== 标题 ==========
    elements.append(Paragraph("天眼查企业数据汇总报告", styles['title']))
    elements.append(Paragraph("盐南高新区 | 经开区", styles['subtitle']))
    elements.append(Paragraph(f"生成时间: {timestamp} | 数据截止时间: {cutoff_date}", styles['normal']))
    elements.append(Spacer(1, 5*mm))

    elements.append(Paragraph("一、按区县汇总", styles['section']))

    region_stats = query_region_stats(conn)

    # 两区汇总行
    total_company_all = sum(r['total_company'] for r in region_stats)
    phone_company_all = sum(r['phone_company'] for r in region_stats)
    new_company_2026_all = sum(r['new_company_2026'] for r in region_stats)
    total_individual_all = sum(r['total_individual'] for r in region_stats)
    phone_individual_all = sum(r['phone_individual'] for r in region_stats)
    new_individual_2026_all = sum(r['new_individual_2026'] for r in region_stats)

    # 汇总行（7列）
    region_summary_col_widths = [28*mm, 28*mm, 22*mm, 28*mm, 24*mm, 22*mm, 28*mm]
    check_table_width(region_summary_col_widths, "区县汇总表")

    region_summary_data = [
        ['汇总', '总企业数\n（不含个体）', '其中有\n联系方式', '26年新注册\n（不含个体）', '总个体\n工商户', '其中有\n联系方式', '26年新注册\n个体工商户'],
        ['盐南高新区\n+经开区', str(total_company_all), str(phone_company_all), str(new_company_2026_all), str(total_individual_all), str(phone_individual_all), str(new_individual_2026_all)],
    ]
    region_summary_table = Table(region_summary_data, colWidths=region_summary_col_widths)
    region_summary_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1F4E79')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('FONTNAME', (0, 0), (-1, -1), CHINESE_FONT),
        ('FONTSIZE', (0, 0), (-1, -1), 8),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ('BACKGROUND', (0, 1), (-1, 1), colors.HexColor('#E2EFDA')),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
    ]))
    elements.append(region_summary_table)
    elements.append(Spacer(1, 3*mm))

    # 分区县数据（7列）
    region_col_widths = [28*mm, 28*mm, 22*mm, 28*mm, 24*mm, 22*mm, 28*mm]
    check_table_width(region_col_widths, "区县表")

    region_data = [
        ['区县', '总企业数\n（不含个体）', '其中有\n联系方式', '26年新注册\n（不含个体）', '总个体\n工商户', '其中有\n联系方式', '26年新注册\n个体工商户'],
    ]

    for r in region_stats:
        region_data.append([
            r['region'],
            str(r['total_company']),
            str(r['phone_company']),
            str(r['new_company_2026']),
            str(r['total_individual']),
            str(r['phone_individual']),
            str(r['new_individual_2026']),
        ])

    region_table = Table(region_data, colWidths=region_col_widths)

    region_style_list = [
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#2E75B6')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('FONTNAME', (0, 0), (-1, -1), CHINESE_FONT),
        ('FONTSIZE', (0, 0), (-1, -1), 8),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
        ('TOPPADDING', (0, 0), (-1, -1), 3),
    ]

    for i, r in enumerate(region_stats, 1):
        bg = colors.HexColor('#D6E4F0') if r['region'] == '盐南高新区' else colors.HexColor('#E2EFDA')
        region_style_list.append(('BACKGROUND', (0, i), (-1, i), bg))

    region_table.setStyle(TableStyle(region_style_list))
    elements.append(region_table)

    # ========== Page 2: 街道汇总 ==========
    elements.append(Spacer(1, 8*mm))
    elements.append(Paragraph("二、按街道汇总", styles['section']))
    elements.append(Paragraph("排序规则：区县 → 街道 → 数量倒序 | 联系方式：11位手机号码", styles['small']))
    elements.append(Spacer(1, 3*mm))

    street_stats = query_street_stats(conn)

    # 两区汇总
    total_company_all_s = sum(s['total_company'] for s in street_stats)
    phone_company_all_s = sum(s['phone_company'] for s in street_stats)
    new_company_2026_all_s = sum(s['new_company_2026'] for s in street_stats)
    total_individual_all_s = sum(s['total_individual'] for s in street_stats)
    phone_individual_all_s = sum(s['phone_individual'] for s in street_stats)
    new_individual_2026_all_s = sum(s['new_individual_2026'] for s in street_stats)

    # 汇总行（7列）
    summary_col_widths = [28*mm, 28*mm, 22*mm, 28*mm, 24*mm, 22*mm, 28*mm]
    check_table_width(summary_col_widths, "汇总表")

    summary_data = [
        ['汇总', '总企业数\n（不含个体）', '其中有\n联系方式', '26年新注册\n（不含个体）', '总个体\n工商户', '其中有\n联系方式', '26年新注册\n个体工商户'],
        ['盐南高新区\n+经开区', str(total_company_all_s), str(phone_company_all_s), str(new_company_2026_all_s), str(total_individual_all_s), str(phone_individual_all_s), str(new_individual_2026_all_s)],
    ]
    summary_table = Table(summary_data, colWidths=summary_col_widths)
    summary_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1F4E79')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('FONTNAME', (0, 0), (-1, -1), CHINESE_FONT),
        ('FONTSIZE', (0, 0), (-1, -1), 8),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ('BACKGROUND', (0, 1), (-1, 1), colors.HexColor('#E2EFDA')),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
    ]))
    elements.append(summary_table)
    elements.append(Spacer(1, 3*mm))

    # 分街道数据（8列）
    street_col_widths = [25*mm, 25*mm, 25*mm, 20*mm, 25*mm, 22*mm, 20*mm, 24*mm]
    check_table_width(street_col_widths, "街道表")

    street_data = [
        ['区县', '街道', '总企业数\n（不含个体）', '其中有\n联系方式', '26年新注册\n（不含个体）', '总个体\n工商户', '其中有\n联系方式', '26年新注册\n个体工商户'],
    ]

    for s in street_stats:
        street_data.append([
            s['region'],
            s['street'],
            str(s['total_company']),
            str(s['phone_company']),
            str(s['new_company_2026']),
            str(s['total_individual']),
            str(s['phone_individual']),
            str(s['new_individual_2026']),
        ])

    street_table = Table(street_data, colWidths=street_col_widths)

    style_list = [
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#2E75B6')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('FONTNAME', (0, 0), (-1, -1), CHINESE_FONT),
        ('FONTSIZE', (0, 0), (-1, -1), 8),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
        ('TOPPADDING', (0, 0), (-1, -1), 3),
    ]

    for i, s in enumerate(street_stats, 1):
        bg = colors.HexColor('#D6E4F0') if s['region'] == '盐南高新区' else colors.HexColor('#E2EFDA')
        style_list.append(('BACKGROUND', (0, i), (-1, i), bg))

    street_table.setStyle(TableStyle(style_list))
    elements.append(street_table)

    # ========== Page 3: 大楼宇汇总（另起一页） ==========
    elements.append(PageBreak())
    elements.append(Paragraph("三、按大楼宇汇总", styles['section']))
    elements.append(Paragraph("排序规则：区县 → 街道 → 数量倒序 | 联系方式：11位手机号码", styles['small']))
    elements.append(Spacer(1, 3*mm))

    building_stats = query_building_stats(conn)

    # 大楼宇表格(9列)
    building_col_widths = [30*mm, 20*mm, 20*mm, 20*mm, 18*mm, 20*mm, 18*mm, 18*mm, 22*mm]
    check_table_width(building_col_widths, "大楼宇表")

    building_data = [
        ['大楼宇', '区县', '街道', '总企业数\n(不含个体)', '其中有\n联系方式', '26年新注册\n(不含个体)', '总个体\n工商户', '其中有\n联系方式', '26年新注册\n个体工商户'],
    ]

    for b in building_stats:
        building_data.append([
            b['building'],
            b['region'],
            b['street'],
            str(b['total_company']),
            str(b['phone_company']),
            str(b['new_company_2026']),
            str(b['total_individual']),
            str(b['phone_individual']),
            str(b['new_individual_2026']),
        ])

    building_table = Table(building_data, colWidths=building_col_widths)

    style_list2 = [
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1F4E79')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('FONTNAME', (0, 0), (-1, -1), CHINESE_FONT),
        ('FONTSIZE', (0, 0), (-1, -1), 7),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
        ('TOPPADDING', (0, 0), (-1, -1), 2),
    ]

    for i, b in enumerate(building_stats, 1):
        bg = colors.HexColor('#D6E4F0') if b['region'] == '盐南高新区' else colors.HexColor('#E2EFDA')
        style_list2.append(('BACKGROUND', (0, i), (-1, i), bg if i % 2 == 1 else colors.white))

    building_table.setStyle(TableStyle(style_list2))
    elements.append(building_table)

    # 页脚
    elements.append(Spacer(1, 5*mm))
    total_all = total_company_all + total_individual_all
    elements.append(Paragraph(f"统计范围:盐南高新区+经开区 | 总记录: {total_all} | 生成时间: {timestamp} | 数据截止时间: {cutoff_date}", styles['small']))

    doc.build(elements)
    conn.close()
    return output_path

if __name__ == "__main__":
    output_path = generate_pdf()
    print(f"✅ PDF 已生成: {output_path}")