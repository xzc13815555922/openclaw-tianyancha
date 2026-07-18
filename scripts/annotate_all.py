#!/usr/bin/env python3
"""
天眼查数据标注模块
功能：根据地址字段，标注 region、street、building、small_building

匹配规则基于：天眼查匹配规则审.xlsx

使用方式：
    python3 scripts/annotate_all.py              # 全量标注（按顺序执行Step1-8）
    python3 scripts/gen_report.py                # 生成汇总报告

执行顺序：
    Step 1: 清空 region, street, building, small_building
    Step 2: 区县匹配（county + address关键词）
    Step 3: 街道匹配（address关键词）
    Step 4: 街道补匹配区县（街道→区域映射）
    Step 5: 大楼宇匹配（address关键词）
    Step 6: 大楼宇补匹配region/street（大楼宇固定映射，强制覆盖）
    Step 7: 同一大楼宇多街道修正（按规则统一）
    Step 8: 小楼宇匹配（特殊规则+通用规则）
"""

import sqlite3
import re
import sys
from pathlib import Path
import yaml

# 规则配置文件路径
RULES_PATH = Path(__file__).parent.parent / "annotate_rules.yaml"

def load_rules():
    """从 YAML 加载标注规则"""
    with open(RULES_PATH, "r") as f:
        return yaml.safe_load(f)

DB_PATH = Path(__file__).parent.parent / "tianyancha.db"

# ── 区县匹配规则 ─────────────────────────────────────────────────────────────
# 规则1: county = 盐城经济技术开发区 → 经开区
# 规则2: county in (亭湖区, 盐都区) AND address含盐南/城南 → 盐南高新区

# ── 街道/乡镇映射 ────────────────────────────────────────────────────────────
# ── 街道规则（从 YAML 配置加载）────────────────────────────────
_RULES = load_rules()
STREET_RULES = [(s['name'], s['keywords']) for s in _RULES.get('streets', [])]
STREET_TO_REGION = {s['name']: s['region'] for s in _RULES.get('streets', [])}

# ── 大楼宇规则（从 YAML 配置加载）────────────────────────────────
# building_rules: [(name, keywords, optional_street_cond)]
BUILDING_RULES = []
for b in _RULES.get('buildings', []):
    item = [b['name'], b['keywords']]
    if 'street_cond' in b:
        item.append(f"street={b['street_cond']}")
    BUILDING_RULES.append(tuple(item) if len(item) > 2 else (b['name'], b['keywords']))

# building → (region, street) 映射
BUILDING_REGION_MAP = {b['name']: (b['region'], b['street']) for b in _RULES.get('buildings', [])}

# 同一大楼宇多街道修正（从 YAML 配置加载）────────────────────
BUILDING_STREET_FIX = _RULES.get('street_fix', {})

# ── 小楼宇提取逻辑 ─────────────────────────────────────────────────────────────

# ── 小楼宇匹配（每个大楼宇独立函数 + 调度字典）────────────────

def _single_building_simple(building):
    """通用单楼宇处理：返回楼宇名本身"""
    return building

def _extract_guoji_chuangtou(addr):
    """国际创投中心"""
    if "南楼" in addr:
        return "国际创投中心南楼"
    if "北楼" in addr:
        return "国际创投中心北楼"
    return "国际创投中心"

def _extract_ziwei(addr):
    """紫薇广场"""
    addr = addr.replace("紫微", "紫薇")
    if "B幢" in addr or "紫薇广场B区" in addr:
        return "紫薇广场B区"
    if "5幢" in addr or "紫薇广场A区" in addr:
        return "紫薇广场A区"
    if "A区" in addr:
        return "紫薇广场A区"
    if "C区" in addr or "C3区" in addr or "C1区" in addr or "C座" in addr or "C2区" in addr:
        return "紫薇广场C区"
    if "G区" in addr:
        return "紫薇广场G区"
    m = re.search(r"紫薇广场(\d+)号楼", addr)
    if m:
        return f"紫薇广场{m.group(1)}号楼"
    m = re.search(r"紫薇广场(\d+)幢", addr)
    if m:
        return f"紫薇广场{m.group(1)}号楼"
    if "紫薇国际广场" in addr:
        if "C座" in addr:
            return "紫薇广场C区"
        if "C2区" in addr:
            return "紫薇广场C区"
    return ""

def _extract_yongxin(addr):
    """涌鑫经贸中心"""
    m = re.search(r"涌鑫.+?(\d+)[号楼幢]", addr)
    if m:
        return f"涌鑫经贸中心{m.group(1)}号楼"
    return ""

def _extract_bigdata(addr):
    """大数据产业园"""
    if "创新大厦南楼" in addr:
        return "创新大厦南楼"
    if "创新大厦北楼" in addr:
        return "创新大厦北楼"
    if "创新大厦A楼" in addr:
        return "创新大厦A楼"
    if "创新大厦B楼" in addr:
        return "创新大厦B楼"
    if "创新大厦" in addr:
        return "创新大厦"
    m = re.search(r"学海路29号(\d+)幢(北楼|南楼)?", addr)
    if m:
        n, s = m.group(1), (m.group(2) or "")
        return f"大数据{n}号楼{s}"
    m = re.search(r"数梦小镇(\d+)[号楼幢栋]", addr)
    if m:
        return f"数梦小镇{m.group(1)}号楼"
    m = re.search(r"学海路29号数梦小镇(\d+)幢", addr)
    if m:
        return f"数梦小镇{m.group(1)}号楼"
    if "苏港产业创新中心" in addr or "无人系统与苏港合作产业区" in addr:
        m = re.search(r"(?:苏港产业创新中心|无人系统与苏港合作产业区)(\d+)号楼", addr)
        if m:
            return f"苏港产业创新中心{m.group(1)}号楼"
        return "苏港产业创新中心"
    addr_up = addr.upper()
    m = re.search(r"大数据产业园([A-Z])-?(\d+)号楼", addr_up)
    if m:
        return f"大数据{m.group(1)}{m.group(2)}号楼"
    m = re.search(r"大数据产业园([A-Z]\d+)号楼", addr_up)
    if m:
        return f"大数据{m.group(1)}号楼"
    m = re.search(r"([A-Z]\d+)楼", addr_up)
    if m:
        return f"大数据{m.group(1)}号楼"
    m = re.search(r"([A-Z]-?\d+)栋", addr_up)
    if m:
        return f"大数据{m.group(1).replace('-', '')}号楼"
    m = re.search(r"大数据产业园([A-Z]\d+)幢", addr_up)
    if m:
        return f"大数据{m.group(1)}号楼"
    m = re.search(r"大数据产业园([A-Z])区(\d+)栋", addr_up)
    if m:
        return f"大数据{m.group(1)}{m.group(2)}号楼"
    m = re.search(r"大数据产业园(?:北区|南区|北)([A-Z]\d+)", addr_up)
    if m:
        return f"大数据{m.group(1)}号楼"
    m = re.search(r"[数大]据产业园([A-Z]\d+)[^\d]*?层", addr_up)
    if m:
        return f"大数据{m.group(1)}号楼"
    if "科创大厦" in addr:
        if "南楼" in addr:
            return "科创大厦南楼"
        if "北楼" in addr:
            return "科创大厦北楼"
        return "科创大厦"
    if "网易联合创新中心" in addr:
        return "网易联合创新中心"
    if "菁英公寓" in addr:
        return "菁英公寓"
    m = re.search(r"A26(东侧|西侧)单元", addr_up)
    if m:
        return f"大数据A26号楼{m.group(1)}单元"
    if "卫生室" in addr or "警务室" in addr:
        return ""
    return ""

def _extract_huabang(addr):
    """华邦国际"""
    for zone in ["A区", "B区", "C区"]:
        if "西厦" in addr and zone in addr:
            return f"华邦西厦{zone}"
        if "东厦" in addr and zone in addr:
            return f"华邦东厦{zone}"
    if "西厦" in addr:
        m = re.search(r"(\d+)幢", addr)
        if m:
            return f"华邦西厦{m.group(1)}号楼"
        m = re.search(r"(\d+)栋", addr)
        if m:
            return f"华邦西厦{m.group(1)}号楼"
        return "华邦西厦"
    if "东厦" in addr:
        m = re.search(r"(\d+)幢", addr)
        if m:
            return f"华邦东厦{m.group(1)}号楼"
        m = re.search(r"(\d+)栋", addr)
        if m:
            return f"华邦东厦{m.group(1)}号楼"
        return "华邦东厦"
    if "西楼" in addr:
        return "华邦西楼"
    if "东楼" in addr:
        return "华邦东楼"
    return ""

def _extract_fenghuang(addr):
    """凤凰文化广场"""
    m = re.search(r"凤凰文化广场(\d+)[幢号楼#]", addr)
    if m:
        return f"凤凰文化广场{m.group(1)}号楼"
    m = re.search(r"凤凰文化广场(\d+)#楼", addr)
    if m:
        return f"凤凰文化广场{m.group(1)}号楼"
    return ""

def _extract_jinying(addr):
    """金鹰天地"""
    m = re.search(r"金鹰天地(\d+)[号楼幢]", addr)
    if m: return f"金鹰天地{m.group(1)}号楼"
    m = re.search(r"金鹰天地(\d+)#", addr)
    if m: return f"金鹰天地{m.group(1)}号楼"
    m = re.search(r"金鹰龙湖商业广场([A-Z])幢", addr)
    if m: return f"金鹰天地{m.group(1)}号楼"
    m = re.search(r"(\d+)[号楼幢]", addr)
    if m: return f"金鹰天地{m.group(1)}号楼"
    if "金鹰国际购物中心" in addr or "金鹰购物中心" in addr or "金鹰聚龙湖购物中心" in addr:
        m = re.search(r"([A-Z]?\d+[F层楼])", addr)
        if m: return f"金鹰购物中心{m.group(1)}"
        return "金鹰聚龙湖购物中心"
    m = re.search(r"金鹰天地.+(\d+)[F层]", addr)
    if m: return f"金鹰天地{m.group(1)}号楼"
    return ""

def _extract_qianjiang(addr):
    """钱江财富广场"""
    if "钱江方洲小区" in addr:
        return ""
    m = re.search(r"钱江财富广场[第]?(\d+)[幢号楼]", addr)
    if m: return f"钱江财富广场{m.group(1)}号楼"
    if "钱江商业街" in addr:
        return "钱江商业街"
    return ""

def _extract_xifuhe(addr):
    """西伏河园区"""
    m = re.search(r"西伏河[^\d]*([A-Z]\d+)#?号?[楼栋]", addr)
    if m: return f"西伏河{m.group(1)}号楼"
    m = re.search(r"西伏河[^\d]*([A-Z]\d)楼", addr)
    if m: return f"西伏河{m.group(1)}号楼"
    m = re.search(r"西伏河[^\d]*(\d+)号楼", addr)
    if m: return f"西伏河{m.group(1)}号楼"
    m = re.search(r"西伏河[^\d]*(\d+)幢", addr)
    if m: return f"西伏河{m.group(1)}号楼"
    m = re.search(r"西伏河[^\d]*([A-Z]\d)#号楼", addr)
    if m: return f"西伏河{m.group(1)}号楼"
    m = re.search(r"西伏河[^\d]*(\d+)[栋楼]", addr)
    if m: return f"西伏河{m.group(1)}号楼"
    m = re.search(r"西伏河园区([A-Z]\d+)厂房", addr)
    if m: return f"西伏河{m.group(1)}厂房"
    m = re.search(r"文港南路(?:49|77)号(\d+)号楼", addr)
    if m: return f"西伏河{m.group(1)}号楼"
    m = re.search(r"文港南路(?:49|77)号(\d+)幢", addr)
    if m: return f"西伏河{m.group(1)}号楼"
    cn_num_xf = {"一": "1", "二": "2", "三": "3", "四": "4", "五": "5", "六": "6",
                  "七": "7", "八": "8", "九": "9", "十": "10"}
    for cn, num in cn_num_xf.items():
        if f"{cn}幢" in addr: return f"西伏河{num}号楼"
        if f"{cn}号楼" in addr and ("文港南路49号" in addr or "文港南路77号" in addr):
            return f"西伏河{num}号楼"
    m = re.search(r"文港南路49号(\d+)号楼", addr)
    if m: return f"西伏河{m.group(1)}号楼"
    if "创客小镇" in addr and "文港南路49号" in addr:
        m = re.search(r"创客小镇.*?(\d+)幢", addr)
        if m: return f"西伏河{m.group(1)}号楼"
        return ""
    m = re.search(r"文港南路75号[^\d]*(\d+)", addr)
    if m: return f"西伏河{m.group(1)}号楼"
    if "机器人产业园" in addr or "绿色低碳科创园" in addr or "机器人产业集聚区" in addr:
        m = re.search(r"([A-Z]\d+)#?[号楼栋]", addr)
        if m: return f"西伏河{m.group(1)}号楼"
        m = re.search(r"(\d+)[栋楼]", addr)
        if m: return f"西伏河{m.group(1)}号楼"
        return ""
    if "产学研协同创新中心" in addr:
        m = re.search(r"中心(\d+)", addr)
        if m: return f"西伏河{m.group(1)}号楼"
    if "展示中心楼" in addr: return "西伏河展示中心楼"
    if "组楼" in addr: return "西伏河组楼"
    m = re.search(r"文港南路(?:49|77)号[^\d]*(\d+)室", addr)
    if m: return ""
    if "文港南路77号" in addr or "文港南路49号" in addr:
        if not re.search(r"[号楼幢栋]", addr):
            return ""
    return ""

def _extract_future(addr):
    """未来科技城"""
    addr_stripped = addr.replace("（CNX）", "").replace("(CNX)", "")
    m = re.search(r"(?:国际)?软件园(\d+)[幢号楼]", addr_stripped)
    if m: return f"未来科技城{m.group(1)}号楼"
    m = re.search(r"未来科技城([A-Z])[座楼]", addr_stripped)
    if m: return f"未来科技城{m.group(1)}座"
    m = re.search(r"([A-Z]\d)楼", addr_stripped)
    if m: return f"未来科技城{m.group(1)}号楼"
    m = re.search(r"未来科技城(\d+)[幢号楼]", addr_stripped)
    if m: return f"未来科技城{m.group(1)}号楼"
    m = re.search(r"希望大道南路\d+号(\d+)[幢号楼]", addr_stripped)
    if m: return f"未来科技城{m.group(1)}号楼"
    if "希望大道南路5号" in addr_stripped:
        return ""
    return ""

def _extract_xinlong(addr):
    """新龙广场"""
    cn_num = {"一": "1", "二": "2", "三": "3", "四": "4", "五": "5",
              "六": "6", "七": "7", "八": "8", "九": "9", "十": "10"}
    for cn, num in cn_num.items():
        if f"新龙广场{cn}号楼" in addr or f"新龙广场{cn}幢" in addr or f"新龙广场{cn}号" in addr:
            return f"新龙广场{num}号楼"
    for pat in [r"新龙广场(\d+)号\d+楼", r"新龙广场([A-Z]\d)楼", r"新龙广场(\d+)#?楼",
                r"新龙广场(\d+)幢", r"新龙广场(\d+)-(\d+)", r"新龙广场(\d+)号",
                r"盐城新龙广场(\d+)号楼", r"新龙广场(\d+)#"]:
        m = re.search(pat, addr)
        if m:
            g = m.groups()
            return f"新龙广场{g[0]}号楼"
    for pat in [r"新龙(?:商务中心|广场)?([D]\d*)(?:座|楼|#楼)?", r"新龙(?:商务中心|广场)?([B]\d*)楼"]:
        m = re.search(pat, addr)
        if m: return f"新龙广场{m.group(1)}号楼"
    m = re.search(r"人民南路38号(\d+)幢", addr)
    if m: return f"新龙广场{m.group(1)}号楼"
    if "人力资源服务产业园" in addr: return "盐城人力资源服务产业园"
    if "港府洲际酒店" in addr or "港府洲际酒店服务公寓" in addr:
        m = re.search(r"(\d+)号楼港府洲际酒店", addr)
        if m: return f"新龙广场{m.group(1)}号楼"
        return "港府洲际酒店"
    if "新弄里" in addr:
        m = re.search(r"新弄里(\d+)[号楼幢]", addr)
        if m: return f"新龙广场{m.group(1)}号楼"
        m = re.search(r"新弄里(\d+)-(\d+)", addr)
        if m: return f"新龙广场{m.group(1)}号楼"
    return ""

def _extract_jinrong(addr):
    """金融城"""
    cn_num = {"一": "1", "二": "2", "三": "3", "四": "4", "五": "5",
              "六": "6", "七": "7", "八": "8", "九": "9", "十": "10",
              "十一": "11", "十二": "12", "十三": "13", "十四": "14", "十五": "15",
              "十六": "16", "十七": "17"}
    if "金融城二期" in addr:
        m = re.search(r"金融城二期(\d+)号楼", addr)
        if m: return f"金融城{m.group(1)}号楼"
    m = re.search(r"金融城(\d+)[号楼]", addr)
    if m: return f"金融城{m.group(1)}号楼"
    m = re.search(r"金融城第?(\d+)[幢撞]", addr)
    if m: return f"金融城{m.group(1)}号楼"
    m = re.search(r"金融城(\d+)栋", addr)
    if m: return f"金融城{m.group(1)}号楼"
    m = re.search(r"金融城(\d+)#", addr)
    if m: return f"金融城{m.group(1)}号楼"
    for cn, num in cn_num.items():
        if f"金融城{cn}号楼" in addr or f"金融城{cn}幢" in addr:
            return f"金融城{num}号楼"
    if "世纪大道5号" in addr or "戴庄路2号" in addr:
        m = re.search(r"金融城[^\d]*?(\d+)[#幢号楼栋]", addr)
        if m: return f"金融城{m.group(1)}号楼"
    m = re.search(r"金融城(\d+)-\d+[室号]", addr)
    if m: return f"金融城{m.group(1)}号楼"
    m = re.search(r"金融城(\d+)-(\d+)号楼", addr)
    if m: return f"金融城{m.group(1)}-{m.group(2)}号楼"
    m = re.search(r"金融城(\d+)号", addr)
    if m: return f"金融城{m.group(1)}号楼"
    m = re.search(r"金融智慧谷(\d+)#?楼", addr)
    if m: return f"金融城{m.group(1)}号楼"
    if "金融新天地" in addr or "金融城新天地" in addr:
        m = re.search(r"金融新天地(?:二期|2期)?(?:商业)?(\d+)号楼", addr)
        if m: return f"金融城{m.group(1)}号楼"
        m = re.search(r"金融(?:城|新天地)?新天地(?:二期)?(\d+)[#号楼幢]", addr)
        if m: return f"金融城{m.group(1)}号楼"
        return ""
    if "A栋" in addr or "A幢" in addr:
        return "金融城A栋"
    m = re.search(r"世纪大道5号(\d+)幢", addr)
    if m: return f"金融城{m.group(1)}号楼"
    if "滨河漫步道" in addr or "圆房子" in addr or "对面" in addr or "水幕墙" in addr:
        return ""
    return ""

def _extract_xindu(addr):
    """新都社区商务楼"""
    if "香苑小区" in addr: return ""
    m = re.search(r"新都社区商务楼(\d+)楼", addr)
    if m: return f"新都社区商务楼{m.group(1)}号楼"
    m = re.search(r"新都商务楼(\d+)楼", addr)
    if m: return f"新都社区商务楼{m.group(1)}号楼"
    return ""

def _extract_zonglvquan(addr):
    """棕榈泉"""
    m = re.search(r"棕榈泉广场(\d+-\d+)[幢栋]", addr)
    if m: return f"棕榈泉{m.group(1)}号楼"
    m = re.search(r"棕榈泉广场(\d+)[号楼幢]", addr)
    if m: return f"棕榈泉{m.group(1)}号楼"
    if "商务楼" in addr: return "棕榈泉商务楼"
    if "4#" in addr: return "棕榈泉4#楼"
    return ""

def _extract_quanmin(addr):
    """全民创业园"""
    m = re.search(r"(\d+)#\d+#?", addr)
    if m: return f"全民创业园{m.group(1)}号楼"
    m = re.search(r"全民创业园(\d+)号楼", addr)
    if m: return f"全民创业园{m.group(1)}号楼"
    return "全民创业园"

def _extract_caifugang(addr):
    """财富港"""
    m = re.search(r"依云小镇商城(\d+)[幢号楼]", addr)
    if m: return f"财富港{m.group(1)}号楼"
    m = re.search(r"财富港(\d+)[幢号楼]", addr)
    if m: return f"财富港{m.group(1)}号楼"
    return "财富港"

def _extract_hanzi(addr):
    """韩资工业园"""
    m = re.search(r"韩资工业园(\d+)幢", addr)
    if m: return f"韩资工业园{m.group(1)}号楼"
    m = re.search(r"韩资工业园(\d+)号楼", addr)
    if m: return f"韩资工业园{m.group(1)}号楼"
    m = re.search(r"标准厂房(\d+)", addr)
    if m: return f"韩资工业园{m.group(1)}号标准厂房"
    m = re.search(r"#标厂", addr)
    if m: return "韩资工业园标厂"
    if "邻里中心" in addr: return "韩资工业园邻里中心"
    return ""

def _extract_guangdian(addr):
    """光电产业园"""
    m = re.search(r"光电产业园(\d+)幢", addr)
    if m: return f"光电产业园{m.group(1)}号楼"
    m = re.search(r"光电产业园(\d+)号楼", addr)
    if m: return f"光电产业园{m.group(1)}号楼"
    if "研发中心" in addr: return "光电产业园研发中心"
    if "办公楼" in addr: return "光电产业园办公楼"
    return ""

def _extract_xinnengyuan(addr):
    """新能源汽车产业园"""
    m = re.search(r"新能源汽车产业园(\d+)幢", addr)
    if m: return f"新能源汽车产业园{m.group(1)}号楼"
    m = re.search(r"新能源汽车产业园(\d+)号楼", addr)
    if m: return f"新能源汽车产业园{m.group(1)}号楼"
    if "研发" in addr: return "新能源汽车产业园研发中心"
    if "办公" in addr: return "新能源汽车产业园办公楼"
    return ""

def _extract_wuyouxianshi(addr):
    """伍佑新型显示产业园"""
    if "伍佑科技园" in addr or "园区路9号" in addr:
        m = re.search(r"厂房(\d+)[-—]?(\d*)", addr)
        if m:
            unit = m.group(1)
            suffix = m.group(2)
            return f"伍佑科技园厂房{unit}-{suffix}室" if suffix else f"伍佑科技园厂房{unit}室"
        return ""
    m = re.search(r"构港村(?:\D{1,2})组(?:\d+)号?(\d+)幢", addr)
    if m: return f"构港村{m.group(1)}号楼"
    if "构港村部" in addr or "构港村办公楼" in addr: return "构港村办公楼"
    m = re.search(r"园区路(\d+)号", addr)
    if m: return ""
    m = re.search(r"工业园区(\d+)号", addr)
    if m: return ""
    return ""

def _extract_baoshui(addr):
    """盐城综合保税区"""
    if "保税区大楼" in addr or "保税区商务楼" in addr: return "保税区商务楼"
    if "综合保税区商务楼" in addr: return "综合保税区商务楼"
    m = re.search(r"(\d+)号标准厂房", addr)
    if m: return f"{m.group(1)}号标准厂房"
    m = re.search(r"(\d+)号保税仓库", addr)
    if m: return f"{m.group(1)}号保税仓库"
    m = re.search(r"C区(\d+)号标准厂房", addr)
    if m: return f"C区{m.group(1)}号厂房"
    m = re.search(r"C区([A-Z]\d+)标准厂房", addr)
    if m: return f"C区{m.group(1)}号厂房"
    if "罗马路" in addr or "纽约路" in addr or "纽约北路" in addr or "纽约中路" in addr or "柏林路" in addr or "香港路" in addr:
        return ""
    return ""

def _extract_nanhai(addr):
    """南海流量经济港"""
    m = re.search(r"南海流量经济港(\d+)号楼", addr)
    if m: return f"南海流量经济港{m.group(1)}号楼"
    m = re.search(r"南海流量经济港(\d+)层", addr)
    if m: return ""
    cn_num = {"一": "1", "二": "2", "两": "2"}
    for cn, num in cn_num.items():
        if f"{cn}号楼" in addr: return f"南海流量经济港{num}号楼"
    if "B栋" in addr: return "南海流量经济港B栋"
    return ""

def _extract_xiangcheng(addr):
    """香城美地"""
    m = re.search(r"香城美地花园(\d+)幢", addr)
    if m: return f"香城美地{m.group(1)}号楼"
    m = re.search(r"香城美地(\d+-\d+)", addr)
    if m: return f"香城美地{m.group(1)}"
    m = re.search(r"香城美地(\d+)", addr)
    if m: return f"香城美地{m.group(1)}号楼"
    return ""

def _small_building_fallback(addr, building):
    """通用兜底规则"""
    b_pos = addr.find(building)
    if b_pos == -1:
        return ""
    remainder = addr[b_pos + len(building):]
    for pat in [r"(\d+)号楼", r"(\d+)幢", r"(\d+)栋", r"(\d+)#楼"]:
        m = re.search(pat, remainder)
        if m:
            num = m.group(1)
            if len(num) >= 4:  # 超过3位基本是房间号
                continue
            return building + num + "号楼"
    return ""

# 小楼宇提取调度字典
SMALL_BUILDING_DISPATCH = {
    "中信大厦": lambda a: "中信大厦",
    "鹿鸣广场": lambda a: "鹿鸣广场",
    "中远商务楼": lambda a: "中远商务楼",
    "丽晶大厦": lambda a: "丽晶大厦",
    "国际创投中心": _extract_guoji_chuangtou,
    "紫薇广场": _extract_ziwei,
    "涌鑫经贸中心": _extract_yongxin,
    "大数据产业园": _extract_bigdata,
    "华邦国际": _extract_huabang,
    "凤凰文化广场": _extract_fenghuang,
    "金鹰天地": _extract_jinying,
    "钱江财富广场": _extract_qianjiang,
    "西伏河园区": _extract_xifuhe,
    "未来科技城": _extract_future,
    "新龙广场": _extract_xinlong,
    "金融城": _extract_jinrong,
    "新都社区商务楼": _extract_xindu,
    "棕榈泉": _extract_zonglvquan,
    "全民创业园（步凤）": _extract_quanmin,
    "财富港": _extract_caifugang,
    "韩资工业园": _extract_hanzi,
    "光电产业园": _extract_guangdian,
    "新能源汽车产业园": _extract_xinnengyuan,
    "伍佑新型显示产业园": _extract_wuyouxianshi,
    "盐城综合保税区": _extract_baoshui,
    "南海流量经济港": _extract_nanhai,
    "香城美地": _extract_xiangcheng,
}


def extract_small_building(addr, building):
    """从注册地址中提取小楼宇名称（调度到各楼宇专用函数）"""
    if not addr or not building:
        return ""
    addr = str(addr)
    handler = SMALL_BUILDING_DISPATCH.get(building)
    if handler:
        return handler(addr)

    # 不在调度表中的大楼宇，使用通用兜底规则
    # （应列在大楼宇匹配规则中，此处兜底）
    return _small_building_fallback(addr, building)

    # 国际创投中心 - 南北楼
    if building == "国际创投中心":
        if "南楼" in addr:
            return "国际创投中心南楼"
        if "北楼" in addr:
            return "国际创投中心北楼"
        return "国际创投中心"

    # 紫薇广场 - A/B/C区
    if building == "紫薇广场":
        # 注意：可能有错字"紫微"
        addr_fixed = addr.replace("紫微", "紫薇")
        # B幢/区
        if "B幢" in addr_fixed or "紫薇广场B区" in addr_fixed:
            return "紫薇广场B区"
        if "5幢" in addr_fixed or "紫薇广场A区" in addr_fixed:
            return "紫薇广场A区"
        if "A区" in addr_fixed:
            return "紫薇广场A区"
        if "C区" in addr_fixed or "C3区" in addr_fixed or "C1区" in addr_fixed or "C座" in addr_fixed or "C2区" in addr_fixed:
            return "紫薇广场C区"
        if "G区" in addr_fixed:
            return "紫薇广场G区"
        m = re.search(r"紫薇广场(\d+)号楼", addr_fixed)
        if m:
            return f"紫薇广场{m.group(1)}号楼"
        # 紫薇广场X幢格式
        m = re.search(r"紫薇广场(\d+)幢", addr_fixed)
        if m:
            return f"紫薇广场{m.group(1)}号楼"
        # 紫薇国际广场特殊
        if "紫薇国际广场" in addr_fixed:
            if "C座" in addr_fixed:
                return "紫薇广场C区"
            if "C2区" in addr_fixed:
                return "紫薇广场C区"
        return ""

    # 涌鑫经贸中心 - 子名称统一
    if building == "涌鑫经贸中心":
        m = re.search(r"涌鑫.+?(\d+)[号楼幢]", addr)
        if m:
            return f"涌鑫经贸中心{m.group(1)}号楼"
        return ""

    # 大数据 - 创新大厦子楼宇
    if building == "大数据产业园":
        if "创新大厦南楼" in addr:
            return "创新大厦南楼"
        if "创新大厦北楼" in addr:
            return "创新大厦北楼"
        if "创新大厦A楼" in addr:
            return "创新大厦A楼"
        if "创新大厦B楼" in addr:
            return "创新大厦B楼"
        if "创新大厦" in addr:
            return "创新大厦"
        # 学海路29号X幢格式（含北楼/南楼后缀）
        m = re.search(r"学海路29号(\d+)幢(北楼|南楼)?", addr)
        if m:
            num = m.group(1)
            suffix = m.group(2) or ""
            return f"大数据{num}号楼{suffix}"
        # 数梦小镇X号楼/幢/栋
        m = re.search(r"数梦小镇(\d+)[号楼幢栋]", addr)
        if m:
            return f"数梦小镇{m.group(1)}号楼"
        m = re.search(r"学海路29号数梦小镇(\d+)幢", addr)
        if m:
            return f"数梦小镇{m.group(1)}号楼"
        # 苏港产业创新中心
        if "苏港产业创新中心" in addr or "无人系统与苏港合作产业区" in addr:
            m = re.search(r"(?:苏港产业创新中心|无人系统与苏港合作产业区)(\d+)号楼", addr)
            if m:
                return f"苏港产业创新中心{m.group(1)}号楼"
            return "苏港产业创新中心"
        # 以下匹配使用大写归一化处理
        addr_up = addr.upper()
        # 大数据产业园A-8号楼、B-21号楼等带连字符格式
        m = re.search(r"大数据产业园([A-Z])-?(\d+)号楼", addr_up)
        if m:
            return f"大数据{m.group(1)}{m.group(2)}号楼"
        # A1号楼格式
        m = re.search(r"大数据产业园([A-Z]\d+)号楼", addr_up)
        if m:
            return f"大数据{m.group(1)}号楼"
        # B12二层203室, A1号楼格式
        m = re.search(r"([A-Z]\d+)楼", addr_up)
        if m:
            return f"大数据{m.group(1)}号楼"
        # 栋格式：B15栋、A8栋、B-11栋、A-1栋等
        m = re.search(r"([A-Z]-?\d+)栋", addr_up)
        if m:
            num = m.group(1).replace("-", "")
            return f"大数据{num}号楼"
        # A8幢（幢代替楼）
        m = re.search(r"大数据产业园([A-Z]\d+)幢", addr_up)
        if m:
            return f"大数据{m.group(1)}号楼"
        # B区10栋格式
        m = re.search(r"大数据产业园([A-Z])区(\d+)栋", addr_up)
        if m:
            return f"大数据{m.group(1)}{m.group(2)}号楼"
        # 北区A8/南区A9格式
        m = re.search(r"大数据产业园(?:北区|南区|北)([A-Z]\d+)", addr_up)
        if m:
            return f"大数据{m.group(1)}号楼"
        # 字母+数字+层（无楼/栋后缀）
        m = re.search(r"[数大]据产业园([A-Z]\d+)[^\d]*?层", addr_up)
        if m:
            return f"大数据{m.group(1)}号楼"
        # 科创大厦
        if "科创大厦" in addr:
            if "南楼" in addr:
                return "科创大厦南楼"
            if "北楼" in addr:
                return "科创大厦北楼"
            return "科创大厦"
        # 网易联合创新中心
        if "网易联合创新中心" in addr:
            return "网易联合创新中心"
        # 菁英公寓
        if "菁英公寓" in addr:
            return "菁英公寓"
        # A26东侧/西侧单元
        m = re.search(r"A26(东侧|西侧)单元", addr_up)
        if m:
            return f"大数据A26号楼{m.group(1)}单元"
        # 卫生室、警务室等公共设施 - 不映射具体楼栋
        if "卫生室" in addr or "警务室" in addr:
            return ""
        return ""

    # 华邦
    if building == "华邦国际":
        # 注意：地址可能是"华邦国际西厦"而非"华邦西厦"
        for zone in ["A区", "B区", "C区"]:
            if "西厦" in addr and zone in addr:
                return f"华邦西厦{zone}"
            if "东厦" in addr and zone in addr:
                return f"华邦东厦{zone}"
        if "西厦" in addr:
            m = re.search(r"(\d+)幢", addr)
            if m:
                return f"华邦西厦{m.group(1)}号楼"
            m = re.search(r"(\d+)栋", addr)
            if m:
                return f"华邦西厦{m.group(1)}号楼"
            return "华邦西厦"
        if "东厦" in addr:
            m = re.search(r"(\d+)幢", addr)
            if m:
                return f"华邦东厦{m.group(1)}号楼"
            m = re.search(r"(\d+)栋", addr)
            if m:
                return f"华邦东厦{m.group(1)}号楼"
            return "华邦东厦"
        if "西楼" in addr:
            return "华邦西楼"
        if "东楼" in addr:
            return "华邦东楼"
        return ""

    # 凤凰文化广场
    if building == "凤凰文化广场":
        m = re.search(r"凤凰文化广场(\d+)[幢号楼#]", addr)
        if m:
            return f"凤凰文化广场{m.group(1)}号楼"
        m = re.search(r"凤凰文化广场(\d+)#楼", addr)
        if m:
            return f"凤凰文化广场{m.group(1)}号楼"
        return ""

    # 金鹰天地
    if building == "金鹰天地":
        m = re.search(r"金鹰天地(\d+)[号楼幢]", addr)
        if m:
            return f"金鹰天地{m.group(1)}号楼"
        m = re.search(r"金鹰天地(\d+)#", addr)
        if m:
            return f"金鹰天地{m.group(1)}号楼"
        # 金鹰龙湖商业广场A幢/B幢
        m = re.search(r"金鹰龙湖商业广场([A-Z])幢", addr)
        if m:
            return f"金鹰天地{m.group(1)}号楼"
        m = re.search(r"(\d+)[号楼幢]", addr)
        if m:
            return f"金鹰天地{m.group(1)}号楼"
        if "金鹰国际购物中心" in addr or "金鹰购物中心" in addr or "金鹰聚龙湖购物中心" in addr:
            m = re.search(r"([A-Z]?\d+[F层楼])", addr)
            if m:
                return f"金鹰购物中心{m.group(1)}"
            return "金鹰聚龙湖购物中心"
        m = re.search(r"金鹰天地.+(\d+)[F层]", addr)
        if m:
            return f"金鹰天地{m.group(1)}号楼"
        return ""

    # 钱江财富广场
    if building == "钱江财富广场":
        if "钱江方洲小区" in addr:
            return ""
        m = re.search(r"钱江财富广场[第]?(\d+)[幢号楼]", addr)
        if m:
            return f"钱江财富广场{m.group(1)}号楼"
        if "钱江商业街" in addr:
            return "钱江商业街"
        return ""

    # 西伏河园区
    if building == "西伏河园区":
        # 含"西伏河"字样的地址：B11楼、数字号楼等
        m = re.search(r"西伏河[^\d]*([A-Z]\d+)#?号?[楼栋]", addr)
        if m:
            return f"西伏河{m.group(1)}号楼"
        m = re.search(r"西伏河[^\d]*([A-Z]\d)楼", addr)
        if m:
            return f"西伏河{m.group(1)}号楼"
        m = re.search(r"西伏河[^\d]*(\d+)号楼", addr)
        if m:
            return f"西伏河{m.group(1)}号楼"
        m = re.search(r"西伏河[^\d]*(\d+)幢", addr)
        if m:
            return f"西伏河{m.group(1)}号楼"
        m = re.search(r"西伏河[^\d]*([A-Z]\d)#号楼", addr)
        if m:
            return f"西伏河{m.group(1)}号楼"
        m = re.search(r"西伏河[^\d]*(\d+)[栋楼]", addr)
        if m:
            return f"西伏河{m.group(1)}号楼"
        # C26厂房格式
        m = re.search(r"西伏河园区([A-Z]\d+)厂房", addr)
        if m:
            return f"西伏河{m.group(1)}厂房"
        # 文港南路49号/77号格式：匹配X号楼/X幢
        m = re.search(r"文港南路(?:49|77)号(\d+)号楼", addr)
        if m:
            return f"西伏河{m.group(1)}号楼"
        m = re.search(r"文港南路(?:49|77)号(\d+)幢", addr)
        if m:
            return f"西伏河{m.group(1)}号楼"
        # 二幢/三幢（中文数字）
        cn_num_map_xf = {"一": "1", "二": "2", "三": "3", "四": "4", "五": "5", "六": "6",
                          "七": "7", "八": "8", "九": "9", "十": "10"}
        for cn, num in cn_num_map_xf.items():
            if f"文港南路(?:49|77)号{cn}幢" or f"{cn}幢" in addr:
                if f"{cn}幢" in addr:
                    return f"西伏河{num}号楼"
            if f"{cn}号楼" in addr and ("文港南路49号" in addr or "文港南路77号" in addr):
                return f"西伏河{num}号楼"
        # 文港南路49号+数字号楼（如从地址文港南路49号提取）
        m = re.search(r"文港南路49号(\d+)号楼", addr)
        if m:
            return f"西伏河{m.group(1)}号楼"
        # 创客小镇格式
        if "创客小镇" in addr and "文港南路49号" in addr:
            m = re.search(r"创客小镇.*?(\d+)幢", addr)
            if m:
                return f"西伏河{m.group(1)}号楼"
            if "三单元" in addr or "二单元" in addr:
                return ""
        # 文港南路75号格式
        m = re.search(r"文港南路75号[^\d]*(\d+)", addr)
        if m:
            return f"西伏河{m.group(1)}号楼"
        # 机器人产业园/绿色低碳科创园格式
        if "机器人产业园" in addr or "绿色低碳科创园" in addr or "机器人产业集聚区" in addr:
            m = re.search(r"([A-Z]\d+)#?[号楼栋]", addr)
            if m:
                return f"西伏河{m.group(1)}号楼"
            m = re.search(r"(\d+)[栋楼]", addr)
            if m:
                return f"西伏河{m.group(1)}号楼"
            return ""
        # 产学研协同创新中心
        if "产学研协同创新中心" in addr:
            m = re.search(r"中心(\d+)", addr)
            if m:
                return f"西伏河{m.group(1)}号楼"
        if "展示中心楼" in addr:
            return "西伏河展示中心楼"
        if "组楼" in addr:
            return "西伏河组楼"
        # 仅文港南路XX号-XX室，无楼栋号
        m = re.search(r"文港南路(?:49|77)号[^\d]*(\d+)室", addr)
        if m:
            return ""
        # 纯门牌号无楼栋号
        if "文港南路77号" in addr or "文港南路49号" in addr:
            m = re.search(r"文港南路77号(\d+)$", addr.replace("（CNX）", "").replace("(CNX)", ""))
            if m:
                return ""
            if "文港南路77号" in addr and not re.search(r"[号楼幢栋]", addr):
                return ""
            if "文港南路49号" in addr and not re.search(r"[号楼幢栋]", addr):
                return ""
        return ""

    # 未来科技城 - 修复：优先匹配软件园/未来科技城后面的楼号，排除"希望大道南路5号"
    if building == "未来科技城":
        # 优先：国际软件园/软件园 + 楼号
        m = re.search(r"(?:国际)?软件园(\d+)[幢号楼]", addr)
        if m:
            return f"未来科技城{m.group(1)}号楼"
        # C2楼/C3楼格式或B座格式
        m = re.search(r"未来科技城([A-Z])[座楼]", addr)
        if m:
            return f"未来科技城{m.group(1)}座"
        m = re.search(r"([A-Z]\d)楼", addr)
        if m:
            return f"未来科技城{m.group(1)}号楼"
        # 未来科技城X号楼/幢
        m = re.search(r"未来科技城(\d+)[幢号楼]", addr)
        if m:
            return f"未来科技城{m.group(1)}号楼"
        # 希望大道+楼栋（排除，只匹配5号以后的楼栋）
        m = re.search(r"希望大道南路\d+号(\d+)[幢号楼]", addr)
        if m:
            return f"未来科技城{m.group(1)}号楼"
        # 仅门牌号（希望大道南路5号+房间号）— 这类多是住宅/小商铺，无明确楼栋号
        if "希望大道南路5号" in addr:
            return ""
        return ""

    # 新龙广场（含新弄里别名）
    if building == "新龙广场":
        # 中文数字（四号楼/五号楼）
        cn_num_map_newlong = {"一": "1", "二": "2", "三": "3", "四": "4", "五": "5", "六": "6",
                              "七": "7", "八": "8", "九": "9", "十": "10"}
        for cn, num in cn_num_map_newlong.items():
            if f"新龙广场{cn}号楼" in addr or f"新龙广场{cn}幢" in addr or f"新龙广场{cn}号" in addr:
                return f"新龙广场{num}号楼"
        # 新龙广场X号/号楼/幢/-#楼格式
        m = re.search(r"新龙广场(\d+)号\d+楼", addr)
        if m:
            return f"新龙广场{m.group(1)}号楼"
        m = re.search(r"新龙广场([A-Z]\d)楼", addr)
        if m:
            return f"新龙广场{m.group(1)}号楼"
        m = re.search(r"新龙广场(\d+)#?楼", addr)
        if m:
            return f"新龙广场{m.group(1)}号楼"
        m = re.search(r"新龙广场(\d+)幢", addr)
        if m:
            return f"新龙广场{m.group(1)}号楼"
        m = re.search(r"新龙广场(\d+)-(\d+)", addr)
        if m:
            return f"新龙广场{m.group(1)}号楼"
        m = re.search(r"新龙广场(\d+)号", addr)
        if m:
            return f"新龙广场{m.group(1)}号楼"
        # "盐城新龙广场"10号楼 格式
        m = re.search(r"盐城新龙广场(\d+)号楼", addr)
        if m:
            return f"新龙广场{m.group(1)}号楼"
        # 新龙广场X#格式（如12#905-07室）
        m = re.search(r"新龙广场(\d+)#", addr)
        if m:
            return f"新龙广场{m.group(1)}号楼"
        # D1座/D1楼/D2#楼/D2楼/B1楼/B3楼
        m = re.search(r"新龙(?:商务中心|广场)?([D]\d*)(?:座|楼|#楼)?", addr)
        if m:
            return f"新龙广场{m.group(1)}号楼"
        m = re.search(r"新龙(?:商务中心|广场)?([B]\d*)楼", addr)
        if m:
            return f"新龙广场{m.group(1)}号楼"
        # 人民南路38号X幢
        m = re.search(r"人民南路38号(\d+)幢", addr)
        if m:
            return f"新龙广场{m.group(1)}号楼"
        # 人力资源服务产业园
        if "人力资源服务产业园" in addr:
            return "盐城人力资源服务产业园"
        # 港府洲际酒店
        if "港府洲际酒店" in addr or "港府洲际酒店服务公寓" in addr:
            m = re.search(r"(\d+)号楼港府洲际酒店", addr)
            if m:
                return f"新龙广场{m.group(1)}号楼"
            return "港府洲际酒店"
        # 新弄里别名
        if "新弄里" in addr:
            m = re.search(r"新弄里(\d+)号楼", addr)
            if m:
                return f"新龙广场{m.group(1)}号楼"
            m = re.search(r"新弄里(\d+)幢", addr)
            if m:
                return f"新龙广场{m.group(1)}号楼"
            m = re.search(r"新弄里(\d+)-(\d+)", addr)
            if m:
                return f"新龙广场{m.group(1)}号楼"
        return ""

    # 金融城 - 新增规则
    if building == "金融城":
        # 金融城二期X号楼
        if "金融城二期" in addr:
            m = re.search(r"金融城二期(\d+)号楼", addr)
            if m:
                return f"金融城{m.group(1)}号楼"
        # 金融城X号楼
        m = re.search(r"金融城(\d+)[号楼]", addr)
        if m:
            return f"金融城{m.group(1)}号楼"
        # 金融城第X幢 / 金融城X幢 / 金融城X撞（错别字）
        m = re.search(r"金融城第?(\d+)[幢撞]", addr)
        if m:
            return f"金融城{m.group(1)}号楼"
        # 金融城X栋
        m = re.search(r"金融城(\d+)栋", addr)
        if m:
            return f"金融城{m.group(1)}号楼"
        # 金融城X#XXXX室
        m = re.search(r"金融城(\d+)#", addr)
        if m:
            return f"金融城{m.group(1)}号楼"
        # 中文数字
        cn_num_map = {"一": "1", "二": "2", "三": "3", "四": "4", "五": "5", "六": "6", 
                      "七": "7", "八": "8", "九": "9", "十": "10", "十一": "11", "十二": "12",
                      "十三": "13", "十四": "14", "十五": "15", "十六": "16", "十七": "17"}
        for cn, num in cn_num_map.items():
            if f"金融城{cn}号楼" in addr or f"金融城{cn}幢" in addr:
                return f"金融城{num}号楼"
        # 世纪大道5号+金融城后面的数字楼栋
        if "世纪大道5号" in addr or "戴庄路2号" in addr:
            m = re.search(r"金融城[^\d]*?(\d+)[#幢号楼栋]", addr)
            if m:
                return f"金融城{m.group(1)}号楼"
        # 金融城X-XXXX室/号（如金融城6-1810室、金融城3-1号）
        m = re.search(r"金融城(\d+)-\d+[室号]", addr)
        if m:
            return f"金融城{m.group(1)}号楼"
        # 金融城X-X号楼（如金融城12-1号楼、金融城12-2号楼、金融城1-2号楼）
        m = re.search(r"金融城(\d+)-(\d+)号楼", addr)
        if m:
            return f"金融城{m.group(1)}-{m.group(2)}号楼"
        # 金融城X号
        m = re.search(r"金融城(\d+)号", addr)
        if m:
            return f"金融城{m.group(1)}号楼"
        # 金融智慧谷X#楼
        m = re.search(r"金融智慧谷(\d+)#?楼", addr)
        if m:
            return f"金融城{m.group(1)}号楼"
        # 金融新天地/金融城新天地（商铺为主）
        if "金融新天地" in addr or "金融城新天地" in addr:
            m = re.search(r"金融新天地(?:二期|2期)?(?:商业)?(\d+)号楼", addr)
            if m:
                return f"金融城{m.group(1)}号楼"
            m = re.search(r"金融(?:城|新天地)?新天地(?:二期)?(\d+)[#号楼幢]", addr)
            if m:
                return f"金融城{m.group(1)}号楼"
            # 二期XX室/商铺时无明确楼栋号
            return ""
        # A栋/A幢
        if "A栋" in addr or "A幢" in addr:
            return "金融城A栋"
        # 世纪大道5号X幢
        m = re.search(r"世纪大道5号(\d+)幢", addr)
        if m:
            return f"金融城{m.group(1)}号楼"
        # 无明确楼栋的地标
        if "滨河漫步道" in addr or "圆房子" in addr or "对面" in addr or "水幕墙" in addr:
            return ""
        return ""

    # 新都社区商务楼 - 新增"新都商务楼"别名
    if building == "新都社区商务楼":
        if "香苑小区" in addr:
            return ""
        # 优先匹配"新都社区商务楼"后面的楼号
        m = re.search(r"新都社区商务楼(\d+)楼", addr)
        if m:
            return f"新都社区商务楼{m.group(1)}号楼"
        # 新都商务楼别名
        m = re.search(r"新都商务楼(\d+)楼", addr)
        if m:
            return f"新都社区商务楼{m.group(1)}号楼"
        return ""

    # 棕榈泉
    if building == "棕榈泉":
        m = re.search(r"棕榈泉广场(\d+-\d+)[幢栋]", addr)
        if m:
            return f"棕榈泉{m.group(1)}号楼"
        m = re.search(r"棕榈泉广场(\d+)[号楼幢]", addr)
        if m:
            return f"棕榈泉{m.group(1)}号楼"
        if "商务楼" in addr:
            return "棕榈泉商务楼"
        if "4#" in addr:
            return "棕榈泉4#楼"
        return ""

    # 全民创业园
    if building == "全民创业园（步凤）":
        m = re.search(r"(\d+)#\d+#?", addr)
        if m:
            return f"全民创业园{m.group(1)}号楼"
        m = re.search(r"全民创业园(\d+)号楼", addr)
        if m:
            return f"全民创业园{m.group(1)}号楼"
        return "全民创业园"

    # 财富港 - 修复：只匹配"依云小镇商城"/"财富港"后面的楼号
    if building == "财富港":
        # 优先：依云小镇商城后面的楼号
        m = re.search(r"依云小镇商城(\d+)[幢号楼]", addr)
        if m:
            return f"财富港{m.group(1)}号楼"
        # 财富港X号楼
        m = re.search(r"财富港(\d+)[幢号楼]", addr)
        if m:
            return f"财富港{m.group(1)}号楼"
        return "财富港"

    # 韩资工业园 - 新增规则
    if building == "韩资工业园":
        m = re.search(r"韩资工业园(\d+)幢", addr)
        if m:
            return f"韩资工业园{m.group(1)}号楼"
        m = re.search(r"韩资工业园(\d+)号楼", addr)
        if m:
            return f"韩资工业园{m.group(1)}号楼"
        m = re.search(r"标准厂房(\d+)", addr)
        if m:
            return f"韩资工业园{m.group(1)}号标准厂房"
        m = re.search(r"#标厂", addr)
        if m:
            return "韩资工业园标厂"
        if "邻里中心" in addr:
            return "韩资工业园邻里中心"
        return ""

    # 光电产业园 - 新增规则
    if building == "光电产业园":
        m = re.search(r"光电产业园(\d+)幢", addr)
        if m:
            return f"光电产业园{m.group(1)}号楼"
        m = re.search(r"光电产业园(\d+)号楼", addr)
        if m:
            return f"光电产业园{m.group(1)}号楼"
        if "研发中心" in addr:
            return "光电产业园研发中心"
        if "办公楼" in addr:
            return "光电产业园办公楼"
        return ""

    # 新能源汽车产业园 - 新增规则
    if building == "新能源汽车产业园":
        m = re.search(r"新能源汽车产业园(\d+)幢", addr)
        if m:
            return f"新能源汽车产业园{m.group(1)}号楼"
        m = re.search(r"新能源汽车产业园(\d+)号楼", addr)
        if m:
            return f"新能源汽车产业园{m.group(1)}号楼"
        if "研发" in addr:
            return "新能源汽车产业园研发中心"
        if "办公" in addr:
            return "新能源汽车产业园办公楼"
        return ""

    # 伍佑新型显示产业园 - 新增规则
    if building == "伍佑新型显示产业园":
        # 园区路9号伍佑科技园厂房
        if "伍佑科技园" in addr or "园区路9号" in addr:
            m = re.search(r"厂房(\d+)[-—]?(\d*)", addr)
            if m:
                unit = m.group(1)
                suffix = m.group(2)
                if suffix:
                    return f"伍佑科技园厂房{unit}-{suffix}室"
                return f"伍佑科技园厂房{unit}室"
            return ""  # 科技园无厂房编号时留空
        # 构港村X组X号X幢
        m = re.search(r"构港村(?:\D{1,2})组(?:\d+)号?(\d+)幢", addr)
        if m:
            return f"构港村{m.group(1)}号楼"
        # 构港村部办公楼
        if "构港村部" in addr or "构港村办公楼" in addr:
            return "构港村办公楼"
        # 构港村园区路X号（无楼栋号）
        m = re.search(r"园区路(\d+)号", addr)
        if m:
            return ""  # 仅门牌号，无明确楼栋
        # 构港村工业园区X号
        m = re.search(r"工业园区(\d+)号", addr)
        if m:
            return ""
        # 构港村X组（纯组地址，无楼栋）
        return ""

    # 盐城综合保税区 - 新增规则
    if building == "盐城综合保税区":
        # 保税区大楼/商务楼+房间号
        if "保税区大楼" in addr or "保税区商务楼" in addr:
            return "保税区商务楼"
        if "综合保税区商务楼" in addr:
            return "综合保税区商务楼"
        # X号标准厂房
        m = re.search(r"(\d+)号标准厂房", addr)
        if m:
            return f"{m.group(1)}号标准厂房"
        # X号保税仓库
        m = re.search(r"(\d+)号保税仓库", addr)
        if m:
            return f"{m.group(1)}号保税仓库"
        # C区标准厂房
        m = re.search(r"C区(\d+)号标准厂房", addr)
        if m:
            return f"C区{m.group(1)}号厂房"
        m = re.search(r"C区([A-Z]\d+)标准厂房", addr)
        if m:
            return f"C区{m.group(1)}号厂房"
        # 国际道路（罗马路、纽约路、柏林路等）+ 门牌号 -> 无明确楼栋
        if "罗马路" in addr or "纽约路" in addr or "纽约北路" in addr or "纽约中路" in addr or "柏林路" in addr or "香港路" in addr:
            return ""
        return ""

    # 南海流量经济港 - 新增规则
    if building == "南海流量经济港":
        # X号楼（阿拉伯数字）
        m = re.search(r"南海流量经济港(\d+)号楼", addr)
        if m:
            return f"南海流量经济港{m.group(1)}号楼"
        m = re.search(r"南海流量经济港(\d+)层", addr)
        if m:
            return ""  # 无明确楼栋号
        # 一号楼/二号楼（中文数字）
        m = re.search(r"([一二两])号楼", addr)
        cn_num = {"一": "1", "二": "2", "两": "2"}
        for cn, num in cn_num.items():
            if f"{cn}号楼" in addr:
                return f"南海流量经济港{num}号楼"
        # B栋
        if "B栋" in addr:
            return "南海流量经济港B栋"
        return ""

    # 香城美地 - 新增规则（地址格式：香城美地花园X幢XX室）
    if building == "香城美地":
        m = re.search(r"香城美地花园(\d+)幢", addr)
        if m:
            return f"香城美地{m.group(1)}号楼"
        m = re.search(r"香城美地(\d+-\d+)", addr)
        if m:
            return f"香城美地{m.group(1)}"
        m = re.search(r"香城美地(\d+)", addr)
        if m:
            return f"香城美地{m.group(1)}号楼"
        return ""

    # 丽晶大厦 - 新增规则（单楼宇）
    if building == "丽晶大厦":
        return "丽晶大厦"

    # 通用规则（兜底）
    b_pos = addr.find(building)
    if b_pos == -1:
        return ""

    remainder = addr[b_pos + len(building):]

    for pat in [r"(\d+)号楼", r"(\d+)幢", r"(\d+)栋", r"(\d+)#楼"]:
        m = re.search(pat, remainder)
        if m:
            num = m.group(1)
            rest = remainder[m.end():]
            # 号楼后面跟4位纯数字且后跟非"室"才跳过（是房间号不是楼栋号）
            # 如"3304-219室"中的"3304"是房间号前缀但要跳过
            # "6幢2103室"中的"2103"不是楼号而是房间号，但也需要跳过吗？
            # 不对！"6幢2103室"中"6幢"本身就是楼栋号，"2103室"是房间号
            # 真正的楼栋号应该只有1-2位数字，超过3位的基本都是房间号
            if len(num) >= 4:
                continue
            return ""

def main():
    conn = sqlite3.connect(str(DB_PATH))
    try:
        _annotate(conn)
    finally:
        conn.close()

def _annotate(conn):
    print("=" * 50)
    print("  天眼查数据标注（修复版）")
    print("=" * 50)

    # Step 1: 清空
    print("\nStep 1: 清空数据治理字段...")
    conn.execute("UPDATE enterprise_detail SET region = '', street = '', building = '', small_building = ''")
    conn.commit()

    # Step 2: 区县匹配
    print("Step 2: 区县匹配...")
    # 规则1: county=盐城经济技术开发区
    conn.execute("""
        UPDATE enterprise_detail 
        SET region = '经开区' 
        WHERE county = '盐城经济技术开发区'
    """)
    # 规则2: county in (亭湖区, 盐都区) AND address含盐南/城南
    conn.execute("""
        UPDATE enterprise_detail 
        SET region = '盐南高新区' 
        WHERE region = '' 
          AND county IN ('亭湖区', '盐都区')
          AND (registered_address LIKE '%盐南%' OR registered_address LIKE '%城南%')
    """)
    # 规则3: 白名单地址补经开区（排除非盐城经开的开发区）
    conn.execute("""
        UPDATE enterprise_detail 
        SET region = '经开区' 
        WHERE (region IS NULL OR region = '')
          AND (
              registered_address LIKE '%盐城经济技术开发区%'
              OR registered_address LIKE '%盐城经济开发区%'
              OR registered_address LIKE '%盐城市经济技术开发区%'
              OR registered_address LIKE '%盐城市经济开发区%'
              OR registered_address LIKE '%江苏省盐城市经济开发区%'
              OR registered_address LIKE '%江苏省盐城市经济技术开发区%'
          )
          AND registered_address NOT LIKE '%亭湖经济开发区%'
          AND registered_address NOT LIKE '%南洋经济开发区%'
          AND registered_address NOT LIKE '%南亭湖经济开发区%'
          AND registered_address NOT LIKE '%新洋经济开发区%'
          AND registered_address NOT LIKE '%盐东经济开发区%'
          AND registered_address NOT LIKE '%农村经济开发区%'
          AND registered_address NOT LIKE '%大丰经济开发区%'
          AND registered_address NOT LIKE '%建湖经济开发区%'
          AND registered_address NOT LIKE '%阜宁经济开发区%'
          AND registered_address NOT LIKE '%滨海经济开发区%'
          AND registered_address NOT LIKE '%射阳经济开发区%'
          AND registered_address NOT LIKE '%东台经济开发区%'
          AND registered_address NOT LIKE '%响水经济开发区%'
    """)
    conn.commit()
    r_count = conn.execute("SELECT COUNT(*) FROM enterprise_detail WHERE region != ''").fetchone()[0]
    print(f"  区县已填充: {r_count}")

    # Step 3: 街道匹配
    print("Step 3: 街道匹配...")
    for street, keywords in STREET_RULES:
        for kw in keywords:
            conn.execute("UPDATE enterprise_detail SET street = ? WHERE street = '' AND registered_address LIKE ?", (street, f"%{kw}%"))
    conn.commit()
    s_count = conn.execute("SELECT COUNT(*) FROM enterprise_detail WHERE street != ''").fetchone()[0]
    print(f"  街道已填充: {s_count}")

    # Step 4: 街道补匹配区县
    print("Step 4: 街道补匹配区县...")
    for street, region in STREET_TO_REGION.items():
        conn.execute("UPDATE enterprise_detail SET region = ? WHERE street = ? AND (region = '' OR region IS NULL)", (region, street))
    conn.commit()
    r_count2 = conn.execute("SELECT COUNT(*) FROM enterprise_detail WHERE region != ''").fetchone()[0]
    print(f"  区县已填充（含补匹配）: {r_count2}")

    # Step 5: 大楼宇匹配
    print("Step 5: 大楼宇匹配...")
    for item in BUILDING_RULES:
        building = item[0]
        kws = item[1]
        street_cond = item[2] if len(item) > 2 else None
        for kw in kws:
            if street_cond and "步凤" in street_cond:
                conn.execute("UPDATE enterprise_detail SET building = ? WHERE building = '' AND street = '步凤镇' AND registered_address LIKE ?", (building, f"%{kw}%"))
            else:
                if building == "大数据产业园":
                    conn.execute("UPDATE enterprise_detail SET building = ? WHERE building = '' AND registered_address LIKE ? AND registered_address NOT LIKE '%张庄%' AND registered_address NOT LIKE '%盐龙%'", (building, f"%{kw}%"))
                else:
                    conn.execute("UPDATE enterprise_detail SET building = ? WHERE building = '' AND registered_address LIKE ?", (building, f"%{kw}%"))
    conn.commit()
    b_count = conn.execute("SELECT COUNT(*) FROM enterprise_detail WHERE building != ''").fetchone()[0]
    print(f"  大楼宇已填充: {b_count}")

    # Step 6: 大楼宇补匹配region/street（强制覆盖，确保所有building记录都有region和street）
    print("Step 6: 大楼宇补匹配region/street（强制覆盖）...")
    for building, (region, street) in BUILDING_REGION_MAP.items():
        conn.execute("UPDATE enterprise_detail SET region = ?, street = ? WHERE building = ?", (region, street, building))
    conn.commit()

    # Step 7: 同一大楼宇多街道修正（强制统一）
    print("Step 7: 同一大楼宇多街道修正...")
    for building, correct_street in BUILDING_STREET_FIX.items():
        count = conn.execute("SELECT COUNT(*) FROM enterprise_detail WHERE building = ?", (building,)).fetchone()[0]
        if count > 0:
            conn.execute("UPDATE enterprise_detail SET street = ? WHERE building = ?", (correct_street, building))
            print(f"  修正 {building}: {count}条 → {correct_street}")

    conn.commit()

    # Step 8: 小楼宇匹配
    print("Step 8: 小楼宇匹配...")
    rows = conn.execute("SELECT id, registered_address, building FROM enterprise_detail WHERE building != ''").fetchall()
    matched = 0
    for row_id, addr, building in rows:
        sb = extract_small_building(addr, building)
        if sb:
            conn.execute("UPDATE enterprise_detail SET small_building = ? WHERE id = ?", (sb, row_id))
            matched += 1
    conn.commit()
    print(f"  小楼宇已匹配: {matched}")

    # 统计汇总
    print("\n" + "=" * 50)
    print("  标注完成")
    print("=" * 50)
    total = conn.execute("SELECT COUNT(*) FROM enterprise_detail").fetchone()[0]
    for field, name in [("region", "区县"), ("street", "街道"), ("building", "大楼宇"), ("small_building", "小楼宇")]:
        cnt = conn.execute(f"SELECT COUNT(*) FROM enterprise_detail WHERE {field} != ''").fetchone()[0]
        pct = cnt / total * 100
        print(f"  {name}: {cnt:,} / {total:,} ({pct:.1f}%)")

if __name__ == "__main__":
    main()