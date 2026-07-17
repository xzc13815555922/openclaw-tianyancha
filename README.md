# 天眼查企业数据管理 Skill

> 大龙虾科技公司 · 天眼查专项项目  
> 负责盐南高新区 + 经开区企业全量数据的导入、标注、查询和报告生成

## 数据概况

| 项目 | 数值 |
|------|------|
| 数据库总量 | 200,001 条 |
| 数据来源 | 天眼查商业查询平台 |
| 标注范围 | 盐南高新区 + 经开区 |
| 已标注区域 | 70,528 条 (35.3%) |
| 联系方式覆盖 | 115,166 条 (57.6%) |

## 技能架构

### 主要脚本

| 脚本 | 功能 | 产出 |
|------|------|------|
| `scripts/import_excel.py` | Excel 数据导入 + 统一社会信用代码去重 + 自动标注 | 入库 |
| `scripts/annotate_all.py` | 地址标注：区县/街道/大楼宇/小楼宇 | 治理字段填充 |
| `scripts/query.py` | 多条件查询统计 + Excel 44列导出 | 控制台统计 / Excel |
| `scripts/query_summary_pdf.py` | PDF 汇总报告（按街道+大楼宇维度） | PDF |
| `scripts/query_pdf.py` | 多条件筛选明细 PDF 报告 | PDF |
| `scripts/gen_report.py` | HTML 数据治理报告 | HTML |

### 数据库

- **位置**：`tianyancha.db`（SQLite，约 1.2GB）
- **主表**：`enterprise_detail`（45 列）
- **核心字段**：企业名称、统一社会信用代码、法定代表人、注册地址、联系方式、行业分类
- **数据治理字段**：`region`（区县）、`street`（街道）、`building`（大楼宇）、`small_building`（小楼宇）
- **去重规则**：以 `unified_social_credit_code` 为唯一键，重复覆盖

## 使用方式

### 导入新数据

```bash
cd ~/.openclaw/plugin-skills/tianyancha
python3 scripts/import_excel.py /path/to/【天眼查】高级搜索YYYYMMDD(...).xlsx
```

导入后自动触发地址标注。

### 查询与报表

```bash
# 全量统计
python3 scripts/query.py

# 按区域筛选
python3 scripts/query.py --region 盐南高新区

# 按大楼宇筛选 + Excel 导出
python3 scripts/query.py --building 金融城 --output excel

# 组合筛选（2026年6月起新注册，排除个体户）
python3 scripts/query.py --region 盐南高新区 --year 2026 --from-month 6 --not 个体 --output excel

# PDF 汇总报告
python3 scripts/query_summary_pdf.py

# 明细 PDF
python3 scripts/query_pdf.py --region 盐南高新区 --year 2026 --not 个体
```

### 查询参数

| 参数 | 说明 | 示例 |
|------|------|------|
| `--region` | 按区县筛选 | `--region 盐南高新区` |
| `--street` | 按街道筛选 | `--street 黄海街道` |
| `--building` | 按大楼宇筛选 | `--building 金融城` |
| `--small_building` | 按小楼宇筛选 | `--small_building 大数据B12号楼` |
| `--year` | 按注册年份筛选 | `--year 2026` |
| `--month` | 按注册月份筛选（需配合--year） | `--month 6` |
| `--from-month` | 按月份起筛选（需配合--year） | `--from-month 6` |
| `--not` | 排除关键词 | `--not 个体` |
| `--output` | 输出格式（excel/pdf） | `--output excel` |
| `--group` | 分组统计 | `--group street` |
| `--limit` | 限制条数 | `--limit 100` |

### Excel 输出规范

- **列数**：44 列（7 固定列序 + 37 其余字段）
- **列序**：区县 → 街道 → 大楼宇 → 小楼宇 → 注册时间 → 企业名称 → 联系号码 → ... → 注册地址 → ... → 入库时间
- **排序**：区县 → 街道 → 大楼宇 → 注册时间倒序
- **文件命名**：`{区县}_{街道}_{大楼宇}_{年份}年{月起}新注册_{是否排除个体}_{数量}条.xlsx`

## 地址标注规则

### 区县匹配

- `county` = 盐城经济技术开发区 → 经开区
- `county` in (亭湖区, 盐都区) AND 地址含"盐南"/"城南" → 盐南高新区
- `street` 映射区域（黄海/新都/科城/新河/伍佑 → 盐南高新区，新城/步凤 → 经开区）

### 大楼宇匹配（27 栋）

**盐南高新区（20 栋）**：金融城、大数据产业园、新龙广场、华邦国际、紫薇广场、金鹰天地、钱江财富广场、西伏河园区、国际创投中心、凤凰文化广场、财富港、中信大厦、中远商务楼、新都社区商务楼、鹿鸣广场、香城美地、棕榈泉、丽晶大厦、南海流量经济港、伍佑新型显示产业园

**经开区（7 栋）**：未来科技城、全民创业园（步凤）、涌鑫经贸中心、韩资工业园、光电产业园、新能源汽车产业园、盐城综合保税区

## 依赖

- Python 3.8+
- SQLite 3
- openpyxl（Excel 导出）
- reportlab（PDF 生成）
- Pillow

## 项目来源

大龙虾科技 · 天眼查企业数据库项目  
数据仅用于内部业务分析
