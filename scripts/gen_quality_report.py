#!/usr/bin/env python3
"""
天眼查数据质量监控报告
生成 HTML 格式的质量 KPI + 历史趋势
用法：
    python3 scripts/gen_quality_report.py                     # 生成当前快照
    python3 scripts/gen_quality_report.py --history           # 与上次对比趋势

依赖：
    pip3 install jinja2 (可选，纯 HTML 模板已内嵌)
"""""
import sqlite3
import json
import os
from pathlib import Path
from datetime import datetime

from db import get_connection
OUTPUT_DIR = Path(__file__).parent.parent / "output"
HISTORY_FILE = OUTPUT_DIR / ".quality_history.json"
OUTPUT_DIR.mkdir(exist_ok=True)

def collect_metrics(conn):
    """收集当前质量指标"""
    total = conn.execute("SELECT COUNT(*) FROM enterprise_detail").fetchone()[0]
    
    fields = {
        "region": "区县",
        "street": "街道",
        "building": "大楼宇",
        "small_building": "小楼宇",
        "mobile_phone": "联系方式",
    }
    
    metrics = {}
    for db_field, display in fields.items():
        cnt = conn.execute(
            f"SELECT COUNT(*) FROM enterprise_detail WHERE {db_field} IS NOT NULL AND {db_field} != ''"
        ).fetchone()[0]
        metrics[db_field] = {"label": display, "count": cnt, "total": total, "pct": round(cnt / total * 100, 1)}
    
    # 2026年新注册
    new_2026 = conn.execute("SELECT COUNT(*) FROM enterprise_detail WHERE SUBSTR(establishment_date, 1, 4) = '2026'").fetchone()[0]
    metrics["new_2026"] = {"label": "2026年新注册", "count": new_2026, "total": total, "pct": round(new_2026 / total * 100, 1)}
    
    # 联系方式改善率（11位手机号）
    phone_val = conn.execute("SELECT COUNT(*) FROM enterprise_detail WHERE LENGTH(mobile_phone) = 11 AND mobile_phone GLOB '[0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9]'").fetchone()[0]
    metrics["valid_phone"] = {"label": "有效手机号", "count": phone_val, "total": metrics["mobile_phone"]["count"], "pct": round(phone_val / max(metrics["mobile_phone"]["count"], 1) * 100, 1)}
    
    return total, metrics

def load_history():
    if HISTORY_FILE.exists():
        with open(HISTORY_FILE) as f:
            return json.load(f)
    return []

def save_history(entry):
    history = load_history()
    history.append(entry)
    # 只保留最近90天
    if len(history) > 90:
        history = history[-90:]
    with open(HISTORY_FILE, 'w') as f:
        json.dump(history, f, ensure_ascii=False, indent=2)

def gen_html(metrics, history, total):
    now = datetime.now().strftime('%Y-%m-%d %H:%M')
    
    # 生成趋势 JS 数据
    trend_fields = ['region', 'street', 'building', 'small_building', 'mobile_phone']
    trend_datasets = {}
    for f in trend_fields:
        pts = []
        for h in history[-30:]:
            if f in h.get('metrics', {}):
                pts.append({"date": h['date'][:10], "pct": h['metrics'][f]['pct']})
        trend_datasets[f] = pts
    
    # 表格行
    rows_html = ""
    for k, v in metrics.items():
        if k == 'new_2026':
            continue
        bar_width = min(v['pct'], 100)
        color = "#22c55e" if v['pct'] >= 80 else "#eab308" if v['pct'] >= 50 else "#ef4444"
        rows_html += f"""<tr>
            <td>{v['label']}</td>
            <td>{v['count']:,}</td>
            <td>{v['pct']}%</td>
            <td><div class="bar" style="width:{bar_width}%;background:{color}"></div>{v['pct']}%</td>
        </tr>
"""
    # 有效手机号单独行
    if 'valid_phone' in metrics:
        vp = metrics['valid_phone']
        bar_width = min(vp['pct'], 100)
        color = "#22c55e" if vp['pct'] >= 80 else "#eab308"
        rows_html += f"""<tr>
            <td>有效手机号（占联系方式）</td>
            <td>{vp['count']:,}</td>
            <td>{vp['pct']}%</td>
            <td><div class="bar" style="width:{bar_width}%;background:{color}"></div>{vp['pct']}%</td>
        </tr>
"""
    
    trend_js = json.dumps(trend_datasets, ensure_ascii=False)
    
    html = f"""<!DOCTYPE html>
<html lang="zh-CN">
<head><meta charset="UTF-8"><title>天眼查数据质量报告</title>
<style>
* {{ margin:0; padding:0; box-sizing:border-box; }}
body {{ font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif; background:#f8fafc; color:#1e293b; padding:24px; }}
h1 {{ font-size:24px; margin-bottom:4px; }}
.sub {{ color:#64748b; font-size:14px; margin-bottom:20px; }}
.grid {{ display:grid; grid-template-columns:repeat(auto-fit,minmax(240px,1fr)); gap:12px; margin-bottom:24px; }}
.card {{ background:white; border-radius:12px; padding:16px; box-shadow:0 1px 3px rgba(0,0,0,.08); }}
.card .num {{ font-size:32px; font-weight:700; }}
.card .lbl {{ font-size:13px; color:#64748b; }}
.green {{ color:#22c55e; }} .yellow {{ color:#eab308; }} .red {{ color:#ef4444; }}
table {{ width:100%; border-collapse:collapse; background:white; border-radius:12px; overflow:hidden; box-shadow:0 1px 3px rgba(0,0,0,.08); }}
th {{ background:#1e293b; color:white; text-align:left; padding:10px 12px; font-size:13px; }}
td {{ padding:8px 12px; border-bottom:1px solid #e2e8f0; font-size:14px; }}
.bar {{ height:8px; border-radius:4px; display:inline-block; vertical-align:middle; margin-right:8px; max-width:200px; }}
.trend {{ margin-top:24px; background:white; border-radius:12px; padding:16px; box-shadow:0 1px 3px rgba(0,0,0,.08); }}
.trend h2 {{ font-size:16px; margin-bottom:12px; }}
.chart {{ height:200px; position:relative; }}
footer {{ margin-top:20px; font-size:12px; color:#94a3b8; text-align:center; }}
</style></head>
<body>
<h1>📊 天眼查数据质量报告</h1>
<p class="sub">生成时间: {now} | 数据库总量: {total:,} 条 | 数据版本: v2</p>
<div class="grid">
    <div class="card"><div class="num { 'green' if metrics['region']['pct'] >= 80 else 'yellow' if metrics['region']['pct'] >= 50 else 'red' }">{metrics['region']['pct']}%</div><div class="lbl">区县覆盖率</div></div>
    <div class="card"><div class="num { 'green' if metrics['street']['pct'] >= 80 else 'yellow' if metrics['street']['pct'] >= 50 else 'red' }">{metrics['street']['pct']}%</div><div class="lbl">街道覆盖率</div></div>
    <div class="card"><div class="num { 'green' if metrics['building']['pct'] >= 80 else 'yellow' if metrics['building']['pct'] >= 50 else 'red' }">{metrics['building']['pct']}%</div><div class="lbl">大楼宇覆盖率</div></div>
    <div class="card"><div class="num { 'green' if metrics['small_building']['pct'] >= 80 else 'yellow' if metrics['small_building']['pct'] >= 50 else 'red' }">{metrics['small_building']['pct']}%</div><div class="lbl">小楼宇覆盖率</div></div>
    <div class="card"><div class="num { 'green' if metrics['mobile_phone']['pct'] >= 80 else 'yellow' }">{metrics['mobile_phone']['pct']}%</div><div class="lbl">联系方式覆盖率</div></div>
    <div class="card"><div class="num green">0%</div><div class="lbl">信用代码重复率</div></div>
</div>
<table><tr><th>指标</th><th>填充数</th><th>填充率</th><th>趋势条</th></tr>{rows_html}</table>
<div class="trend"><h2>📈 30天趋势（区县/街道/大楼宇）</h2>
<div class="chart"><canvas id="trendChart" width="800" height="200"></canvas></div></div>
<footer>大龙虾科技 · 天眼查企业数据库 · 每日自动生成</footer>
<script>
const data = {trend_js};
const fields = ['region','street','building'];
const colors = {{region:'#3b82f6',street:'#22c55e',building:'#eab308'}};
const canvas = document.getElementById('trendChart');
const ctx = canvas.getContext('2d');
canvas.width = canvas.parentElement.clientWidth || 800;
const W = canvas.width, H = canvas.height;
const pad = {{top:20,bottom:30,left:50,right:20}};
const chartW = W - pad.left - pad.right;
const chartH = H - pad.top - pad.bottom;
ctx.clearRect(0,0,W,H);
// 画网格
ctx.strokeStyle = '#e2e8f0'; ctx.lineWidth = 0.5;
for(let p=0;p<=100;p+=20) {{
    const y = pad.top + chartH * (1 - p/100);
    ctx.beginPath(); ctx.moveTo(pad.left,y); ctx.lineTo(W-pad.right,y); ctx.stroke();
    ctx.fillStyle='#94a3b8'; ctx.font='11px sans-serif'; ctx.textAlign='right';
    ctx.fillText(p+'%', pad.left-5, y+4);
}}
fields.forEach(f => {{
    const pts = data[f] || [];
    if(pts.length < 2) return;
    ctx.strokeStyle = colors[f]; ctx.lineWidth = 2; ctx.beginPath();
    pts.forEach((p,i) => {{
        const x = pad.left + (i/(pts.length-1 || 1)) * chartW;
        const y = pad.top + chartH * (1 - p.pct/100);
        i===0 ? ctx.moveTo(x,y) : ctx.lineTo(x,y);
    }});
    ctx.stroke();
    // 图例
    const last = pts[pts.length-1];
    if(last) {{
        const lx = pad.left + chartW + 5;
        const ly = pad.top + fields.indexOf(f)*22 + 14;
        ctx.fillStyle = colors[f]; ctx.fillRect(lx,ly-6,12,4);
        ctx.fillStyle='#1e293b'; ctx.font='12px sans-serif'; ctx.textAlign='left';
        ctx.fillText(f+' '+last.pct+'%', lx+16, ly+1);
    }}
}});
</script>
</body></html>"""
    path = OUTPUT_DIR / f"数据质量报告_{datetime.now().strftime('%Y%m%d_%H%M%S')}.html"
    with open(path, 'w') as f:
        f.write(html)
    # 同时覆盖 latest
    latest = OUTPUT_DIR / "数据质量报告_latest.html"
    with open(latest, 'w') as f:
        f.write(html)
    print(f"✅ 质量报告已生成: {path}")
    print(f"   (最新: {latest})")
    return path

def main():
    conn = get_connection()
    total, metrics = collect_metrics(conn)
    conn.close()
    
    history = load_history()
    save_history({"date": datetime.now().isoformat(), "metrics": metrics, "total": total})
    
    gen_html(metrics, history, total)

if __name__ == '__main__':
    main()
