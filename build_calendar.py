"""新版日历:8/20 起整周 + 9 月,带 Excel/CSV 导出,数据全部从 DB 取"""
import sqlite3, json, calendar, base64, io
from datetime import date, timedelta
from collections import Counter

DB = "/workspace/cninfo_shareholders.db"
OUT_HTML = "/workspace/calendar_v2.html"

# 视图范围:8/24(周一)起到 9/30,大约 5 周多
VIEW_START = date(2026, 8, 24)
VIEW_END = date(2026, 9, 30)

con = sqlite3.connect(DB)
con.row_factory = sqlite3.Row
cur = con.cursor()

# 拉视图范围内的所有会议(包含两份 source_run 合并)
cur.execute("""
    SELECT m.meeting_date, m.meeting_time, m.weekday_zh, m.ampm,
           m.location_raw, m.province, m.city, m.district,
           c.sec_code, c.sec_name, a.title, a.pdf_url, a.publish_date
    FROM meeting m
    JOIN announcement a ON a.id = m.announcement_id
    JOIN company c ON c.sec_code = a.sec_code
    WHERE m.meeting_date IS NOT NULL
      AND m.meeting_date BETWEEN ? AND ?
    ORDER BY m.meeting_date, m.meeting_time
""", (VIEW_START.isoformat(), VIEW_END.isoformat()))

rows = cur.fetchall()
con.close()

# 补算 weekday
WEEKDAY_FULL = ["", "星期一", "星期二", "星期三", "星期四",
                "星期五", "星期六", "星期日"]
for r in rows:
    d = r["meeting_date"]
    try:
        from datetime import datetime
        wd = datetime.strptime(d, "%Y-%m-%d").weekday() + 1
        r = dict(r)
        r["_wd_full"] = WEEKDAY_FULL[wd]
        # 用 dict 替代
    except Exception:
        pass

# 重新读为 dict list
rows = [dict(r) for r in rows]
for r in rows:
    try:
        from datetime import datetime
        wd = datetime.strptime(r["meeting_date"], "%Y-%m-%d").weekday() + 1
        r["_wd_full"] = WEEKDAY_FULL[wd]
    except Exception:
        r["_wd_full"] = ""

# 按日期分组
by_date = {}
for r in rows:
    by_date.setdefault(r["meeting_date"], []).append(r)

# 生成 8/24 起的连续 6 周(42 天)
WEEKDAY_CN = ["一", "二", "三", "四", "五", "六", "日"]


def fmt_day_cell(d, meetings):
    iso = d.isoformat()
    items = meetings.get(iso, [])
    n = len(items)
    is_weekend = d.weekday() >= 5
    bg = "#fff8f0" if is_weekend else "#ffffff"
    is_aug = d.month == 8
    if is_aug:
        bg = "#fafafa"
    title_color = "#666" if is_aug else "#222"

    count_html = ""
    if n:
        count_bg = "#ff5722" if n >= 10 else ("#ff9800" if n >= 5 else "#4caf50")
        count_html = f'<div class="count" style="background:{count_bg};">{n} 家</div>'

    busy_class = " busy" if n >= 6 else ""
    list_html = f'<div class="list-wrap"><div class="list">'
    for it in items:
        time_str = it["meeting_time"] or "—"
        city = it["city"] or it["province"] or "—"
        list_html += (
            f'<div class="item" '
            f'data-sec="{it["sec_code"]}" '
            f'data-name="{it["sec_name"]}" '
            f'data-time="{time_str}" '
            f'data-loc="{it["location_raw"]}" '
            f'data-pdf="{it["pdf_url"]}" '
            f'data-title="{it["title"]}">'
            f'<span class="t">{time_str}</span> '
            f'<span class="n">{it["sec_name"]}</span>'
            f'<span class="c">{city}</span>'
            f'</div>'
        )
    list_html += "</div></div>"

    return f"""
    <div class="day{busy_class}" style="background:{bg};">
      <div class="head" style="color:{title_color};">
        <span class="dn">{d.day}</span>
        {count_html}
      </div>
      {list_html}
    </div>
    """


# 构造网格:从第一个有会议的日期所在周的周一开始
# 如果 VIEW_START 当天或之后 7 天内有会议,用那个;否则用 VIEW_START
first_meet_date = min(
    (r["meeting_date"] for r in rows if r.get("meeting_date")),
    default=VIEW_START.isoformat(),
)
first_meet = date.fromisoformat(first_meet_date)
# 找到 first_meet 所在周的周一
cal_start = first_meet - timedelta(days=first_meet.weekday())
# 取到 VIEW_END 所在周的周日
cal_end = VIEW_END + timedelta(days=(6 - VIEW_END.weekday()))
# 总天数,向上取整到 7 的倍数
total_days = (cal_end - cal_start).days + 1
total_days = ((total_days + 6) // 7) * 7

all_grid = []
d = cal_start
for _ in range(total_days):
    all_grid.append(d)
    d += timedelta(days=1)
days_html = "".join(fmt_day_cell(dd, by_date) for dd in all_grid)
print(f"  日历范围: {cal_start} ~ {cal_end} ({total_days} 天)")


# === 顶部统计 ===
total = len(rows)
in_aug = sum(1 for r in rows if r["meeting_date"].startswith("2026-08"))
in_sep = sum(1 for r in rows if r["meeting_date"].startswith("2026-09"))
single_max = max((len(v) for v in by_date.values()), default=0)

# 城市分布(排除未识别)
city_count = Counter()
for r in rows:
    k = r["city"] or r["province"]
    if k: city_count[k] += 1
top_cities = city_count.most_common(10)


# === 预生成 Excel 和 CSV(用 openpyxl) ===
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

wb = openpyxl.Workbook()
header_font = Font(bold=True, color="FFFFFF", size=11)
header_fill = PatternFill("solid", fgColor="2E7D32")
center = Alignment(horizontal="center", vertical="center", wrap_text=True)
left = Alignment(horizontal="left", vertical="center", wrap_text=True)
thin = Side(border_style="thin", color="CCCCCC")
border = Border(left=thin, right=thin, top=thin, bottom=thin)

ws1 = wb.active
ws1.title = "股东大会列表"
headers1 = ["序号", "会议日期", "星期", "会议时间", "证券代码", "证券简称",
            "省份", "城市", "区县", "详细地址", "完整地点", "公告标题",
            "公告日期", "PDF链接"]
ws1.append(headers1)
for c, h in enumerate(headers1, 1):
    cell = ws1.cell(row=1, column=c)
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = center
    cell.border = border

for i, r in enumerate(rows, 1):
    loc = r["location_raw"] or ""
    detail = loc
    for prefix in (r["province"] or "", r["city"] or "", r["district"] or ""):
        if prefix and detail.startswith(prefix):
            detail = detail[len(prefix):]
    detail = detail.lstrip(" ,;:。;；,、")
    ws1.append([
        i, r["meeting_date"], r["_wd_full"], r["meeting_time"] or "",
        r["sec_code"], r["sec_name"],
        r["province"] or "", r["city"] or "", r["district"] or "",
        detail, loc, r["title"], r["publish_date"] or "", r["pdf_url"],
    ])

widths1 = [6, 12, 10, 10, 10, 14, 10, 10, 12, 40, 40, 50, 12, 60]
for idx, w in enumerate(widths1, 1):
    ws1.column_dimensions[get_column_letter(idx)].width = w
for row in ws1.iter_rows(min_row=2, max_row=ws1.max_row, max_col=len(headers1)):
    for cell in row:
        cell.border = border
        cell.alignment = center if cell.column in (1, 2, 3, 4, 5, 6, 7, 8, 9, 13) else left
ws1.freeze_panes = "A2"

# Sheet 2: 按日汇总
ws2 = wb.create_sheet("按日汇总")
ws2.append(["会议日期", "星期", "场次", "城市列表"])
for c in range(1, 5):
    cell = ws2.cell(row=1, column=c)
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = center
    cell.border = border

for d_iso in sorted(by_date.keys()):
    items = by_date[d_iso]
    wd = items[0]["_wd_full"]
    cities = Counter(it["city"] or it["province"] or "未识别" for it in items)
    city_str = ", ".join(f"{c}×{n}" for c, n in cities.most_common())
    ws2.append([d_iso, wd, len(items), city_str])
    for c in range(1, 5):
        cell = ws2.cell(row=ws2.max_row, column=c)
        cell.border = border
        cell.alignment = left if c == 4 else center

ws2.column_dimensions["A"].width = 12
ws2.column_dimensions["B"].width = 10
ws2.column_dimensions["C"].width = 8
ws2.column_dimensions["D"].width = 60
ws2.freeze_panes = "A2"

# Sheet 3: 城市分布
ws3 = wb.create_sheet("城市分布")
ws3.append(["城市/省份", "场次", "占比"])
for c in range(1, 4):
    cell = ws3.cell(row=1, column=c)
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = center
    cell.border = border

all_city = Counter()
for r in rows:
    all_city[r["city"] or r["province"] or "未识别"] += 1
tot = sum(all_city.values())
for c, n in all_city.most_common():
    ws3.append([c, n, f"{n*100/tot:.1f}%"])
    for col in range(1, 4):
        cell = ws3.cell(row=ws3.max_row, column=col)
        cell.border = border
        cell.alignment = center
ws3.column_dimensions["A"].width = 16
ws3.column_dimensions["B"].width = 10
ws3.column_dimensions["C"].width = 10
ws3.freeze_panes = "A2"

# Sheet 4: 公告日期分布
ws4 = wb.create_sheet("公告日期分布")
ws4.append(["公告发布日期", "场次"])
for c in range(1, 3):
    cell = ws4.cell(row=1, column=c)
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = center
    cell.border = border
pub_cnt = Counter(r["publish_date"] for r in rows if r["publish_date"])
for d, n in pub_cnt.most_common():
    ws4.append([d, n])
    for col in range(1, 3):
        cell = ws4.cell(row=ws4.max_row, column=col)
        cell.border = border
        cell.alignment = center
ws4.column_dimensions["A"].width = 16
ws4.column_dimensions["B"].width = 10

# 保存为 base64
buf = io.BytesIO()
wb.save(buf)
xlsx_b64 = base64.b64encode(buf.getvalue()).decode("ascii")

# CSV
import csv
csv_buf = io.StringIO()
writer = csv.writer(csv_buf)
writer.writerow(headers1)
for i, r in enumerate(rows, 1):
    loc = r["location_raw"] or ""
    detail = loc
    for prefix in (r["province"] or "", r["city"] or "", r["district"] or ""):
        if prefix and detail.startswith(prefix):
            detail = detail[len(prefix):]
    detail = detail.lstrip(" ,;:。;；,、")
    writer.writerow([
        i, r["meeting_date"], r["_wd_full"], r["meeting_time"] or "",
        r["sec_code"], r["sec_name"],
        r["province"] or "", r["city"] or "", r["district"] or "",
        detail, loc, r["title"], r["publish_date"] or "", r["pdf_url"],
    ])
csv_b64 = base64.b64encode(csv_buf.getvalue().encode("utf-8-sig")).decode("ascii")

print(f"xlsx b64: {len(xlsx_b64)} chars")
print(f"csv b64:  {len(csv_b64)} chars")


# === 动态变量(标题/文件名/页面元数据) ===
from datetime import date as _date
from datetime import datetime as _dt
today = _date.today().isoformat()
updated_at = _dt.now().strftime("%Y-%m-%d %H:%M")

# 公告日期范围(从 announcement.publish_date)
ann_dates = sorted(r["publish_date"] for r in rows if r.get("publish_date"))
if ann_dates:
    ann_date_range = f"{ann_dates[0]} ~ {ann_dates[-1]} ({len(set(ann_dates))} 天)"
else:
    ann_date_range = "—"

# 会议日期范围
meet_dates = sorted(r["meeting_date"] for r in rows if r.get("meeting_date"))
if meet_dates:
    meet_date_range = f"{meet_dates[0]} ~ {meet_dates[-1]}"
else:
    meet_date_range = "—"


# === HTML ===
html = f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<title>股东大会日历</title>
<style>
  body {{
    font-family: -apple-system, "PingFang SC", "Microsoft YaHei", sans-serif;
    background: #fafafa; color: #222; margin: 0; padding: 24px;
  }}
  h1 {{ margin: 0 0 6px; font-size: 22px; }}
  .meta {{ color: #666; font-size: 13px; margin-bottom: 18px; }}
  .meta b {{ color: #d32f2f; }}
  .cal {{
    display: grid; grid-template-columns: repeat(7, 1fr);
    gap: 4px; background: #ddd; border: 1px solid #ddd;
    max-width: 1500px;
  }}
  .cal-header {{
    background: #263238; color: #fff; padding: 8px; text-align: center;
    font-weight: 600; font-size: 13px;
  }}
  .cal-header.weekend {{ background: #455a64; }}
  .day {{
    min-height: 110px; padding: 6px; border: 1px solid transparent;
    font-size: 12px; overflow: hidden;
  }}
  .day.busy {{ min-height: 380px; }}
  .day .list-wrap {{ max-height: 350px; overflow-y: auto; }}
  .day .list-wrap::-webkit-scrollbar {{ width: 4px; }}
  .day .list-wrap::-webkit-scrollbar-thumb {{ background: #ccc; border-radius: 2px; }}
  .day .head {{
    display: flex; justify-content: space-between; align-items: center;
    font-weight: 600; margin-bottom: 4px;
  }}
  .day .head .dn {{ font-size: 14px; }}
  .day .count {{
    color: #fff; font-size: 11px; padding: 2px 6px; border-radius: 10px;
    font-weight: 500;
  }}
  .day .list {{ display: flex; flex-direction: column; gap: 2px; }}
  .day .item {{
    display: flex; gap: 4px; cursor: pointer;
    padding: 2px 4px; border-radius: 3px; background: #f5f9ff;
    font-size: 11px; line-height: 1.4;
  }}
  .day .item:hover {{ background: #fff3e0; }}
  .day .item .t {{ color: #d84315; font-weight: 600; min-width: 36px; }}
  .day .item .n {{ flex: 1; color: #1565c0; font-weight: 500;
    overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }}
  .day .item .c {{ color: #666; font-size: 10px;
    overflow: hidden; text-overflow: ellipsis; white-space: nowrap; max-width: 70px; }}
  #tip {{
    position: fixed; background: #fff; border: 1px solid #ccc;
    box-shadow: 0 4px 14px rgba(0,0,0,0.15); padding: 12px;
    border-radius: 6px; max-width: 380px; font-size: 13px; line-height: 1.6;
    display: none; z-index: 1000;
  }}
  #tip h3 {{ margin: 0 0 6px; font-size: 14px; color: #1565c0; }}
  #tip .row {{ margin: 2px 0; }}
  #tip .label {{ color: #888; display: inline-block; width: 60px; }}
  #tip a {{ color: #d84315; text-decoration: none; font-size: 12px; }}
  .stats {{
    display: flex; gap: 20px; margin: 18px 0; flex-wrap: wrap;
    max-width: 1500px;
  }}
  .stat-card {{
    background: #fff; border: 1px solid #e0e0e0; border-radius: 6px;
    padding: 12px 18px; min-width: 120px;
  }}
  .stat-card .v {{ font-size: 22px; font-weight: 700; color: #d32f2f; }}
  .stat-card .l {{ font-size: 12px; color: #666; }}
  .cities {{
    max-width: 1500px; background: #fff; border: 1px solid #e0e0e0;
    border-radius: 6px; padding: 12px 18px; margin-bottom: 18px;
    font-size: 13px;
  }}
  .cities .chip {{
    display: inline-block; background: #e3f2fd; color: #1565c0;
    padding: 3px 10px; border-radius: 12px; margin: 2px 4px;
  }}
  .cities .chip b {{ color: #d32f2f; margin-left: 4px; }}
  .legend {{ font-size: 12px; color: #666; margin: 10px 0; }}
  .legend span {{ display: inline-block; margin-right: 14px; }}
  .legend .dot {{
    display: inline-block; width: 10px; height: 10px; border-radius: 50%;
    margin-right: 4px; vertical-align: middle;
  }}
  .export-btn {{
    float: right; margin-left: 8px;
    background: #2e7d32; color: #fff; border: none; padding: 8px 16px;
    border-radius: 6px; font-size: 14px; font-weight: 500; cursor: pointer;
    box-shadow: 0 2px 6px rgba(46,125,50,0.3);
  }}
  .export-btn:hover {{ background: #1b5e20; }}
  .export-btn.secondary {{ background: #1565c0; }}
  .export-btn.secondary:hover {{ background: #0d47a1; }}
  .export-btn:disabled {{ background: #999; }}
  #exportToast {{
    position: fixed; bottom: 30px; right: 30px; background: #323232;
    color: #fff; padding: 12px 20px; border-radius: 6px; font-size: 14px;
    box-shadow: 0 4px 14px rgba(0,0,0,0.3); opacity: 0; transition: opacity 0.3s;
    z-index: 2000;
  }}
  #exportToast.show {{ opacity: 1; }}
  .month-tag {{
    display: inline-block; padding: 2px 8px; border-radius: 4px;
    font-size: 11px; margin-left: 4px; font-weight: 600;
  }}
  .month-8 {{ background: #ffe0b2; color: #e65100; }}
  .month-9 {{ background: #bbdefb; color: #0d47a1; }}
</style>
</head>
<body>

<h1>📅 股东大会日历
  <button id="exportXlsx" class="export-btn">📊 导出 Excel</button>
  <button id="exportCsv" class="export-btn secondary">📄 导出 CSV</button>
</h1>
<div class="meta">
  数据源: <b>巨潮资讯网</b> · 共 <b>{total}</b> 场会议 · 最近更新: <b>{updated_at}</b>
</div>

<div class="stats">
  <div class="stat-card"><div class="v">{total}</div><div class="l">总场次</div></div>
  <div class="stat-card"><div class="v">{in_sep}</div><div class="l">9 月召开</div></div>
  <div class="stat-card"><div class="v">{len(by_date)}</div><div class="l">有会议天数</div></div>
  <div class="stat-card"><div class="v">{single_max}</div><div class="l">单日最多</div></div>
</div>

<div class="cities">
  <b>📍 城市分布 Top 10:</b>
  {"".join(f'<span class="chip">{c}<b>{n}</b></span>' for c, n in top_cities)}
</div>

{"" if in_aug == 0 else f'''
<div class="aug-section">
  <b>📌 8 月召开(共 {in_aug} 场):</b> 灰底显示在日历左上方。
</div>
'''}

<div class="legend">
  <span><span class="dot" style="background:#4caf50;"></span>1-4 家</span>
  <span><span class="dot" style="background:#ff9800;"></span>5-9 家</span>
  <span><span class="dot" style="background:#ff5722;"></span>≥ 10 家</span>
  <span class="month-tag month-9">9月</span>
</div>

<div class="cal">
  {"".join(f'<div class="cal-header{" weekend" if i>=5 else ""}">周{w}</div>' for i, w in enumerate(WEEKDAY_CN))}
  {days_html}
</div>

<div id="tip"></div>
<div id="exportToast"></div>

<script>
const XLSX_B64 = "{xlsx_b64}";
const CSV_B64  = "{csv_b64}";

function b64ToBlob(b64, mime) {{
  const bin = atob(b64);
  const arr = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) arr[i] = bin.charCodeAt(i);
  return new Blob([arr], {{ type: mime }});
}}
function download(blob, filename) {{
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = filename; a.click();
  setTimeout(() => URL.revokeObjectURL(url), 100);
}}
function showToast(msg) {{
  const t = document.getElementById('exportToast');
  t.textContent = msg;
  t.classList.add('show');
  setTimeout(() => t.classList.remove('show'), 2200);
}}
document.getElementById('exportXlsx').addEventListener('click', e => {{
  e.target.disabled = true;
  try {{
    const blob = b64ToBlob(XLSX_B64,
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    download(blob, '股东大会_{total}场_{today}.xlsx');
    showToast('✅ Excel 已下载 · 4 个 Sheet');
  }} finally {{
    e.target.disabled = false;
  }}
}});
document.getElementById('exportCsv').addEventListener('click', e => {{
  e.target.disabled = true;
  try {{
    const blob = b64ToBlob(CSV_B64, 'text/csv;charset=utf-8');
    download(blob, '股东大会_{total}场_{today}.csv');
    showToast('✅ CSV 已下载');
  }} finally {{
    e.target.disabled = false;
  }}
}});

const tip = document.getElementById('tip');
document.querySelectorAll('.item').forEach(el => {{
  el.addEventListener('mouseenter', e => {{
    const d = el.dataset;
    tip.innerHTML = `
      <h3>${{d.sec}} ${{d.name}}</h3>
      <div class="row"><span class="label">时间</span>${{d.time}}</div>
      <div class="row"><span class="label">地点</span>${{d.loc}}</div>
      <div class="row"><span class="label">公告</span><a href="${{d.pdf}}" target="_blank">查看 PDF →</a></div>
      <div class="row" style="color:#888;font-size:11px;margin-top:4px;">${{d.title}}</div>
    `;
    tip.style.display = 'block';
    const r = el.getBoundingClientRect();
    let left = r.right + 8, top = r.top;
    if (left + 400 > window.innerWidth) left = r.left - 400;
    if (top + 200 > window.innerHeight) top = window.innerHeight - 220;
    tip.style.left = left + 'px';
    tip.style.top = top + 'px';
  }});
  el.addEventListener('mouseleave', () => {{ tip.style.display = 'none'; }});
}});
</script>
</body>
</html>
"""

with open(OUT_HTML, "w", encoding="utf-8") as f:
    f.write(html)
print(f"\n已生成: {OUT_HTML} ({len(html)} bytes)")
print(f"会议: 8月{in_aug} + 9月{in_sep} = 总{total}")
