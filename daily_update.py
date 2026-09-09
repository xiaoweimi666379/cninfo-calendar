"""
GitHub Actions 版每日更新脚本

环境约定:
  - 仓库根目录运行
  - 数据库路径: $DB_PATH (默认 ./data/cninfo.db)
  - HTML 输出:   $HTML_OUT (默认 ./dist/index.html)
  - 工作目录:   $WORK_DIR  (默认 ./work,放临时 PDF)

清理规则:
  - 删除 meeting_date < (今天 - EXPIRE_DAYS) 的会议(默认 0 = 删今天之前全部)
  - 同时删除没关联的 announcement
  - 保留 company 表
"""
import os
import re
import json
import time
import io
import shutil
import sqlite3
import urllib.request
import urllib.parse
import subprocess
import base64
from datetime import date, timedelta, datetime
from concurrent.futures import ThreadPoolExecutor, as_completed
from collections import Counter

# === 路径配置 ===
DB_PATH = os.environ.get("DB_PATH", "./data/cninfo.db")
HTML_OUT = os.environ.get("HTML_OUT", "./dist/index.html")
WORK_DIR = os.environ.get("WORK_DIR", "./work")
# 清理策略:删除 meeting_date < 今天的会议(已开过的)
# 也可设 EXPIRE_DAYS=1 表示"保留今天还在开的"
EXPIRE_DAYS = int(os.environ.get("EXPIRE_DAYS", "0"))

URL_API = "http://www.cninfo.com.cn/new/hisAnnouncement/query"
TODAY = date.today()
YESTERDAY = TODAY - timedelta(days=1)

os.makedirs(os.path.dirname(DB_PATH) or ".", exist_ok=True)
os.makedirs(os.path.dirname(HTML_OUT) or ".", exist_ok=True)
os.makedirs(WORK_DIR, exist_ok=True)
PDF_DIR = f"{WORK_DIR}/pdfs"

print(f"=== 每日更新 {YESTERDAY} ({datetime.now().strftime('%Y-%m-%d %H:%M:%S')}) ===")
print(f"  DB:   {DB_PATH}")
print(f"  HTML: {HTML_OUT}")
print(f"  工作:  {WORK_DIR}")


# === 抓清单 ===
KEEP = re.compile(r"(召开.{0,30}股东(大会|会).{0,8}(通知|提示性公告)|股东(大会|会)会议资料|会议资料)")
DROP = re.compile(r"(决议公告|法律意见|延长.*有效期|关于取消|关于未能|关于未)")


def meeting_key(title, code):
    m = re.search(r"(第[一二三四五六七八九十\d]+次)(临时|年度|)?", title)
    suffix = (m.group(1) + (m.group(2) or "")) if m else ""
    return (code, suffix)


def fetch_list(d_str):
    all_items, page = [], 1
    while True:
        data = {
            "pageNum": str(page), "pageSize": "30",
            "column": "sse", "tabName": "fulltext",
            "plate": "", "stock": "", "searchkey": "股东大会",
            "secid": "", "category": "category_gddh_szsh;",
            "trade": "", "seDate": f"{d_str}~{d_str}",
            "sortName": "time", "sortType": "desc", "isHLtitle": "true",
        }
        req = urllib.request.Request(
            URL_API, data=urllib.parse.urlencode(data).encode("utf-8"),
            headers={"User-Agent": "Mozilla/5.0",
                     "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
                     "Origin": "http://www.cninfo.com.cn",
                     "Referer": "http://www.cninfo.com.cn/new/commonUrl/pageOfSearch?url=disclosure/list/search"},
            method="POST",
        )
        with urllib.request.urlopen(req, timeout=30) as r:
            obj = json.loads(r.read().decode("utf-8"))
        items = obj.get("announcements") or []
        if not items: break
        all_items.extend(items)
        if len(all_items) >= int(obj.get("totalAnnouncement") or 0): break
        page += 1
        time.sleep(0.2)
    return all_items


DATES = [YESTERDAY.isoformat()]
SOURCE = YESTERDAY.isoformat()

raw_all = []
for d in DATES:
    items = fetch_list(d)
    raw_all.extend(items)
    print(f"  [抓] {d}: {len(items)} 条")

seen, uniq = set(), []
for it in raw_all:
    if it.get("adjunctUrl") not in seen:
        seen.add(it.get("adjunctUrl")); uniq.append(it)

best = {}
for it in uniq:
    title = (it.get("announcementTitle") or "").replace("<em>", "").replace("</em>", "")
    if DROP.search(title) or not KEEP.search(title): continue
    k = meeting_key(title, it.get("secCode"))
    m = re.search(r"finalpage/(\d{4}-\d{2}-\d{2})/", it.get("adjunctUrl", ""))
    pub = m.group(1) if m else ""
    rec = {"publish_date": pub, "secCode": it.get("secCode"),
           "secName": it.get("secName"), "title": title,
           "pdf": it.get("adjunctUrl")}
    if k not in best:
        best[k] = rec
    elif "通知" in rec["title"] and "通知" not in best[k]["title"]:
        best[k] = rec

notice_list = list(best.values())
print(f"  [过] 通知类: {len(notice_list)} 条")


# === 解析 PDF ===
TIME_RE = re.compile(
    r"(\d{4})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日"
    r"(?:\s*[（(]([^\）)]+)[）)])?"
    r"\s*((?:上午|下午|凌晨|晚上|中午)?\s*\d{1,2}\s*[:：点]\s*\d{1,2}\s*分?)"
)
TIME_RE_HOUR_ONLY = re.compile(
    r"(\d{4})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日"
    r"(?:\s*[（(]([^\）)]+)[）)])?"
    r"\s*((?:上午|下午|凌晨|晚上|中午)?\s*\d{1,2}\s*点(?!\s*\d))"
)
LOC_RE = re.compile(r"会议地点[：:]\s*([^\n\r。;；]+)")
LOC_RE2 = re.compile(r"(?:现场|召开)\s*会议地点[：:]\s*([^\n\r。;；]+)")
LOC_RE3 = re.compile(r"召开地点[：:]\s*([^\n\r。;；]+)")
LOC_RE_NEXTLINE = re.compile(r"会议地点\s*\n+\s*([^\n]+)")


def _fmt_time(y, mo, d, weekday, hm):
    hm = re.sub(r"\s+", "", hm)
    hm = re.sub(r"(\d{1,2})\s*[:：点]\s*(\d{1,2})\s*分?", r"\1:\2", hm)
    hm = re.sub(r"(\d{1,2})\s*点(?!\s*\d)", r"\1:00", hm)
    return f"{y}年{int(mo)}月{int(d)}日 {hm}"


def extract_time(text):
    for line in text.splitlines():
        line = line.strip()
        if "现场" in line and "时间" in line and "年" in line and "日" in line:
            m = TIME_RE.search(line) or TIME_RE_HOUR_ONLY.search(line)
            if m: return _fmt_time(*m.groups())
    for line in text.splitlines():
        line = line.strip()
        if ("现场会议召开" in line or "会议召开" in line) and "日期" in line and "年" in line:
            m = TIME_RE.search(line) or TIME_RE_HOUR_ONLY.search(line)
            if m: return _fmt_time(*m.groups())
    m = TIME_RE.search(text) or TIME_RE_HOUR_ONLY.search(text)
    if m: return _fmt_time(*m.groups())
    return None


def extract_loc(text):
    for pat in (LOC_RE, LOC_RE2, LOC_RE3):
        m = pat.search(text)
        if m:
            loc = m.group(1).strip()
            loc = re.split(r"[。;；\n\r]", loc)[0].strip()
            return loc
    m = LOC_RE_NEXTLINE.search(text)
    return m.group(1).strip() if m else None


import pypdf


def download(item):
    code = item["secCode"]
    pdf_name = item["pdf"].split("/")[-1]
    path = f"{PDF_DIR}/{code}_{pdf_name}"
    if os.path.exists(path) and os.path.getsize(path) > 1000:
        return item, path, None
    url = "http://static.cninfo.com.cn/" + item["pdf"]
    for _ in range(2):
        try:
            req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
            with urllib.request.urlopen(req, timeout=30) as r:
                with open(path, "wb") as f: f.write(r.read())
            return item, path, None
        except Exception:
            time.sleep(0.5)
    return item, None, "下载失败"


results = []
errs = 0
with ThreadPoolExecutor(max_workers=8) as ex:
    futs = [ex.submit(download, it) for it in notice_list]
    for i, fut in enumerate(as_completed(futs), 1):
        item, path, err = fut.result()
        if err:
            errs += 1
            if errs <= 3:
                print(f"  [下载失败样本] {item.get('secCode')} {item.get('pdf')}: {err}")
            results.append({**item, "time": None, "loc": None}); continue
        try:
            r = pypdf.PdfReader(path)
            text = "\n".join(p.extract_text() for p in r.pages)
        except Exception as e:
            text = ""
        results.append({**item, "time": extract_time(text), "loc": extract_loc(text)})
if errs:
    print(f"  [下载失败总数] {errs}/{len(notice_list)}")

print(f"  [解] {len(results)} 条, "
      f"时间 {sum(1 for r in results if r.get('time'))}, "
      f"地点 {sum(1 for r in results if r.get('loc'))}")


# === 入库 ===
def parse_meeting_time(s):
    if not s: return None, None
    m = re.match(
        r"(\d{4})年(\d{1,2})月(\d{1,2})日"
        r"(?:\s*[（(]([^\）)]+)[）)])?"
        r"\s*(上午|下午|凌晨|晚上|中午)?\s*(\d{1,2}):(\d{2})", s,
    )
    if not m: return None, None
    y, mo, d, wd, ampm, hh, mm = m.groups()
    if ampm == "下午" and int(hh) < 12: hh = str(int(hh) + 12)
    elif ampm == "上午" and int(hh) == 12: hh = "00"
    elif ampm == "晚上" and int(hh) < 12: hh = str(int(hh) + 12)
    return f"{int(y):04d}-{int(mo):02d}-{int(d):02d}", f"{int(hh):02d}:{mm}"


def parse_location(s):
    if not s: return None, None, None, s
    province, city, district, detail = None, None, None, s
    m = re.match(r"(.+?[省自治区])(.+?[市])(.*)$", s)
    if m:
        province = m.group(1); rest = m.group(2) + m.group(3)
        m2 = re.match(r"(.+?[市])(.*)$", rest)
        if m2: city = m2.group(1); rest = m2.group(2)
        m3 = re.match(r"(.+?[区县])(.*)$", rest)
        if m3: district = m3.group(1); detail = m3.group(2)
        else: detail = rest
    else:
        m = re.match(r"(北京市|上海市|天津市|重庆市)(.*)$", s)
        if m:
            province = m.group(1); rest = m.group(2)
            m3 = re.match(r"(.+?[区县])(.*)$", rest)
            if m3: district = m3.group(1); detail = m3.group(2)
            else: detail = rest
        else:
            m = re.match(r"(.+?[市])(.*)$", s)
            if m:
                city = m.group(1); rest = m.group(2)
                m3 = re.match(r"(.+?[区县])(.*)$", rest)
                if m3: district = m3.group(1); detail = m3.group(2)
                else: detail = rest
    return province, city, district, detail.strip(" ,;:。;；,、")


con = sqlite3.connect(DB_PATH)
cur = con.cursor()

# 建表(首次运行时)
cur.executescript("""
CREATE TABLE IF NOT EXISTS company (
    sec_code TEXT PRIMARY KEY,
    sec_name TEXT NOT NULL,
    full_name TEXT
);
CREATE TABLE IF NOT EXISTS announcement (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    sec_code TEXT NOT NULL,
    title TEXT NOT NULL,
    pdf_url TEXT NOT NULL,
    publish_date TEXT,
    source_run TEXT
);
CREATE TABLE IF NOT EXISTS meeting (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    announcement_id INTEGER NOT NULL,
    meeting_date TEXT,
    meeting_time TEXT,
    weekday INTEGER,
    weekday_zh TEXT,
    ampm TEXT,
    location_raw TEXT,
    province TEXT,
    city TEXT,
    district TEXT,
    detail TEXT,
    source_run TEXT,
    FOREIGN KEY(announcement_id) REFERENCES announcement(id)
);
CREATE INDEX IF NOT EXISTS idx_meeting_date ON meeting(meeting_date);
CREATE INDEX IF NOT EXISTS idx_meeting_city ON meeting(city);
""")
con.commit()

# 清理过期
cutoff = (TODAY - timedelta(days=EXPIRE_DAYS)).isoformat()
cur.execute("DELETE FROM meeting WHERE meeting_date IS NOT NULL AND meeting_date < ?", (cutoff,))
deleted_m = cur.rowcount
cur.execute("DELETE FROM announcement WHERE id NOT IN (SELECT DISTINCT announcement_id FROM meeting)")
deleted_a = cur.rowcount
con.commit()
after_m = con.execute("SELECT COUNT(*) FROM meeting").fetchone()[0]
print(f"  [清] 截止 {cutoff}, 删 {deleted_m} 场过期, 删 {deleted_a} 条无关联公告 (剩 {after_m} 场)")

# 灌本批
ins_ann = ins_meet = 0
for it in results:
    sec_code = it["secCode"]; sec_name = it["secName"]
    title = it["title"]
    pdf_url = "http://static.cninfo.com.cn/" + it["pdf"]
    pub = it.get("publish_date", "")
    cur.execute("INSERT OR IGNORE INTO company(sec_code, sec_name) VALUES (?, ?)",
                (sec_code, sec_name))
    cur.execute(
        "SELECT id FROM announcement WHERE sec_code=? AND title=? AND source_run=?",
        (sec_code, title, SOURCE),
    )
    row = cur.fetchone()
    if row:
        ann_id = row[0]
    else:
        cur.execute(
            "INSERT INTO announcement(sec_code, title, pdf_url, publish_date, source_run)"
            " VALUES (?, ?, ?, ?, ?)",
            (sec_code, title, pdf_url, pub, SOURCE),
        )
        ann_id = cur.lastrowid
        ins_ann += 1
    date_iso, hhmm = parse_meeting_time(it.get("time"))
    prov, city, dist, detail = parse_location(it.get("loc"))
    cur.execute(
        "SELECT id FROM meeting WHERE announcement_id=? AND meeting_date=? "
        "AND meeting_time=? AND source_run=?",
        (ann_id, date_iso, hhmm, SOURCE),
    )
    if not cur.fetchone():
        cur.execute(
            "INSERT INTO meeting(announcement_id, meeting_date, meeting_time, "
            "weekday, weekday_zh, ampm, location_raw, province, city, district, "
            "detail, source_run) VALUES (?,?,?,?,?,?,?,?,?,?,?,?)",
            (ann_id, date_iso, hhmm, None, "", None,
             it.get("loc"), prov, city, dist, detail, SOURCE),
        )
        ins_meet += 1

con.commit()
con.close()
print(f"  [库] 本批新增 {ins_meet} 场, 累计 {after_m} 场")

# 重建 HTML(直接 subprocess 调 build_calendar.py,环境变量已设置)
print("  [重] 重建日历 HTML ...")
ret = subprocess.run(
    ["python3", "build_calendar.py"],
    capture_output=True, text=True, timeout=120,
    env={**os.environ, "DB_PATH": DB_PATH, "HTML_OUT": HTML_OUT},
)
if ret.returncode != 0:
    print("  [重] ❌ 失败:")
    print(ret.stdout[-500:])
    print(ret.stderr[-500:])
    raise SystemExit(1)
print(f"  [重] ✓ 已生成 {HTML_OUT}")

print(f"\n=== 完成 {datetime.now().strftime('%H:%M:%S')} ===")
print(f"  数据库: {DB_PATH} ({after_m} 场)")
print(f"  日历:   {HTML_OUT}")
