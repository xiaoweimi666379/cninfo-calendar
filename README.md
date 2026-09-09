# 📅 股东大会日历 · 巨潮资讯网

每天 23:00(北京时间)自动从巨潮资讯网抓取前一天发布的股东大会公告,提取会议时间和地点,生成日历 HTML 并部署到 GitHub Pages。

## 🌐 在线访问

**GitHub Pages**: `https://<你的用户名>.github.io/cninfo-calendar/`

## ✨ 功能

- 📅 日历视图,按周展示(每月自动扩展)
- 🖱️ 鼠标悬停看完整公司/时间/地点/PDF
- 📊 一键导出 Excel / CSV
- 🗄️ SQLite 数据库自动增量更新
- 🧹 自动清理过期会议(默认保留 30 天)

## 🔧 本地运行

```bash
# 装依赖
pip install pypdf openpyxl

# 抓指定一天
python daily_update.py 2026-08-29

# 抓"昨天"(默认行为)
python daily_update.py
```

## ⚙️ 配置

通过环境变量:

| 变量 | 默认 | 说明 |
|------|------|------|
| `DB_PATH` | `./data/cninfo.db` | SQLite 数据库路径 |
| `HTML_OUT` | `./dist/index.html` | 生成的 HTML 路径 |
| `WORK_DIR` | `./work` | 临时 PDF 缓存目录 |
| `EXPIRE_DAYS` | `30` | 会议保留天数 |

## 🚀 GitHub Actions 部署

本仓库用 GitHub Actions **每天 UTC 15:00(北京时间 23:00)** 自动跑:

1. 抓取昨天发布的股东大会公告
2. 解析 PDF 提取会议时间+地点
3. 增量更新到 SQLite 数据库
4. 清理 30 天前的过期数据
5. 重建日历 HTML
6. 部署到 GitHub Pages
7. 把新数据库存为 artifact(下次任务下载用)

### 第一次配置

1. Fork/创建这个仓库(可以是 **private** 的,数据不会泄露)
2. 进入 **Settings → Pages → Source**:选 `GitHub Actions`
3. 等 23:00(北京时间)自动跑,或到 **Actions** 标签手动触发
4. 跑完后 Settings → Pages 会有访问链接

## 📂 项目结构

```
cninfo-calendar/
├── .github/workflows/daily.yml   # GitHub Actions 配置
├── daily_update.py                # 每日增量更新入口
├── build_calendar.py              # 日历 HTML 生成
├── data/                          # SQLite 数据库
├── work/                          # 临时 PDF 缓存
└── dist/                          # 生成的 HTML(部署到 Pages)
```

## 📊 数据库结构

**company**: `sec_code` (PK), `sec_name`
**announcement**: `id` (PK), `sec_code`, `title`, `pdf_url`, `publish_date`, `source_run`
**meeting**: `id` (PK), `announcement_id` (FK), `meeting_date`, `meeting_time`, `weekday_zh`, `location_raw`, `province`, `city`, `district`, `detail`, `source_run`

## 🧪 手动触发

进入仓库 **Actions** → 选 "每日股东大会数据更新" → Run workflow → 可填日期 → 运行
