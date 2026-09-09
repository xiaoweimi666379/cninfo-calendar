#!/bin/bash
# 一键推送到 GitHub 的脚本
# 用法:
#   1) 先在 GitHub 上创建空仓库(不要勾 README/.gitignore/license)
#   2) bash push.sh https://github.com/你的用户名/cninfo-calendar.git

set -e
URL="$1"
if [ -z "$URL" ]; then
  echo "用法: bash push.sh <仓库 URL>"
  echo "例如: bash push.sh https://github.com/yourname/cninfo-calendar.git"
  exit 1
fi

cd "$(dirname "$0")"

# 确保 dist/ 有内容(GitHub Pages 部署需要)
if [ ! -f dist/index.html ]; then
  echo "dist/ 是空的,先本地跑一次 build_calendar 生成 HTML..."
  DB_PATH=./data/cninfo.db HTML_OUT=./dist/index.html python3 build_calendar.py
fi

git init -b main
git add .
git commit -m "init: 股东大会日历 + GitHub Actions 每日更新"
git remote add origin "$URL"
echo ""
echo "准备 push,可能需要你输入 GitHub 凭据..."
git push -u origin main
echo ""
echo "✅ 推送完成!"
echo ""
echo "接下来:"
echo "  1. 打开 https://github.com/你的用户名/cninfo-calendar"
echo "  2. Settings → Pages → Source: 选 GitHub Actions"
echo "  3. 等 Actions 跑完(可手动触发测试)"
echo "  4. 访问 https://你的用户名.github.io/cninfo-calendar/"
