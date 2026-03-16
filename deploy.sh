#!/bin/bash
# 运营宣推数据看板 - 一键部署到 GitHub Pages
# 用法：./deploy.sh [提交说明]
cd "$(dirname "$0")"
msg="${1:-更新看板}"
git add -A
if git diff --cached --quiet; then
  echo "无变更，跳过部署。"
  exit 0
fi
git commit -m "$msg"
git push origin main
echo "已推送到 https://echo7qi.github.io/cursor-demo/ ，约 1–2 分钟后生效。"
