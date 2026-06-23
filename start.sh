#!/bin/bash
DIR="$(cd "$(dirname "$0")" && pwd)"
echo ""
echo "  ================================================"
echo "    小升初英语 - 趣味学习工具"
echo "  ================================================"
echo ""

if command -v python3 &> /dev/null; then
  echo "  使用 Python 3 启动..."
  (sleep 1 && open http://localhost:8080) &
  cd "$DIR" && python3 -m http.server 8080
  exit 0
elif command -v python &> /dev/null; then
  echo "  使用 Python 启动..."
  (sleep 1 && open http://localhost:8080) &
  cd "$DIR" && python -m http.server 8080
  exit 0
elif command -v node &> /dev/null; then
  echo "  使用 Node.js 启动..."
  (sleep 1 && open http://localhost:8080) &
  cd "$DIR" && node server.js
  exit 0
else
  echo "  [错误] 未找到 Python 或 Node.js！"
  echo "  请安装: https://python.org 或 https://nodejs.org"
  exit 1
fi
