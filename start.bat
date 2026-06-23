@echo off
chcp 65001 >nul
title 小升初英语 - 趣味学习工具

echo.
echo  ================================================
echo    小升初英语 - 趣味学习工具
echo  ================================================
echo.
echo  正在启动服务...

rem 尝试 Python（最常见，内置模块无需安装）
where python >nul 2>nul
if %errorlevel% equ 0 (
  echo  使用 Python 启动...
  echo.
  start http://localhost:8080
  python -m http.server 8080 -d "%~dp0"
  pause
  goto :end
)

rem 尝试 Node.js
where node >nul 2>nul
if %errorlevel% equ 0 (
  echo  使用 Node.js 启动...
  echo.
  start http://localhost:8080
  node "%~dp0server.js"
  pause
  goto :end
)

rem 尝试 Python (python3)
where python3 >nul 2>nul
if %errorlevel% equ 0 (
  echo  使用 Python 3 启动...
  echo.
  start http://localhost:8080
  python3 -m http.server 8080 -d "%~dp0"
  pause
  goto :end
)

echo.
echo  [错误] 未找到 Python 或 Node.js！
echo  请安装其中任意一个：
echo    Python: https://python.org
echo    Node.js: https://nodejs.org
echo.
pause

:end
