@echo off
chcp 65001 >nul
cd /d "%~dp0"
title 工期预警管控系统 Web 服务
echo ============================================
echo   工期预警管控系统 Web 服务
echo ============================================
echo.
echo 启动中，请在浏览器访问: http://127.0.0.1:8080
echo.
echo 按 Ctrl+C 停止服务
echo.
python app.py
pause
