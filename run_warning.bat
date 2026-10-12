@echo off
chcp 65001 >nul
cd /d "%~dp0"
python schedule_warning.py >> warning_log.txt 2>&1
