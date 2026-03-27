@echo off
chcp 65001>/dev/null
echo 製令查詢系統（除錯模式）...
"%~dp0_python\python.exe" "%~dp0_app\main.py"
pause
