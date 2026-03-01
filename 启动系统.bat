@echo off
chcp 65001 >nul
title ZURU 排期录入系统
cd /d "%~dp0"
echo 正在关闭旧进程...
for /f "tokens=5" %%a in ('netstat -aon ^| findstr :5000 ^| findstr LISTENING') do taskkill /F /PID %%a >nul 2>&1
timeout /t 1 >nul
start "" http://localhost:5000
"C:\Users\Administrator\AppData\Local\Programs\Python\Python312\python.exe" app.py
pause
