@echo off
chcp 65001 >nul
title ZURU 排期录入系统
cd /d "%~dp0"

REM ===== 自动查找Python =====
set "PYTHON="

REM 方法1: py启动器（强制64位）
where py >nul 2>&1 && (
    for /f "tokens=*" %%p in ('py -3-64 -c "import sys; print(sys.executable)" 2^>nul') do set "PYTHON=%%p"
)
if defined PYTHON goto :found

REM 方法2: PATH中的python（验证64位）
where python >nul 2>&1 && (
    for /f "tokens=*" %%p in ('python -c "import sys;print(sys.executable) if sys.maxsize>2**32 else None" 2^>nul') do set "PYTHON=%%p"
)
if defined PYTHON goto :found

REM 方法3: 常见安装位置
for %%d in (
    "%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
    "%LOCALAPPDATA%\Programs\Python\Python311\python.exe"
    "%LOCALAPPDATA%\Programs\Python\Python310\python.exe"
    "%PROGRAMFILES%\Python312\python.exe"
    "%PROGRAMFILES%\Python311\python.exe"
    "%PROGRAMFILES%\Python310\python.exe"
) do (
    if exist %%d (
        set "PYTHON=%%~d"
        goto :found
    )
)

echo [错误] 找不到Python，请先安装Python 3.10+
echo 下载地址: https://www.python.org/downloads/
pause
exit /b 1

:found
echo 使用Python: %PYTHON%

REM ===== 首次运行自动安装依赖 =====
if exist requirements.txt (
    if not exist "data\.deps_installed" (
        echo 首次运行，正在安装依赖...
        "%PYTHON%" -m pip install -r requirements.txt
        if not errorlevel 1 (
            mkdir data 2>nul
            echo ok > "data\.deps_installed"
        )
    )
)

REM ===== 读取端口配置 =====
set "PORT=5000"
if exist "data\config.json" (
    for /f "tokens=2 delims=:, " %%a in ('findstr "port" "data\config.json" 2^>nul') do (
        set "PORT=%%~a"
    )
)

REM ===== 关闭旧进程 =====
echo 正在关闭旧进程...
for /f "tokens=5" %%a in ('netstat -aon ^| findstr :%PORT% ^| findstr LISTENING') do taskkill /F /PID %%a >nul 2>&1
timeout /t 1 >nul

REM ===== 启动 =====
start "" http://localhost:%PORT%
"%PYTHON%" app.py
pause
