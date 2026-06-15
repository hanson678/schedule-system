@echo off
chcp 65001 2>NUL
title ZURU - Install Dependencies
cd /d "%~dp0"

echo ============================================
echo   ZURU - Dependencies Installer
echo ============================================
echo.

set "PYTHON="

for %%p in (
    "%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
    "%LOCALAPPDATA%\Programs\Python\Python311\python.exe"
    "%LOCALAPPDATA%\Programs\Python\Python310\python.exe"
    "%PROGRAMFILES%\Python312\python.exe"
    "%PROGRAMFILES%\Python311\python.exe"
    "%PROGRAMFILES%\Python310\python.exe"
) do (
    if exist %%p (
        set "PYTHON=%%~p"
        goto :found
    )
)

where py 2>NUL
if not %errorlevel%==0 goto :trypath
for /f "tokens=*" %%p in ('py -3-64 -c "import sys; print(sys.executable)" 2^>NUL') do set "PYTHON=%%p"
if defined PYTHON goto :found

:trypath
where python 2>NUL
if not %errorlevel%==0 goto :notfound
for /f "tokens=*" %%p in ('python -c "import sys; print(sys.executable)" 2^>NUL') do set "PYTHON=%%p"
if defined PYTHON goto :found

:notfound
echo [ERROR] Python not found. Install Python 3.10+ 64-bit
echo https://www.python.org/downloads/
pause
exit /b 1

:found
echo [OK] Python: %PYTHON%
echo.
echo [1/3] Upgrading pip...
"%PYTHON%" -m pip install --upgrade pip -q
echo.
echo [2/3] Installing packages...
"%PYTHON%" -m pip install -r requirements.txt
if %errorlevel% neq 0 (
    echo [ERROR] Install failed
    pause
    exit /b 1
)
echo.
echo [3/3] Registering pywin32 COM...
"%PYTHON%" Scripts\pywin32_postinstall.py -install 2>NUL
"%PYTHON%" -c "import pythoncom" 2>NUL
echo.
echo ============================================
echo   Done! Run vbs to start.
echo ============================================
pause
