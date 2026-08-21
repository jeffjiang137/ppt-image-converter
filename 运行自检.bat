@echo off
chcp 65001 >nul
cd /d "%~dp0"
where py >nul 2>nul
if %errorlevel%==0 (
    set "PYTHON_CMD=py"
) else (
    set "PYTHON_CMD=python"
)
%PYTHON_CMD% 自检.py
if errorlevel 1 (
    echo.
    echo 自检发现异常，请先运行“安装依赖.bat”。
)
pause
