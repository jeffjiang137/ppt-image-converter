@echo off
chcp 65001 >nul
cd /d "%~dp0"
where py >nul 2>nul
if %errorlevel%==0 (
    set "PYTHON_CMD=py"
) else (
    set "PYTHON_CMD=python"
)
%PYTHON_CMD% ppt_image_tool.py
if errorlevel 1 (
    echo.
    echo 启动失败，请先双击“安装依赖.bat”。
    pause
)
