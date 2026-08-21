@echo off
chcp 65001 >nul
cd /d "%~dp0"
where py >nul 2>nul
if %errorlevel%==0 (
    set "PYTHON_CMD=py"
) else (
    set "PYTHON_CMD=python"
)

echo 正在安装 PPT 图片互转工具所需组件...
%PYTHON_CMD% -m pip install --upgrade pip
%PYTHON_CMD% -m pip install -r requirements.txt
if errorlevel 1 (
    echo.
    echo 安装失败。请确认已安装 Python 3.10 或更高版本，并在安装时勾选 Add Python to PATH。
    pause
    exit /b 1
)
echo.
echo 安装完成，可以双击“启动工具.bat”。
pause
