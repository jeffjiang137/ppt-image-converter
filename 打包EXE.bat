@echo off
chcp 65001 >nul
cd /d "%~dp0"
where py >nul 2>nul
if %errorlevel%==0 (
    set "PYTHON_CMD=py"
) else (
    set "PYTHON_CMD=python"
)

echo 正在打包 Windows 单文件 EXE...
%PYTHON_CMD% -m PyInstaller --noconfirm --clean --onefile --windowed ^
  --name "PPT图片互转工具" ^
  --icon "app_icon.ico" ^
  --add-data "app_icon.png;." ^
  --add-data "README_使用说明.md;." ^
  --collect-all pptx ^
  --collect-all PIL ^
  --collect-all fitz ^
  ppt_image_tool.py
if errorlevel 1 (
    echo.
    echo 打包失败，请先运行“安装依赖.bat”。
    pause
    exit /b 1
)
echo.
echo 打包完成：dist\PPT图片互转工具.exe
pause
