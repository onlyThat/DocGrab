@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"

echo ============================================
echo   DocGrab 打包脚本（PyInstaller 单文件 exe）
echo ============================================
echo.

python -m PyInstaller --noconfirm --clean ^
  --name DocGrab ^
  --onefile ^
  --windowed ^
  --icon "docgrab.ico" ^
  --version-file "version_info.txt" ^
  --add-data "docgrab.ico;." ^
  --collect-all tkinterdnd2 ^
  --hidden-import openpyxl ^
  --hidden-import docx ^
  --hidden-import pptx ^
  --hidden-import xlrd ^
  --hidden-import win32com.client ^
  --exclude-module matplotlib ^
  --exclude-module numpy ^
  --exclude-module pandas ^
  --exclude-module PIL ^
  "main.py"

if errorlevel 1 (
  echo.
  echo [失败] 打包出错，请查看上面的信息。
  pause
  exit /b 1
)

echo.
echo [成功] 产物：dist\DocGrab.exe
echo.
endlocal