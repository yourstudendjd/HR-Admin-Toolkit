@echo off
chcp 65001 >nul
echo ============================================
echo   考勤与宿舍管理工具集 - 打包为 EXE
echo ============================================
echo.

REM --- 1. Install dependencies ---
echo [1/3] 安装 Python 依赖...
pip install -r requirements.txt -q
if %errorlevel% neq 0 (
    echo 错误: 依赖安装失败。
    pause
    exit /b 1
)
echo 完成.
echo.

REM --- 2. Build with PyInstaller ---
echo [2/3] 使用 PyInstaller 打包...
python -m PyInstaller --onefile --windowed --name "AttendanceSuite" --add-data "attendance_processor.py;." --add-data "dormitory_processor.py;." main.py
if %errorlevel% neq 0 (
    echo 错误: PyInstaller 打包失败。
    pause
    exit /b 1
)
echo 完成.
echo.

REM --- 3. Show output ---
echo [3/3] 打包完成!
echo.
echo EXE 位置: %CD%\dist\AttendanceSuite.exe
echo.

start "" "%CD%\dist"

echo.
pause
