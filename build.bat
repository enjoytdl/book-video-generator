@echo off
chcp 65001 >nul
echo ============================================
echo       书单视频生成器 - Windows 打包脚本
echo ============================================
echo.

REM 检查 Python 是否安装
python --version >nul 2>&1
if errorlevel 1 (
    echo ❌ 错误: 未检测到 Python
    echo 请先安装 Python 3.10+: https://python.org/downloads
    pause
    exit /b 1
)

echo ✅ Python 已安装
echo.

REM 安装依赖
echo 📦 正在安装依赖包...
pip install pyinstaller openpyxl pandas pyqt5 pillow numpy opencv-python -q
echo ✅ 依赖安装完成
echo.

REM 执行打包
echo 🔨 正在打包...
pyinstaller --name="书单视频生成器" --windowed --onefile --add-data="fonts;fonts" --add-data="config;config" --hidden-import=PyQt5.QtCore --hidden-import=PyQt5.QtGui --hidden-import=PyQt5.QtWidgets --hidden-import=openpyxl --hidden-import=pandas --hidden-import=PIL --hidden-import=cv2 --hidden-import=numpy book_video_generator.py

echo.
echo ============================================
echo ✅ 打包完成！
echo ============================================
echo 📁 文件位置: dist\书单视频生成器.exe
echo.
echo ⚠️  注意：
echo    - 请将 fonts 文件夹与 exe 放在同一目录
echo    - 首次运行可能较慢（需要解压依赖）
echo    - 杀毒软件可能报警，点击"允许"即可
echo.
pause
