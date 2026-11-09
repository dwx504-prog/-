@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo 正在安装打包依赖...
py -m pip install --upgrade pip
py -m pip install -r requirements.txt
if errorlevel 1 (
  echo 依赖安装失败，请检查 Python 和网络。
  pause
  exit /b 1
)
echo 正在生成 EXE...
py -m PyInstaller --noconfirm --clean --onefile --windowed --name "Word公式图片转换助手" word_formula_to_image.py
if errorlevel 1 (
  echo 打包失败。
  pause
  exit /b 1
)
echo.
echo 打包完成：dist\Word公式图片转换助手.exe
pause
