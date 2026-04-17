@echo off
chcp 65001 > nul
echo ========================================
echo 合格证自动化系统 - Windows打包脚本
echo ========================================
echo.

REM 检查Python是否安装
python --version >nul 2>&1
if errorlevel 1 (
    echo ❌ 错误: 未找到Python，请先安装Python
    pause
    exit /b 1
)

echo [1/4] 创建虚拟环境...
python -m venv venv
call venv\Scripts\activate.bat

echo [2/4] 安装依赖...
pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple

echo [3/4] 清理旧构建文件...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist __pycache__ rmdir /s /q __pycache__

echo [4/4] 开始打包（约需1-2分钟）...
pyinstaller 合格证自动化系统.spec --clean

echo.
echo ========================================
echo ✅ 打包完成！
echo ========================================
echo.
echo 输出位置: dist\合格证自动化系统.exe
echo.
echo 提示: 建议将exe上传到 VirusTotal 检测误报情况
echo.

REM 询问是否打开输出文件夹
choice /c YN /m "是否打开输出文件夹?"
if errorlevel 2 goto end
explorer dist

:end
pause
