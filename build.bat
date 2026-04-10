@echo off
chcp 65001 >nul
echo ==============================================
echo Excel比对工具打包脚本
echo ==============================================
echo.

:: 检查Python是否安装
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ❌ 错误：未检测到Python，请先安装Python并添加到环境变量
    pause
    exit /b 1
)

:: 检查pyinstaller是否安装
pyinstaller --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ⚠️  未检测到pyinstaller，正在自动安装...
    pip install pyinstaller -i https://pypi.tuna.tsinghua.edu.cn/simple
    if %errorlevel% neq 0 (
        echo ❌ pyinstaller安装失败，请手动执行：pip install pyinstaller
        pause
        exit /b 1
    )
)

echo 🚀 开始打包...
echo.

:: 执行打包，加--clean清除缓存，--distpath指定输出目录
pyinstaller -F -w --clean -i NONE --distpath ./输出文件 main.py

if %errorlevel% equ 0 (
    echo.
    echo ==============================================
    echo ✅ 打包成功！
    echo 📂 生成的EXE文件在 输出文件 文件夹中
    echo ==============================================
) else (
    echo.
    echo ❌ 打包失败，请检查错误信息
)

echo.
pause