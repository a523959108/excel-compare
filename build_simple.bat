@echo off
chcp 65001 >nul
echo 正在简化打包，排除复杂参数...
pyinstaller -F main.py
echo.
if exist dist\main.exe (
    echo 打包成功！EXE在dist目录下
) else (
    echo 打包失败，请把上面的错误信息发我
)
pause