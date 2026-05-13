@echo off
REM Build standalone app for Windows
REM Usage: build_win.bat

echo === 发票 Word 生成器 - Windows 打包 ===

cd /d "%~dp0"

pip install -r requirements.txt

echo Packaging with flet...
flet pack main.py ^
    --name "发票Word生成器" ^
    --icon assets\icon.png ^
    --add-data "engine.py;." ^
    --product-name "发票Word生成器" ^
    --product-version "1.0.0" ^
    --copyright "FSAE Team"

echo.
echo === 打包完成 ===
echo 输出位置: dist\
echo.
echo 如果没有 icon.png，可以去掉 --icon 参数重新运行。
pause
