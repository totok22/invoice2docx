@echo off
REM Build standalone app for Windows
REM Usage: build_win.bat

echo === 发票 Word 生成器 - Windows 打包 ===

cd /d "%~dp0"

set PYTHON_BIN=python
if exist ".venv\Scripts\python.exe" set PYTHON_BIN=.venv\Scripts\python.exe
set FLET_BIN=flet
if exist ".venv\Scripts\flet.exe" set FLET_BIN=.venv\Scripts\flet.exe

%PYTHON_BIN% -m pip install -r requirements.txt

echo Packaging with flet...
%FLET_BIN% pack main.py ^
    --yes ^
    --name "发票Word生成器" ^
    --add-data "engine.py;." ^
    --add-data "默认报账说明模板.docx;." ^
    --add-data "默认验收单模板.docx;." ^
    --product-name "发票Word生成器" ^
    --product-version "0.1.0" ^
    --file-version "0.1.0.0" ^
    --company-name "BITFSAE" ^
    --copyright "BITFSAE"

echo.
echo === 打包完成 ===
echo 输出位置: dist\
echo.
pause
