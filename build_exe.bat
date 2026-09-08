@echo off
REM build_exe.bat - Windows 一键构建 EXE + 安装包（双击即可）
chcp 65001 > nul
cd /d "%~dp0"

REM 定位 Python：优先使用已验证可用的系统 Python 3.9（带 tkinter + 已装依赖）
REM 注意：直接双击时系统 PATH 里可能没有 python 命令，必须用完整路径
set "PYTHON="
if exist "D:\Python\Python39\python.exe" (
    set "PYTHON=D:\Python\Python39\python.exe"
) else (
    set "PYTHON=python"
)
echo 使用 Python：%PYTHON%
echo.

echo ====================================
echo  ① 安装依赖
echo ====================================
"%PYTHON%" -m pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple
if errorlevel 1 (
    echo 依赖安装失败，请检查 Python 与网络
    pause
    exit /b 1
)

echo.
echo ====================================
echo  ② 打包 EXE（首次约需 1-3 分钟，请耐心等待）
echo ====================================
"%PYTHON%" build_exe.py
if errorlevel 1 (
    echo 打包失败
    pause
    exit /b 1
)

echo.
echo ====================================
echo  ③ 生成安装包（首次约需 30-60 秒）
echo ====================================
REM 定位 NSIS
set "MAKENSIS="
if exist "C:\Program Files (x86)\NSIS\makensis.exe" (
    set "MAKENSIS=C:\Program Files (x86)\NSIS\makensis.exe"
) else (
    where makensis >nul 2>nul
    if not errorlevel 1 set "MAKENSIS=makensis"
)
if "%MAKENSIS%"=="" (
    echo 未找到 NSIS。请到 https://nsis.sourceforge.io/Download 安装 NSIS 3.x 后重试。
    pause
    exit /b 1
)

REM NSIS 要求脚本带 UTF-8 BOM，自动写入（避免编辑器去掉 BOM 后报错）
    powershell -NoProfile -Command "$p='installer.nsi'; if (Test-Path $p) { $b=[System.IO.File]::ReadAllBytes($p); if ($b.Length -lt 3 -or $b[0] -ne 0xEF -or $b[1] -ne 0xBB -or $b[2] -ne 0xBF) { [System.IO.File]::WriteAllBytes($p, (0xEF,0xBB,0xBF + [byte[]]$b)) } }"
"%MAKENSIS%" /V2 installer.nsi
if errorlevel 1 (
    echo 安装包生成失败
    pause
    exit /b 1
)

echo.
echo ====================================
echo  构建完成！
echo   EXE：    dist\价格牌处理工具.exe
echo   安装包：  dist\价格牌处理工具_安装程序_v1.3.exe
echo ====================================
explorer dist
pause