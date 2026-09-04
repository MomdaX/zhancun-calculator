@echo off
chcp 65001 >nul
title 站存计算器 - 同步启动
echo ============================================
echo   站存计算器
echo   正在从 SIM2.9 同步数据...
echo ============================================

set "SRC=C:\Program Files\SIM2.9\TranData"
set "DST=%~dp0HTML\TranData"

if not exist "%SRC%" (
    echo [警告] 源目录不存在: %SRC%
    echo 请确认 SIM2.9 已安装，或手动修改本文件中的 SRC 路径
    pause
    start "" "%~dp0HTML\index.html"
    exit /b 0
)

if not exist "%DST%" mkdir "%DST%"

robocopy "%SRC%" "%DST%" *.xls *.xlsm *.xlsx /MIR /R:3 /W:3 /NP /NDL /NJH /NJS

if %ERRORLEVEL% GEQ 8 (
    echo [错误] 同步失败！
    pause
) else (
    echo [完成] 同步成功，正在打开...
    timeout /t 1 /nobreak >nul
)

start "" "%~dp0HTML\index.html"
exit /b 0