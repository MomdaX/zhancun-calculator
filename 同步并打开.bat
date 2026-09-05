@echo off
chcp 65001 >nul
setlocal EnableExtensions

:: ===== 1. 管理员自提升（Program Files 下的数据目录常需管理员权限才能读）=====
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo 需要管理员权限，正在请求提权...
    powershell -NoProfile -Command "Start-Process -FilePath '%~f0' -Verb RunAs"
    exit /b
)

title 站存计算器 - 同步启动

:: ===== 2. 探测真实的 TranData 源路径 =====
set "SRC="
for %%D in (
    "C:\Program Files\SIM2.9\TranData"
    "C:\Program Files (x86)\SIM2.9\TranData"
) do (
    if exist "%%~D" set "SRC=%%~D"
)
:: 手动覆盖：在脚本同目录新建 src.txt，首行写入真实路径即可（例如 D:\SIM2.9\TranData）
if exist "%~dp0src.txt" set /p SRC=<"%~dp0src.txt"

if not defined SRC (
    cls
    echo [错误] 未找到 SIM2.9 的 TranData 源目录。
    echo 已尝试：
    echo   C:\Program Files\SIM2.9\TranData
    echo   C:\Program Files (x86)\SIM2.9\TranData
    echo.
    echo 解决办法（任选其一）：
    echo   1. 在脚本同目录新建 src.txt，首行写入真实路径，例如：
    echo      D:\SIM2.9\TranData
    echo   2. 打开资源管理器，找到 SIM2.9 实际安装位置下的 TranData，把它的路径写进 src.txt。
    pause
    exit /b 1
)

:: ===== 3. 同步到本地 HTML\TranData（应用只读这个本地副本）=====
set "DST=%~dp0HTML\TranData"
if not exist "%DST%" mkdir "%DST%"

echo ============================================
echo   站存计算器
echo   源: %SRC%
echo   目标: %DST%
echo   正在同步...
echo ============================================
robocopy "%SRC%" "%DST%" *.xls *.xlsm *.xlsx /MIR /R:3 /W:3 /NP /NDL /NJH /NJS

if %ERRORLEVEL% GEQ 8 (
    echo [错误] 同步失败（ERRORLEVEL=%ERRORLEVEL%）
    pause
) else (
    echo [完成] 同步成功，正在打开...
    timeout /t 1 /nobreak >nul
    start "" "%~dp0HTML\index.html"
)
exit /b 0
