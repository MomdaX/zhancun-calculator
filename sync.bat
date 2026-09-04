@echo off
chcp 65001 >nul
set "SRC=C:\Program Files\SIM2.9\TranData"
set "DST=%~dp0HTML\TranData"

if not exist "%SRC%" (
    echo [错误] 源目录不存在: %SRC%
    pause
    exit /b 1
)

if not exist "%DST%" mkdir "%DST%"
robocopy "%SRC%" "%DST%" *.xls *.xlsm *.xlsx /MIR /R:3 /W:3 /NP /NDL /NJH /NJS
echo 完成！
timeout /t 2 /nobreak >nul
exit /b 0