@echo off
echo 启动 XJ 能源平衡计算系统...
echo.
echo 请选择要运行的版本：
echo 1. 英文名称版本
echo 2. 中文名称版本
echo.
set /p choice=请输入选择 (1 或 2): 

if "%choice%"=="1" (
    start "" "dist\XJ-EnergyBalanceSystem.exe"
) else if "%choice%"=="2" (
    start "" "dist\XJ-能源平衡计算系统.exe"
) else (
    echo 无效选择，启动英文名称版本...
    start "" "dist\XJ-EnergyBalanceSystem.exe"
)

echo 程序已启动，请稍候...
pause