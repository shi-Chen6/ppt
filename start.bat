@echo off
setlocal EnableDelayedExpansion
cd /d "%~dp0"
title PPT 智能生成工具

echo ============================================
echo    PPT 智能生成工具 - 一键启动
echo ============================================
echo.

rem ---- 1) 若服务已在运行，直接打开页面 ----
curl -s -o nul -m 2 http://127.0.0.1:5000/ 2>nul
if not errorlevel 1 (
    echo 服务已在运行，直接打开页面...
    start "" http://127.0.0.1:5000
    exit /b 0
)

rem ---- 2) 选择 Python（优先系统自带的 3.12，依赖最全）----
set "PY=C:\Users\Administrator\AppData\Local\Programs\Python\Python312\python.exe"
if not exist "!PY!" (
    where python >nul 2>nul
    if not errorlevel 1 ( set "PY=python" ) else ( set "PY=" )
)
if not defined PY (
    echo [错误] 未找到 Python，请先安装 Python 3.10 及以上版本。
    pause
    exit /b 1
)
echo 使用 Python: !PY!

rem ---- 3) 检查依赖，缺失则自动安装 ----
"!PY!" -c "import flask, requests, pptx, docx, PIL, matplotlib" 2>nul
if errorlevel 1 (
    echo 检测到依赖缺失，正在自动安装（首次较慢，请稍候）...
    "!PY!" -m pip install -r requirements.txt
    if errorlevel 1 (
        echo [错误] 依赖安装失败，请检查网络后重试。
        pause
        exit /b 1
    )
)

rem ---- 4) 启动服务，稍后自动打开浏览器 ----
echo 正在启动服务，稍后会自动打开 http://127.0.0.1:5000
echo （启动后请勿关闭本窗口，按 Ctrl+C 可停止服务）
start "" /min cmd /c "timeout /t 4 >nul & start http://127.0.0.1:5000"
"!PY!" app.py

echo.
echo 服务已停止。按任意键关闭窗口...
pause >nul
