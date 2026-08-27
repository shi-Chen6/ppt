@echo off
:: ============================================================
:: push.bat  ——  将本地最新文件推送到 GitHub 仓库
:: 仓库: git@github.com:shi-Chen6/ppt.git  (分支: main)
:: 用法:
::   push.bat            使用时间戳作为提交说明
::   push.bat "说明文字"  使用自定义提交说明
:: 说明: 自动遵守 .gitignore，不会推送 venv / __pycache__ /
::       output / .workbuddy 等被忽略的内容。
:: ============================================================
setlocal
cd /d "%~dp0"

echo [1/3] 暂存所有变更（遵循 .gitignore）...
git add -A

echo [2/3] 检查是否有需要提交的更改...
git diff --cached --quiet
if errorlevel 1 (
    if "%~1"=="" (
        git commit -m "Update: %DATE% %TIME%"
    ) else (
        git commit -m "%~1"
    )
    echo       已提交本地更改。
) else (
    echo       没有需要提交的更改，直接尝试推送。
)

echo [3/3] 推送到 origin/main ...
git push origin main
if errorlevel 1 (
    echo [错误] 推送失败。可能远程已有更新，请先运行 pull.bat 再试。
    pause
    exit /b 1
)
echo [完成] 已推送到 GitHub。
pause
