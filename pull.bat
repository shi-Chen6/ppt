@echo off
:: ============================================================
:: pull.bat  ——  从 GitHub 仓库拉取最新文件到本地
:: 仓库: git@github.com:shi-Chen6/ppt.git  (分支: main)
:: 用法:
::   pull.bat
:: 说明: 将远程 main 分支的最新内容合并到本地当前分支。
:: ============================================================
setlocal
cd /d "%~dp0"

echo [1/2] 从 origin/main 拉取最新文件...
git pull origin main
if errorlevel 1 (
    echo [错误] 拉取失败。请检查网络连接或 SSH 密钥是否已加入 GitHub。
    pause
    exit /b 1
)

echo [完成] 已拉取最新文件到本地。
pause
