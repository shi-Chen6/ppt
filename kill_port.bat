@echo off
chcp 65001 >nul
setlocal EnableDelayedExpansion
cd /d "%~dp0"
title 释放端口

:: ============================================================
:: kill_port.bat —— 结束占用指定端口的进程，释放该端口
:: 用法:
::   kill_port.bat            (默认释放 5000 端口)
::   kill_port.bat 8080       (释放 8080 端口)
::   kill_port.bat 5000 /q    (静默模式，结束不暂停，供其他脚本调用)
:: 说明: 只终止处于 LISTENING 状态的进程，TIME_WAIT 等残留连接
::       会在几十秒内自动消失，无需处理。
:: ============================================================

set "PORT=%~1"
if not defined PORT set "PORT=5000"

echo ============================================
echo    释放端口 %PORT%
echo ============================================
echo.

set "TMPF=%TEMP%\kill_port_%PORT%_%RANDOM%.tmp"
if exist "%TMPF%" del "%TMPF%" >nul 2>nul

set "COUNT=0"
set "FAILED=0"

rem ---- 逐个处理监听该端口的进程 ----
for /f "tokens=2,5" %%a in ('netstat -ano ^| findstr /C:":%PORT% " ^| findstr /C:"LISTENING"') do (
    set "ADDR=%%a"
    set "PID=%%b"

    rem 取地址最后一个冒号后的部分，避免 5000 误匹配 50000
    set "LAST=!ADDR::= !"
    set "P="
    for %%p in (!LAST!) do set "P=%%p"

    if "!P!"=="%PORT%" (
        rem 同一 PID 可能占多行，去重只杀一次
        findstr /X /C:"!PID!" "%TMPF%" >nul 2>nul
        if errorlevel 1 (
            >>"%TMPF%" echo !PID!

            rem 查询进程名，方便确认杀的是什么
            set "PNAME=未知进程"
            for /f "tokens=1" %%n in ('tasklist /FI "PID eq !PID!" /NH 2^>nul') do (
                if not "%%n"=="信息:" set "PNAME=%%n"
            )

            echo   发现 !PNAME!  ^(PID=!PID!^)  监听 !ADDR!
            taskkill /PID !PID! /F >nul 2>nul
            if errorlevel 1 (
                echo     [失败] 无法终止，可能权限不足，请以管理员身份重试。
                set "FAILED=1"
            ) else (
                echo     [已终止]
                set /a COUNT+=1
            )
        )
    )
)

if exist "%TMPF%" del "%TMPF%" >nul 2>nul

echo.
if "!FAILED!"=="0" (
    if "!COUNT!"=="0" (
        echo 端口 %PORT% 当前没有被任何进程监听，无需处理。
    ) else (
        echo 共终止 !COUNT! 个进程。
    )
) else (
    if "!COUNT!"=="0" (
        echo [未完成] 有进程占用端口但无法终止，请以管理员身份重新运行本脚本。
    ) else (
        echo 共终止 !COUNT! 个进程，另有部分进程终止失败，请以管理员身份重试。
    )
)

rem ---- 验证端口是否真正释放 ----
echo.
echo 正在验证...
ping -n 2 127.0.0.1 >nul
netstat -ano | findstr /C:":%PORT% " | findstr /C:"LISTENING" >nul
if errorlevel 1 (
    echo [完成] 端口 %PORT% 已释放。
) else (
    echo [警告] 端口 %PORT% 仍被以下进程占用：
    netstat -ano | findstr /C:":%PORT% " | findstr /C:"LISTENING"
    echo 建议以管理员身份重新运行本脚本。
)

echo.
if /i not "%~2"=="/q" pause
