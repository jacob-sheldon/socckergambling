@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion

echo ============================================================
echo   足球彩票分析工具 - 一键安装 (Windows)
echo ============================================================
echo.

:: Check Python
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [错误] 未检测到 Python，请先安装 Python 3.12+
    echo 下载地址: https://www.python.org/downloads/
    echo 安装时务必勾选 "Add Python to PATH"
    pause
    exit /b 1
)
echo [OK] Python 已检测到

:: Install uv
uv --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [安装] 正在安装 uv 包管理器...
    powershell -Command "irm https://astral.sh/uv/install.ps1 | iex"
    if %errorlevel% neq 0 (
        echo [错误] uv 安装失败，请手动安装: https://github.com/astral-sh/uv
        pause
        exit /b 1
    )
)
echo [OK] uv 已就绪

:: Sync dependencies
echo [安装] 正在安装项目依赖...
cd /d "%~dp0"
uv sync
if %errorlevel% neq 0 (
    echo [错误] 依赖安装失败
    pause
    exit /b 1
)
echo [OK] 依赖安装完成

:: Install Chromium
echo [安装] 正在安装 Chromium 浏览器（约 150MB）...
uv run playwright install chromium
if %errorlevel% neq 0 (
    echo [警告] Chromium 安装可能失败，启动后可在网页中手动安装
)
echo [OK] Chromium 安装完成

echo.
echo ============================================================
echo   安装完成！现在启动工具...
echo ============================================================
echo.

uv run python launcher.py
pause
