@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo 正在启动足球彩票分析工具...

:: Check if dependencies are installed
uv run python -c "import flask" >nul 2>&1
if %errorlevel% neq 0 (
    echo 首次运行，正在安装依赖...
    uv sync
    if %errorlevel% neq 0 (
        echo 依赖安装失败，请检查网络后重试
        pause
        exit /b 1
    )
)

:: Check if Chromium is installed
uv run python -c "from browser_bet_scraper import _get_default_playwright_cache_dir; import sys; cache = _get_default_playwright_cache_dir(); sys.exit(0 if cache.exists() and list(cache.glob('chromium-*')) else 1)" >nul 2>&1
if %errorlevel% neq 0 (
    echo 首次运行，正在安装 Chromium 浏览器（约 150MB）...
    uv run playwright install chromium
)

uv run python launcher.py
