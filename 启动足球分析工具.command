#!/bin/bash
cd "$(dirname "$0")"

echo "正在启动足球彩票分析工具..."

# Check if dependencies are installed
if ! uv run python -c "import flask" 2>/dev/null; then
    echo "首次运行，正在安装依赖..."
    uv sync
    if [ $? -ne 0 ]; then
        echo "依赖安装失败，请检查网络后重试"
        read -p "按回车键退出..."
        exit 1
    fi
fi

# Check if Chromium is installed
if ! uv run python -c "
from browser_bet_scraper import _get_default_playwright_cache_dir
import sys
cache = _get_default_playwright_cache_dir()
sys.exit(0 if cache.exists() and list(cache.glob('chromium-*')) else 1)
" 2>/dev/null; then
    echo "首次运行，正在安装 Chromium 浏览器（约 150MB）..."
    uv run playwright install chromium
fi

uv run python launcher.py
