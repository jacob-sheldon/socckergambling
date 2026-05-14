@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo 正在启动足球彩票分析工具...
uv run python launcher.py
pause
