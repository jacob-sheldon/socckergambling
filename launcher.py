"""
Cross-platform launcher for the Soccer Betting Analysis web app.
Starts the Flask server and opens the default browser.
"""
import subprocess
import sys
import time
import webbrowser
from pathlib import Path


def main():
    port = 5100
    url = f"http://127.0.0.1:{port}"

    script_dir = Path(__file__).resolve().parent
    web_app = script_dir / "web_app.py"

    print("=" * 60)
    print("  足球彩票分析工具 - Web 启动器")
    print("=" * 60)
    print(f"  访问地址: {url}")
    print(f"  浏览器即将自动打开...")
    print("=" * 60)

    proc = subprocess.Popen(
        [sys.executable, str(web_app), "--port", str(port)],
        cwd=str(script_dir),
    )

    # Give the server a moment to start, then open the browser
    time.sleep(1.5)

    try:
        webbrowser.open(url)
    except Exception:
        print(f"  无法自动打开浏览器，请手动访问: {url}")

    print(f"  服务运行中 (PID: {proc.pid})")
    print(f"  按 Ctrl+C 停止服务")
    print("=" * 60)

    try:
        proc.wait()
    except KeyboardInterrupt:
        print("\n  正在停止服务...")
        proc.terminate()
        proc.wait()
        print("  服务已停止")


if __name__ == "__main__":
    main()
