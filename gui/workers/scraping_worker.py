"""
Background worker for scraping match data using Playwright.
Runs in a separate QThread to avoid blocking the GUI.
"""

import asyncio
import os
import sys
from pathlib import Path
from PyQt6.QtCore import QThread, pyqtSignal
from typing import List, Optional


def _get_default_playwright_cache_dir() -> Path:
    """Return the default Playwright browser cache directory for this platform."""
    if sys.platform.startswith("win"):
        base = Path(os.environ.get("LOCALAPPDATA", Path.home() / "AppData" / "Local"))
        return base / "ms-playwright"
    if sys.platform == "darwin":
        return Path.home() / "Library" / "Caches" / "ms-playwright"
    return Path.home() / ".cache" / "ms-playwright"


def _setup_bundled_browser():
    """
    Configure Playwright to use the bundled browser in the .app package.
    This allows the app to run without requiring separate Playwright installation.
    """
    # Check if we're running from a PyInstaller bundle
    if not getattr(sys, 'frozen', False):
        return False

    # Try multiple strategies to find the bundled browsers directory
    browsers_path = None

    # Strategy 1: Check sys._MEIPASS (works for onedir builds)
    if hasattr(sys, '_MEIPASS'):
        path = Path(sys._MEIPASS) / 'ms-playwright'
        if path.exists():
            browsers_path = path
            print(f"[DEBUG] Found browsers via _MEIPASS: {browsers_path}")

    # Strategy 2: For macOS .app bundles, find Resources via executable path
    # .app structure: AppName.app/Contents/MacOS/executable -> ../Resources
    if browsers_path is None and 'MacOS' in Path(sys.executable).parts:
        exec_path = Path(sys.executable)
        resources_path = exec_path.parent.parent / 'Resources' / 'ms-playwright'
        if resources_path.exists():
            browsers_path = resources_path
            print(f"[DEBUG] Found browsers via executable path: {browsers_path}")

    # Strategy 3: Try looking relative to the executable (for onefile builds)
    if browsers_path is None:
        exec_dir = Path(sys.executable).parent
        path = exec_dir / 'ms-playwright'
        if path.exists():
            browsers_path = path
            print(f"[DEBUG] Found browsers via exec_dir: {browsers_path}")

    if browsers_path:
        os.environ['PLAYWRIGHT_BROWSERS_PATH'] = str(browsers_path)  # Directory containing chromium-*
        print(f"[DEBUG] Set PLAYWRIGHT_BROWSERS_PATH to: {browsers_path}")
        return True

    # Fall back to system cache if browsers were installed separately.
    cache_dir = _get_default_playwright_cache_dir()
    if cache_dir.exists():
        os.environ.setdefault('PLAYWRIGHT_BROWSERS_PATH', str(cache_dir))
        print(f"[DEBUG] Using Playwright cache from system: {cache_dir}")
        return True

    print(f"[DEBUG] Could not find bundled browsers. sys.frozen={sys.frozen}, _MEIPASS={getattr(sys, '_MEIPASS', 'N/A')}, executable={sys.executable}")
    return False


class ScrapingWorker(QThread):
    """
    Worker thread for fetching match data using Playwright browser automation.
    Emits signals to communicate progress and results back to the main thread.
    """

    # Signals
    progress_updated = pyqtSignal(str)  # Status message
    match_fetched = pyqtSignal(object)  # Individual MatchData object
    scraping_complete = pyqtSignal(list)  # List of all MatchData objects
    error_occurred = pyqtSignal(str)  # Error message

    def __init__(
        self,
        url: str,
        headless: bool = True,
        max_matches: Optional[int] = None,
        fetch_enhanced_odds: bool = False,
        fetch_asian_handicap: bool = False
    ):
        super().__init__()
        self.url = url
        self.headless = headless
        self.max_matches = max_matches
        self.fetch_enhanced_odds = fetch_enhanced_odds
        self.fetch_asian_handicap = fetch_asian_handicap
        self._is_running = True

    def stop(self):
        """Stop the worker thread."""
        self._is_running = False

    def run(self):
        """
        Called when thread starts. Creates a new event loop and runs async scraping.
        """
        try:
            # Create new event loop for this thread
            loop = asyncio.new_event_loop()
            asyncio.set_event_loop(loop)

            # Run async function
            matches = loop.run_until_complete(
                self._scrape_matches()
            )

            if self._is_running:
                self.scraping_complete.emit(matches)

        except Exception as e:
            self.error_occurred.emit(f"抓取失败: {str(e)}")
        finally:
            loop.close()

    async def _scrape_matches(self):
        """
        Async scraping logic. Imports from browser_bet_scraper module.
        """
        # Setup bundled browser path if running from PyInstaller bundle
        _setup_bundled_browser()

        from browser_bet_scraper import (
            fetch_matches_with_browser,
            fetch_asian_handicap_data,
            fetch_euro_kelly_data,
        )

        self.progress_updated.emit("正在初始化浏览器...")

        # Fetch basic match data
        try:
            matches = await fetch_matches_with_browser(
                self.url,
                self.headless
            )
            # Debug logging
            if matches:
                print(f"[DEBUG GUI] Fetched {len(matches)} matches")
                print(f"[DEBUG GUI] First match: {matches[0].match_id} | is_fallback: {matches[0].is_fallback}")
        except Exception as e:
            print(f"[DEBUG GUI] Exception in fetch_matches_with_browser: {type(e).__name__}: {e}")
            import traceback
            traceback.print_exc()
            self.error_occurred.emit(f"浏览器自动化异常: {str(e)}")
            return []

        if not self._is_running:
            return []

        # Check if we received fallback data
        if matches and matches[0].is_fallback:
            print(f"[DEBUG GUI] Fallback data detected! is_fallback={matches[0].is_fallback}")
            try:
                import browser_bet_scraper
                last_error = getattr(browser_bet_scraper, "LAST_SCRAPE_ERROR", "")
            except Exception:
                last_error = ""
            detail = ""
            if last_error:
                detail = f"\n\n详细错误:\n{last_error[:600]}"
            error_msg = (
                "浏览器自动化失败，已回退到示例数据。\n\n"
                "这可能是因为:\n"
                "• 网络连接问题\n"
                "• Playwright 浏览器未安装\n"
                "• 目标网站无法访问或被拦截\n\n"
                "如果已安装仍失败，可删除浏览器缓存后重试。\n"
                "Windows 缓存路径: %LOCALAPPDATA%\\ms-playwright\n"
                "macOS 缓存路径: ~/Library/Caches/ms-playwright"
            )
            self.error_occurred.emit(error_msg + detail)
            return []

        # Limit matches if specified
        if self.max_matches is not None:
            print(f"[DEBUG GUI] Limiting matches to {self.max_matches} (original count: {len(matches)})")
            matches = matches[:self.max_matches]

        # Emit each match as it's fetched
        for match in matches:
            if not self._is_running:
                break
            self.match_fetched.emit(match)

        # Optionally fetch Asian handicap data
        if self.fetch_asian_handicap and self._is_running and matches:
            self.progress_updated.emit("正在获取亚洲盘口数据...")
            matches = await fetch_asian_handicap_data(matches, self.headless)

            # Re-emit matches with updated data
            for match in matches:
                if not self._is_running:
                    break
                self.match_fetched.emit(match)

            if self._is_running:
                self.progress_updated.emit("正在获取百家欧赔即时凯利数据...")
                matches = await fetch_euro_kelly_data(matches, self.headless)

        self.progress_updated.emit(f"完成！共获取 {len(matches)} 场比赛")
        return matches
