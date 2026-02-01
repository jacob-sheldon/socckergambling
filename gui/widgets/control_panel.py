"""
Control panel widget with input fields, checkboxes, and action buttons.
"""

import io
import os
import sys
from contextlib import redirect_stderr, redirect_stdout
from pathlib import Path
from PyQt6.QtWidgets import (
    QWidget, QGridLayout, QLabel, QLineEdit, QPushButton,
    QSpinBox, QCheckBox, QFileDialog, QProgressBar
)
from PyQt6.QtCore import pyqtSignal, QProcess, QThread


class _StreamEmitter(io.TextIOBase):
    """Stream wrapper that emits text chunks via callback."""

    def __init__(self, emit_callback):
        super().__init__()
        self._emit = emit_callback

    def write(self, s):
        if s:
            self._emit(str(s))
        return len(s)

    def flush(self):
        return None


class PlaywrightInstallThread(QThread):
    """Run Playwright install in-process to avoid external CLI dependency."""

    output = pyqtSignal(str)
    error = pyqtSignal(str)
    done = pyqtSignal(int)

    def run(self):
        exit_code = 1
        try:
            from playwright.__main__ import main as playwright_main
        except Exception as exc:
            self.error.emit(f"无法导入 Playwright: {exc}")
            self.done.emit(exit_code)
            return

        argv_backup = sys.argv[:]
        sys.argv = ["playwright", "install", "chromium"]
        try:
            stdout_stream = _StreamEmitter(self.output.emit)
            stderr_stream = _StreamEmitter(self.error.emit)
            with redirect_stdout(stdout_stream), redirect_stderr(stderr_stream):
                try:
                    result = playwright_main()
                    exit_code = 0 if result is None else int(result)
                except SystemExit as exc:
                    if isinstance(exc.code, int):
                        exit_code = exc.code
                    else:
                        exit_code = 1
        except Exception as exc:
            self.error.emit(f"安装异常: {exc}")
            exit_code = 1
        finally:
            sys.argv = argv_backup
            self.done.emit(exit_code)


class ControlPanel(QWidget):
    """
    Control panel with user input controls for the scraping application.
    """

    # Signals
    start_scraping = pyqtSignal(dict)  # Emits dict with options
    export_excel = pyqtSignal()

    def __init__(self, parent=None):
        super().__init__(parent)

        # Default values
        self.default_url = "https://live.500.com/"
        self.default_output = os.path.expanduser("~/Desktop/live_betting_template.xlsx")
        self.browser_ready = False
        self._scraping_enabled = True
        self._installing_browser = False
        self._install_process = None
        self._install_thread = None

        self._init_ui()
        self._update_browser_status()

    def _init_ui(self):
        """Initialize the UI components."""
        layout = QGridLayout()
        layout.setSpacing(10)
        layout.setContentsMargins(10, 10, 10, 10)

        # Row 0: Output filename
        layout.addWidget(QLabel("输出文件:"), 0, 0)
        self.output_edit = QLineEdit(self.default_output)
        layout.addWidget(self.output_edit, 0, 1)

        self.browse_btn = QPushButton("浏览...")
        self.browse_btn.clicked.connect(self._browse_output_file)
        layout.addWidget(self.browse_btn, 0, 2)

        # Row 1: URL
        layout.addWidget(QLabel("数据源 URL:"), 1, 0)
        self.url_edit = QLineEdit(self.default_url)
        layout.addWidget(self.url_edit, 1, 1, 1, 2)

        # Row 2: Max matches
        layout.addWidget(QLabel("最大比赛数:"), 2, 0)
        self.max_matches_spin = QSpinBox()
        self.max_matches_spin.setRange(0, 500)  # 0 means "全部" (all)
        self.max_matches_spin.setValue(50)
        self.max_matches_spin.setSpecialValueText("全部")
        self.max_matches_spin.setValue(0)  # 0 means all
        layout.addWidget(self.max_matches_spin, 2, 1)

        # Row 3: Options (two checkboxes side by side)
        self.headless_check = QCheckBox("无头模式 (隐藏浏览器)")
        self.headless_check.setChecked(True)
        layout.addWidget(self.headless_check, 3, 1)

        self.enhanced_odds_check = QCheckBox("增强赔率数据 (较慢)")
        self.enhanced_odds_check.setVisible(False)

        self.asian_handicap_check = QCheckBox("亚洲盘口分析 (慢)")
        layout.addWidget(self.asian_handicap_check, 3, 2)  # Moved to column 2, row 3

        # Row 4: Browser status label
        self.browser_status_label = QLabel("")
        layout.addWidget(self.browser_status_label, 4, 1)

        # Progress bar (hidden by default, shares position with start_btn)
        self.install_progress = QProgressBar()
        self.install_progress.setTextVisible(True)
        self.install_progress.setRange(0, 100)  # 0-100%
        self.install_progress.setValue(0)
        self.install_progress.setVisible(False)
        layout.addWidget(self.install_progress, 5, 1)  # Shares position with start_btn

        # Row 5: Action buttons
        self.start_btn = QPushButton("开始抓取")
        self.start_btn.setFixedHeight(40)
        self.start_btn.clicked.connect(self._on_start_clicked)
        layout.addWidget(self.start_btn, 5, 1)

        self.export_btn = QPushButton("导出 Excel")
        self.export_btn.setFixedHeight(40)
        self.export_btn.setEnabled(False)
        self.export_btn.clicked.connect(self.export_excel.emit)
        layout.addWidget(self.export_btn, 5, 2)

        self.setLayout(layout)

    def _browse_output_file(self):
        """Open file dialog to select output file."""
        filename, _ = QFileDialog.getSaveFileName(
            self,
            "选择输出文件",
            self.output_edit.text(),
            "Excel Files (*.xlsx)"
        )
        if filename:
            self.output_edit.setText(filename)

    def _on_start_clicked(self):
        """Emit signal with scraping options."""
        if self._installing_browser:
            return
        if not self.browser_ready:
            self._install_browser()
            return
        options = {
            'url': self.url_edit.text(),
            'output': self.output_edit.text(),
            'headless': self.headless_check.isChecked(),
            'max_matches': self.max_matches_spin.value() or None,
            'enhanced_odds': self.enhanced_odds_check.isChecked(),
            'asian_handicap': self.asian_handicap_check.isChecked(),
        }
        self.start_scraping.emit(options)

    def set_scraping_enabled(self, enabled: bool):
        """Enable or disable the start button during scraping."""
        self._scraping_enabled = enabled
        self._update_start_button()

    def set_export_enabled(self, enabled: bool):
        """Enable or disable the export button."""
        self.export_btn.setEnabled(enabled)

    def get_output_filename(self) -> str:
        """Get the current output filename."""
        return self.output_edit.text()

    def _update_start_button(self):
        """Update start button state based on browser/install/scrape status."""
        if self._installing_browser:
            self.start_btn.setEnabled(False)
            self.start_btn.setText("安装浏览器中...")
            return

        if not self.browser_ready:
            self.start_btn.setEnabled(True)
            self.start_btn.setText("安装浏览器")
            return

        if self._scraping_enabled:
            self.start_btn.setEnabled(True)
            self.start_btn.setText("开始抓取")
        else:
            self.start_btn.setEnabled(False)
            self.start_btn.setText("抓取中...")

    def _get_playwright_cache_dir(self) -> Path:
        """Get Playwright cache directory for current platform."""
        if sys.platform.startswith("win"):
            base = Path(os.environ.get("LOCALAPPDATA", Path.home() / "AppData" / "Local"))
            return base / "ms-playwright"
        if sys.platform == "darwin":
            return Path.home() / "Library" / "Caches" / "ms-playwright"
        return Path.home() / ".cache" / "ms-playwright"

    def _is_browser_installed(self) -> bool:
        """Check if Playwright Chromium is installed."""
        cache_dir = self._get_playwright_cache_dir()
        if not cache_dir.exists():
            return False

        for item in cache_dir.iterdir():
            if not item.is_dir():
                continue
            if item.name.startswith("chromium-") and "headless" not in item.name:
                for sub in item.iterdir():
                    if sub.is_dir() and sub.name.startswith("chrome-"):
                        return True
        return False

    def _update_browser_status(self):
        """Update browser status label and button states."""
        self.browser_ready = self._is_browser_installed()
        if self.browser_ready:
            self.browser_status_label.setText("浏览器已安装")
        else:
            self.browser_status_label.setText("未安装浏览器 (需安装 Chromium)")

        self._update_start_button()

    def _install_browser(self):
        """Install Playwright Chromium browser."""
        if self._installing_browser:
            return

        self._installing_browser = True
        self._update_start_button()

        # Hide start button and show progress bar
        self.start_btn.setVisible(False)
        self.install_progress.setVisible(True)
        self.install_progress.setValue(0)
        self.install_progress.setFormat("准备下载...")

        if getattr(sys, "frozen", False):
            # Run Playwright install via Python module to avoid external CLI dependency.
            self._install_thread = PlaywrightInstallThread()
            self._install_thread.output.connect(self._on_install_output_text)
            self._install_thread.error.connect(self._on_install_error_text)
            self._install_thread.done.connect(self._on_install_finished)
            self._install_thread.finished.connect(self._install_thread.deleteLater)
            self._install_thread.start()
        else:
            # Use python -m playwright for source runs to avoid thread/GUI edge cases.
            self._install_process = QProcess(self)
            self._install_process.setProgram(sys.executable)
            self._install_process.setArguments(["-m", "playwright", "install", "chromium"])
            self._install_process.readyReadStandardOutput.connect(self._on_install_output_process)
            self._install_process.readyReadStandardError.connect(self._on_install_error_process)
            self._install_process.finished.connect(self._on_install_finished)
            self._install_process.start()

    def _handle_install_output(self, output: str):
        """Parse install output and update progress."""
        if not output:
            return

        output_lower = output.lower()

        if "downloading" in output_lower or "download" in output_lower:
            self.install_progress.setFormat("下载中...")
            self.install_progress.setValue(30)
        elif "extracting" in output_lower or "extract" in output_lower:
            self.install_progress.setFormat("解压中...")
            self.install_progress.setValue(60)
        elif "installing" in output_lower or "installed" in output_lower:
            self.install_progress.setFormat("安装中...")
            self.install_progress.setValue(90)
        elif "chromium" in output_lower and "installed" in output_lower:
            self.install_progress.setValue(100)
            self.install_progress.setFormat("完成")

    def _is_ignorable_install_message(self, text: str) -> bool:
        """Filter out known non-fatal warnings (e.g., Node deprecations)."""
        if not text:
            return True
        text_upper = text.upper()
        if "DEP0169" in text_upper:
            return True
        if "DEPRECATIONWARNING" in text_upper and "DEP0169" in text_upper:
            return True
        return False

    def _on_install_output_text(self, output: str):
        """Handle installation output text."""
        self._handle_install_output(output.strip())

    def _on_install_error_text(self, error: str):
        """Handle installation error output text."""
        if error and not self._is_ignorable_install_message(error):
            self.install_progress.setFormat(f"错误: {error.strip()[:50]}...")

    def _on_install_output_process(self):
        """Handle installation output from QProcess."""
        data = self._install_process.readAllStandardOutput()
        output = bytes(data).decode("utf-8", errors="ignore")
        self._handle_install_output(output.strip())

    def _on_install_error_process(self):
        """Handle installation error output from QProcess."""
        data = self._install_process.readAllStandardError()
        error = bytes(data).decode("utf-8", errors="ignore").strip()
        if error and not self._is_ignorable_install_message(error):
            self.install_progress.setFormat(f"错误: {error[:50]}...")

    def _on_install_finished(self, exit_code, _exit_status=None):
        """Handle browser install completion."""
        self._installing_browser = False
        self._install_thread = None
        self._install_process = None
        self._install_thread = None

        if exit_code == 0:
            self.install_progress.setValue(100)
            self.install_progress.setFormat("安装完成")
        else:
            self.install_progress.setFormat("安装失败")

        # Hide progress bar after a delay and show start button
        from PyQt6.QtCore import QTimer
        QTimer.singleShot(3000, self._restore_start_button)

        self._update_browser_status()

    def _restore_start_button(self):
        """Restore start button and hide progress bar after installation."""
        self.install_progress.setVisible(False)
        self.start_btn.setVisible(True)
