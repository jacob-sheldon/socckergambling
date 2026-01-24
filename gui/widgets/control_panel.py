"""
Control panel widget with input fields, checkboxes, and action buttons.
"""

import os
import sys
from pathlib import Path
from PyQt6.QtWidgets import (
    QWidget, QGridLayout, QLabel, QLineEdit, QPushButton,
    QSpinBox, QCheckBox, QFileDialog
)
from PyQt6.QtCore import pyqtSignal, QProcess


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

        # Row 3: Options
        self.headless_check = QCheckBox("无头模式 (隐藏浏览器)")
        self.headless_check.setChecked(True)
        layout.addWidget(self.headless_check, 3, 1)

        self.enhanced_odds_check = QCheckBox("增强赔率数据 (较慢)")
        self.enhanced_odds_check.setVisible(False)

        self.asian_handicap_check = QCheckBox("亚洲盘口分析 (慢)")
        layout.addWidget(self.asian_handicap_check, 4, 1)

        # Row 4: Browser install
        self.browser_status_label = QLabel("")
        layout.addWidget(self.browser_status_label, 4, 2)

        self.install_browser_btn = QPushButton("安装浏览器")
        self.install_browser_btn.clicked.connect(self._install_browser)
        layout.addWidget(self.install_browser_btn, 4, 0)

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

        self.install_browser_btn.setVisible(not self.browser_ready)
        self.install_browser_btn.setEnabled(not self.browser_ready and not self._installing_browser)
        self._update_start_button()

    def _install_browser(self):
        """Install Playwright Chromium browser."""
        if self._installing_browser:
            return

        self._installing_browser = True
        self.install_browser_btn.setEnabled(False)
        self._update_start_button()
        self.browser_status_label.setText("正在安装浏览器...")

        self._install_process = QProcess(self)
        if getattr(sys, "frozen", False):
            self._install_process.setProgram("playwright")
            self._install_process.setArguments(["install", "chromium"])
        else:
            self._install_process.setProgram(sys.executable)
            self._install_process.setArguments(["-m", "playwright", "install", "chromium"])

        self._install_process.finished.connect(self._on_install_finished)
        self._install_process.start()

    def _on_install_finished(self, exit_code, _exit_status):
        """Handle browser install completion."""
        self._installing_browser = False
        self._update_browser_status()
        if exit_code == 0:
            self.browser_status_label.setText("浏览器安装完成")
        else:
            self.browser_status_label.setText("浏览器安装失败，请重试")
