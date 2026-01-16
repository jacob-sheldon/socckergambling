"""
Control panel widget with input fields, checkboxes, and action buttons.
"""

import os
from PyQt6.QtWidgets import (
    QWidget, QGridLayout, QLabel, QLineEdit, QPushButton,
    QSpinBox, QCheckBox, QFileDialog
)
from PyQt6.QtCore import pyqtSignal


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

        self._init_ui()

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
        layout.addWidget(self.enhanced_odds_check, 3, 2)

        self.asian_handicap_check = QCheckBox("亚洲盘口分析 (慢)")
        layout.addWidget(self.asian_handicap_check, 4, 1)

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
        self.start_btn.setEnabled(enabled)
        if enabled:
            self.start_btn.setText("开始抓取")
        else:
            self.start_btn.setText("抓取中...")

    def set_export_enabled(self, enabled: bool):
        """Enable or disable the export button."""
        self.export_btn.setEnabled(enabled)

    def get_output_filename(self) -> str:
        """Get the current output filename."""
        return self.output_edit.text()
