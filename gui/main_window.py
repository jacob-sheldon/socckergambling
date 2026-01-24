"""
Main application window for the Soccer Betting Analysis Tool.
"""

import os
import sys
from PyQt6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QSplitter,
    QMessageBox, QFileDialog, QStatusBar
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QAction

from gui.widgets.control_panel import ControlPanel
from gui.widgets.match_table import MatchTable
from gui.widgets.progress_dialog import ProgressDialog
from gui.workers.scraping_worker import ScrapingWorker


class MainWindow(QMainWindow):
    """
    Main application window with control panel and match table.
    """

    def __init__(self):
        super().__init__()

        self.worker = None
        self.matches = []

        self._init_ui()
        self._connect_signals()

    def _init_ui(self):
        """Initialize the main window UI."""
        self.setWindowTitle("足球彩票分析工具")
        self.setMinimumSize(1200, 800)

        # Create menu bar
        self._create_menu_bar()

        # Create central widget with splitter
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        layout = QVBoxLayout(central_widget)
        layout.setContentsMargins(10, 10, 10, 10)

        # Create splitter for resizable panels
        splitter = QSplitter(Qt.Orientation.Vertical)

        # Control panel (top)
        self.control_panel = ControlPanel()
        splitter.addWidget(self.control_panel)

        # Match table (bottom)
        self.match_table = MatchTable()
        splitter.addWidget(self.match_table)

        # Set splitter proportions
        splitter.setStretchFactor(0, 0)  # Control panel doesn't stretch
        splitter.setStretchFactor(1, 1)  # Table stretches

        layout.addWidget(splitter)

        # Create status bar
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("就绪")

    def _create_menu_bar(self):
        """Create the application menu bar."""
        menubar = self.menuBar()

        # File menu
        file_menu = menubar.addMenu("文件")

        export_action = QAction("导出 Excel", self)
        export_action.setShortcut("Ctrl+S")
        export_action.triggered.connect(self._export_excel)
        export_action.setEnabled(False)
        self.export_action = export_action
        file_menu.addAction(export_action)

        file_menu.addSeparator()

        exit_action = QAction("退出", self)
        exit_action.setShortcut("Ctrl+Q")
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)

        # Help menu
        help_menu = menubar.addMenu("帮助")

        about_action = QAction("关于", self)
        about_action.triggered.connect(self._show_about)
        help_menu.addAction(about_action)

    def _connect_signals(self):
        """Connect signals from child widgets."""
        self.control_panel.start_scraping.connect(self._on_start_scraping)
        self.control_panel.export_excel.connect(self._export_excel)

    def _on_start_scraping(self, options: dict):
        """Handle start scraping signal from control panel."""
        # Clear previous data
        self.match_table.clear_matches()
        self.matches = []
        self.export_action.setEnabled(False)
        self.control_panel.set_export_enabled(False)

        # Show progress dialog
        self.progress_dialog = ProgressDialog(self)
        self.progress_dialog.show()

        # Create and start worker thread
        self.worker = ScrapingWorker(
            url=options['url'],
            headless=options['headless'],
            max_matches=options['max_matches'],
            fetch_enhanced_odds=options['enhanced_odds'],
            fetch_asian_handicap=options['asian_handicap']
        )

        # Connect worker signals
        self.worker.progress_updated.connect(self._on_progress_updated)
        self.worker.match_fetched.connect(self._on_match_fetched)
        self.worker.scraping_complete.connect(self._on_scraping_complete)
        self.worker.error_occurred.connect(self._on_error_occurred)

        # Disable controls during scraping
        self.control_panel.set_scraping_enabled(False)

        # Start worker
        self.worker.start()

    def _on_progress_updated(self, message: str):
        """Handle progress update from worker."""
        self.progress_dialog.set_status(message)
        self.status_bar.showMessage(message)

    def _on_match_fetched(self, match):
        """Handle individual match fetched from worker."""
        # Upsert by match_id to avoid duplicates when data refreshes.
        for idx, existing in enumerate(self.matches):
            if existing.match_id == match.match_id:
                self.matches[idx] = match
                break
        else:
            self.matches.append(match)

        self.match_table.add_match(match)

    def _on_scraping_complete(self, matches):
        """Handle scraping completion from worker."""
        self.progress_dialog.close()
        self.control_panel.set_scraping_enabled(True)
        self.export_action.setEnabled(True)
        self.control_panel.set_export_enabled(True)

        # Refresh table to ensure latest data (asian handicap / euro kelly).
        self.matches = matches
        self.match_table.clear_matches()
        for match in matches:
            self.match_table.add_match(match)
        self.match_table.scrollToTop()

        self.status_bar.showMessage(f"完成！共获取 {len(matches)} 场比赛", 5000)

        self.worker = None

    def _on_error_occurred(self, error_message: str):
        """Handle error from worker."""
        self.progress_dialog.close()
        self.control_panel.set_scraping_enabled(True)

        # Check if this is a fallback data warning
        if "示例数据" in error_message:
            QMessageBox.warning(
                self,
                "数据获取失败",
                error_message
            )
        else:
            QMessageBox.critical(
                self,
                "抓取失败",
                f"抓取数据时出错:\n\n{error_message}"
            )

        self.status_bar.showMessage("抓取失败", 5000)
        self.worker = None

    def _export_excel(self):
        """Export matches to Excel file."""
        if not self.matches:
            QMessageBox.warning(
                self,
                "没有数据",
                "没有可导出的比赛数据。请先抓取数据。"
            )
            return

        # Get output filename
        output_file = self.control_panel.get_output_filename()

        # Show file dialog
        filename, _ = QFileDialog.getSaveFileName(
            self,
            "导出 Excel 文件",
            output_file,
            "Excel Files (*.xlsx)"
        )

        if not filename:
            return

        try:
            from browser_bet_scraper import (
                create_template_workbook,
                set_column_widths,
                merge_header_cells,
                style_header_rows,
                add_match_data
            )

            # Update status
            self.status_bar.showMessage("正在导出 Excel...")

            # Create workbook
            wb, ws = create_template_workbook()

            # Apply formatting
            set_column_widths(ws)
            merge_header_cells(ws)
            style_header_rows(ws)

            # Add match data
            current_row = 3
            for match in self.matches:
                rows_added = add_match_data(ws, current_row, match)
                current_row += rows_added

            # Freeze header rows
            ws.freeze_panes = "A3"

            # Save workbook
            wb.save(filename)

            QMessageBox.information(
                self,
                "导出成功",
                f"Excel 文件已保存至:\n{filename}"
            )

            self.status_bar.showMessage(f"已导出 {len(self.matches)} 场比赛到 Excel", 5000)

        except Exception as e:
            QMessageBox.critical(
                self,
                "导出失败",
                f"导出 Excel 文件时出错:\n\n{str(e)}"
            )
            self.status_bar.showMessage("导出失败", 5000)

    def _show_about(self):
        """Show about dialog."""
        QMessageBox.about(
            self,
            "关于",
            """<h3>足球彩票分析工具</h3>
            <p>版本: 1.0.0</p>
            <p>使用浏览器自动化从 live.500.com 抓取实时比赛数据，
            并生成专业的足球彩票分析 Excel 模板。</p>
            <p><b>技术栈:</b> PyQt6 + Playwright</p>"""
        )

    def closeEvent(self, event):
        """Handle window close event."""
        # Stop worker if running
        if self.worker and self.worker.isRunning():
            reply = QMessageBox.question(
                self,
                "确认退出",
                "数据抓取正在进行中，确定要退出吗？",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )

            if reply == QMessageBox.StandardButton.Yes:
                self.worker.stop()
                self.worker.wait()
                event.accept()
            else:
                event.ignore()
        else:
            event.accept()
