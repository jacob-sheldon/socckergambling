"""
Progress dialog for showing scraping progress.
"""

from PyQt6.QtWidgets import QDialog, QVBoxLayout, QLabel, QProgressBar, QPushButton
from PyQt6.QtCore import Qt


class ProgressDialog(QDialog):
    """
    Modal dialog showing progress during data scraping.
    """

    def __init__(self, parent=None):
        super().__init__(parent)

        self.setWindowTitle("正在抓取数据")
        self.setModal(True)
        self.setFixedWidth(400)

        self._init_ui()

    def _init_ui(self):
        """Initialize the UI components."""
        layout = QVBoxLayout()
        layout.setSpacing(15)
        layout.setContentsMargins(20, 20, 20, 20)

        # Status label
        self.status_label = QLabel("正在初始化...")
        self.status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.status_label)

        # Progress bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 0)  # Indeterminate progress
        self.progress_bar.setTextVisible(False)
        layout.addWidget(self.progress_bar)

        # Cancel button
        self.cancel_btn = QPushButton("取消")
        self.cancel_btn.clicked.connect(self.reject)
        layout.addWidget(self.cancel_btn)

        self.setLayout(layout)

    def set_status(self, message: str):
        """Update the status message."""
        self.status_label.setText(message)

    def set_progress(self, value: int, maximum: int):
        """Set progress bar to determinate mode with specific value."""
        self.progress_bar.setRange(0, maximum)
        self.progress_bar.setValue(value)
        self.progress_bar.setTextVisible(True)
