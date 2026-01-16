"""
Main entry point for the GUI application.
"""

import sys
from PyQt6.QtWidgets import QApplication
from PyQt6.QtCore import Qt
from gui.main_window import MainWindow


def main():
    """Launch the GUI application."""
    # Create application
    app = QApplication(sys.argv)
    app.setApplicationName("足球彩票分析工具")
    app.setApplicationDisplayName("足球彩票分析工具")
    app.setOrganizationName("SoccerGambling")

    # Load macOS stylesheet if available
    try:
        from pathlib import Path
        style_path = Path(__file__).parent / "styles" / "macos_style.qss"
        if style_path.exists():
            with open(style_path, 'r', encoding='utf-8') as f:
                app.setStyleSheet(f.read())
    except Exception:
        pass  # Continue without stylesheet

    # Create and show main window
    window = MainWindow()
    window.show()

    # Run application
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
