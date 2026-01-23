"""
py2app setup configuration for building macOS .app bundle.

Usage:
    uv run python setup.py py2app
"""

from setuptools import setup
import sys
from pathlib import Path


# Get the playwright browsers path
venv_path = Path(sys.prefix)
playwright_browsers_path = venv_path / "bin" / "ms-playwright"

# Check if playwright browsers are installed
if playwright_browsers_path.exists():
    # Include all playwright browser data
    data_files = []
    for browser_dir in playwright_browsers_path.iterdir():
        if browser_dir.is_dir():
            # Collect all files in browser directory
            browser_files = []
            for item in browser_dir.rglob("*"):
                if item.is_file():
                    rel_path = item.relative_to(playwright_browsers_path)
                    browser_files.append(str(item))

            if browser_files:
                data_files.append(
                    (f"ms-playwright/{browser_dir.name}", browser_files)
                )
else:
    data_files = []
    print("Warning: Playwright browsers not found. Run: uv run playwright install chromium")

APP = ['gui/main.py']

DATA_FILES = data_files + [
    ('gui/styles', ['gui/styles/macos_style.qss']),
]

OPTIONS = {
    'argv_emulation': False,
    'iconfile': None,  # Add path to .icns file if available
    'plist': {
        'CFBundleName': '足球彩票分析工具',
        'CFBundleDisplayName': '足球彩票分析工具',
        'CFBundleGetInfoString': 'Soccer Betting Analysis Tool',
        'CFBundleIdentifier': 'com.soccergambling.app',
        'CFBundleVersion': '1.0.0',
        'CFBundleShortVersionString': '1.0.0',
        'NSHumanReadableCopyright': '© 2025',
        'NSHighResolutionCapable': True,
        'LSMinimumSystemVersion': '10.13.0',
        'NSRequiresAquaSystemAppearance': False,  # Support Dark Mode
    },
    'packages': [
        'openpyxl',
        'playwright',
        'PyQt6',
        'gui',
        'gui.widgets',
        'gui.workers',
    ],
    'includes': [
        'asyncio',
        'json',
        'dataclasses',
        'PyQt6.QtCore',
        'PyQt6.QtGui',
        'PyQt6.QtWidgets',
        'playwright.async_api',
    ],
    'excludes': [],
    'strip': True,
    'optimize': 1,
}

setup(
    name='SoccerGambling',
    app=APP,
    data_files=DATA_FILES,
    options={'py2app': OPTIONS},
)
