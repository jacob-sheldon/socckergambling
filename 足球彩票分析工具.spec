# -*- mode: python ; coding: utf-8 -*-
import os
import shutil
from pathlib import Path

# Toggle to bundle Playwright browsers into the app package.
BUNDLE_PLAYWRIGHT = False


chromium_path = None
headless_shell_path = None

if BUNDLE_PLAYWRIGHT:
    # Get Playwright browser paths (macOS default cache)
    playwright_cache = Path.home() / 'Library' / 'Caches' / 'ms-playwright'

    # Find browser directories
    if playwright_cache.exists():
        for item in playwright_cache.iterdir():
            if item.is_dir():
                if item.name.startswith('chromium-') and 'headless' not in item.name:
                    chromium_path = item
                elif item.name.startswith('chromium_headless_shell-'):
                    headless_shell_path = item

# Build datas list
datas_list = [
    ('gui/styles/macos_style.qss', 'gui/styles'),
]

# Store browser paths for post-build (we'll copy them after PyInstaller)
BROWSER_SRC_PATHS = []
if chromium_path and chromium_path.exists():
    chrome_path = chromium_path / 'chrome-mac-arm64'
    if chrome_path.exists():
        BROWSER_SRC_PATHS.append(('chromium', chromium_path.name, chrome_path))

if headless_shell_path and headless_shell_path.exists():
    shell_path = headless_shell_path / 'chrome-headless-shell-mac-arm64'
    if shell_path.exists():
        BROWSER_SRC_PATHS.append(('headless', headless_shell_path.name, shell_path))

if BROWSER_SRC_PATHS:
    total_size = 0
    for browser_type, name, path in BROWSER_SRC_PATHS:
        size_mb = sum(f.stat().st_size for f in path.rglob('*') if f.is_file()) / 1024 / 1024
        total_size += size_mb
        print(f"[PyInstaller] Will bundle Playwright {browser_type} from: {path} ({size_mb:.1f} MB)")
    print(f"[PyInstaller] Total browser size: {total_size:.1f} MB")
else:
    if BUNDLE_PLAYWRIGHT:
        print("[PyInstaller] WARNING: Playwright browsers not found. Bundle will require separate browser installation.")
    else:
        print("[PyInstaller] Playwright browsers not bundled. Users must install Chromium on first run.")


a = Analysis(
    ['gui/main.py'],
    pathex=[],
    binaries=[],
    datas=datas_list,
    hiddenimports=[
        'playwright',
        'playwright.async_api',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='足球彩票分析工具',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='足球彩票分析工具',
)
app = BUNDLE(
    coll,
    name='足球彩票分析工具.app',
    icon='resources/icons/app_icon.icns',
    bundle_identifier='com.soccergambling.app',
    info_plist={
        'CFBundleName': '足球彩票分析工具',
        'CFBundleDisplayName': '足球彩票分析工具',
        'CFBundleGetInfoString': 'Soccer Betting Analysis Tool',
        'CFBundleVersion': '1.0.0',
        'CFBundleShortVersionString': '1.0.0',
        'NSHumanReadableCopyright': '© 2025',
        'NSHighResolutionCapable': True,
        'LSMinimumSystemVersion': '10.13.0',
    },
)

# Post-build: Copy Playwright browsers into the app bundle
def copy_browser_to_bundle():
    """Copy the Playwright browsers into the built .app bundle."""
    if not BUNDLE_PLAYWRIGHT or not BROWSER_SRC_PATHS:
        print("[Post-Build] No browsers to bundle (not found or paths not set)")
        return

    # The app bundle path after PyInstaller build
    app_bundle = Path('dist/足球彩票分析工具.app')
    resources_dir = app_bundle / 'Contents' / 'Resources'

    if not app_bundle.exists():
        print(f"[Post-Build] App bundle not found at {app_bundle}")
        return

    print(f"[Post-Build] Copying browsers to bundle...")

    for browser_type, browser_dir_name, browser_src in BROWSER_SRC_PATHS:
        browser_dest = resources_dir / 'ms-playwright' / browser_dir_name
        dest_path = browser_dest / browser_src.name

        # Create destination directory
        dest_path.mkdir(parents=True, exist_ok=True)

        print(f"  Copying {browser_type} browser...")
        print(f"    From: {browser_src}")
        print(f"    To: {dest_path}")

        # Copy all files from browser to destination
        for item in browser_src.iterdir():
            dest_item = dest_path / item.name
            if item.is_dir():
                if dest_item.exists():
                    shutil.rmtree(dest_item)
                shutil.copytree(item, dest_item)
            else:
                shutil.copy2(item, dest_item)

        print(f"    ✓ {browser_type} browser copied successfully!")

    print(f"[Post-Build] All browsers copied successfully!")
    print(f"[Post-Build] Bundle size: {sum(f.stat().st_size for f in app_bundle.rglob('*') if f.is_file()) / 1024 / 1024:.1f} MB")

# Execute post-build step
copy_browser_to_bundle()
