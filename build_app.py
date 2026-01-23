#!/usr/bin/env python3
"""
Build script for creating an all-in-one macOS app bundle.
This script creates a standalone .app that includes the Playwright browser.
"""

import os
import shutil
import subprocess
from pathlib import Path


def run_command(cmd, description=""):
    """Run a command and print output."""
    print(f"\n{'='*60}")
    print(f"Running: {description or cmd}")
    print(f"{'='*60}")
    result = subprocess.run(cmd, shell=True, capture_output=False, text=True)
    if result.returncode != 0:
        print(f"ERROR: Command failed with code {result.returncode}")
        return False
    return True


def main():
    """Build the all-in-one app bundle."""
    print("="*60)
    print("Building 足球彩票分析工具 All-In-One Bundle")
    print("="*60)

    # Project root
    root = Path(__file__).parent
    os.chdir(root)

    # Step 1: Clean previous build
    print("\n[1/4] Cleaning previous build...")
    dist_dir = root / "dist"
    build_dir = root / "build"
    if dist_dir.exists():
        shutil.rmtree(dist_dir)
    if build_dir.exists():
        shutil.rmtree(build_dir)

    # Step 2: Check Playwright browser
    print("\n[2/4] Checking Playwright browser installation...")
    playwright_cache = Path.home() / 'Library' / 'Caches' / 'ms-playwright'
    chromium_path = None

    if playwright_cache.exists():
        for item in playwright_cache.iterdir():
            if item.is_dir() and item.name.startswith('chromium-') and 'headless' not in item.name:
                chromium_path = item
                break

    if chromium_path and chromium_path.exists():
        browser_src = chromium_path / 'chrome-mac-arm64'
        size_mb = sum(f.stat().st_size for f in browser_src.rglob('*') if f.is_file()) / 1024 / 1024
        print(f"  Found browser at: {browser_src}")
        print(f"  Browser size: {size_mb:.1f} MB")
    else:
        print("  WARNING: Playwright browser not found!")
        print("  Run: uv run playwright install chromium")
        return 1

    # Step 3: Build with PyInstaller
    print("\n[3/4] Building with PyInstaller...")
    if not run_command("uv run pyinstaller 足球彩票分析工具.spec --clean", "PyInstaller Build"):
        print("ERROR: PyInstaller build failed!")
        return 1

    # Step 4: Verify the bundle
    print("\n[4/4] Verifying bundle...")
    app_path = root / "dist" / "足球彩票分析工具.app"

    if not app_path.exists():
        print("ERROR: .app bundle was not created!")
        return 1

    # Calculate bundle size
    bundle_size = sum(f.stat().st_size for f in app_path.rglob('*') if f.is_file()) / 1024 / 1024
    print(f"  Bundle created: {app_path}")
    print(f"  Bundle size: {bundle_size:.1f} MB")

    # Check for bundled browser
    bundled_browser = app_path / "Contents" / "Resources" / "ms-playwright"
    if bundled_browser.exists():
        print(f"  Browser bundled: YES")
    else:
        print(f"  WARNING: Browser not found in bundle!")

    print("\n" + "="*60)
    print("Build Complete!")
    print("="*60)
    print(f"\nApp bundle ready at:")
    print(f"  {app_path}")
    print(f"\nTo distribute:")
    print(f"  1. Right-click the .app and select 'Compress' to create a ZIP")
    print(f"  2. Send the ZIP file to the user")
    print(f"  3. User extracts and double-clicks the .app to run")
    print(f"\nNote: First launch may take a few seconds to unpack.")

    return 0


if __name__ == "__main__":
    exit(main())
