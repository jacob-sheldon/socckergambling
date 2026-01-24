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

    bundle_playwright = False

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

    # Step 2: Build with PyInstaller
    print("\n[2/4] Building with PyInstaller...")
    if not run_command("uv run pyinstaller 足球彩票分析工具.spec --clean", "PyInstaller Build"):
        print("ERROR: PyInstaller build failed!")
        return 1

    # Step 3: Verify the bundle
    print("\n[3/4] Verifying bundle...")
    app_path = root / "dist" / "足球彩票分析工具.app"

    if not app_path.exists():
        print("ERROR: .app bundle was not created!")
        return 1

    # Calculate bundle size
    bundle_size = sum(f.stat().st_size for f in app_path.rglob('*') if f.is_file()) / 1024 / 1024
    print(f"  Bundle created: {app_path}")
    print(f"  Bundle size: {bundle_size:.1f} MB")

    # Check for bundled browser (optional)
    bundled_browser = app_path / "Contents" / "Resources" / "ms-playwright"
    if bundled_browser.exists():
        print(f"  Browser bundled: YES")
    else:
        print(f"  Browser bundled: NO")

    print("\n" + "="*60)
    print("Build Complete!")
    print("="*60)
    print(f"\nApp bundle ready at:")
    print(f"  {app_path}")
    print(f"\nTo distribute:")
    print(f"  1. Right-click the .app and select 'Compress' to create a ZIP")
    print(f"  2. Send the ZIP file to the user")
    print(f"  3. User extracts and double-clicks the .app to run")
    if not bundled_browser.exists():
        print(f"\nNote: Users must install Playwright Chromium on first run.")

    return 0


if __name__ == "__main__":
    exit(main())
