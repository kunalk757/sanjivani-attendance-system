#!/usr/bin/env python3
"""
Sanjivani Attendance System - Build & Packaging Script
Compiles the standalone Windows Executable using PyInstaller with all assets bundled.
"""

import os
import sys
import time
import subprocess
import shutil

BASE_DIR = os.path.abspath(os.path.dirname(__file__))

def print_banner():
    print("=" * 70)
    print("  SANJIVANI ATTENDANCE SYSTEM - STANDALONE BUILD SCRIPT")
    print("  Packaging Windows Executable (Sanjivani-Attendance-System.exe)")
    print("=" * 70)

def check_dependencies():
    print("\n[1/5] Checking Python dependencies...")
    required_packages = ["cv2", "numpy", "openpyxl", "PIL", "PyInstaller"]
    missing = []
    for pkg in required_packages:
        try:
            if pkg == "cv2":
                import cv2
                # Check LBPH face recognizer availability
                if not hasattr(cv2, "face") or not hasattr(cv2.face, "LBPHFaceRecognizer_create"):
                    print("  [!] WARNING: cv2.face is missing! Ensure opencv-contrib-python is installed.")
            elif pkg == "PIL":
                from PIL import Image
            else:
                __import__(pkg)
            print(f"  [OK] {pkg} found.")
        except ImportError:
            missing.append(pkg)
            print(f"  [X] {pkg} MISSING!")

    if missing:
        print(f"\n[ERROR] Missing required packages: {', '.join(missing)}")
        print("Please install them using: pip install -r requirements.txt")
        sys.exit(1)

def ensure_icon():
    print("\n[2/6] Generating and verifying multi-resolution application icon (sanjivani.ico)...")
    from create_icon import create_app_icon
    create_app_icon()
    sanjivani_ico = os.path.join(BASE_DIR, "sanjivani.ico")
    if os.path.exists(sanjivani_ico):
        print(f"  [OK] sanjivani.ico generated successfully ({os.path.getsize(sanjivani_ico)} bytes).")
    else:
        print(f"[ERROR] Failed to generate {sanjivani_ico}")
        sys.exit(1)

def clean_previous_builds():
    print("\n[3/6] Cleaning previous build artifacts...")
    build_dir = os.path.join(BASE_DIR, "build")
    if os.path.exists(build_dir):
        try:
            if sys.platform == "win32":
                subprocess.run(
                    ["powershell", "-NoProfile", "-Command", f"Remove-Item -Recurse -Force '{build_dir}' -ErrorAction SilentlyContinue"],
                    cwd=BASE_DIR,
                    check=False
                )
            if os.path.exists(build_dir):
                shutil.rmtree(build_dir, ignore_errors=True)
            print("  [OK] Cleaned build directory.")
        except Exception as e:
            print(f"  [!] Note: {e}")

def run_pyinstaller():
    print("\n[4/6] Running PyInstaller build with app.spec...")
    spec_path = os.path.join(BASE_DIR, "app.spec")
    if not os.path.exists(spec_path):
        print(f"[ERROR] Spec file not found at: {spec_path}")
        sys.exit(1)

    cmd = [sys.executable, "-m", "PyInstaller", "--noconfirm", "--clean", spec_path]
    print(f"  Executing: {' '.join(cmd)}")
    result = subprocess.run(cmd, cwd=BASE_DIR)
    if result.returncode != 0:
        print(f"\n[ERROR] PyInstaller failed with exit code: {result.returncode}")
        sys.exit(result.returncode)
    print("  [OK] PyInstaller compilation completed successfully.")

def find_iscc():
    """Locate Inno Setup compiler executable (ISCC.exe)."""
    candidates = [
        os.path.expandvars(r"%LOCALAPPDATA%\Programs\Inno Setup 6\ISCC.exe"),
        os.path.expandvars(r"%LOCALAPPDATA%\Programs\Inno Setup 7\ISCC.exe"),
        r"C:\Program Files (x86)\Inno Setup 6\ISCC.exe",
        r"C:\Program Files\Inno Setup 6\ISCC.exe",
        r"C:\Program Files (x86)\Inno Setup 5\ISCC.exe",
        r"C:\Program Files\Inno Setup 5\ISCC.exe",
    ]
    for path in candidates:
        if os.path.exists(path):
            return path
    # Check PATH
    iscc_in_path = shutil.which("ISCC.exe") or shutil.which("iscc")
    if iscc_in_path:
        return iscc_in_path
    return None

def build_installer():
    print("\n[5/6] Building Inno Setup Installer package...")
    iss_path = os.path.join(BASE_DIR, "installer.iss")
    if not os.path.exists(iss_path):
        print(f"  [!] Skipping installer: {iss_path} not found.")
        return

    iscc_path = find_iscc()
    if not iscc_path:
        print("  [!] Inno Setup compiler (ISCC.exe) not found in standard paths.")
        print("      To build the setup installer, ensure Inno Setup is installed.")
        return

    print(f"  Found ISCC compiler: {iscc_path}")
    cmd = [iscc_path, iss_path]
    result = subprocess.run(cmd, cwd=BASE_DIR)
    if result.returncode != 0:
        print(f"\n[ERROR] Inno Setup compilation failed with exit code: {result.returncode}")
    else:
        print("  [OK] Windows Installer compilation completed successfully.")

def verify_output():
    print("\n[6/6] Verifying output binaries...")
    exe_path = os.path.join(BASE_DIR, "dist", "Sanjivani Attendance.exe")
    installer_path = os.path.join(BASE_DIR, "dist", "Sanjivani-Attendance-Setup.exe")
    
    if os.path.exists(exe_path):
        size_mb = os.path.getsize(exe_path) / (1024 * 1024)
        print(f"  [OK] Standalone Executable : {exe_path} ({size_mb:.2f} MB)")
    else:
        print(f"  [ERROR] Standalone Executable not found at: {exe_path}")

    if os.path.exists(installer_path):
        size_mb = os.path.getsize(installer_path) / (1024 * 1024)
        print(f"  [OK] Setup Installer       : {installer_path} ({size_mb:.2f} MB)")
    else:
        print(f"  [NOTE] Setup Installer not generated or skipped.")

    print("\n" + "=" * 70)
    print("  BUILD COMPLETE: Ready for deployment with unified multi-resolution icon!")
    print("=" * 70 + "\n")

if __name__ == "__main__":
    print_banner()
    check_dependencies()
    ensure_icon()
    clean_previous_builds()
    run_pyinstaller()
    build_installer()
    verify_output()
