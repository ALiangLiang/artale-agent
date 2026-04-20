import shutil
import subprocess
import sys
from pathlib import Path

from artale_agent.utils import _project_root


def clean() -> None:
    root = Path(_project_root())

    # Clean dist & build
    for folder in ["dist", "build"]:
        p = root / folder
        if p.exists():
            shutil.rmtree(p)

    # Clean log & spec
    build_log = root.joinpath("artale_agent.log")
    if build_log.exists():
        build_log.unlink()
    build_spec = root.joinpath("ArtaleAgent.spec")
    if build_spec.exists():
        build_spec.unlink()


def build_win():
    # 0. Check sys.platform
    if sys.platform != "win32":
        print("Not Windows environment.")
        exit(1)

    root = Path(_project_root())

    # 1. Kill process（Windows only）
    subprocess.run(
        ["taskkill", "/F", "/IM", "ArtaleAgent.exe"],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )

    # 2. Clean
    clean()

    # 3. PyInstaller command
    cmd = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--onefile",
        "--name",
        "ArtaleAgent",
        "--icon",
        "assets/app_icon.ico",
        "--add-data",
        "assets;assets",
        "--add-data",
        "vendor/Tesseract-OCR;Tesseract-OCR",
        "--add-data",
        "VERSION;.",
        "--paths",
        "src",
        "--hidden-import",
        "psutil",
        "--hidden-import",
        "pynput.keyboard._win32",
        "--hidden-import",
        "win32process",
        "--hidden-import",
        "win32file",
        "--hidden-import",
        "PyQt6.QtWebSockets",
        "--hidden-import",
        "sip",
        "--collect-all",
        "PyQt6",
        "--clean",
        "--noconsole",
        "--noupx",
        "src/artale_agent/main.py",
    ]

    result = subprocess.run(cmd, cwd=root)

    exit(result.returncode)


def build_mac():
    if sys.platform != "darwin":
        print("Not macOS environment.")
        exit(1)

    root = Path(_project_root())

    # 1. Kill process (macOS)
    subprocess.run(
        ["pkill", "-f", "ArtaleAgent"],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )

    # 2. Clean
    clean()

    # 3. PyInstaller command
    cmd = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--onefile",
        "--name",
        "ArtaleAgent",
        "--add-data",
        "assets:assets",
        "--add-data",
        "VERSION:.",
        "--paths",
        "src",
        "--hidden-import",
        "psutil",
        "--hidden-import",
        "pynput.keyboard._darwin",
        "--hidden-import",
        "PyQt6.QtWebSockets",
        "--hidden-import",
        "sip",
        "--collect-all",
        "PyQt6",
        "--clean",
        "--noconsole",
        "--noupx",
        "src/artale_agent/main.py",
    ]

    result = subprocess.run(cmd, cwd=root)
    exit(result.returncode)
