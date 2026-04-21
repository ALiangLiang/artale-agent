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
    build_log = root.joinpath("artale_agent_build.log")
    if build_log.exists():
        build_log.unlink()
    build_spec = root.joinpath("ArtaleAgent.spec")
    if build_spec.exists():
        build_spec.unlink()


def _generate_version_info(version_str: str, root_path: Path):
    """
    生成 Windows 執行檔的版本資訊資源檔。
    """
    import re
    
    # 1. 解析基礎版本號 (Major, Minor, Patch)
    # 例如 v0.3.3-alpha.3 -> nums = ['0', '3', '3', '3']
    nums = re.findall(r'\d+', version_str)
    major = int(nums[0]) if len(nums) > 0 else 0
    minor = int(nums[1]) if len(nums) > 1 else 0
    patch = int(nums[2]) if len(nums) > 2 else 0
    
    # 2. 決定第四位數字 (v4)
    v4 = 0
    version_lower = version_str.lower()
    
    if "alpha" in version_lower:
        # Alpha 版：起始值為 0，若有後綴數字 (如 alpha.3) 則為 3
        v4 = int(nums[3]) if len(nums) > 3 else 0
    elif "beta" in version_lower:
        # Beta 版：起始值為 500，若有後綴數字 (如 beta.1) 則為 501
        v4 = 500 + (int(nums[3]) if len(nums) > 3 else 0)
    elif "-" not in version_str:
        # 正式版 (無任何橫槓後綴)：固定 999
        v4 = 999
    else:
        # 其他後綴情況 (例如 dev 或 rc)
        v4 = int(nums[3]) if len(nums) > 3 else 0

    ver_tuple = (major, minor, patch, v4)
    ver_dotted = f"{major}.{minor}.{patch}"

    content = f"""# UTF-8
VSVersionInfo(
  ffi=FixedFileInfo(
    filevers={ver_tuple},
    prodvers={ver_tuple},
    mask=0x3f,
    flags=0x0,
    OS=0x40004,
    fileType=0x1,
    subtype=0x0,
    date=(0, 0)
    ),
  kids=[
    StringFileInfo(
      [
      StringTable(
        u'040404B0',
        [StringStruct(u'CompanyName', u'Artale Agent Team'),
        StringStruct(u'FileDescription', u'Artale 瑞士刀 - 遊戲輔助工具'),
        StringStruct(u'FileVersion', u'{ver_dotted}'),
        StringStruct(u'InternalName', u'artale-agent'),
        StringStruct(u'LegalCopyright', u'Copyright (C) 2026'),
        StringStruct(u'OriginalFilename', u'ArtaleAgent.exe'),
        StringStruct(u'ProductName', u'Artale 瑞士刀'),
        StringStruct(u'ProductVersion', u'{version_str}')])
      ]), 
    VarFileInfo([VarStruct(u'Translation', [1028, 1200])])
  ]
)
"""
    info_path = root_path / "file_version_info.txt"
    info_path.write_text(content, encoding="utf-8")
    return info_path


def build_win():
    # 0. Check sys.platform
    if sys.platform != "win32":
        print("Not Windows environment.")
        exit(1)

    root = Path(_project_root())
    
    # 0.1 讀取版本號並生成資訊檔
    version_file = root / "VERSION"
    version_str = version_file.read_text(encoding="utf-8").strip() if version_file.exists() else "v0.0.0"
    version_info_path = _generate_version_info(version_str, root)

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
        f"--version-file={version_info_path}",
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
        "--clean",
        "--noconsole",
        "--noupx",
        "src/artale_agent/main.py",
    ]

    result = subprocess.run(cmd, cwd=root)
    
    # 清理暫存的版本資訊檔
    if version_info_path.exists():
        version_info_path.unlink()

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
