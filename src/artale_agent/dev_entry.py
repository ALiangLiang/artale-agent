"""Lightweight entry point for the ``dev`` console script.

The uv/pip trampoline launcher (``dev.exe``) alters the Windows DLL search
order in a way that prevents PyQt6 from loading its native libraries when
the project lives on a mapped network drive (UNC path).

This module intentionally avoids importing **any** Qt code.  Instead it
re-launches the application through ``python -m artale_agent``, which
starts a clean process where DLL loading works correctly.
"""

import os
import subprocess
import sys


def run_app():
    # Detect platform
    is_windows = sys.platform == "win32"
    
    # Resolve the root directory (3 levels up from dev_entry.py)
    script_dir = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    
    python = None

    if is_windows:
        # Check if we are currently in a directory with a drive letter (e.g., Z:\)
        # We prefer this path over any UNC path resolved by the uv shim.
        cwd = os.getcwd()
        if len(cwd) > 2 and cwd[1] == ":" and cwd[2] == "\\":
            drive_letter = cwd[:2]  # e.g., "Z:"
            
            # Find the .venv relative to this script's known location in the source tree
            # (src/artale_agent/dev_entry.py -> ../../../.venv)
            # We use the relative path parts to join with the drive letter path
            try:
                # Get the absolute path of the current file
                this_file = os.path.abspath(__file__)
                # If the file path is UNC (starts with \\), we try to find its relative
                # path to the project root and then prepend the drive letter.
                if this_file.startswith("\\\\"):
                    # Find where 'src' starts to identify the project structure
                    src_marker = os.path.join(os.sep, "src", "artale_agent")
                    if src_marker.lower() in this_file.lower():
                        root_part = this_file.lower().split(src_marker.lower())[0]
                        # Construct a relative path from the UNC root
                        rel_to_root = os.path.relpath(this_file, root_part)
                        # The project root on the drive-letter path is CWD's root-relative part
                        # But simpler: if CWD is the project root (common for uv run), use it.
                        if os.path.exists(os.path.join(cwd, "pyproject.toml")):
                            python = os.path.join(cwd, ".venv", "Scripts", "python.exe")
            except Exception:
                pass

        if not python or not os.path.isfile(python):
            # Fallback 1: Use script_dir directly (might be UNC, but we try)
            python = os.path.join(script_dir, ".venv", "Scripts", "python.exe")
        
        if not os.path.isfile(python):
            # Fallback 2: Check same dir as current executable
            exe_dir = os.path.dirname(os.path.abspath(sys.executable))
            python = os.path.join(exe_dir, "python.exe")
    else:
        # macOS / Linux
        # Try to find the python in the venv
        python = os.path.join(script_dir, ".venv", "bin", "python")
        if not os.path.isfile(python):
            python = sys.executable

    # Final fallback
    if not python or not os.path.isfile(python):
        python = sys.executable

    # Use subprocess.call to wait for completion and pass exit code
    # Relaunch as module: python -m artale_agent
    result = subprocess.call([python, "-m", "artale_agent"] + sys.argv[1:])
    sys.exit(result)
