"""Auto-select platform implementation based on sys.platform."""

import sys

from .base import ScreenCapture, WindowManager, FocusTracker, AudioPlayer, WindowInfo

if sys.platform == "win32":
    from .windows import (
        WinScreenCapture as ScreenCaptureImpl,
        WinWindowManager as WindowManagerImpl,
        WinFocusTracker as FocusTrackerImpl,
        WinAudioPlayer as AudioPlayerImpl,
    )
elif sys.platform == "darwin":
    from .darwin import (
        MacScreenCapture as ScreenCaptureImpl,
        MacWindowManager as WindowManagerImpl,
        MacFocusTracker as FocusTrackerImpl,
        MacAudioPlayer as AudioPlayerImpl,
    )
else:
    raise RuntimeError(f"Unsupported platform: {sys.platform}")

__all__ = [
    "ScreenCapture", "WindowManager", "FocusTracker", "AudioPlayer",
    "WindowInfo",
    "ScreenCaptureImpl", "WindowManagerImpl", "FocusTrackerImpl", "AudioPlayerImpl",
]
