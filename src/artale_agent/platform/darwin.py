"""macOS platform layer using pyobjc (Quartz + AppKit)."""

import os
import subprocess
import threading
from collections.abc import Callable
from typing import override

import cv2
import numpy as np

from artale_agent.platform.base import AudioPlayer, FocusTracker, ScreenCapture, WindowInfo, WindowManager

try:
    import Quartz
    from Quartz import (
        CGRectNull,
        CGWindowListCopyWindowInfo,
        CGWindowListCreateImage,
        kCGNullWindowID,
        kCGWindowListOptionIncludingWindow,
        kCGWindowListOptionOnScreenOnly,
    )
except ImportError:
    Quartz = None

try:
    from AppKit import NSWorkspace
except ImportError:
    NSWorkspace = None


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _get_window_info_list(option: int, window_id: int = 0) -> list:
    """Return the raw CGWindowInfo list for the given option/window_id."""
    if Quartz is None:
        return []
    wid = window_id if window_id else kCGNullWindowID
    result = CGWindowListCopyWindowInfo(option, wid)
    return list(result) if result else []


def _find_window_dict(window_id: int) -> dict | None:
    """Return the CGWindowInfo dict for a single on-screen window by ID."""
    windows = _get_window_info_list(kCGWindowListOptionIncludingWindow, window_id)
    return windows[0] if windows else None


def _bounds_to_xywh(bounds: dict) -> tuple[int, int, int, int]:
    """Convert a CGWindowBounds dict to (x, y, width, height) integers."""
    return (
        int(bounds.get("X", 0)),
        int(bounds.get("Y", 0)),
        int(bounds.get("Width", 0)),
        int(bounds.get("Height", 0)),
    )


# ---------------------------------------------------------------------------
# WindowManager
# ---------------------------------------------------------------------------


class MacWindowManager(WindowManager):
    """macOS implementation of WindowManager using CGWindowList APIs."""

    @override
    def find_game_window(
        self, title_pattern: str, process_name: str
    ) -> WindowInfo | None:
        own_pid = os.getpid()
        windows = _get_window_info_list(kCGWindowListOptionOnScreenOnly)

        for win in windows:
            pid = int(win.get("kCGWindowOwnerPID", -1))
            if pid == own_pid:
                continue

            name: str = win.get("kCGWindowName") or ""
            owner: str = win.get("kCGWindowOwnerName") or ""

            title_match = title_pattern and title_pattern.lower() in name.lower()
            process_match = process_name and process_name.lower() in owner.lower()

            if not (title_match or process_match):
                continue

            window_id = int(win.get("kCGWindowNumber", 0))
            bounds = win.get("kCGWindowBounds", {})
            x, y, w, h = _bounds_to_xywh(bounds)

            return WindowInfo(
                window_id=window_id,
                title=name,
                pid=pid,
                x=x,
                y=y,
                width=w,
                height=h,
            )

        return None

    @override
    def get_client_rect(self, window_id: int) -> tuple[int, int, int, int]:
        win = _find_window_dict(window_id)
        if win is None:
            return (0, 0, 0, 0)
        bounds = win.get("kCGWindowBounds", {})
        return _bounds_to_xywh(bounds)

    @override
    def is_minimized(self, window_id: int) -> bool:
        win = _find_window_dict(window_id)
        if win is None:
            return False
        bounds = win.get("kCGWindowBounds", {})
        w = int(bounds.get("Width", 0))
        h = int(bounds.get("Height", 0))
        return w == 0 or h == 0

    @override
    def is_maximized(self, window_id: int) -> bool:
        # macOS does not have the same maximize concept as Windows.
        return False

    @override
    def is_valid(self, window_id: int) -> bool:
        return _find_window_dict(window_id) is not None

    @override
    def get_window_title(self, window_id: int) -> str:
        win = _find_window_dict(window_id)
        if win is None:
            return ""
        return win.get("kCGWindowName") or ""

    @override
    def client_to_screen(self, window_id: int, x: int, y: int) -> tuple[int, int]:
        # kCGWindowBounds is already in global screen coordinates on macOS.
        win = _find_window_dict(window_id)
        if win is None:
            return (x, y)
        bounds = win.get("kCGWindowBounds", {})
        bx = int(bounds.get("X", 0))
        by = int(bounds.get("Y", 0))
        return (bx + x, by + y)


# ---------------------------------------------------------------------------
# ScreenCapture
# ---------------------------------------------------------------------------


class MacScreenCapture(ScreenCapture):
    """macOS screen capture using CGWindowListCreateImage at ~1 FPS."""

    _TARGET_FPS = 1.0

    def __init__(self) -> None:
        self._stop_event = threading.Event()
        self._thread: threading.Thread | None = None
        self._active = False

    @override
    def start(
        self, window_info: WindowInfo, on_frame: Callable[[np.ndarray], None]
    ) -> None:
        if self._active:
            return

        self._stop_event.clear()
        self._active = True
        window_id = window_info.window_id

        def _capture_loop() -> None:
            interval = 1.0 / self._TARGET_FPS
            try:
                while not self._stop_event.wait(interval):
                    frame = self._capture_frame(window_id)
                    if frame is not None:
                        on_frame(frame)
            finally:
                self._active = False

        self._thread = threading.Thread(target=_capture_loop, daemon=True)
        self._thread.start()

    @override
    def stop(self) -> None:
        self._stop_event.set()
        if self._thread is not None:
            self._thread.join(timeout=5.0)
            self._thread = None

    @override
    def is_active(self) -> bool:
        return self._active

    # ------------------------------------------------------------------
    # Internal helpers
    # ------------------------------------------------------------------

    @staticmethod
    def _capture_frame(window_id: int) -> np.ndarray | None:
        if Quartz is None:
            return None

        try:
            cg_image = CGWindowListCreateImage(
                CGRectNull,
                kCGWindowListOptionIncludingWindow,
                window_id,
                0,
            )
            if cg_image is None:
                return None

            width = Quartz.CGImageGetWidth(cg_image)
            height = Quartz.CGImageGetHeight(cg_image)
            bytes_per_row = Quartz.CGImageGetBytesPerRow(cg_image)

            if width == 0 or height == 0:
                return None

            data_provider = Quartz.CGImageGetDataProvider(cg_image)
            raw_data = Quartz.CGDataProviderCopyData(data_provider)
            if raw_data is None:
                return None

            buf = bytes(raw_data)
            arr = np.frombuffer(buf, dtype=np.uint8)
            arr = arr.reshape((height, bytes_per_row // 4, 4))
            arr = arr[:, :width, :]  # trim padding columns

            # CGImage pixel format is BGRA on macOS; drop alpha for BGR.
            bgr = cv2.cvtColor(arr, cv2.COLOR_BGRA2BGR)
            return bgr

        except Exception:
            return None


# ---------------------------------------------------------------------------
# FocusTracker
# ---------------------------------------------------------------------------


class MacFocusTracker(FocusTracker):
    """Tracks whether the target game process is the frontmost application."""

    def __init__(self) -> None:
        self._game_active = False
        self._target: str = ""
        self._observer: object | None = None

    @override
    def start(self, target_process_name: str) -> None:
        self._target = target_process_name

        if NSWorkspace is None:
            return

        workspace = NSWorkspace.sharedWorkspace()

        # Check current frontmost app immediately.
        frontmost = workspace.frontmostApplication()
        if frontmost is not None:
            self._game_active = self._matches(frontmost.localizedName())

        # Register for app-activation notifications.
        notification_center = workspace.notificationCenter()
        notification_center.addObserver_selector_name_object_(
            self,
            "_onAppActivated:",
            "NSWorkspaceDidActivateApplicationNotification",
            None,
        )
        self._observer = notification_center

    @override
    def stop(self) -> None:
        if self._observer is not None:
            try:
                self._observer.removeObserver_(self)
            except Exception:
                pass
            self._observer = None

    @override
    @property
    def is_game_active(self) -> bool:
        return self._game_active

    # ------------------------------------------------------------------
    # Notification callback (called by Cocoa runtime)
    # ------------------------------------------------------------------

    def _onAppActivated_(self, notification) -> None:  # noqa: N802 – Cocoa naming
        user_info = notification.userInfo()
        if user_info is None:
            return
        app = user_info.get("NSWorkspaceApplicationKey")
        if app is not None:
            self._game_active = self._matches(app.localizedName())

    def _matches(self, app_name: str | None) -> bool:
        if not app_name or not self._target:
            return False
        return self._target.lower() in app_name.lower()


# ---------------------------------------------------------------------------
# AudioPlayer
# ---------------------------------------------------------------------------


class MacAudioPlayer(AudioPlayer):
    """macOS audio player using the system afplay utility."""

    _SOUND_FILE = "/System/Library/Sounds/Tink.aiff"

    @override
    def beep(self, frequency: int, duration_ms: int) -> None:
        # frequency and duration_ms are ignored; afplay plays the fixed sound.
        try:
            subprocess.Popen(
                ["afplay", self._SOUND_FILE],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
            )
        except FileNotFoundError:
            pass
