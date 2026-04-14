"""Windows platform implementation of the platform abstraction layer."""

import ctypes
import ctypes.wintypes
import logging
import os
from collections.abc import Callable

import numpy as np
import psutil

try:
    import win32con
    import win32gui
    import win32process
except ImportError:
    win32gui = win32process = win32con = None

try:
    import winsound
except ImportError:
    winsound = None

try:
    from windows_capture import WindowsCapture
except ImportError:
    WindowsCapture = None

from .base import AudioPlayer, FocusTracker, ScreenCapture, WindowInfo, WindowManager

logger = logging.getLogger(__name__)

# Win32 event hook constants
EVENT_SYSTEM_FOREGROUND = 0x0003
WINEVENT_OUTOFCONTEXT = 0x0000
SW_SHOWMINIMIZED = 2
SW_SHOWMAXIMIZED = 3

# Callback type for SetWinEventHook
WinEventProcType = ctypes.WINFUNCTYPE(
    None,
    ctypes.wintypes.HANDLE,
    ctypes.wintypes.DWORD,
    ctypes.wintypes.HWND,
    ctypes.wintypes.LONG,
    ctypes.wintypes.LONG,
    ctypes.wintypes.DWORD,
    ctypes.wintypes.DWORD,
)


class WinWindowManager(WindowManager):
    """Windows implementation of WindowManager using win32gui."""

    def find_game_window(
        self, title_pattern: str, process_name: str
    ) -> WindowInfo | None:
        """Find game window by title substring, with psutil process fallback."""
        if not win32gui:
            return None

        my_pid = os.getpid()
        found_hwnds = []

        def _enum_callback(hwnd, _extra):
            try:
                if not win32gui.IsWindowVisible(hwnd):
                    return True
                _, pid = win32process.GetWindowThreadProcessId(hwnd)
                if pid == my_pid:
                    return True
                title = win32gui.GetWindowText(hwnd)
                if title_pattern.lower() in title.lower():
                    found_hwnds.append(hwnd)
            except Exception:
                pass
            return True

        try:
            win32gui.EnumWindows(_enum_callback, None)
        except Exception as e:
            err_str = str(e)
            if not any(code in err_str for code in ["(2,", "(1400,"]):
                logger.debug("[WinWindowManager] EnumWindows failed: %s", e)

        if not found_hwnds:
            # Fallback: search by process name via psutil
            try:
                for proc in psutil.process_iter(["pid", "name"]):
                    if (
                        proc.info["name"]
                        and proc.info["name"].lower() == process_name.lower()
                    ):
                        proc_hwnds = []

                        def _proc_callback(hwnd, extra):
                            if win32gui.IsWindowVisible(hwnd):
                                _, pid = win32process.GetWindowThreadProcessId(hwnd)
                                if pid == proc.info["pid"]:
                                    extra.append(hwnd)
                            return True

                        win32gui.EnumWindows(_proc_callback, proc_hwnds)
                        if proc_hwnds:
                            # Prefer window with the longest title (usually the game window)
                            found_hwnds.append(
                                max(
                                    proc_hwnds,
                                    key=lambda h: len(win32gui.GetWindowText(h)),
                                )
                            )
                            break
            except Exception as e:
                logger.debug("[WinWindowManager] Process fallback failed: %s", e)

        if not found_hwnds:
            return None

        hwnd = found_hwnds[0]
        try:
            title = win32gui.GetWindowText(hwnd)
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            x, y, width, height = self.get_client_rect(hwnd)
            return WindowInfo(
                window_id=hwnd,
                title=title,
                pid=pid,
                x=x,
                y=y,
                width=width,
                height=height,
            )
        except Exception as e:
            logger.debug("[WinWindowManager] Failed to build WindowInfo: %s", e)
            return None

    def get_client_rect(self, window_id: int) -> tuple[int, int, int, int]:
        """Return (x, y, width, height) of client area in screen coordinates."""
        if not win32gui:
            return (0, 0, 0, 0)
        try:
            rect = win32gui.GetClientRect(window_id)
            # rect is (left, top, right, bottom) in client coords
            width = rect[2] - rect[0]
            height = rect[3] - rect[1]
            # Convert top-left client origin to screen coords
            screen_x, screen_y = win32gui.ClientToScreen(window_id, (0, 0))
            return (screen_x, screen_y, width, height)
        except Exception as e:
            logger.debug("[WinWindowManager] get_client_rect failed: %s", e)
            return (0, 0, 0, 0)

    def is_minimized(self, window_id: int) -> bool:
        """Check if the window is minimized."""
        if not win32gui:
            return False
        try:
            placement = win32gui.GetWindowPlacement(window_id)
            return placement[1] == SW_SHOWMINIMIZED
        except Exception:
            return False

    def is_maximized(self, window_id: int) -> bool:
        """Check if the window is maximized."""
        if not win32gui:
            return False
        try:
            placement = win32gui.GetWindowPlacement(window_id)
            return placement[1] == SW_SHOWMAXIMIZED
        except Exception:
            return False

    def is_valid(self, window_id: int) -> bool:
        """Check if the window handle is still valid."""
        if not win32gui:
            return False
        try:
            return bool(win32gui.IsWindow(window_id))
        except Exception:
            return False

    def get_window_title(self, window_id: int) -> str:
        """Get the window title string."""
        if not win32gui:
            return ""
        try:
            return win32gui.GetWindowText(window_id)
        except Exception:
            return ""

    def client_to_screen(self, window_id: int, x: int, y: int) -> tuple[int, int]:
        """Convert client-area coordinates to screen coordinates."""
        if not win32gui:
            return (x, y)
        try:
            return win32gui.ClientToScreen(window_id, (x, y))
        except Exception as e:
            logger.debug("[WinWindowManager] client_to_screen failed: %s", e)
            return (x, y)


class WinScreenCapture(ScreenCapture):
    """Windows screen capture using the windows_capture library."""

    def __init__(self):
        self._active = False

    def start(
        self, window_info: WindowInfo, on_frame: Callable[[np.ndarray], None]
    ) -> None:
        """Start capturing frames from the given window."""
        if not WindowsCapture:
            logger.warning(
                "[WinScreenCapture] windows_capture not installed; capture disabled."
            )
            return

        self._active = True

        cap_config = {
            "window_name": window_info.title,
            "cursor_capture": False,
            "minimum_update_interval": 1000,
        }

        try:
            capture = WindowsCapture(draw_border=False, **cap_config)
        except Exception as e:
            if "Toggling the capture border is not supported" in str(e):
                capture = WindowsCapture(draw_border=True, **cap_config)
            else:
                logger.error("[WinScreenCapture] Failed to create capture: %s", e)
                self._active = False
                return

        @capture.event
        def on_frame_arrived(frame, _control):
            if not self._active:
                return
            try:
                bgr = np.asarray(frame.frame_buffer)
                bgr = bgr[..., :3]  # Drop alpha channel (BGRA -> BGR)
                on_frame(bgr)
            except Exception as exc:
                logger.debug("[WinScreenCapture] Frame conversion error: %s", exc)

        @capture.event
        def on_closed():
            logger.info("[WinScreenCapture] Capture session closed.")
            self._active = False

        try:
            capture.start_free_threaded()
        except Exception as e:
            logger.error("[WinScreenCapture] start_free_threaded failed: %s", e)
            self._active = False

    def stop(self) -> None:
        """Stop the capture session."""
        self._active = False

    def is_active(self) -> bool:
        """Whether the capture session is currently running."""
        return self._active


class WinFocusTracker(FocusTracker):
    """Windows focus tracker using SetWinEventHook."""

    TARGET_PROCESS_DEFAULT = "msw.exe"

    def __init__(self):
        self._is_game_active = False
        self._target_process: str | None = None
        self._callback = None
        self._hook = None

    def start(self, target_process_name: str) -> None:
        """Start tracking focus for the given process name."""
        self._target_process = target_process_name.lower()

        user32 = ctypes.windll.user32

        self._callback = WinEventProcType(self._on_focus_change)
        self._hook = user32.SetWinEventHook(
            EVENT_SYSTEM_FOREGROUND,
            EVENT_SYSTEM_FOREGROUND,
            0,
            self._callback,
            0,
            0,
            WINEVENT_OUTOFCONTEXT,
        )

        # Check initial foreground state
        self._check_foreground()

    def stop(self) -> None:
        """Stop tracking (hook is cleaned up on process exit)."""
        # No explicit unhook needed; the OS cleans up on process exit.
        # If needed in future, call user32.UnhookWinEvent(self._hook) here.
        pass

    @property
    def is_game_active(self) -> bool:
        """Whether the game is currently the foreground window."""
        return self._is_game_active

    def _on_focus_change(
        self,
        hWinEventHook,
        event,
        hwnd,
        idObject,
        idChild,
        dwEventThread,
        dwmsEventTime,
    ) -> None:
        self._check_foreground(hwnd)

    def _check_foreground(self, hwnd=None) -> None:
        """Update is_game_active based on the foreground window's process name."""
        try:
            user32 = ctypes.windll.user32

            if hwnd is None:
                hwnd = user32.GetForegroundWindow()
            if not hwnd or hwnd <= 0:
                self._is_game_active = False
                return

            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            # Treat as 32-bit unsigned integer
            pid = pid & 0xFFFFFFFF

            # Sanity check: PIDs are almost never > 1,000,000 in reality
            if pid <= 0 or pid > 2**31:
                self._is_game_active = False
                return

            try:
                proc = psutil.Process(pid)
                p_name = proc.name().lower()
            except (psutil.NoSuchProcess, psutil.AccessDenied, ValueError):
                self._is_game_active = False
                return

            target = self._target_process or self.TARGET_PROCESS_DEFAULT
            was_active = self._is_game_active
            self._is_game_active = p_name == target

            if self._is_game_active != was_active:
                status = "focused" if self._is_game_active else "lost focus"
                logger.info("[WinFocusTracker] %s %s", target, status)

        except Exception:
            # Silent fallback to avoid console spam
            self._is_game_active = False


class WinAudioPlayer(AudioPlayer):
    """Windows audio player using winsound.Beep."""

    def beep(self, frequency: int, duration_ms: int) -> None:
        """Play a beep tone at the given frequency for duration_ms milliseconds."""
        if not winsound:
            logger.debug("[WinAudioPlayer] winsound not available; beep skipped.")
            return
        try:
            winsound.Beep(frequency, duration_ms)
        except Exception as e:
            logger.debug("[WinAudioPlayer] Beep failed: %s", e)
