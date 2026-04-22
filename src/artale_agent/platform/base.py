"""Platform abstraction layer for cross-platform window management and screen capture."""

from abc import ABC, abstractmethod
from collections.abc import Callable
from dataclasses import dataclass

import numpy as np


@dataclass
class WindowInfo:
    """Platform-agnostic window information."""

    window_id: int  # HWND on Windows, CGWindowID on macOS
    title: str
    pid: int
    x: int  # Screen coordinates (top-left)
    y: int
    width: int  # Client area dimensions
    height: int


class ScreenCapture(ABC):
    """Captures frames from a specific game window."""

    @abstractmethod
    def start(
        self, window_info: WindowInfo, on_frame: Callable[[np.ndarray], None]
    ) -> None:
        """Start capturing frames. on_frame receives BGR numpy arrays."""

    @abstractmethod
    def stop(self) -> None:
        """Stop the capture session."""

    @abstractmethod
    def is_active(self) -> bool:
        """Whether the capture session is currently running."""


class WindowManager(ABC):
    """Finds and queries game windows."""

    @abstractmethod
    def find_game_window(
        self, title_pattern: str, process_name: str
    ) -> WindowInfo | None:
        """Find game window by title substring or process name."""

    @abstractmethod
    def get_client_rect(self, window_id: int) -> tuple[int, int, int, int]:
        """Return (x, y, width, height) of the client area in screen coords."""

    @abstractmethod
    def is_minimized(self, window_id: int) -> bool:
        """Check if the window is minimized."""

    @abstractmethod
    def is_maximized(self, window_id: int) -> bool:
        """Check if the window is maximized."""

    @abstractmethod
    def is_valid(self, window_id: int) -> bool:
        """Check if the window handle/ID is still valid."""

    @abstractmethod
    def get_window_title(self, window_id: int) -> str:
        """Get the window title string."""

    @abstractmethod
    def client_to_screen(self, window_id: int, x: int, y: int) -> tuple[int, int]:
        """Convert client-area coordinates to screen coordinates."""

    @abstractmethod
    def set_topmost(self, window_id: int, topmost: bool) -> None:
        """Set whether the window stays on top of others."""

    @abstractmethod
    def get_foreground_process_id(self) -> int:
        """Get the PID of the process that owns the foreground window."""


class FocusTracker(ABC):
    """Tracks whether the target game window is focused."""

    @abstractmethod
    def start(self, target_process_name: str) -> None:
        """Start tracking focus for the given process name."""

    @abstractmethod
    def stop(self) -> None:
        """Stop tracking."""

    @property
    @abstractmethod
    def is_game_active(self) -> bool:
        """Whether the game is currently the foreground window."""


class AudioPlayer(ABC):
    """Platform-specific audio playback."""

    @abstractmethod
    def beep(self, frequency: int, duration_ms: int) -> None:
        """Play a beep tone."""
