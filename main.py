import sys
import ctypes
import ctypes.wintypes
import win32process
import psutil
from overlay import ArtaleOverlay, SettingsWindow, ConfigManager
from PyQt6.QtWidgets import QApplication
from pynput import keyboard, mouse

# --- Game Focus Tracker using SetWinEventHook ---
EVENT_SYSTEM_FOREGROUND = 0x0003
WINEVENT_OUTOFCONTEXT = 0x0000

user32 = ctypes.windll.user32

# Callback type: void(HWINEVENTHOOK, DWORD, HWND, LONG, LONG, DWORD, DWORD)
WinEventProcType = ctypes.WINFUNCTYPE(
    None, ctypes.wintypes.HANDLE, ctypes.wintypes.DWORD,
    ctypes.wintypes.HWND, ctypes.wintypes.LONG,
    ctypes.wintypes.LONG, ctypes.wintypes.DWORD,
    ctypes.wintypes.DWORD
)

class GameFocusTracker:
    TARGET_PROCESS = "msw.exe"
    
    def __init__(self):
        self.is_game_active = False
        self._callback = WinEventProcType(self._on_focus_change)
        self._hook = user32.SetWinEventHook(
            EVENT_SYSTEM_FOREGROUND, EVENT_SYSTEM_FOREGROUND,
            0, self._callback, 0, 0, WINEVENT_OUTOFCONTEXT
        )
        # Check initial state
        self._check_foreground()
    
    def _on_focus_change(self, hWinEventHook, event, hwnd, idObject,
                         idChild, dwEventThread, dwmsEventTime):
        self._check_foreground(hwnd)
    
    def _check_foreground(self, hwnd=None):
        try:
            if hwnd is None:
                hwnd = user32.GetForegroundWindow()
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            proc = psutil.Process(pid)
            was_active = self.is_game_active
            self.is_game_active = (proc.name().lower() == self.TARGET_PROCESS)
            if self.is_game_active != was_active:
                status = "focused" if self.is_game_active else "lost focus"
                print(f"[Focus] {self.TARGET_PROCESS} {status}")
        except:
            self.is_game_active = False

def start_keyboard_listener(overlay, settings_window, focus_tracker):
    # --- Mouse Listener for Right Click Cancellation ---
    def on_click(x, y, button, pressed):
        if button == mouse.Button.right and pressed:
            if focus_tracker.is_game_active:
                overlay.check_right_click(x, y)

    mouse_l = mouse.Listener(on_click=on_click)
    mouse_l.start()

    # --- Keyboard Listener ---
    # Variable to hold current triggers
    current_config = ConfigManager.load_config()

    def update_local_config():
        nonlocal current_config
        current_config = ConfigManager.load_config()
        print(f"[Config] Updated triggers: {current_config['triggers']}")

    # Connect settings signal to update listener
    settings_window.config_updated.connect(update_local_config)

    def on_press(key):
        try:
            # 1. Check for Pause Break to show settings (always works)
            if key == keyboard.Key.pause:
                print("[Input] Pause Break pressed. Emitting show signal.")
                settings_window.request_show.emit()
                return

            # 2. Only trigger timers when msw.exe is focused
            if not focus_tracker.is_game_active:
                return

            # 3. Check for dynamic triggers from config
            k_name = ""
            if hasattr(key, 'name'): k_name = key.name
            elif hasattr(key, 'char'):
                k_name = key.char
                if k_name: k_name = k_name.lower()
            
            # Normalize: alt_l/alt_r -> alt, shift_r -> shift, ctrl_l -> ctrl
            if k_name:
                for base in ["alt", "shift", "ctrl"]:
                    if k_name.startswith(base):
                        k_name = base
                        break
            
            if k_name and k_name in current_config["triggers"]:
                trigger_data = current_config["triggers"][k_name]
                if isinstance(trigger_data, dict):
                    seconds = trigger_data.get("seconds", 10)
                    icon = trigger_data.get("icon", "")
                else:
                    seconds = trigger_data
                    icon = ""
                overlay.timer_request.emit(k_name, seconds, icon if icon else "")
                
        except Exception as e:
            print(f"[Error] Listener: {e}")

    listener = keyboard.Listener(on_press=on_press)
    listener.start()
    print("[Input] Listener active. Press 'Pause Break' for settings.")

def run_app():
    app = QApplication(sys.argv)
    app.setQuitOnLastWindowClosed(False)
    
    overlay = ArtaleOverlay()
    settings_window = SettingsWindow(overlay)
    focus_tracker = GameFocusTracker()
    
    start_keyboard_listener(overlay, settings_window, focus_tracker)
    
    print("[Main] Artale Helper initialized. Waiting for input...")
    sys.exit(app.exec())

if __name__ == "__main__":
    run_app()
