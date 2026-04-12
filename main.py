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
        except Exception as e:
            print(f"[Input Error] {e}")
            self.is_game_active = False

import time

def start_keyboard_listener(overlay, settings_window, focus_tracker):
    # --- Mouse Listener for Right Click Cancellation ---
    def on_click(x, y, button, pressed):
        if button == mouse.Button.right and pressed:
            if focus_tracker.is_game_active:
                overlay.check_right_click(x, y)

    mouse_l = mouse.Listener(on_click=on_click)
    mouse_l.start()

    # --- Keyboard Listener ---
    current_config = ConfigManager.load_config()

    # State for double-press detection (Profile switching F1-F9)
    last_key = None
    last_time = 0
    DOUBLE_PRESS_DELAY = 0.35 # seconds
    is_globally_enabled = True

    def update_local_config():
        nonlocal current_config, is_globally_enabled
        current_config = ConfigManager.load_config()
        active = current_config.get("active_profile", "F1")
        p_data = current_config["profiles"].get(active, {"triggers": {}})
        triggers = p_data.get("triggers", {})
        is_globally_enabled = True # Re-enable on profile switch
        print(f"[Config] Switched to {active}. Triggers: {list(triggers.keys())}")

    # Connect settings signal to update listener
    settings_window.config_updated.connect(update_local_config)

    def on_press(key):
        nonlocal last_key, last_time, current_config, is_globally_enabled
        try:
            k_name = None
            # 1. Try to get name or char
            if hasattr(key, 'name'):
                k_name = key.name
            elif hasattr(key, 'char') and key.char:
                k_name = key.char.lower()
            
            # 2. Hard fallback for VK codes (Numpad <96>-<105>)
            if k_name is None:
                vk = getattr(key, 'vk', None)
                if vk is None:
                    # Fallback for some pynput versions/OS where vk is hidden in string
                    s_key = str(key)
                    if s_key.startswith('<') and s_key.endswith('>'):
                        try: vk = int(s_key[1:-1])
                        except: pass
                
                if vk is not None:
                    if 96 <= vk <= 105:
                        k_name = str(vk - 96)
                    elif vk == 110: # Numpad dot
                        k_name = "."
            
            if k_name:
                for base in ["alt", "shift", "ctrl"]:
                    if k_name.startswith(base):
                        k_name = base
                        break
            
            # 3. Profile Switching (Double Press F1-F9) or Disable (F12)
            now = time.time()
            
            # 2. Always-on controls
            if k_name == 'pause':
                print("[Input] Pause Break pressed. Emitting show signal.")
                settings_window.request_show.emit()
                return
            
            # --- BLOCK TRIGGERS IF RECORDING ---
            if settings_window.is_recording or settings_window.recording_global_key:
                return

            # Use configurable hotkeys
            hks = current_config.get("hotkeys", {})
            
            if k_name == hks.get("reset", "f12"):
                print(f"[Input] {k_name.upper()} Reset Triggered.")
                is_globally_enabled = False
                overlay.clear_request.emit()
                return

            if k_name == hks.get("exp_toggle", "f10"):
                overlay.toggle_exp_request.emit()
                return

            if k_name == hks.get("exp_pause", "f11"):
                overlay.toggle_pause_request.emit()
                return

            # --- RJPQ SMART HOTKEYS ---
            if k_name in ["1", "2", "3", "4"]:
                # Check if RJPQ is active and connected
                if hasattr(settings_window, 'rjpq_tab') and settings_window.rjpq_tab.client.is_connected:
                    if settings_window.rjpq_tab.mark_by_hotkey(int(k_name) - 1):
                        return # Consume the key if it was used for RJPQ

            # 3. Profile Switching (Double Press F1-F9)
            now = time.time()
            if k_name and k_name.startswith('f') and len(k_name) <= 3:
                f_num = k_name[1:]
                if f_num.isdigit() and 1 <= int(f_num) <= 9:
                    if last_key == k_name and (now - last_time) < DOUBLE_PRESS_DELAY:
                        # Success! Switch profile
                        p_key = f"F{f_num}"
                        config = ConfigManager.load_config()
                        config["active_profile"] = p_key
                        ConfigManager.save_config(config)
                        update_local_config()
                        overlay.profile_switch_request.emit()
                        last_key = None # Reset
                        return
                    last_key = k_name
                    last_time = now

            # 4. Only trigger timers when enabled AND msw.exe is focused
            if not is_globally_enabled or not focus_tracker.is_game_active:
                return

            # Access current profile triggers
            active_p = current_config.get("active_profile", "F1")
            prof_data = current_config["profiles"].get(active_p, {"triggers": {}})
            triggers = prof_data.get("triggers", {})
            
            if k_name and k_name in triggers:
                trigger_data = triggers[k_name]
                if isinstance(trigger_data, dict):
                    seconds = trigger_data.get("seconds", 10)
                    icon = trigger_data.get("icon", "")
                    sound = trigger_data.get("sound", True)
                else:
                    seconds = trigger_data
                    icon = ""
                    sound = True
                overlay.timer_request.emit(k_name, seconds, icon if icon else "", sound)
                
        except Exception as e:
            print(f"[Error] Listener: {e}")

    listener = keyboard.Listener(on_press=on_press)
    listener.start()
    print("[Input] Listener active. Press 'Pause Break' for Control Center (🍁).")

def check_network_drive():
    try:
        # Get the directory of the executable
        app_path = os.path.abspath(sys.argv[0])
        drive = os.path.splitdrive(app_path)[0]
        if drive:
            import win32file
            drive_type = win32file.GetDriveType(drive + "\\")
            if drive_type == win32file.DRIVE_REMOTE:
                from PyQt6.QtWidgets import QMessageBox
                msg = QMessageBox()
                msg.setIcon(QMessageBox.Icon.Warning)
                msg.setWindowTitle("環境建議 - Artale Agent")
                msg.setText("偵測到程式正在網路磁碟機 (Samba) 上執行。")
                msg.setInformativeText("在網路硬碟上執行可能會導致視窗捕捉失敗 (0x80070490)。\n\n建議將程式複製到「本機磁碟」(如桌面或 C 槽) 以獲得最佳穩定性。")
                msg.setStandardButtons(QMessageBox.StandardButton.Ok)
                msg.exec()
    except Exception:
        pass

def run_app():
    app = QApplication(sys.argv)
    app.setQuitOnLastWindowClosed(False)
    
    # Check for Samba/Network drive issues
    check_network_drive()
    
    overlay = ArtaleOverlay()
    settings_window = SettingsWindow(overlay)
    
    # Ship Reminder and notifications
    settings_window.timer_request.connect(overlay.timer_request)
    settings_window.notification_request.connect(overlay.notification_request)
    
    focus_tracker = GameFocusTracker()
    
    # Tray Icon Connection
    overlay.settings_show_request.connect(settings_window.safe_show)
    
    start_keyboard_listener(overlay, settings_window, focus_tracker)
    
    print("[Main] Artale Agent initialized. Waiting for input...")
    sys.exit(app.exec())

if __name__ == "__main__":
    run_app()
