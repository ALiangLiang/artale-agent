import sys
import logging
import ctypes
import ctypes.wintypes
import time
import os
import win32process
import platform
import psutil
import win32file
import sentry_sdk

# Force import for PyInstaller visibility and runtime thread-safety (sip.isdeleted)
try:
    from PyQt6 import QtWebSockets, QtNetwork
    import sip
except ImportError:
    pass

from sentry_sdk.integrations.logging import LoggingIntegration
from PyQt6.QtWidgets import QApplication, QMessageBox
from PyQt6.QtCore import Qt
from pynput import keyboard, mouse

# Local imports
import overlay
from overlay import ArtaleOverlay, SettingsWindow, ConfigManager
from utils import get_version

# Initialize logger
logger = logging.getLogger(__name__)

def check_dynamic_console():
    """Enable console window if --debug or --console argument is present"""
    if "--debug" in sys.argv or "--console" in sys.argv:
        try:
            # Attach to parent console or allocate a new one
            if not ctypes.windll.kernel32.AttachConsole(-1):
                ctypes.windll.kernel32.AllocConsole()
            
            # Re-map standard I/O to the new console
            sys.stdout = open("CONOUT$", "w", encoding='utf-8')
            sys.stderr = open("CONOUT$", "w", encoding='utf-8')
            
            # Ensure the console charset handles special characters
            os.system('chcp 65001 > nul')
            
            print("\n" + "="*50)
            print("🚀 Artale 瑞士刀 - Debug Console Enabled")
            print("="*50 + "\n")
        except Exception as e:
            logger.error(f"[Main] Failed to allocate console: {e}")

check_dynamic_console()

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
            if not hwnd or hwnd <= 0:
                self.is_game_active = False
                return
                
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            # Ensure PID is treated as a 32-bit unsigned integer
            pid = pid & 0xFFFFFFFF
            
            # Sanity check: PIDs are almost never > 1,000,000 in reality
            # Avoid passing extreme trash PIDs to psutil
            if pid <= 0 or pid > 2**31: 
                self.is_game_active = False
                return
            
            try:
                proc = psutil.Process(pid)
                p_name = proc.name().lower()
            except (psutil.NoSuchProcess, psutil.AccessDenied, ValueError):
                self.is_game_active = False
                return
                
            was_active = self.is_game_active
            self.is_game_active = (p_name == self.TARGET_PROCESS)
            if self.is_game_active != was_active:
                status = "focused" if self.is_game_active else "lost focus"
                logger.info(f"[Focus] {self.TARGET_PROCESS} {status}")
        except Exception:
            # Silent fallback for focus tracking to avoid console spam
            self.is_game_active = False

def start_keyboard_listener(overlay, settings_window, focus_tracker):
    # --- Mouse Listener for Right Click Cancellation ---
    def on_click(x, y, button, pressed):
        if not pressed: return
        if not focus_tracker.is_game_active: return
        
        if button == mouse.Button.right:
            overlay.check_right_click(x, y)
        elif button == mouse.Button.left:
            overlay.check_left_click(x, y)

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
        logger.info(f"[Config] Switched to {active}. Triggers: {list(triggers.keys())}")

    # Connect settings signal to update listener
    settings_window.config_updated.connect(update_local_config)

    def on_press(key):
        nonlocal last_key, last_time, current_config, is_globally_enabled
        try:
            k_name = None
            # 1. Prioritize VK for Numpad keys (96-105, 110) to distinguish from top row
            vk = getattr(key, 'vk', None)
            if vk is None:
                s_key = str(key)
                if s_key.startswith('<') and s_key.endswith('>'):
                    try: vk = int(s_key[1:-1])
                    except Exception as e: 
                        logger.debug(f"[Input] Failed to parse VK: {e}")
            
            if vk is not None:
                if 96 <= vk <= 105:
                    k_name = f"num_{vk - 96}"
                elif vk == 110: # Numpad dot
                    k_name = "num_dot"
            
            # 2. Try to get name or char if not a numpad key
            if k_name is None:
                if hasattr(key, 'name'):
                    k_name = key.name
                elif hasattr(key, 'char') and key.char:
                    k_name = key.char.lower()
                
                pass
            
            if k_name:
                for base in ["alt", "shift", "ctrl"]:
                    if k_name.startswith(base):
                        k_name = base
                        break
            
            # 3. Profile Switching (Double Press F1-F9) or Disable (F12)
            now = time.time()
            
            # Use configurable hotkeys
            hks = current_config.get("hotkeys", {})

            # 2. Always-on controls
            if k_name == hks.get("show_settings", "pause"):
                logger.info(f"[Input] {k_name.upper()} pressed. Emitting show signal.")
                settings_window.request_show.emit()
                return
            
            # --- BLOCK TRIGGERS IF RECORDING ---
            if settings_window.is_recording or settings_window.recording_global_key:
                return
            
            if k_name == hks.get("reset", "f12"):
                logger.info(f"[Input] {k_name.upper()} Reset Triggered.")
                is_globally_enabled = False
                overlay.clear_request.emit()
                return

            if k_name == hks.get("exp_toggle", "f10"):
                overlay.toggle_exp_request.emit()
                return

            if k_name == hks.get("exp_pause", "f11"):
                overlay.toggle_pause_request.emit()
                return

            if k_name == hks.get("exp_report", "f12"):
                overlay.export_report_request.emit()
                return

            # --- RJPQ SMART HOTKEYS ---
            rjpq_keys = {
                hks.get("rjpq_1", "num_1"): 0,
                hks.get("rjpq_2", "num_2"): 1,
                hks.get("rjpq_3", "num_3"): 2,
                hks.get("rjpq_4", "num_4"): 3,
            }
            
            if k_name in rjpq_keys:
                col_idx = rjpq_keys[k_name]
                # Check if RJPQ is active and connected
                if hasattr(settings_window, 'rjpq_tab') and settings_window.rjpq_tab.client.is_connected:
                    if settings_window.rjpq_tab.mark_by_hotkey(col_idx):
                        return # Consume the key if it was used for RJPQ

            # 3. Profile Switching (Double Press F1-F9)
            now = time.time()
            if k_name and k_name.startswith('f') and len(k_name) <= 3:
                f_num = k_name[1:]
                if f_num.isdigit() and 1 <= int(f_num) <= 8:
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
            logger.error(f"[Error] Listener: {e}")

    listener = keyboard.Listener(on_press=on_press)
    listener.start()
    logger.info("[Input] Listener active. Press 'Pause Break' for Control Center (🍁).")

def check_network_drive():
    try:
        # Get the directory of the executable
        app_path = os.path.abspath(sys.argv[0])
        drive = os.path.splitdrive(app_path)[0]
        if drive:
            drive_type = win32file.GetDriveType(drive + "\\")
            if drive_type == win32file.DRIVE_REMOTE:
                msg = QMessageBox()
                msg.setIcon(QMessageBox.Icon.Warning)
                msg.setWindowTitle("環境建議 - Artale 瑞士刀")
                msg.setText("偵測到程式正在網路磁碟機 (Samba) 上執行，這可能導致 OCR 效能嚴重下降或辨識失敗。請考慮將程式移動到本機硬碟執行以獲得最佳效能。")
                msg.setInformativeText("在網路硬碟上執行可能會導致視窗捕捉失敗 (0x80070490)。\n\n建議將程式複製到「本機磁碟」(如桌面或 C 槽) 以獲得最佳穩定性。")
                msg.setStandardButtons(QMessageBox.StandardButton.Ok)
                msg.exec()
    except Exception as e:
        logger.debug(f"[Main] Network drive check skipped or failed: {e}")

def run_app():
    # Setup logging
    log_file = os.path.join(os.getcwd(), "artale_agent.log")
    logging.basicConfig(
        level=logging.DEBUG,
        format='%(asctime)s [%(levelname)s] %(message)s',
        datefmt='%H:%M:%S',
        handlers=[
            logging.StreamHandler(),
            logging.FileHandler(log_file, encoding='utf-8', mode='w')
        ]
    )
    logger = logging.getLogger("Artale")
    logger.info(f"--- Artale Agent Initializing (Log: {log_file}) ---")
    logger.info(f"[System] OS: {platform.platform()}")

    # --- Initialize Sentry (Only in bundled/production mode) ---
    if getattr(sys, 'frozen', False):
        sentry_sdk.init(
            dsn="https://b120418a69ec5d8ccd74a0bb4d2acacf@o4511210222452736.ingest.us.sentry.io/4511210254565376",
            integrations=[
                LoggingIntegration(
                    level=logging.INFO,
                    event_level=logging.ERROR
                ),
            ],
            traces_sample_rate=1.0,
            enable_logs=True,
            release=f"artale-agent@{get_version()}"
        )
    else:
        logger.info(f"[Main] Dev mode: Sentry disabled.")

    # --- Enable High DPI Awareness ---
    # Qt 6 defaults to PerMonitorAwareV2, so manual ctypes calls are redundant and cause "Access Denied" errors.
    try:
        QApplication.setHighDpiScaleFactorRoundingPolicy(Qt.HighDpiScaleFactorRoundingPolicy.PassThrough)
    except Exception as e:
        logger.debug(f"[Main] Rounding policy set failed: {e}")

    app = QApplication(sys.argv)
    
    from PyQt6.QtGui import QFont
    font = QFont()
    font.setFamilies(["Microsoft JhengHei", "微軟正黑體"])
    app.setFont(font)
    app.setQuitOnLastWindowClosed(False)
    
    # Check for Samba/Network drive issues
    # check_network_drive()
    
    main_overlay = ArtaleOverlay()
    settings_window = main_overlay.settings_window
    focus_tracker = GameFocusTracker()
    
    # Tray Icon Connection
    main_overlay.settings_show_request.connect(settings_window.safe_show)
    
    start_keyboard_listener(main_overlay, settings_window, focus_tracker)
    
    # Auto-show Control Center on startup
    settings_window.safe_show()
    
    logger.info("[Main] Artale 瑞士刀 initialized. Waiting for input...")
    sys.exit(app.exec())

if __name__ == "__main__":
    run_app()
