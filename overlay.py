import sys
import json
import os
import threading
import time
import re
import logging
import webbrowser
import subprocess
import urllib.request
import datetime
import psutil
try:
    import win32gui
    import win32api
    import win32con
    import win32process
    import winsound
except ImportError:
    win32gui = win32api = win32con = win32process = winsound = None

import cv2
import numpy as np
import pytesseract
from PyQt6 import sip
from PyQt6.QtWidgets import (QApplication, QWidget, QVBoxLayout, QHBoxLayout, 
                             QLabel, QLineEdit, QPushButton, QScrollArea, QFrame,
                             QGridLayout, QDialog, QTabWidget, QComboBox, QSlider, QCheckBox,
                             QSystemTrayIcon, QMenu, QGroupBox)
from PyQt6.QtCore import Qt, QPoint, QRect, QTimer, pyqtSignal, QSize, QRectF, QUrl, QObject, QStandardPaths
from PyQt6.QtGui import QFont, QColor, QPainter, QPen, QPixmap, QIcon, QPainterPath, QAction, QImage, QLinearGradient
try:
    from PyQt6.QtWebSockets import QWebSocket
except ImportError:
    QWebSocket = None
from PyQt6.QtNetwork import QAbstractSocket

try:
    from windows_capture import WindowsCapture, Frame
except ImportError:
    WindowsCapture = None

# Local imports
from rjpq_tool import RJPQSyncClient, RJPQTabContent, draw_rjpq_panel
from skill_timer import IconSelectorDialog, PositionHandle, TimerManager
from settings_window import SettingsWindow

# Initialize logger
logger = logging.getLogger(__name__)

# Using utils.py for VERSION, REPO_URL, resource_path, ConfigManager
from utils import VERSION, REPO_URL, resource_path, ConfigManager

# Tesseract Portable Setup (LOCAL ONLY)
def get_tess_cmd():
    """Detect Tesseract-OCR path and handle DLL loading for PyInstaller"""
    import os, sys
    
    executable_path = None
    
    # 1. Check for PyInstaller internal bundle path
    if hasattr(sys, '_MEIPASS'):
        bundle_dir = sys._MEIPASS
        executable_path = os.path.join(bundle_dir, "Tesseract-OCR", "tesseract.exe")
        if os.path.exists(executable_path):
            # IMPORTANT: For Tesseract to find grouped DLLs, its folder must be in PATH
            tess_dir = os.path.dirname(executable_path)
            if tess_dir not in os.environ["PATH"]:
                os.environ["PATH"] = tess_dir + os.pathsep + os.environ["PATH"]
            return executable_path

    # 2. Check for local folder (for portable/dev use)
    base_dir = os.path.dirname(os.path.abspath(sys.argv[0] if getattr(sys, 'frozen', False) else __file__))
    local_tess = os.path.join(base_dir, "Tesseract-OCR", "tesseract.exe")
    if os.path.exists(local_tess):
        tess_dir = os.path.dirname(local_tess)
        if tess_dir not in os.environ["PATH"]:
            os.environ["PATH"] = tess_dir + os.pathsep + os.environ["PATH"]
        return local_tess
    
    # 3. Last fallback (standard path)
    common_p = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
    if os.path.exists(common_p):
        return common_p
    
    return None

if pytesseract:
    pytesseract.pytesseract.tesseract_cmd = get_tess_cmd()

# Using resource_path from utils.py (imported above)

# ConfigManager moved to utils.py

# PositionHandle and IconSelectorDialog moved to skill_timer.py

# SettingsWindow moved to settings_window.py

# SettingsWindow moved to settings_window.py

class ArtaleOverlay(QWidget):
    # 1080p Calibration Reference (Bottom-Left Anchored)
    BASE_W, BASE_H = 1920, 1080
    X_OFF_FROM_LEFT = 1084   # Fixed horizontal offset from left
    Y_OFF_FROM_BOTTOM = 66   # Fixed vertical offset from bottom
    BASE_CW, BASE_CH = 240, 22
    
    LV_X_OFF_FROM_LEFT = 100
    LV_Y_OFF_FROM_BOTTOM = 46
    LV_BASE_CW, LV_BASE_CH = 75, 26
    
    timer_request = pyqtSignal(str, int, str, bool) 
    clear_request = pyqtSignal()
    notification_request = pyqtSignal(str)
    profile_switch_request = pyqtSignal()
    exp_update_request = pyqtSignal(dict)    # For Stats data
    exp_visual_request = pyqtSignal(dict)    # For Debug Images: exp, lv, coin
    lv_update_request = pyqtSignal(dict)
    toggle_exp_request = pyqtSignal()
    toggle_pause_request = pyqtSignal()
    toggle_rjpq_request = pyqtSignal()
    settings_show_request = pyqtSignal()
    rjpq_cell_clicked = pyqtSignal(int)
    export_report_request = pyqtSignal()
    update_found = pyqtSignal(str, str)
    money_update_request = pyqtSignal(int)
    
    def __init__(self, target_window_title="MapleStory Worlds-Artale (繁體中文版)"):
        super().__init__()
        self.target_window_title = target_window_title
        self.timer_manager = TimerManager(self)
        self.timer_manager.updated.connect(self.update)
        self.click_zones = {}  
        self.is_active = False # For timers compat
        self.show_preview = False
        self.active_profile_name = "F1"
        self._is_running = True
        
        # Load configs early
        config = ConfigManager.load_config()
        self.show_exp_panel = config.get("show_exp", False)
        self.show_money_log = config.get("show_money_log", True)
        self.exp_paused = False # New state
        self.total_pause_time = 0 # Cumulative pause duration
        self.pause_start_time = 0
        self.needs_calibration = False # Flag for resume frame
        self.show_rjpq_panel = False # Default off
        self.show_debug = config.get("show_debug", False)
        self.base_opacity = config.get("opacity", 0.5)
        
        self.msg_text = ""; self.msg_opacity = 0
        self.x_offset = 0; self.y_offset = 0
        self.exp_x_offset = 0; self.exp_y_offset = 0
        self.current_exp_data = {"text": "---", "value": 0, "percent": 0.0, "gained_10m": 0, "percent_10m": 0.0}
        self.exp_history = [] 
        self.exp_initial_val = None 
        self.selected_color = -1
        self.cumulative_gain = 0
        self.cumulative_pct = 0.0
        self.max_10m_exp = 0 
        self.last_exp_pct = 0.0
        self.money_history = [] 
        self.detected_coins = [] # Deduplication: (timestamp, x, y, val)
        self.cumulative_money = 0
        self.rjpq_data = [4] * 40
        self.rjpq_x_offset = -400
        self.rjpq_y_offset = 0
        self.rjpq_click_zones = {}
        self.last_capture_time = 0
        self.current_lv = None 
        self.last_confirmed_lv = None # To detect level up
        self.exp_rate_history = [] # Rolling 10m rates for the graph
        self.exp_tracker_event = threading.Event()
        self._tesseract_error_shown = False
        self.last_crop_info = None
        self.last_lv_crop_info = None
        
        self.timer_request.connect(self.timer_manager.start_timer)
        self.clear_request.connect(self.timer_manager.clear_all)
        self.notification_request.connect(self.show_notification)
        self.profile_switch_request.connect(self.load_profile_immediately)
        self.exp_update_request.connect(self.on_exp_update)
        self.toggle_exp_request.connect(self.on_toggle_exp)
        self.toggle_pause_request.connect(self.on_toggle_pause)
        self.toggle_rjpq_request.connect(self.on_toggle_rjpq)
        self.export_report_request.connect(self.export_exp_report)
        self.update_found.connect(self.on_update_found)
        self.money_update_request.connect(self.on_money_update)
        
        # Instantiate SettingsWindow (now from separate module)
        self.settings_window = SettingsWindow(self)
        self.settings_show_request.connect(self.settings_window.request_show.emit)
        self.settings_window.config_updated.connect(self.load_profile_immediately)
        self.settings_window.timer_request.connect(self.timer_manager.start_timer)
        self.settings_window.notification_request.connect(self.show_notification)
        
        # Explicitly connect tracking updates to settings window
        self.exp_visual_request.connect(self.settings_window.update_debug_img)
        self.lv_update_request.connect(self.settings_window.update_lv_debug_img)
        self.update_found.connect(self.settings_window.show_update_banner)
        
        # We'll use this signal for tray to talk to settings_window
        self.request_show_settings_signal = pyqtSignal()
        if hasattr(self, 'request_show_settings_signal'):
             # Handle signal in main.py later
             pass
        
        self.tracking_timer = QTimer(self); self.tracking_timer.timeout.connect(self.sync_with_game_window); self.tracking_timer.start(100)
        self.world_timers = {} 
        
        # Initialize ExpTracker (Only start if panel is requested on startup)
        if WindowsCapture and self.show_exp_panel:
            self.exp_tracker_thread = threading.Thread(target=self.run_exp_tracker, daemon=True)
            self.exp_tracker_thread.start()
        elif not WindowsCapture:
            logger.warning("[Warning] windows-capture is not installed. EXP tracking disabled.")
        
        frame_p = resource_path("buff_pngs/skill_frame.png")
        self.icon_frame = QPixmap(frame_p) if os.path.exists(frame_p) else None
        self.last_coin_pos = None # (x, y, w, h) in client coords
        self.last_coin_info_pos = None
        self.last_coin_ocr = ""
        self.load_profile_immediately()
        self.init_tray()
        
        # Load coin template for matching
        self.coin_tpl = None
        if os.path.exists("coin.png"):
            self.coin_tpl = cv2.imread("coin.png")
            if self.coin_tpl is not None:
                logger.info(f"[ExpTracker] Loaded coin template: {self.coin_tpl.shape}")
        
        self.init_ui()

    def init_tray(self):
        """Initialize System Tray Icon with context menu"""
        self.tray_icon = QSystemTrayIcon(self)
        
        # Use our beautiful generated app icon
        icon_path = resource_path("app_icon.png")
        if os.path.exists(icon_path):
            self.tray_icon.setIcon(QIcon(icon_path))
        
        # Create context menu
        tray_menu = QMenu()
        
        show_settings_action = QAction("🚀 開啟控制中心 (Pause)", self)
        show_settings_action.triggered.connect(self.request_show_settings)
        tray_menu.addAction(show_settings_action)
        
        reset_exp_action = QAction("📊 重置經驗值統計", self)
        reset_exp_action.triggered.connect(self.reset_exp_stats)
        tray_menu.addAction(reset_exp_action)
        
        tray_menu.addSeparator()
        
        quit_action = QAction("❌ 結束程式", self)
        quit_action.triggered.connect(QApplication.instance().quit)
        tray_menu.addAction(quit_action)
        
        self.tray_icon.setContextMenu(tray_menu)
        self.tray_icon.setToolTip("Artale 瑞士刀")
        
        # Click to toggle settings
        self.tray_icon.activated.connect(self.on_tray_activated)
        self.tray_icon.show()

    def on_tray_activated(self, reason):
        if reason == QSystemTrayIcon.ActivationReason.Trigger:
            self.request_show_settings()

    def request_show_settings(self):
        self.settings_show_request.emit()

    def reset_exp_stats(self, silent=False):
        """Reset EXP tracking baseline"""
        self.exp_history = []
        self.exp_rate_history = []
        self.cumulative_gain = 0
        self.cumulative_pct = 0.0
        self.max_10m_exp = 0
        self.exp_initial_val = None
        self.last_exp_pct = 0.0
        if not silent:
            self.show_notification("📊 經驗值統計已重置")

    def on_update_found(self, tag, url):
        self._latest_version_info = (tag, url)
        # We don't need any special logic here, 
        # but if SettingsWindow is already open, it should know.
        pass

    def check_for_updates(self, auto=False):
        """Check GitHub for new releases"""
        def _check():
            try:
                import urllib.request, json, webbrowser
                url = f"https://api.github.com/repos/{REPO_URL}/releases/latest"
                req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
                with urllib.request.urlopen(req, timeout=5) as response:
                    data = json.loads(response.read().decode())
                    latest_tag = data.get("tag_name", VERSION)
                    if latest_tag != VERSION:
                        html_url = data.get("html_url", f"https://github.com/{REPO_URL}/releases")
                        self.update_found.emit(latest_tag, html_url)
                        msg = f"✨ 發現新版本: {latest_tag}！請下載更新"
                        self.notification_request.emit(msg)
                        if not auto: webbrowser.open(html_url)
                    else:
                        if not auto: self.notification_request.emit("✅ 目前已是最新版本")
            except Exception as e:
                logger.debug(f"[Update] Check failed: {e}")
                if not auto: self.notification_request.emit(f"❌ 檢查失敗: {e}")
        
        threading.Thread(target=_check, daemon=True).start()


    def init_ui(self):
        self.setWindowFlags(Qt.WindowType.WindowStaysOnTopHint | Qt.WindowType.FramelessWindowHint | Qt.WindowType.WindowTransparentForInput | Qt.WindowType.Tool)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.setAttribute(Qt.WidgetAttribute.WA_AlwaysStackOnTop)
        
        # Ensure we cover ALL monitors correctly across the entire Virtual Desktop
        v_rect = QApplication.primaryScreen().virtualGeometry()
        self.setGeometry(v_rect)
        
        # Move to the absolute top-left of the virtual desktop (handles negative coords)
        self.move(v_rect.topLeft())
        
        # Subtle trick for Windows composition: 0.99 opacity can force full-window rendering
        self.setWindowOpacity(0.99)
        
        logger.debug(f"[Debug] Overlay spans: {v_rect.x()}, {v_rect.y()} to {v_rect.width()}, {v_rect.height()}")
        self.show()

    def play_sound(self, times=1):
        self.timer_manager.play_sound(times)

    def sync_with_game_window(self):
        # We no longer move the Overlay window; it stays full-screen on the virtual desktop.
        # This keeps the UI stable and independent of the game's movement.
        if not win32gui: return
        hwnd = 0
        try:
            my_pid = os.getpid()
            def callback(h, extra):
                nonlocal hwnd
                try:
                    if not win32gui.IsWindowVisible(h): return True
                    _, pid = win32process.GetWindowThreadProcessId(h)
                    if pid == my_pid: return True
                    title = win32gui.GetWindowText(h).lower()
                    if self.target_window_title.lower() in title:
                        hwnd = h; return False
                except: pass
                return True
            win32gui.EnumWindows(callback, 0)
        except Exception as e:
            # Error 2 or 1400 are common during window transitions, log only other errors
            err_str = str(e)
            if not any(code in err_str for code in ["(2,", "(1400,"]):
                logger.debug(f"[Overlay] Window search failed: {e}")
            self.game_hwnd = None
        
        if hwnd:
            try:
                # Update anchor point bx, by internally based on current game position
                # so that relative elements (like debug box) know where the game is.
                rect = win32gui.GetClientRect(hwnd)
                client_h = rect[3]
                bl_point = win32gui.ClientToScreen(hwnd, (0, client_h))
                
                # DPI Scaling Correction:
                # win32gui returns physical pixels. Qt uses logical pixels.
                dpr = self.screen().devicePixelRatio()
                logical_gl_pt = QPoint(int(bl_point[0] / dpr), int(bl_point[1] / dpr))
                
                local_bl = self.mapFromGlobal(logical_gl_pt)
                self.bx, self.by = local_bl.x(), local_bl.y()
            except Exception as e:
                logger.debug(f"[Overlay] ClientToScreen mapping failed: {e}")
            
        # Ensure overlay is visible and on top, but DON'T change its geometry
        if not self.isVisible(): self.show()
        self.raise_()

    def update_offset(self, gx, gy):
        local = self.mapFromGlobal(QPoint(gx, gy))
        self.x_offset = local.x() - self.rect().center().x()
        self.y_offset = local.y() - self.rect().center().y()
        self.click_zones = {}; self.update()

    def update_exp_offset(self, gx, gy):
        local = self.mapFromGlobal(QPoint(gx, gy))
        # Account for the logic in draw_exp_panel:
        # bx = center_x + exp_x_offset
        # Top-Right corner Y = (center_y + exp_y_offset) - 120 - ph // 2
        ph = 115
        self.exp_x_offset = local.x() - self.rect().center().x()
        self.exp_y_offset = local.y() - self.rect().center().y() + 120 + (ph // 2)
        self.update()

    def update_rjpq_offset(self, gx, gy):
        local = self.mapFromGlobal(QPoint(gx, gy))
        self.rjpq_x_offset = local.x() - self.rect().center().x()
        self.rjpq_y_offset = local.y() - self.rect().center().y()
        self.update()

    def clear_all_timers(self, show_msg=True):
        self.timer_manager.clear_all()
        if show_msg: self.show_notification("⚠️ 已強制關閉並重設計時器")

    def check_left_click(self, gx, gy):
        p = QPoint(gx, gy)
        if self.show_rjpq_panel:
            for idx, rect in self.rjpq_click_zones.items():
                if rect.contains(p):
                    self.rjpq_cell_clicked.emit(idx)
                    return True
        return False

    def check_right_click(self, gx, gy):
        p = QPoint(gx, gy)
        for key, rect in list(self.click_zones.items()):
            if rect.contains(p):
                if key in self.timer_manager.active_timers: 
                    del self.timer_manager.active_timers[key]
                    self.timer_manager.updated.emit()
                return True
        return False

    def on_toggle_exp(self):
        self.show_exp_panel = not self.show_exp_panel
        status = "已啟用" if self.show_exp_panel else "已關閉"
        logger.info(f"[Overlay] EXP Panel toggled: {status}")
        self.show_notification(f"📊 經驗監測系統 {status} (F10)")

        if self.show_exp_panel:
            # Wake up the thread immediately
            self.exp_tracker_event.set()
            # Start/Restart tracker thread if needed
            if not hasattr(self, 'exp_tracker_thread') or not self.exp_tracker_thread.is_alive():
                self.exp_tracker_thread = threading.Thread(target=self.run_exp_tracker, daemon=True)
                self.exp_tracker_thread.start()
        else:
            # Full logic reset on toggle off
            self.current_exp_data = {
                "text": "---", "value": 0, "percent": 0.0, 
                "gained_10m": 0, "percent_10m": 0.0, 
                "is_estimated": True, "tracking_duration": 0, "time_to_level": -1
            }
            self.exp_history = []
            self.exp_rate_history = [] 
            self.exp_initial_val = None
            self.exp_session_start_time = None
            self.cumulative_gain = 0
            self.cumulative_pct = 0.0
            self.last_exp_pct = 0.0
            self.total_pause_time = 0
            self.pause_start_time = 0
            self.exp_paused = False
            self.needs_calibration = False
            # Note: run_exp_tracker will eventually exit because on_frame_arrived_callback 
            # will see self.show_exp_panel is False if we set it as a hard exit condition.
        self.update()

    def on_toggle_pause(self):
        now = time.time()
        self.exp_paused = not self.exp_paused
        status = "已暫停" if self.exp_paused else "已恢復"
        
        if self.exp_paused:
            self.pause_start_time = now
        else:
            # Resuming
            if self.pause_start_time > 0:
                shift = now - self.pause_start_time
                self.total_pause_time += shift
                # Important: Shift the session start time so duration doesn't grow during pause
                if hasattr(self, 'exp_session_start_time') and self.exp_session_start_time:
                    self.exp_session_start_time += shift
                # Shift all history timestamps to maintain efficiency without dilution
                self.exp_history = [(t + shift, v, p) for t, v, p in self.exp_history]
            self.needs_calibration = True # Skip next frame's gain calculation
            
        logger.info(f"[ExpTracker] Recording {status}")
        self.show_notification(f"📊 經驗追蹤 {status} (F11)")
        self.update()

    def on_toggle_rjpq(self):
        self.show_rjpq_panel = not self.show_rjpq_panel
        self.show_notification(f"羅茱路徑面板: {'已啟用' if self.show_rjpq_panel else '已關閉'}")
        self.update()

    def update_rjpq_data(self, data):
        self.rjpq_data = data
        self.update()
        
    def set_rjpq_color(self, color):
        self.selected_color = color
        self.update()
    def set_rjpq_overlay_visible(self, visible):
        self.show_rjpq_panel = visible
        self.update()

    # --- EXP Tracker Logic ---
    def closeEvent(self, event):
        self._is_running = False
        super().closeEvent(event)

    def _perform_enhanced_ocr(self, thresh_img, key, upscale=4, whitelist="0123456789[].% ", psm=7):
        """
        Two-Pass Split Mode:
        Physically separates the numeric value from the percentage via geometric anchor detection.
        """
        ocr_text = ""
        native_conf = 0
        stable_img_attr = f"_stable_img_{key}"
        
        # 1. Padding and Segmentation
        thresh_img = cv2.copyMakeBorder(thresh_img, 10, 10, 50, 50, cv2.BORDER_CONSTANT, value=0)
        contours, _ = cv2.findContours(thresh_img, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        
        raw_boxes = []
        for cnt in contours:
            x, y, w, h = cv2.boundingRect(cnt)
            # Keep tiny symbols like dots (h>=2), 
            # we rely on the bracket_idx logic (h>=14) to ignore commas later.
            if h >= 2: raw_boxes.append([x, y, w, h])
        raw_boxes.sort(key=lambda b: b[0])
        
        merged_boxes = []
        if raw_boxes:
            curr = raw_boxes[0]
            for i in range(1, len(raw_boxes)):
                nxt = raw_boxes[i]
                if nxt[0] <= (curr[0] + curr[2] + 2): 
                    x1 = min(curr[0], nxt[0]); y1 = min(curr[1], nxt[1])
                    x2 = max(curr[0] + curr[2], nxt[0] + nxt[2])
                    y2 = max(curr[1] + curr[3], nxt[1] + nxt[3])
                    curr = [x1, y1, x2 - x1, y2 - y1]
                else:
                    merged_boxes.append(curr); curr = nxt
            merged_boxes.append(curr)
            
        if not merged_boxes:
            return "", 0, cv2.bitwise_not(thresh_img)
            
        # 2. Logic Split: Identify Brackets by Height (Brackets are the tallest in the line)
        bracket_idx = -1
        bracket_end_idx = -1
        
        if key == "exp" and len(merged_boxes) > 4:
            heights = [b[3] for b in merged_boxes]
            max_h = max(heights)
            min_h = min(heights)
            
            # Significant height gap indicates brackets exist (usually 16px vs 14px)
            if max_h >= min_h + 2:
                # First tallest is [, last tallest is ]
                for i, b in enumerate(merged_boxes):
                    if b[3] >= max_h - 1:
                        if bracket_idx == -1: bracket_idx = i
                        bracket_end_idx = i
        
        # 3. Two-Pass Reconstruction
        target_h = 60
        # Increase spacing for level back to 20 (90% conf reached previously)
        char_spacing = 20 if bracket_idx == -1 else 30
        
        def build_pass_canvas(boxes):
            if not boxes: return None
            # Find max height in this pass to calculate uniform scale
            max_h_pass = max([b[3] for b in boxes])
            scale = target_h / max_h_pass if max_h_pass > 0 else 1.0
            
            # Increase horizontal padding: char_spacing + extra 40px on both sides
            total_w = sum([int(b[2] * scale) for b in boxes]) + (len(boxes) + 1) * char_spacing + 80
            # Increase vertical padding: target_h + 120px (60px top/bot)
            canvas = np.ones((target_h + 120, total_w), dtype=np.uint8) * 255
            curr_x = 40 # Start with 40px left padding
            
            # Baseline alignment: move characters down weight (60px top padding)
            baseline_y = target_h + 60
            
            for b in boxes:
                char = thresh_img[b[1]:b[1]+b[3], b[0]:b[0]+b[2]]
                char_inv = cv2.bitwise_not(char)
                nw, nh = int(b[2] * scale), int(b[3] * scale)
                if nw > 0 and nh > 0:
                    resized = cv2.resize(char_inv, (nw, nh), interpolation=cv2.INTER_CUBIC)
                    
                    if key == "lv":
                        # Apply smoothing ONLY for Level (LV) as requested
                        resized = cv2.GaussianBlur(resized, (25, 25), 0)
                        _, resized = cv2.threshold(resized, 180, 255, cv2.THRESH_BINARY)
                    
                    y_off = baseline_y - nh
                    canvas[y_off:y_off+nh, curr_x:curr_x+nw] = resized
                    curr_x += nw + char_spacing
            return canvas

        parts = []
        if bracket_idx != -1:
            # Value part (Purely digits)
            parts.append({'boxes': merged_boxes[:bracket_idx], 'wl': '0123456789'})
            # Percentage part (Numeric and symbols, brackets handled manually)
            parts.append({'boxes': merged_boxes[bracket_idx:], 'wl': '0123456789.%'})
        else:
            # Fallback (e.g. for Level)
            parts.append({'boxes': merged_boxes, 'wl': '0123456789'})

        # 4. OCR Passes & Debug Stitching
        final_texts = []
        conf_sums = []
        canvases = []
        
        # Helper to manually add brackets if we found them
        found_brackets = (bracket_idx != -1)

        for i, p in enumerate(parts):
            boxes_to_use = p['boxes']
            if found_brackets and i == 1:
                start_in_part_b = 1
                end_in_part_b = len(boxes_to_use)
                if bracket_end_idx != -1 and bracket_end_idx > bracket_idx:
                    end_in_part_b = len(boxes_to_use) - 1
                boxes_to_use = boxes_to_use[start_in_part_b:end_in_part_b]

            cvs = build_pass_canvas(boxes_to_use)
            if cvs is not None:
                canvases.append(cvs)
                wl = p['wl'].replace('[', '').replace(']', '')
                config = f'--psm 7 --oem 3 -c tessedit_char_whitelist={wl}'
                try:
                    data = pytesseract.image_to_data(cvs, config=config, output_type=pytesseract.Output.DICT)
                    txts = [t for t in data['text'] if t.strip()]
                    confs = [int(c) for c in data['conf'] if int(c) != -1]
                    
                    if i == 1 and found_brackets:
                        res = "".join(txts)
                        final_texts.append(f"[{res}]")
                    else:
                        final_texts.append("".join(txts))
                    if confs: conf_sums.extend(confs)
                except Exception as e:
                    logger.debug(f"[OCR] Pass Error: {e}")
        
        # Stitch canvases for Control Center debug
        if len(canvases) > 1:
            max_w = max(c.shape[1] for c in canvases)
            total_h = sum(c.shape[0] for c in canvases) + (len(canvases) - 1) * 10
            combined = np.ones((total_h, max_w), dtype=np.uint8) * 255
            curr_y = 0
            for c in canvases:
                h_c, w_c = c.shape
                combined[curr_y:curr_y+h_c, :w_c] = c
                curr_y += h_c + 10 
            processed_img = combined
        elif canvases:
            processed_img = canvases[0]
        else:
            processed_img = cv2.bitwise_not(thresh_img)

        ocr_text = "".join(final_texts)
        native_conf = sum(conf_sums) / len(conf_sums) if conf_sums else 0
        if native_conf <= 0 and any(c.isdigit() for c in ocr_text):
            native_conf = 85.0
            
        print(f"[OCR-DEBUG-SPLIT] '{ocr_text}' | Conf: {native_conf:.1f}%")
        
        # Apply Stability
        curr_score_attr = f"_current_score_{key}"
        if not hasattr(self, stable_img_attr):
            setattr(self, stable_img_attr, None)
            setattr(self, curr_score_attr, 0)
            
        prev_img = getattr(self, stable_img_attr)
        setattr(self, stable_img_attr, thresh_img.copy())
        
        # 4. Pure Neural Reporting
        new_score = native_conf
        setattr(self, curr_score_attr, new_score)
        
        # 6. Logging Low Confidence (Lv: 85%, Exp: 90%)
        warn_thresh = 85 if key == "lv" else 90
        if new_score < warn_thresh and self.show_debug:
            logger.warning(f"[OCR] Low Confidence ({key}): {new_score:.1f}% | Text: {ocr_text}")
            # Save the EXACT image sent to Tesseract for visual debugging
            try:
                if not os.path.exists("tmp"): os.makedirs("tmp")
                clean_text = "".join(c for c in ocr_text if c.isalnum() or c in "._-")
                fname = f"tmp/{key}_{new_score:.1f}_{clean_text}.png"
                cv2.imwrite(fname, processed_img)
            except Exception as e:
                logger.debug(f"[OCR] Failed to save debug image: {e}")
        
        return ocr_text, new_score, processed_img

    def run_exp_tracker(self):
        """Background thread to capture and recognize EXP with adaptive scaling"""
        last_processed = 0
        target_hwnd = None
        
        def on_frame_arrived_callback(frame, control):
            nonlocal last_processed, target_hwnd
            try:
                # Process if panel is ON
                if not self._is_running: return
                if not getattr(self, "show_exp_panel", False):
                    # Hard stop when panel is hidden, even if debug is on (for performance)
                    logger.info("[ExpTracker] Stopping session (Panel Closed)")
                    self._exp_tracker_active = False # Manual break for the inner loop
                    control.stop()
                    return
                
                now = time.time()
                if now - last_processed < 1.0: return
                
                # Window Search (Dynamic lookup)
                if not target_hwnd or not win32gui.IsWindow(target_hwnd):
                    candidate = None
                    for name in ["MapleStory Worlds-Artale (繁體中文版)", "MapleStory Worlds-Artale"]:
                        h_cand = win32gui.FindWindow(None, name)
                        if h_cand: candidate = h_cand; break
                    if not candidate: return
                    target_hwnd = candidate
                
                # Handle handle errors or minimized state silently
                try:
                    placement = win32gui.GetWindowPlacement(target_hwnd)
                    if placement[1] == win32con.SW_SHOWMINIMIZED: return
                except Exception:
                    target_hwnd = None; return

                img_orig = frame.frame_buffer
                img = cv2.cvtColor(img_orig, cv2.COLOR_BGRA2BGR)
                h, w = img.shape[:2]
                
                crect = win32gui.GetClientRect(target_hwnd)
                cw_ref, ch_ref = crect[2], crect[3]
                scale = min(cw_ref / self.BASE_W, ch_ref / self.BASE_H)
                
                # Handle maximized vs windowed border logic
                placement = win32gui.GetWindowPlacement(target_hwnd)
                if placement[1] == win32con.SW_SHOWMAXIMIZED:
                    off_x = 0; off_y = h - ch_ref
                else:
                    border_w = max(0, (w - cw_ref) // 2)
                    off_x = border_w; off_y = max(0, h - ch_ref - border_w)
                
                # 1. LV Recognition
                lv_cx = off_x + int(self.LV_X_OFF_FROM_LEFT * scale)
                lv_cy = off_y + (ch_ref - int(self.LV_Y_OFF_FROM_BOTTOM * scale))
                lv_cw = int(self.LV_BASE_CW * scale); lv_ch = int(self.LV_BASE_CH * scale)
                lv_crop = img[max(0, lv_cy):min(h, lv_cy+lv_ch), max(0, lv_cx):min(w, lv_cx+lv_cw)]
                if lv_crop.size > 0:
                    lv_gray = cv2.cvtColor(lv_crop, cv2.COLOR_BGR2GRAY)
                    r_lv = 60 / lv_gray.shape[0] if lv_gray.shape[0] > 0 else 3
                    lv_gray = cv2.resize(lv_gray, None, fx=r_lv, fy=r_lv, interpolation=cv2.INTER_CUBIC)
                    _, lv_thresh = cv2.threshold(lv_gray, 210, 255, cv2.THRESH_BINARY)
                    lv_text, lv_conf, lv_processed = self._perform_enhanced_ocr(lv_thresh, "lv", upscale=4, whitelist="0123456789")
                    if not sip.isdeleted(self): 
                        self.lv_update_request.emit({"thresh": lv_processed, "level": lv_text, "conf": lv_conf})

                # 2. Coin Recognition (With ROI to avoid static UI icons)
                if self.show_money_log and hasattr(self, 'coin_tpl') and self.coin_tpl is not None:
                    try:
                        # Search only in central area (Avoid status bars and inventory corners)
                        roi_y1, roi_y2 = int(h * 0.1), int(h * 0.85)
                        roi_x1, roi_x2 = int(w * 0.1), int(w * 0.9)
                        img_roi = img[roi_y1:roi_y2, roi_x1:roi_x2]
                        
                        tpl_h, tpl_w = self.coin_tpl.shape[:2]
                        st_w = int(tpl_w * scale); st_h = int(tpl_h * scale)
                        if st_w > 5 and st_h > 5:
                            tpl_resized = cv2.resize(self.coin_tpl, (st_w, st_h))
                            res = cv2.matchTemplate(img_roi, tpl_resized, cv2.TM_CCOEFF_NORMED)
                            _, max_val, _, max_loc = cv2.minMaxLoc(res)
                            if max_val > 0.8:
                                # Adjust coords back to full image
                                real_x, real_y = max_loc[0] + roi_x1, max_loc[1] + roi_y1
                                
                                self.last_coin_pos = (real_x - off_x, real_y - off_y, st_w, st_h)
                                info_w = int(280 * scale); info_h = int(31 * scale)
                                info_ix = real_x + st_w + int(30 * scale)
                                info_iy = real_y + (st_h // 2) - (info_h // 2) + int(1 * scale)
                                self.last_coin_info_pos = (info_ix - off_x, info_iy - off_y, info_w, info_h)
                                
                                ic_crop = img[max(0, info_iy):min(h, info_iy+info_h), max(0, info_ix):min(w, info_ix+info_w)]
                                if ic_crop.size > 0:
                                    ic_gray = cv2.cvtColor(ic_crop, cv2.COLOR_BGR2GRAY)
                                    ic_r = 60 / ic_gray.shape[0] if ic_gray.shape[0] > 0 else 3
                                    ic_gray = cv2.resize(ic_gray, None, fx=ic_r, fy=ic_r, interpolation=cv2.INTER_CUBIC)
                                    _, ic_thresh = cv2.threshold(ic_gray, 130, 255, cv2.THRESH_BINARY)
                                    if self.show_debug and not sip.isdeleted(self):
                                        self.exp_visual_request.emit({"coin": ic_thresh.copy()})
                                    if pytesseract and pytesseract.pytesseract.tesseract_cmd:
                                        try:
                                            ic_padded = cv2.copyMakeBorder(ic_thresh, 10, 10, 10, 10, cv2.BORDER_CONSTANT, value=0)
                                            c_text = pytesseract.image_to_string(ic_padded, config='--psm 7 -c tessedit_char_whitelist=0123456789,').strip()
                                            if c_text:
                                                self.last_coin_ocr = c_text
                                                val_str = c_text.replace(',', '')
                                                if val_str.isdigit():
                                                    m_val = int(val_str)
                                                    if not sip.isdeleted(self): 
                                                        self.money_update_request.emit(m_val)
                                        except: pass
                            else:
                                self.last_coin_pos = None; self.last_coin_info_pos = None; self.last_coin_ocr = ""
                    except: pass

                # 3. EXP OCR
                cx = off_x + int(self.X_OFF_FROM_LEFT * scale)
                cy = off_y + (ch_ref - int(self.Y_OFF_FROM_BOTTOM * scale))
                cw = int(self.BASE_CW * scale); ch = int(self.BASE_CH * scale)

                if self.exp_paused:
                    last_processed = now
                    self.update()
                    return

                crop = img[max(0, cy):min(h, cy+ch), max(0, cx):min(w, cx+cw)]
                if crop.size > 0:
                    self.last_crop_info = (cx, cy, cw, ch, w, h)
                    gray = cv2.cvtColor(crop, cv2.COLOR_BGR2GRAY)
                    r_e = 60 / gray.shape[0] if gray.shape[0] > 0 else 3
                    gray = cv2.resize(gray, None, fx=r_e, fy=r_e, interpolation=cv2.INTER_CUBIC)
                    _, thresh = cv2.threshold(gray, 180, 255, cv2.THRESH_BINARY)
                    
                    text, e_conf, e_processed = self._perform_enhanced_ocr(thresh, "exp", upscale=4, whitelist="0123456789.%")
                    if text:
                        if not sip.isdeleted(self):
                            self.exp_visual_request.emit({"thresh": e_processed, "conf": e_conf})
                        self.parse_and_update_exp(text, thresh, crop, e_conf)
                
                last_processed = now
                if not sip.isdeleted(self): self.update()
            except Exception as e:
                import traceback
                if not any(code in str(e) for code in ["1400", "(2,", "(1400,"]):
                    logger.error(f"[ExpTracker] Callback Critical Error: {e}\n{traceback.format_exc()}")

        while self._is_running:
            # 0. Global Efficiency Gate: Idle thread if panel is hidden
            if not getattr(self, "show_exp_panel", False):
                self._exp_tracker_active = False
                # Efficiently wait for the "Wake up" signal or check again in 1s
                self.exp_tracker_event.wait(timeout=1.0) 
                self.exp_tracker_event.clear()
                continue 
            
            target_hwnd = None
            
            # 1. Primary: Find Artale by precise window title
            my_pid = os.getpid()
            found_hwnds = []
            def enum_handler(hwnd, lparam):
                if win32gui.IsWindowVisible(hwnd):
                    # Check if the window belongs to OUR process
                    _, pid = win32process.GetWindowThreadProcessId(hwnd)
                    if pid == my_pid:
                        return True # Skip our own windows
                        
                    title = win32gui.GetWindowText(hwnd)
                    if "MapleStory Worlds-Artale" in title:
                        found_hwnds.append(hwnd)
            
            try:
                win32gui.EnumWindows(enum_handler, None)
            except: pass
            
            target_hwnd = found_hwnds[0] if found_hwnds else None
            
            # 2. Secondary: Fallback to process search (msw.exe)
            if not target_hwnd:
                try:
                    for proc in psutil.process_iter(['pid', 'name']):
                        if proc.info['name'] and proc.info['name'].lower() == 'msw.exe':
                            def callback(hwnd, extra):
                                if win32gui.IsWindowVisible(hwnd):
                                    _, pid = win32process.GetWindowThreadProcessId(hwnd)
                                    if pid == proc.info['pid']:
                                        extra.append(hwnd)
                                return True
                            found_hwnds = []
                            win32gui.EnumWindows(callback, found_hwnds)
                            if found_hwnds:
                                # Prioritize the one with the longest title (usually the game window)
                                target_hwnd = max(found_hwnds, key=lambda h: len(win32gui.GetWindowText(h)))
                                break
                except: pass

            if target_hwnd:
                precise_name = win32gui.GetWindowText(target_hwnd)
                try:
                    # Shared settings
                    cap_config = {
                        "window_name": precise_name,
                        "cursor_capture": False,
                        "minimum_update_interval": 1000
                    }
                    
                    try:
                        # Try borderless first (Win10 2004+)
                        capture = WindowsCapture(draw_border=False, **cap_config)
                    except Exception as e:
                        if "Toggling the capture border is not supported" in str(e):
                            capture = WindowsCapture(draw_border=True, **cap_config)
                        else:
                            raise e
                    
                    @capture.event
                    def on_frame_arrived(frame, control):
                        on_frame_arrived_callback(frame, control)
                    
                    @capture.event
                    def on_closed():
                        logger.info("[ExpTracker] Session closed.")
                        self._exp_tracker_active = False
                    
                    logger.info(f"[ExpTracker] Starting session for HWND {target_hwnd}")
                    
                    self._exp_tracker_active = True
                    capture.start_free_threaded()
                    
                    # Stay in this inner loop as long as session is healthy and panel is open
                    while self._exp_tracker_active and self._is_running and getattr(self, "show_exp_panel", False):
                        time.sleep(1.0)
                except Exception as e:
                    logger.debug(f"[ExpTracker] Session failed: {e}")
                    time.sleep(2.0)
            else:
                time.sleep(2.0) # Game not found, wait

    def parse_and_update_exp(self, raw_text, debug_img=None, raw_img=None, conf=0):
        try:
            # Prepare data
            data = {"text": "---", "value": 0, "percent": 0.0, "timestamp": time.time(), "thresh": debug_img, "conf": conf}
            
            # (Keeping bytes for cross-process compatibility if needed, but UI wants 'thresh')
            if debug_img is not None:
                try:
                    success, buffer = cv2.imencode('.png', debug_img)
                    if success: data["debug_bytes"] = buffer.tobytes()
                except: pass
            
            if raw_img is not None:
                try:
                    success, buffer = cv2.imencode('.png', raw_img)
                    if success: data["raw_bytes"] = buffer.tobytes()
                except: pass

            # Extract numeric parts (including potential decimals)
            nums = re.findall(r"\d+\.\d+|\d+", raw_text.replace(",", ""))
            if not nums: return

            # Prioritize finding the percentage (chunk with a dot)
            val_str, pct_val = "", 0.0
            dotted = [n for n in nums if "." in n]
            non_dotted = [n for n in nums if "." not in n]
            
            if dotted:
                raw_pct_str = dotted[0]
                if float(raw_pct_str) > 100 and len(raw_pct_str.split(".")[0]) > 2:
                    # Joined case: 19554905.60
                    dot_idx = raw_pct_str.find(".")
                    val_str = raw_pct_str[:dot_idx-2]
                    pct_val = float(raw_pct_str[dot_idx-2:])
                else:
                    pct_val = float(raw_pct_str)
                    val_str = max(non_dotted, key=len) if non_dotted else "0"
            elif len(nums) >= 2:
                # No dots, assume last chunk is percentage
                val_str = max(nums[:-1], key=len)
                pct_val = float(nums[-1])
            else:
                val_str = nums[0]
                pct_val = 0.0

            # Handle Noise: if PCT is like 105.60 because of a bracket '[' seen as '1'
            if pct_val > 100:
                s_pct = f"{pct_val:.2f}"
                if s_pct.startswith("1"):
                    pct_val = float(s_pct[1:])
            
            if not val_str: return
            # Ensure val_str is pure digits for int conversion
            val_str = "".join(filter(str.isdigit, val_str))
            if not val_str: return
            
            val = int(val_str)
            if self.show_debug:
                logger.debug(f"[ExpTracker] Parse OK: {val:,} [{pct_val:.2f}%]")
            
            data["text"] = f"{val:,} [{pct_val:.2f}%]"
            data["value"] = val
            data["percent"] = pct_val
            
            # Confidence Check for Update (Still using 85% but it's now visual stability)
            if data["conf"] >= 85 or self.exp_initial_val is None:
                if not sip.isdeleted(self): self.exp_update_request.emit(data)
        except Exception as e:
            if self.show_debug: 
                logger.warning(f"[ExpTracker] Parse Error: {e} | Raw: {raw_text}")


    def on_exp_update(self, data):
        # Initialize session variables if needed
        if not hasattr(self, 'exp_session_start_time'): self.exp_session_start_time = None
        if not hasattr(self, 'exp_initial_val') or self.exp_initial_val is None: 
            self.exp_initial_val = data["value"]
            logger.info(f"[ExpTracker] Initial baseline: {self.exp_initial_val:,}")

        # Trigger session start on first actual EXP gain
        if self.exp_session_start_time is None and data["value"] > self.exp_initial_val:
            self.exp_session_start_time = data["timestamp"]
            logger.info(f"[ExpTracker] Session triggered! First gain detected.")

        now = data["timestamp"]
        current_exp = data["value"]
        current_pct = data.get("percent", 0.0)
        
        # 0. Level Up Detection
        is_lv_up = False
        if hasattr(self, 'last_exp_val') and self.last_exp_val is not None:
            # If current EXP and PCT dropped significantly, it's a level up
            if current_exp < self.last_exp_val and current_pct < (self.last_exp_pct - 10):
                is_lv_up = True
                logger.info(f"[ExpTracker] Level Up detected! ({self.last_exp_pct:.2f}% -> {current_pct:.2f}%)")
        
        # 0.5 Outlier / Spike Detection (USER REQUEST)
        # If current value < last value, we might be seeing the resolution of a previous OCR spike.
        if not is_lv_up and hasattr(self, 'last_exp_val') and self.last_exp_val is not None:
            if current_exp < self.last_exp_val:
                # If current_exp is still >= the one before the last one, the last one was a spike!
                if len(self.exp_history) >= 2:
                    second_last_val = self.exp_history[-2][1]
                    if current_exp >= second_last_val:
                        fake_gain = self.last_exp_val - second_last_val
                        logger.info(f"[ExpTracker] Correcting history: Removed spike (+{fake_gain:,}). New base: {current_exp:,}")
                        self.exp_history.pop() # Remove spike from history
                        self.cumulative_gain = max(0, self.cumulative_gain - fake_gain)
                        # Corrected: Treat the one before spike as the new comparative base
                        self.last_exp_val = second_last_val 
                    else:
                        # Massive drop that isn't a level up and doesn't match history -> Malformed
                        logger.warning(f"[ExpTracker] Malformed drop ignored: {self.last_exp_val:,} -> {current_exp:,}")
                        return
                else:
                    return # Not enough history to validate
        
        if is_lv_up:
            # On level up, we treat the new value as the new baseline
            self.needs_calibration = True 
            self.exp_history = [] 
        
        # 1. Update baseline history
        self.exp_history.append((now, current_exp, current_pct))
        self.exp_history = [h for h in self.exp_history if h[0] >= now - 3600]
        
        # Accumulate gains ONLY when active
        if not self.needs_calibration and hasattr(self, 'last_exp_val') and self.last_exp_val is not None:
            v_diff = current_exp - self.last_exp_val
            if v_diff > 0: 
                self.cumulative_gain += v_diff
                self.cumulative_pct += (current_pct - self.last_exp_pct)
        
        # Update last state
        self.needs_calibration = False
        self.last_exp_val = current_exp
        self.last_exp_pct = current_pct
        
        # Calculate rates using sliding window
        ten_min_ago = now - 600
        history_ten = [h for h in self.exp_history if h[0] >= ten_min_ago]
        
        gain_10m = 0
        gain_pct_10m = 0.0
        if len(history_ten) >= 2:
            time_diff = history_ten[-1][0] - history_ten[0][0]
            val_diff = int(history_ten[-1][1]) - int(history_ten[0][1])
            pct_diff = float(history_ten[-1][2]) - float(history_ten[0][2])
            
            if time_diff > 5:
                gain_10m = int((val_diff / time_diff) * 600)
                gain_pct_10m = (pct_diff / time_diff) * 600
        
        self.ten_min_gain = max(0, gain_10m)
        
        # 2. Update UI Metadata using sliding window results
        self.current_exp_data["text"] = data["text"]
        self.current_exp_data["value"] = current_exp
        self.current_exp_data["percent"] = data["percent"]
        self.current_exp_data["gained_10m"] = self.ten_min_gain
        self.current_exp_data["percent_10m"] = max(0.0, gain_pct_10m)
        
        # Track session peak 10m efficiency
        if self.ten_min_gain > self.max_10m_exp:
             self.max_10m_exp = self.ten_min_gain
        
        # Update Rate History for Graph (Sampling every 5 second)
        if not hasattr(self, 'last_graph_sample_time'): self.last_graph_sample_time = 0
        if now - self.last_graph_sample_time >= 5:
            # Store PER MINUTE gain (10m gain / 10)
            self.exp_rate_history.append(self.ten_min_gain / 10.0)
            if len(self.exp_rate_history) > 40: self.exp_rate_history.pop(0)
            self.last_graph_sample_time = now
        
        # 3. Level up estimation (New High-Precision Algorithm)
        # 3a. Estimate total EXP pool for this level
        cur_p = data.get("percent", 0.0)
        total_pool = 0
        if cur_p > 0.1: # Only estimate if we have some progress
            total_pool = (current_exp * 100.0) / cur_p
        
        # 3b. Calculate remaining EXP and time
        self.current_exp_data["time_to_level"] = -1
        if gain_10m > 0 and total_pool > current_exp:
            rem_val = total_pool - current_exp
            # Rate per second = gain_10m / 600
            # Seconds = rem_val / (gain_10m / 600)
            self.current_exp_data["time_to_level"] = int(rem_val * 600 / gain_10m)
        
        rt = (now - self.exp_session_start_time) if self.exp_session_start_time else 0
        self.current_exp_data["is_estimated"] = rt < 600
        self.current_exp_data["tracking_duration"] = int(rt)

        self.update()

    def on_money_update(self, total_val):
        now = time.time()
        
        # Initialize baseline on first read
        if not hasattr(self, 'money_initial_val') or self.money_initial_val is None:
            self.money_initial_val = total_val
            self.last_total_money = total_val
            logger.info(f"[MoneyTracker] Initial money baseline: {total_val:,}")
            return

        # Calculate Gain (Delta)
        gain = total_val - self.last_total_money
        
        # Filter: Only record positive gains, ignore drops (spending)
        if gain > 0:
            # SANITY CHECK: Ignore insane spikes (e.g. OCR error reading 100M+ jump in 1s)
            if gain < 50_000_000: 
                self.cumulative_money += gain
                self.money_history.append((now, gain))
        elif gain < -100:
            # User likely spent money. Re-align base to avoid counting the whole wealth after next gain.
            logger.info(f"[MoneyTracker] Money dropped ({gain:,}), re-aligning baseline.")
        
        self.last_total_money = total_val
        
        # Standard logic for 10m math
        self.money_history = [h for h in self.money_history if h[0] >= now - 3600]
        
        # Standardized 10m efficiency (Projected)
        ten_min_ago = now - 600
        history_ten = [h for h in self.money_history if h[0] >= ten_min_ago]
        
        m_gain_10m = 0
        if history_ten:
            total_in_window = sum(h[1] for h in history_ten)
            session_start = getattr(self, "exp_session_start_time", history_ten[0][0])
            if session_start is None: session_start = history_ten[0][0]
            start_t = max(ten_min_ago, session_start)
            time_diff = now - start_t
            if time_diff > 5:
                m_gain_10m = int((total_in_window / time_diff) * 600)
            else:
                m_gain_10m = total_in_window

        self.current_exp_data["money_10m"] = max(0, m_gain_10m)
        self.update()


    def load_profile_immediately(self):
        self.clear_all_timers(show_msg=False)
        config = ConfigManager.load_config()
        active = config.get("active_profile", "F1")
        nickname = config["profiles"].get(active, {}).get("name", active)
        self.active_profile_name = active
        self.x_offset, self.y_offset = config.get("offset", [0, 0])
        self.exp_x_offset, self.exp_y_offset = config.get("exp_offset", [0, 0])
        self.rjpq_x_offset, self.rjpq_y_offset = config.get("rjpq_offset", [-400, 0])
        self.show_notification(f"切換至 {active}: {nickname}"); self.update()

    def show_notification(self, text):
        # Internal Overlay Animation
        try:
            if sip.isdeleted(self): return
            self.msg_text = text; self.msg_opacity = 255
            if hasattr(self, 'fade_timer'):
                try:
                    if self.fade_timer.isActive(): self.fade_timer.stop()
                except (RuntimeError, AttributeError): pass
            self.fade_timer = QTimer(self); self.fade_timer.timeout.connect(self.step_fade)
            QTimer.singleShot(3000, lambda: self.fade_timer.start(16)); self.update()
        except Exception as e:
            logger.debug(f"[Overlay] Notification Error: {e}")

    def step_fade(self):
        if self.msg_opacity > 0: self.msg_opacity = max(0, self.msg_opacity - 5); self.update()
        else: self.fade_timer.stop()

    def paintEvent(self, event):
        painter = QPainter(self); painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        
        # 0. Draw Debug Coin Location
        if self.show_debug and hasattr(self, 'last_coin_pos') and self.last_coin_pos:
            try:
                for name in ["MapleStory Worlds-Artale (繁體中文版)", "MapleStory Worlds-Artale"]:
                    hwnd = win32gui.FindWindow(None, name)
                    if hwnd:
                        cx, cy, cw, ch = self.last_coin_pos
                        # Convert client coords to global screen coords (Physical)
                        screen_pt = win32gui.ClientToScreen(hwnd, (cx, cy))
                        
                        # High DPI Correct:
                        dpr = self.screen().devicePixelRatio()
                        # Map physical global to logical global, then subtract virtual desktop top-left
                        v_rect = QApplication.primaryScreen().virtualGeometry()
                        px = (screen_pt[0] / dpr) - v_rect.x()
                        py = (screen_pt[1] / dpr) - v_rect.y()
                        lw, lh = cw / dpr, ch / dpr
                        
                        painter.setPen(QPen(QColor(255, 255, 0, 200), 2))
                        painter.drawRect(int(px), int(py), int(lw), int(lh))
                        painter.setPen(QColor(255, 255, 0))
                        painter.drawText(int(px), int(py - 5), "🪙 Coin")
                        
                        # Draw Secondary Box (Cyan for distinction)
                        if hasattr(self, 'last_coin_info_pos') and self.last_coin_info_pos:
                            ix, iy, iw, ih = self.last_coin_info_pos
                            # Convert to global (same as above)
                            info_pt = win32gui.ClientToScreen(hwnd, (ix, iy))
                            ipx = (info_pt[0] / dpr) - v_rect.x()
                            ipy = (info_pt[1] / dpr) - v_rect.y()
                            liw, lih = iw / dpr, ih / dpr
                            
                            painter.setPen(QPen(QColor(0, 255, 255, 180), 2)) # Cyan box
                            painter.drawRect(int(ipx), int(ipy), int(liw), int(lih))
                            if hasattr(self, 'last_coin_ocr') and self.last_coin_ocr:
                                painter.setPen(QColor(0, 255, 255))
                                painter.drawText(int(ipx), int(ipy - 5), f"Value: {self.last_coin_ocr}")
                        break
            except: pass

        # 0. Draw EXP Statistics Panel (Top Left)
        if self.show_exp_panel:
            self.draw_exp_panel(painter)
            
        if getattr(self, "show_rjpq_panel", False):
            try:
                # Multi-profile check: Use current session offsets if not reloaded yet
                pw, ph = 180, 320
                # Anchor (ax, ay) is Top-Right
                ax = self.rect().center().x() + self.rjpq_x_offset
                ay = self.rect().center().y() + self.rjpq_y_offset
                
                # Fetch data from RJPQ client if available, else use local cache
                data = getattr(self, "rjpq_data", [4]*40)
                sel_color = getattr(self, "selected_color", -1)
                
                from rjpq_tool import draw_rjpq_panel
                # draw_rjpq_panel uses Top-Left as start, so start_x = ax - pw
                draw_rjpq_panel(painter, ax - pw, ay, pw, ph, self.base_opacity, data, sel_color)
            except Exception as e:
                logger.debug(f"[Overlay] RJPQ Panel draw failed: {e}")
            
            # --- Update Click Zones ---
            self.rjpq_click_zones = {}
            start_x = ax - pw + 35
            start_y = ay + 45
            cell_w, cell_h = 32, 22
            for row in range(10):
                for col in range(4):
                    idx = row * 4 + col
                    cx = start_x + col * 35
                    cy = start_y + row * 25
                    local_rect = QRect(int(cx), int(cy), cell_w, cell_h)
                    # Convert to GLOBAL coordinates for main.py listener
                    global_topleft = self.mapToGlobal(local_rect.topLeft())
                    self.rjpq_click_zones[idx] = QRect(global_topleft, local_rect.size())
        
        # 0.1 Draw Debug Crop Box (Red Outline)
        if self.show_debug and self.last_crop_info:
            try:
                import win32gui
                target_hwnd = None
                for name in ["MapleStory Worlds-Artale (繁體中文版)", "MapleStory Worlds-Artale"]:
                    hwnd = win32gui.FindWindow(None, name)
                    if hwnd and win32gui.IsWindowVisible(hwnd):
                        target_hwnd = hwnd; break
                
                if target_hwnd:
                    # 1. Get Game Client Area size
                    crect = win32gui.GetClientRect(target_hwnd)
                    client_w, client_h = crect[2], crect[3]
                    
                    # 2. Get Global Screen coord of Client BOTTOM-LEFT (Physical)
                    bl_point = win32gui.ClientToScreen(target_hwnd, (0, client_h))
                    
                    # 3. DPI Scaled Map to Overlay Local coordinates
                    dpr = self.screen().devicePixelRatio()
                    logical_gl_pt = QPoint(int(bl_point[0]/dpr), int(bl_point[1]/dpr))
                    local_bl = self.mapFromGlobal(logical_gl_pt)
                    bx, by = local_bl.x(), local_bl.y()
                    
                    # Logical Dimensions
                    lbw, lbh = client_w / dpr, client_h / dpr
                    
                    # Sync using Min Ratio logic
                    visual_scale = min(client_w / self.BASE_W, client_h / self.BASE_H)
                    
                    # A. EXP Zone (Calculated in Logical Units)
                    tx = bx + int(self.X_OFF_FROM_LEFT * visual_scale / dpr)
                    ty = by - int(self.Y_OFF_FROM_BOTTOM * visual_scale / dpr)
                    tw, th = int(self.BASE_CW * visual_scale / dpr), int(self.BASE_CH * visual_scale / dpr)
                    
                    painter.setPen(QPen(QColor(255, 0, 0, 200), 2, Qt.PenStyle.DashLine))
                    painter.setBrush(QColor(255, 0, 0, 40))
                    painter.drawRect(int(tx), int(ty), int(tw), int(th))
                    painter.setPen(QPen(QColor(255, 0, 0)))
                    painter.drawText(int(tx), int(ty - 5), "EXP Zone")

                    # B. LV Zone
                    lvx = bx + int(self.LV_X_OFF_FROM_LEFT * visual_scale / dpr)
                    lvy = by - int(self.LV_Y_OFF_FROM_BOTTOM * visual_scale / dpr)
                    lcw, lch = int(self.LV_BASE_CW * visual_scale / dpr), int(self.LV_BASE_CH * visual_scale / dpr)
                    
                    painter.setPen(QPen(QColor(255, 165, 0, 200), 2, Qt.PenStyle.DashLine))
                    painter.setBrush(QColor(255, 165, 0, 40))
                    painter.drawRect(int(lvx), int(lvy), int(lcw), int(lch))
                    painter.setPen(QPen(QColor(255, 165, 0)))
                    painter.drawText(int(lvx), int(lvy - 5), "LV Zone")
            except: pass

        # Guard: Stop painting if idle and no debug info
        if not self.is_active and not self.show_preview and self.msg_opacity == 0 and not self.show_debug: 
            return
        
        # Base coordinates
        base_x = self.rect().center().x() + self.x_offset
        base_y = self.rect().center().y() + self.y_offset

        # 1. Profile/Action Notification (Centered above anchor)
        if self.msg_opacity > 0:
            font = QFont()
            font.setFamilies(["Microsoft JhengHei", "微軟正黑體"])
            font.setPointSize(18)
            font.setBold(True)
            painter.setFont(font)
            tw = painter.fontMetrics().horizontalAdvance(self.msg_text)
            # Draw notification right-aligned clearly above the timer block
            bg_rect = QRect(base_x - (tw+40), base_y - 70, tw+40, 45)
            # Notification background affected by both fade and settings
            bg_alpha = int(min(200, self.msg_opacity) * (self.base_opacity / 1.0))
            painter.setBrush(QColor(0, 0, 0, bg_alpha))
            painter.setPen(Qt.PenStyle.NoPen); painter.drawRoundedRect(bg_rect, 8, 8)
            color = QColor(255, 100, 100, self.msg_opacity) if "F12" in self.msg_text else QColor(255, 215, 0, self.msg_opacity)
            painter.setPen(color); painter.drawText(bg_rect, Qt.AlignmentFlag.AlignCenter, self.msg_text)

        if not self.timer_manager.active_timers and not self.show_preview: return

        timers_to_draw = []
        if self.timer_manager.active_timers:
            sorted_active = sorted(self.timer_manager.active_timers.items(), key=lambda x: x[1]["seconds"], reverse=True)
            for k, d in sorted_active: timers_to_draw.append((k, d["seconds"], d["pixmap"]))
        
        # Also include world_timers if any are active
        for k, d in self.world_timers.items():
            timers_to_draw.append((k, d["seconds"], d["pixmap"]))

        if not timers_to_draw and self.show_preview:
            timers_to_draw.append(("preview", 300, QPixmap(resource_path("buff_pngs/arrow.png"))))

        new_click_zones = {}; spacing = 56; total_width = len(timers_to_draw) * spacing
        # Right Aligned: The anchor base_x is the RIGHT edge of the timer group
        block_start_x = base_x - total_width
        block_center_y = base_y + 60
        
        for idx, (key, seconds, pixmap) in enumerate(timers_to_draw):
            x_pos = block_start_x + idx * spacing + (spacing // 2); block_center = QPoint(x_pos, block_center_y)
            icon_size = 40; icon_rect = QRect(block_center.x() - 20, block_center.y() - 45, 40, 40)
            text_rect = QRect(block_center.x() - 50, block_center.y() - 13, 100, 50)
            
            # Click zone: Use icon zone if exists, otherwise use text zone
            if key != "preview":
                click_rect = icon_rect if pixmap else text_rect
                new_click_zones[key] = QRect(self.mapToGlobal(click_rect.topLeft()), click_rect.size())
            
            if pixmap:
                if self.icon_frame: painter.drawPixmap(icon_rect.adjusted(-2, -2, 2, 2), self.icon_frame)
                painter.drawPixmap(icon_rect, pixmap.scaled(40, 40, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation))
            display_seconds = max(0, seconds); text = str(display_seconds)
            color = QColor(100, 255, 100) if seconds > 30 else QColor(255, 50, 50)
            if self.show_preview and not self.timer_manager.active_timers: color = QColor(255, 255, 255, 150)
            font = QFont()
            font.setFamilies(["Microsoft JhengHei", "微軟正黑體"])
            font.setPointSize(22 if seconds > 3 else 26)
            font.setBold(True)
            painter.setFont(font)
            text_rect = QRect(block_center.x() - 50, block_center.y() - 13, 100, 50)
            painter.setPen(QPen(QColor(0,0,0,200), 4)); painter.drawText(text_rect, Qt.AlignmentFlag.AlignCenter, text)
            painter.setPen(QPen(color, 2)); painter.drawText(text_rect, Qt.AlignmentFlag.AlignCenter, text)
        self.click_zones = new_click_zones


    def export_exp_report(self):
        """ Render the EXP panel to a static image and save it """
        pw, ph = 330, 220 # Slightly taller for report
        pixmap = QPixmap(pw, ph)
        pixmap.fill(Qt.GlobalColor.transparent)
        
        painter = QPainter(pixmap)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        
        # 1. Background
        rect = QRect(0, 0, pw, ph)
        path = QPainterPath()
        path.addRoundedRect(QRectF(rect).adjusted(2, 2, -2, -2), 15, 15)
        painter.setPen(QPen(QColor(255, 215, 0), 2))
        painter.setBrush(QColor(10, 10, 15, 240)) # More opaque for report
        painter.drawPath(path)
        
        # Add a more visible watermark
        painter.setPen(QPen(QColor(255, 255, 255, 80))) # Increased opacity
        font = QFont("Microsoft JhengHei", 9)
        font.setItalic(True)
        painter.setFont(font)
        painter.drawText(rect.adjusted(0, 0, -15, -10), Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignBottom, "使用 Artale 瑞士刀記錄")

        # Reuse drawing logic but relative to (0,0)
        self._draw_exp_content(painter, 0, 0, pw, ph, is_export=True)
        painter.end()
        
        # Save to Pictures folder
        filename = f"Artale瑞士刀_{int(time.time())}.png"
        pictures_dir = QStandardPaths.writableLocation(QStandardPaths.StandardLocation.PicturesLocation)
        save_path = os.path.join(pictures_dir, filename)
        
        if pixmap.save(save_path, "PNG"):
            # Copy to clipboard
            from PyQt6.QtWidgets import QApplication
            QApplication.clipboard().setPixmap(pixmap)
            
            logger.info(f"[ExpTracker] Report exported to {save_path} and copied to clipboard")
            self.show_notification(f"✅ 成果圖已儲存並複製到剪貼簿！")
            # Try to open the file
            try: subprocess.Popen(f'explorer /select,"{save_path}"')
            except: pass
        else:
            self.show_notification("❌ 產出失敗，請檢查權限")

    def _draw_exp_content(self, painter, px, py, pw, ph, is_export=False):
        now = time.time()
        # 1. Title
        painter.setPen(QColor(255, 255, 255))
        font = QFont("Microsoft JhengHei")
        font.setPointSize(12 if is_export else 11); font.setBold(True)
        painter.setFont(font)
        y = py + (30 if is_export else 25)
        painter.drawText(px + 15, y, f"📊 經驗值監測報告" if is_export else "📊 經驗值監測")
        
        # 2. Sub-info (Duration & Total on one line)
        y += (28 if is_export else 25)
        duration_sec = self.current_exp_data.get("tracking_duration", 0)
        h_dur = duration_sec // 3600; m_dur = (duration_sec % 3600) // 60; s_dur = duration_sec % 60
        val_dur = f"{h_dur:02d}:{m_dur:02d}:{s_dur:02d}" if h_dur > 0 else f"{m_dur:02d}:{s_dur:02d}"
        
        # Duration Part
        painter.setPen(QColor(100, 255, 100))
        font.setPointSize(9); font.setBold(True); painter.setFont(font)
        fm_small = painter.fontMetrics()
        lbl_dur = "紀錄時長:"
        painter.drawText(px + 15, y, lbl_dur)
        painter.setPen(QColor(255, 255, 255))
        painter.drawText(px + 15 + fm_small.horizontalAdvance(lbl_dur) + 5, y, val_dur)
        
        # Total Part (Right Aligned)
        painter.setPen(QColor(100, 255, 100))
        lbl_total = "總累積:"
        # Position label based on typical width or middle
        total_start_x = px + 145
        painter.drawText(total_start_x, y, lbl_total)
        painter.setPen(QColor(255, 255, 255))
        val_total = f"+{self.cumulative_gain:,} ({self.cumulative_pct:+.2f}%)"
        painter.drawText(QRect(px + 15, y - 18, pw - 30, 24), Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter, val_total)

        # 3. Time to Level Up
        y += (32 if is_export else 28)
        painter.setPen(QColor(100, 255, 100))
        font.setPointSize(12 if is_export else 11); font.setWeight(QFont.Weight.DemiBold); painter.setFont(font)
        fm = painter.fontMetrics()
        
        lbl_ttl = "升級預計還需: "
        painter.drawText(px + 15, y, lbl_ttl)
        
        ttl_sec = self.current_exp_data.get("time_to_level", -1)
        val_ttl = f"{ttl_sec // 3600}小時 {(ttl_sec % 3600) // 60}分" if ttl_sec > 0 else "計算速率中..."
        painter.setPen(QColor(255, 255, 255))
        painter.drawText(QRect(px + 15, y - 18, pw - 30, 24), Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter, val_ttl)
        
        # 4. 10min Efficiency
        y += (32 if is_export else 28)
        gain_val = self.current_exp_data.get("gained_10m", 0)
        gain_pct = self.current_exp_data.get("percent_10m", 0.0)
        is_est = self.current_exp_data.get("is_estimated", True)
        
        lbl_eff = f"{'（預估）' if is_est else ''}10分鐘效率: "
        painter.setPen(QColor(100, 255, 100))
        painter.drawText(px + 15, y, lbl_eff)
        
        val_eff = f"+{gain_val:,} ({gain_pct:+.2f}%)"
        painter.setPen(QColor(255, 255, 255))
        painter.drawText(QRect(px + 15, y - 18, pw - 30, 24), Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter, val_eff)

        # 5. Highest 10m Record
        y += (32 if is_export else 28)
        lbl_max = "十分鐘最高經驗: "
        painter.setPen(QColor(100, 255, 100))
        painter.drawText(px + 15, y, lbl_max)
        
        session_sec = now - self.exp_session_start_time if (hasattr(self, 'exp_session_start_time') and self.exp_session_start_time) else 0
        val_max = f"{self.max_10m_exp:,}" if session_sec >= 600 else "(未滿十分鐘)"
        painter.setPen(QColor(255, 255, 255))
        painter.drawText(QRect(px + 15, y - 18, pw - 30, 24), Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter, val_max)
        
        # 6. Mesos Efficiency & Total (ONLY IF ENABLED OR EXPORTING)
        if self.show_money_log or is_export:
            # 6a. 10m Efficiency
            y += (32 if is_export else 28)
            lbl_money_10m = "十分鐘楓幣效率: "
            painter.setPen(QColor(255, 215, 0)) # Gold color
            painter.drawText(px + 15, y, lbl_money_10m)
            money_10m = self.current_exp_data.get("money_10m", 0)
            val_money_10m = f"+{money_10m:,}"
            painter.setPen(QColor(255, 255, 255))
            painter.drawText(QRect(px + 15, y - 18, pw - 30, 24), Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter, val_money_10m)
            
            # 6b. Total Gained Mesos
            y += (28 if is_export else 25)
            lbl_money_total = "累計獲取楓幣: "
            painter.setPen(QColor(255, 215, 0))
            font.setPointSize(9); painter.setFont(font)
            painter.drawText(px + 15, y, lbl_money_total)
            val_total_money = f"+{self.cumulative_money:,}"
            painter.setPen(QColor(255, 255, 255))
            painter.drawText(QRect(px + 15, y - 18, pw - 30, 24), Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter, val_total_money)
        
        last_y = y # For graph positioning

        # 5. Trend Graph
        if len(self.exp_rate_history) > 1:
            max_points = 40
            gh = 50 if is_export else 40; gw = pw - 30
            gx = px + 15; gy = last_y + (15 if is_export else 10)
            
            painter.setPen(QColor(255, 255, 255, 20))
            painter.setBrush(QColor(255, 255, 255, 5))
            painter.drawRoundedRect(gx, gy, gw, gh, 4, 4)
            
            max_v = max(self.exp_rate_history)
            if max_v <= 0: max_v = 1
            
            path = QPainterPath()
            step_x = gw / (max_points - 1)
            
            for i, v in enumerate(self.exp_rate_history):
                vx = gx + i * step_x
                vy = gy + gh - (v / max_v * (gh - 4)) - 2 # Padding inside rect
                if i == 0: path.moveTo(vx, vy)
                else: path.lineTo(vx, vy)
            
            # --- Draw Fill (Area) ---
            fill_path = QPainterPath(path)
            last_idx = len(self.exp_rate_history) - 1
            fill_path.lineTo(gx + last_idx * step_x, gy + gh)
            fill_path.lineTo(gx, gy + gh)
            fill_path.closeSubpath()
            
            # Gradient green matching the theme
            grad = QLinearGradient(gx, gy, gx, gy + gh)
            grad.setColorAt(0, QColor(100, 255, 100, 80))  # Semi-transparent green
            grad.setColorAt(1, QColor(100, 255, 100, 0))   # Fades to transparent
            painter.fillPath(fill_path, grad)
            
            # --- Draw Stroke (Line) ---
            painter.setPen(QPen(QColor(100, 255, 100), 2)) # Solid green line
            painter.setBrush(Qt.BrushStyle.NoBrush)
            painter.drawPath(path)

    def draw_exp_panel(self, painter):
        if not self.show_exp_panel:
            return
            
        # Positional logic
        bx = self.rect().center().x() + self.exp_x_offset
        by = self.rect().center().y() + self.exp_y_offset
        pw, ph = 330, 230 # Increased height for graph and money
        # Right Aligned: bx is the right edge
        panel_rect = QRect(bx - pw, by - 120 - ph // 2, pw, ph)
        px, py = panel_rect.x(), panel_rect.y()
        
        # 1. Background
        path = QPainterPath()
        path.addRoundedRect(QRectF(panel_rect), 12, 12)
        painter.setPen(QPen(QColor(255, 215, 0, 255), 2))
        painter.setBrush(QColor(10, 10, 15, int(self.base_opacity * 255))) 
        painter.drawPath(path)
        
        # Use refactored drawing logic
        self._draw_exp_content(painter, px, py, pw, ph, is_export=False)

        # 1.5 Title & Level
        # (Content moved to _draw_exp_content)
        



