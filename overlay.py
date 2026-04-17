from exp_tracker import StatsData
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
try:
    import win32gui
    import win32process
except ImportError:
    win32gui = win32api = win32con = win32process = winsound = None

import cv2
import pytesseract
from PyQt6 import sip
from PyQt6.QtWidgets import (QApplication, QWidget, QSystemTrayIcon, QMenu)
from PyQt6.QtCore import (Qt, QPoint, QRect, QTimer, pyqtSignal, QRectF, QStandardPaths)
from PyQt6.QtGui import QFont, QColor, QPainter, QPen, QPixmap, QIcon, QPainterPath, QAction, QLinearGradient

try:
    from windows_capture import WindowsCapture
except ImportError:
    WindowsCapture = None

# Local imports
from capture_engine import ArtaleCapture
from ocr_engine import ArtaleOCR
from rjpq_tool import draw_rjpq_panel
from skill_timer import TimerManager
from settings_window import SettingsWindow
from ocr_engine import ArtaleOCR

# 初始化日誌記錄器
logging.getLogger('pytesseract').setLevel(logging.WARNING)
logging.getLogger('PIL').setLevel(logging.WARNING)
logger = logging.getLogger(__name__)

# Using utils.py for VERSION, REPO_URL, resource_path, ConfigManager, EXP_TABLE
from utils import VERSION, REPO_URL, resource_path, ConfigManager, EXP_TABLE

# Tesseract Portable Setup (LOCAL ONLY)
def get_tess_cmd():
    """自動偵測 Tesseract-OCR 路徑，並處理 PyInstaller 的 DLL 載入問題"""
    import os, sys
    
    executable_path = None
    
    # 1. 檢查 PyInstaller 內部的打包路徑 (Internal Bundle Path)
    if hasattr(sys, '_MEIPASS'):
        bundle_dir = sys._MEIPASS
        executable_path = os.path.join(bundle_dir, "Tesseract-OCR", "tesseract.exe")
        if os.path.exists(executable_path):
            # 重要：為了讓 Tesseract 找到關連的 DLL，必須將其資料夾加入環境變數 PATH
            tess_dir = os.path.dirname(executable_path)
            if tess_dir not in os.environ["PATH"]:
                os.environ["PATH"] = tess_dir + os.pathsep + os.environ["PATH"]
            return executable_path

    # 2. 檢查程式所在目錄 (用於便攜版或開發環境)
    base_dir = os.path.dirname(os.path.abspath(sys.argv[0] if getattr(sys, 'frozen', False) else __file__))
    local_tess = os.path.join(base_dir, "Tesseract-OCR", "tesseract.exe")
    if os.path.exists(local_tess):
        tess_dir = os.path.dirname(local_tess)
        if tess_dir not in os.environ["PATH"]:
            os.environ["PATH"] = tess_dir + os.pathsep + os.environ["PATH"]
        return local_tess
    
    # 3. 最後的備選路徑 (標準安裝路徑)
    common_p = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
    if os.path.exists(common_p):
        return common_p
    
    return None

if pytesseract:
    pytesseract.pytesseract.tesseract_cmd = get_tess_cmd()

class ArtaleOverlay(QWidget):
    # 1080p 校準參考 (以左下角為基準錨點)
    BASE_W, BASE_H = 1920, 1080
    X_OFF_FROM_LEFT = 1084   # 固定水平偏移量 (距離左側)
    Y_OFF_FROM_BOTTOM = 66   # 固定垂直偏移量 (距離底部)
    BASE_CW, BASE_CH = 240, 22
    
    LV_X_OFF_FROM_LEFT = 96
    LV_Y_OFF_FROM_BOTTOM = 46
    LV_BASE_CW, LV_BASE_CH = 75, 26
    
    timer_request = pyqtSignal(str, int, str, bool) 
    clear_request = pyqtSignal()
    notification_request = pyqtSignal(str)
    profile_switch_request = pyqtSignal()
    exp_update_request = pyqtSignal(dict)    # 用於統計數據
    exp_visual_request = pyqtSignal(dict)    # 用於除錯影像: exp, lv, coin
    lv_update_request = pyqtSignal(dict)
    toggle_exp_request = pyqtSignal()
    toggle_pause_request = pyqtSignal()
    toggle_rjpq_request = pyqtSignal()
    settings_show_request = pyqtSignal()
    rjpq_cell_clicked = pyqtSignal(int)
    export_report_request = pyqtSignal()
    update_found = pyqtSignal(str, str) # version, download_url
    money_update_request = pyqtSignal(int)
    stats_updated = pyqtSignal(dict) # 新增：接收來自 Tracker 的完整統計數據
    request_show_settings_signal = pyqtSignal()
    
    def __init__(self, target_window_title="MapleStory Worlds-Artale (繁體中文版)"):
        super().__init__()
        self.target_window_title = target_window_title
        self.timer_manager = TimerManager(self)
        self.timer_manager.updated.connect(self.update)
        self.click_zones = {}  # 點擊偵測區域
        self.is_active = False # 用於計時器的相容性
        self.show_preview = False
        self.active_profile_name = "F1"
        self._is_running = True
        
        # 提前載入設定
        config = ConfigManager.load_config()
        self.show_exp_panel = config.get("show_exp", False)
        self.show_money_log = config.get("show_money_log", True)
        self.exp_paused = False # 經驗值暫停狀態
        self.total_pause_time = 0 # 累積暫停時長
        self.pause_start_time = 0
        self.needs_calibration = False # 恢復後校準旗標
        self.show_rjpq_panel = False # 羅茱面板預設關閉
        self.show_debug = config.get("show_debug", False)
        self.base_opacity = config.get("opacity", 0.5)
        
        self.msg_text = ""; self.msg_opacity = 0
        self.x_offset = 0; self.y_offset = 0
        self.exp_x_offset = 0; self.exp_y_offset = 0
        
        # UI 顯示用的實時數據包
        self.current_exp_data = {
            "text": "---", "value": 0, "percent": 0.0, 
            "gained_10m": 0, "percent_10m": 0.0, "tracking_duration": 0
        }
        self.exp_rate_history = [] 
        self.money_rate_history = []
        self.cumulative_gain = 0
        self.cumulative_pct = 0.0
        self.cumulative_money = 0
        self.max_10m_exp = 0
        
        self.selected_color = -1
        self.rjpq_data = [4] * 40
        self.rjpq_x_offset = -400
        self.rjpq_y_offset = 0
        self.rjpq_click_zones = {}
        self.current_lv = None 
        self.last_confirmed_lv = None # 用於升級偵測
        self.last_crop_info = None
        
        self.timer_request.connect(self.timer_manager.start_timer)
        self.clear_request.connect(self.timer_manager.clear_all)
        self.notification_request.connect(self.show_notification)
        self.toggle_exp_request.connect(self.on_toggle_exp)
        self.toggle_pause_request.connect(self.on_toggle_pause)
        self.toggle_rjpq_request.connect(self.on_toggle_rjpq)
        self.stats_updated.connect(self.on_stats_updated)
        
        # Instantiate SettingsWindow (now from separate module)
        self.settings_window = SettingsWindow(self)
        self.settings_show_request.connect(self.settings_window.request_show.emit)
        self.settings_window.timer_request.connect(self.timer_manager.start_timer)
        self.settings_window.notification_request.connect(self.show_notification)
        
        # 明確地將追蹤更新連結到設定視窗 (使用 QueuedConnection 確保執行緒安全)
        self.exp_visual_request.connect(self.settings_window.update_debug_img, Qt.ConnectionType.QueuedConnection)
        self.lv_update_request.connect(self.settings_window.update_lv_debug_img, Qt.ConnectionType.QueuedConnection)
        self.update_found.connect(self.settings_window.show_update_banner, Qt.ConnectionType.QueuedConnection)
        
        # 控制器將從外部賦予，用以協調模組運行
        self.controller = None
        
        self.tracking_timer = QTimer(self); self.tracking_timer.timeout.connect(self.sync_with_game_window); self.tracking_timer.start(1000)
        self.world_timers = {} 
        
        frame_p = resource_path("buff_pngs/skill_frame.png")
        self.icon_frame = QPixmap(frame_p) if os.path.exists(frame_p) else None
        self.last_coin_pos = None # (x, y, w, h) in client coords
        self.last_coin_info_pos = None
        self.last_coin_ocr = ""
        self.init_tray()
        
        # 載入楓幣模板用於影像比對
        self.coin_tpl = None
        if os.path.exists("coin.png"):
            self.coin_tpl = cv2.imread("coin.png")
            if self.coin_tpl is not None:
                logger.info(f"[ExpTracker] Loaded coin template: {self.coin_tpl.shape}")
        
        self.init_ui()

    def init_tray(self):
        """初始化系統匣圖示與右鍵選單"""
        self.tray_icon = QSystemTrayIcon(self)
        
        # 使用生成的精美應用程式圖示
        icon_path = resource_path("app_icon.png")
        if os.path.exists(icon_path):
            self.tray_icon.setIcon(QIcon(icon_path))
        
        # 建立右鍵選單
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
        
        # 點擊圖示可切換設定視窗
        self.tray_icon.activated.connect(self.on_tray_activated)
        self.tray_icon.show()

    def on_tray_activated(self, reason):
        if reason == QSystemTrayIcon.ActivationReason.Trigger:
            self.request_show_settings()

    def request_show_settings(self):
        self.settings_show_request.emit()

    def reset_exp_stats(self, silent=False):
        """重置經驗值追蹤基準點"""
        self.cumulative_gain = 0
        self.cumulative_pct = 0.0
        self.max_10m_exp = 0
        if not silent:
            self.show_notification("📊 經驗值統計已重置")

    def on_update_found(self, tag, url):
        self._latest_version_info = (tag, url)
        pass


    def init_ui(self):
        self.setWindowFlags(Qt.WindowType.WindowStaysOnTopHint | Qt.WindowType.FramelessWindowHint | Qt.WindowType.WindowTransparentForInput | Qt.WindowType.Tool)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.setAttribute(Qt.WidgetAttribute.WA_AlwaysStackOnTop)
        
        # 確保覆蓋虛擬桌面上的所有顯示器
        v_rect = QApplication.primaryScreen().virtualGeometry()
        self.setGeometry(v_rect)
        
        # 移動到虛擬桌面的絕對左上角 (處理負座標情況)
        self.move(v_rect.topLeft())
        
        # Windows 組合渲染小技巧：0.99 透明度可強制執行全室渲染
        self.setWindowOpacity(0.99)
        
        logger.debug(f"[Debug] Overlay spans: {v_rect.x()}, {v_rect.y()} to {v_rect.width()}, {v_rect.height()}")
        self.show()

    def play_sound(self, times=1):
        self.timer_manager.play_sound(times)

    def sync_with_game_window(self):
        # 我們不再移動 Overlay 視窗，它會一直固定在虛擬桌面的全螢幕位置。
        # 這能保持 UI 穩定，不隨遊戲視窗移動而產生抖動。
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
            # Error 2, 18 or 1400 are common during window transitions, log only other errors
            err_str = str(e)
            if not any(code in err_str for code in ["(2,", "(18,", "(1400,"]):
                logger.debug(f"[Overlay] Window search failed: {e}")
            self.game_hwnd = None
        
        if hwnd:
            try:
                # 內部更新遊戲錨點 (bx, by)，
                # 讓開發者模式或除錯視窗能知道遊戲的實體位置。
                rect = win32gui.GetClientRect(hwnd)
                client_h = rect[3]
                bl_point = win32gui.ClientToScreen(hwnd, (0, client_h))
                
                # DPI 縮放修正：
                # win32gui 返回實體像素，而 Qt 使用邏輯像素。
                dpr = self.screen().devicePixelRatio()
                logical_gl_pt = QPoint(int(bl_point[0] / dpr), int(bl_point[1] / dpr))
                
                local_bl = self.mapFromGlobal(logical_gl_pt)
                self.bx, self.by = local_bl.x(), local_bl.y()
            except Exception as e:
                logger.debug(f"[Overlay] ClientToScreen mapping failed: {e}")
            
        # 確保 Overlay 是顯示狀態且位於頂層，但不改變其幾何佈局
        if not self.isVisible(): self.show()
        self.raise_()

    def update_offset(self, gx, gy):
        local = self.mapFromGlobal(QPoint(gx, gy))
        self.x_offset = local.x() - self.rect().center().x()
        self.y_offset = local.y() - self.rect().center().y()
        self.click_zones = {}; self.update()

    def update_exp_offset(self, gx, gy):
        local = self.mapFromGlobal(QPoint(gx, gy))
        # 基於 draw_exp_panel 中的計算邏輯：
        # bx = center_x + exp_x_offset
        # 面板右上角 Y = (center_y + exp_y_offset) - 120 - ph // 2
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

    def on_stats_updated(self, stats: StatsData):
        """接收來自 ExpTracker 的完整數據包並同步至 UI 顯示層"""
        self.current_exp_data = stats
        self.cumulative_gain = stats.get("cumulative_gain", 0)
        self.cumulative_pct = stats.get("cumulative_pct", 0.0)
        self.cumulative_money = stats.get("cumulative_money", 0)
        self.max_10m_exp = stats.get("max_10m_exp", 0)
        self.exp_rate_history = stats.get("exp_rate_history", [])
        self.money_rate_history = stats.get("money_rate_history", [])
        self.update()

    def on_toggle_exp(self):
        self.show_exp_panel = not self.show_exp_panel
        status = "已啟用" if self.show_exp_panel else "已關閉"
        
        # Modular toggle logic via controller
        if self.controller:
            self.controller.toggle_tracking(self.show_exp_panel)
            
        self.show_notification(f"📊 經驗追蹤面板 {status} (F10)")
        self.update()

    def on_toggle_pause(self):
        self.exp_paused = not self.exp_paused
        status = "已暫停" if self.exp_paused else "已恢復"
        
        if self.controller and self.controller.tracker:
            self.controller.tracker.toggle_pause()
            
        self.show_notification(f"📊 經驗追蹤 {status} (F11)")
        self.update()

    def on_toggle_rjpq(self):
        self.show_rjpq_panel = not self.show_rjpq_panel
        self.show_notification(f"羅茱路徑面板: {'已啟用' if self.show_rjpq_panel else '已關閉'}")
        self.update()

    def update_rjpq_data(self, data):
        """更新遠端同步的路徑狀態"""
        self.rjpq_data = data
        self.update()
        
    def set_rjpq_color(self, color):
        self.selected_color = color
        self.update()
    def set_rjpq_overlay_visible(self, visible):
        self.show_rjpq_panel = visible
        self.update()

    def closeEvent(self, event):
        self._is_running = False
        super().closeEvent(event)

    # --- UI 輔助邏輯 ---


    def apply_profile_config(self, active, nickname, offsets):
        """僅更新 UI 層級的配置狀態"""
        self.active_profile_name = active
        self.x_offset, self.y_offset = offsets.get("offset", [0, 0])
        self.exp_x_offset, self.exp_y_offset = offsets.get("exp_offset", [0, 0])
        self.rjpq_x_offset, self.rjpq_y_offset = offsets.get("rjpq_offset", [-400, 0])
        self.show_notification(f"切換至 {active}: {nickname}")
        self.update()

    def show_notification(self, text):
        """內部 Overlay 淡入淡出通知動畫"""
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
        
        # 1. 繪製經驗值統計面板 (左上角)
        if self.show_exp_panel:
            self.draw_exp_panel(painter)
            
        if getattr(self, "show_rjpq_panel", False):
            try:
                # 配置檢查：如果尚未重新載入，使用目前的 Session 偏移量
                pw, ph = 180, 320
                # 錨點 (ax, ay) 為右上角
                ax = self.rect().center().x() + self.rjpq_x_offset
                ay = self.rect().center().y() + self.rjpq_y_offset
                
                # 若可用，從 RJPQ 客戶端獲取數據，否則使用本地快取
                data = getattr(self, "rjpq_data", [4]*40)
                sel_color = getattr(self, "selected_color", -1)
                
                from rjpq_tool import draw_rjpq_panel
                # draw_rjpq_panel 使用左上角作為起點，因此 start_x = ax - pw
                draw_rjpq_panel(painter, ax - pw, ay, pw, ph, self.base_opacity, data, sel_color)
            except Exception as e:
                logger.debug(f"[Overlay] RJPQ Panel draw failed: {e}")
            
            # --- 更新點擊區域 ---
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
                    # 轉換為全域座標供 main.py 監聽器使用
                    global_topleft = self.mapToGlobal(local_rect.topLeft())
                    self.rjpq_click_zones[idx] = QRect(global_topleft, local_rect.size())
        
        # 0.1 繪製除錯截取框 (紅色外框)
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

        # Base coordinates
        base_x = self.rect().center().x() + self.x_offset
        base_y = self.rect().center().y() + self.y_offset

        # 1. 配置/操作通知 (置中顯示於錨點上方)
        if self.msg_opacity > 0:
            font = QFont()
            font.setFamilies(["Microsoft JhengHei", "微軟正黑體"])
            font.setPointSize(18)
            font.setBold(True)
            painter.setFont(font)
            tw = painter.fontMetrics().horizontalAdvance(self.msg_text)
            # 在計時器方塊上方清晰地靠右對齊繪製通知
            bg_rect = QRect(base_x - (tw+40), base_y - 70, tw+40, 45)
            # 通知背景受淡出效果與設定透明度共同影響
            bg_alpha = int(min(200, self.msg_opacity) * (self.base_opacity / 1.0))
            painter.setBrush(QColor(0, 0, 0, bg_alpha))
            painter.setPen(Qt.PenStyle.NoPen); painter.drawRoundedRect(bg_rect, 8, 8)
            color = QColor(255, 100, 100, self.msg_opacity) if "F12" in self.msg_text else QColor(255, 215, 0, self.msg_opacity)
            painter.setPen(color); painter.drawText(bg_rect, Qt.AlignmentFlag.AlignCenter, self.msg_text)

        if not self.timer_manager.active_timers and not self.show_preview and not self.show_debug: return

        timers_to_draw = []
        if self.timer_manager.active_timers:
            sorted_active = sorted(self.timer_manager.active_timers.items(), key=lambda x: x[1]["seconds"], reverse=True)
            for k, d in sorted_active: timers_to_draw.append((k, d["seconds"], d["pixmap"]))
        
        # 如果有任何世界計時器處於活動狀態，也將其加入列表
        for k, d in self.world_timers.items():
            timers_to_draw.append((k, d["seconds"], d["pixmap"]))

        if not timers_to_draw and self.show_preview:
            timers_to_draw.append(("preview", 300, QPixmap(resource_path("buff_pngs/arrow.png"))))

        new_click_zones = {}; spacing = 56; total_width = len(timers_to_draw) * spacing
        # 右對齊：錨點 base_x 是計時器群組的 右邊界
        block_start_x = base_x - total_width
        block_center_y = base_y + 60
        
        for idx, (key, seconds, pixmap) in enumerate(timers_to_draw):
            x_pos = block_start_x + idx * spacing + (spacing // 2); block_center = QPoint(x_pos, block_center_y)
            icon_size = 40; icon_rect = QRect(block_center.x() - 20, block_center.y() - 45, 40, 40)
            text_rect = QRect(block_center.x() - 50, block_center.y() - 13, 100, 50)
            
            # 點擊區域：優先使用圖示區域，否則使用文字區域
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
            font.setPointSize(15 if display_seconds >= 1000 else (22 if seconds > 3 else 26))
            font.setBold(True)
            painter.setFont(font)
            text_rect = QRect(block_center.x() - 50, block_center.y() - 13, 100, 50)
            painter.setPen(QPen(QColor(0,0,0,200), 4)); painter.drawText(text_rect, Qt.AlignmentFlag.AlignCenter, text)
            painter.setPen(QPen(color, 2)); painter.drawText(text_rect, Qt.AlignmentFlag.AlignCenter, text)
        self.click_zones = new_click_zones

        # --- FINAL LAYER: Debug Overlays (Drawn last to stay on top) ---
        if self.show_debug:
            painter.setPen(QColor(255, 0, 0))
            painter.drawText(20, 60, f"[DEBUG MODE ACTIVE] HWND: {getattr(self, 'last_target_hwnd', 'None')}")

            if hasattr(self, 'last_coin_pos') and self.last_coin_pos:
                try:
                    hwnd = getattr(self, 'last_target_hwnd', None)
                    if hwnd and win32gui.IsWindow(hwnd):
                        cx, cy, cw, ch = self.last_coin_pos
                        pt_raw = win32gui.ClientToScreen(hwnd, (cx, cy))
                        dpr = self.screen().devicePixelRatio() if self.screen() else 1.0
                        logical_pt = QPoint(int(pt_raw[0]/dpr), int(pt_raw[1]/dpr))
                        local_pt = self.mapFromGlobal(logical_pt)
                        
                        painter.setPen(QPen(QColor(255, 255, 0, 200), 2))
                        painter.drawRect(int(local_pt.x()), int(local_pt.y()), int(cw/dpr), int(ch/dpr))
                        conf_val = getattr(self, 'last_coin_match_conf', 0)
                        painter.setPen(QColor(255, 255, 0))
                        painter.drawText(local_pt.x(), local_pt.y() - 5, f"Coin ({int(conf_val*100)}%)")
                        
                        if hasattr(self, 'last_coin_info_pos') and self.last_coin_info_pos:
                            ix, iy, iw, ih = self.last_coin_info_pos
                            ipt_raw = win32gui.ClientToScreen(hwnd, (ix, iy))
                            ilogical_pt = QPoint(int(ipt_raw[0]/dpr), int(ipt_raw[1]/dpr))
                            ilocal_pt = self.mapFromGlobal(ilogical_pt)
                            painter.setPen(QPen(QColor(0, 255, 255, 180), 2))
                            painter.drawRect(int(ilocal_pt.x()), int(ilocal_pt.y()), int(iw/dpr), int(ih/dpr))
                            if hasattr(self, 'last_coin_ocr') and self.last_coin_ocr:
                                painter.setPen(QColor(0, 255, 255))
                                painter.drawText(ilocal_pt.x(), ilocal_pt.y() - 5, f"Found: {self.last_coin_ocr}")
                except: pass


    def export_exp_report(self):
        """
        將經驗值面板渲染為靜態圖片並直接儲存（此處保留與 Controller 重複的邏輯以維持原狀）。
        """
        pw, ph = 330, 220 # 報告圖略高一些以容納更多細節
        pixmap = QPixmap(pw, ph)
        pixmap.fill(Qt.GlobalColor.transparent)
        
        painter = QPainter(pixmap)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        
        # 1. 繪製背景
        rect = QRect(0, 0, pw, ph)
        path = QPainterPath()
        path.addRoundedRect(QRectF(rect).adjusted(2, 2, -2, -2), 15, 15)
        painter.setPen(QPen(QColor(255, 215, 0), 2))
        painter.setBrush(QColor(10, 10, 15, 240)) # 報告圖使用較不透明的背景
        painter.drawPath(path)
        
        # 加上浮水印與版權宣告
        painter.setPen(QPen(QColor(255, 255, 255, 80)))
        font = QFont("Microsoft JhengHei", 9)
        font.setItalic(True)
        painter.setFont(font)
        painter.drawText(rect.adjusted(0, 0, -15, -10), Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignBottom, "使用 Artale 瑞士刀記錄")

        # 呼叫共用的繪圖邏輯（座標設為相對原點）
        self._draw_exp_content(painter, 0, 0, pw, ph, is_export=True)
        painter.end()
        
        # 儲存至圖片資料夾
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
        # 1. 標題與標籤
        painter.setPen(QColor(255, 255, 255))
        font = QFont("Microsoft JhengHei")
        font.setPointSize(12 if is_export else 11); font.setBold(True)
        painter.setFont(font)
        y = py + (30 if is_export else 25)
        painter.drawText(px + 15, y, f"📊 經驗值監測報告" if is_export else "📊 經驗值監測")
        
        # 2. 次要資訊 (紀錄時長與累計，整合在同一行)
        y += (28 if is_export else 25)
        duration_sec = self.current_exp_data.get("tracking_duration", 0)
        h_dur = duration_sec // 3600; m_dur = (duration_sec % 3600) // 60; s_dur = duration_sec % 60
        val_dur = f"{h_dur:02d}:{m_dur:02d}:{s_dur:02d}" if h_dur > 0 else f"{m_dur:02d}:{s_dur:02d}"
        
        # 時長部分 (Duration)
        painter.setPen(QColor(100, 255, 100))
        font.setPointSize(9); font.setBold(True); painter.setFont(font)
        fm_small = painter.fontMetrics()
        lbl_dur = "紀錄時長:"
        painter.drawText(px + 15, y, lbl_dur)
        painter.setPen(QColor(255, 255, 255))
        painter.drawText(px + 15 + fm_small.horizontalAdvance(lbl_dur) + 5, y, val_dur)
        
        # 累計部分 (靠右對齊)
        painter.setPen(QColor(100, 255, 100))
        lbl_total = "總累積:"
        # 基於寬度或中間位置來定位標籤
        total_start_x = px + 145
        painter.drawText(total_start_x, y, lbl_total)
        painter.setPen(QColor(255, 255, 255))
        val_total = f"+{self.cumulative_gain:,} ({self.cumulative_pct:+.2f}%)"
        painter.drawText(QRect(px + 15, y - 18, pw - 30, 24), Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter, val_total)

        # 3. 升級預計時長
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
        
        # 4. 十分鍾效率 (Sliding Window)
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

        # 5. 十分鍾歷史最高紀錄 (Record)
        y += (32 if is_export else 28)
        lbl_max = "十分鐘最高經驗: "
        painter.setPen(QColor(100, 255, 100))
        painter.drawText(px + 15, y, lbl_max)
        
        duration_sec = self.current_exp_data.get("tracking_duration", 0)
        val_max = f"{self.max_10m_exp:,}" if duration_sec >= 600 else "(未滿十分鐘)"
        painter.setPen(QColor(255, 255, 255))
        painter.drawText(QRect(px + 15, y - 18, pw - 30, 24), Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter, val_max)
        
        # 6. 楓幣效率與累計 (僅在啟用或匯出時顯示)
        if self.show_money_log or is_export:
            # 6a. 十分鍾效率
            y += (32 if is_export else 28)
            lbl_money_10m = "十分鐘楓幣效率: "
            painter.setPen(QColor(255, 215, 0)) # 黃金色
            painter.drawText(px + 15, y, lbl_money_10m)
            money_10m = self.current_exp_data.get("money_10m", 0)
            val_money_10m = f"+{money_10m:,}"
            painter.setPen(QColor(255, 255, 255))
            painter.drawText(QRect(px + 15, y - 18, pw - 30, 24), Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter, val_money_10m)
            
            # 6b. 累計獲取楓幣 (Total)
            y += (28 if is_export else 25)
            lbl_money_total = "累計獲取楓幣: "
            painter.setPen(QColor(255, 215, 0))
            font.setPointSize(9); painter.setFont(font)
            painter.drawText(px + 15, y, lbl_money_total)
            val_total_money = f"{self.cumulative_money:+,}"
            painter.setPen(QColor(255, 255, 255))
            painter.drawText(QRect(px + 15, y - 18, pw - 30, 24), Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter, val_total_money)
        
        last_y = y # 用於圖表定位參考

        # 5. 趨勢走勢圖 (雙線: 綠色為經驗值, 黃色為楓幣)
        if len(self.exp_rate_history) > 1:
            gh = 50 if is_export else 40; gw = pw - 30
            gx = px + 15; gy = last_y + (15 if is_export else 10)
            
            painter.setPen(QColor(255, 255, 255, 20))
            painter.setBrush(QColor(255, 255, 255, 5))
            painter.drawRoundedRect(gx, gy, gw, gh, 4, 4)
            
            # 繪製專用走勢線的輔助函式
            def draw_line(history, color, alpha_fill):
                if not history: return
                max_v = max(history)
                if max_v <= 0: max_v = 1
                
                path = QPainterPath()
                max_points = 40
                step_x = gw / (max_points - 1)
                
                for i, v in enumerate(history):
                    vx = gx + i * step_x
                    vy = gy + gh - (v / max_v * (gh - 4)) - 2
                    if i == 0: path.moveTo(vx, vy)
                    else: path.lineTo(vx, vy)
                
                # 區域平滑填充 (透明漸層)
                fill_path = QPainterPath(path)
                fill_path.lineTo(gx + (len(history)-1) * step_x, gy + gh)
                fill_path.lineTo(gx, gy + gh)
                fill_path.closeSubpath()
                
                grad = QLinearGradient(gx, gy, gx, gy + gh)
                grad.setColorAt(0, QColor(*color.getRgb()[:3], alpha_fill))
                grad.setColorAt(1, QColor(*color.getRgb()[:3], 0))
                painter.fillPath(fill_path, grad)
                
                # 繪製線條邊框
                painter.setPen(QPen(color, 2))
                painter.setBrush(Qt.BrushStyle.NoBrush)
                painter.drawPath(path)

            # 先繪製楓幣線 (黃色)，使其位於經驗值線後方
            if self.show_money_log and len(self.money_rate_history) > 1:
                draw_line(self.money_rate_history, QColor(255, 215, 0), 40)
            
            # 最上方繪製經驗值線 (綠色)
            draw_line(self.exp_rate_history, QColor(100, 255, 100), 70)

    def draw_exp_panel(self, painter):
        if not self.show_exp_panel:
            return
            
        # 面板幾何佈局邏輯
        bx = self.rect().center().x() + self.exp_x_offset
        by = self.rect().center().y() + self.exp_y_offset
        
        # 根據是否顯示楓幣動態調整高度 (兩行約增加 55px)
        ph = 250 if self.show_money_log else 195
        pw = 330
        
        # 靠右對齊：bx 代表右邊界
        panel_rect = QRect(bx - pw, by - 130 - ph // 2, pw, ph)
        px, py = panel_rect.x(), panel_rect.y()
        
        # 1. Background
        path = QPainterPath()
        path.addRoundedRect(QRectF(panel_rect), 12, 12)
        painter.setPen(QPen(QColor(255, 215, 0, 255), 2))
        painter.setBrush(QColor(10, 10, 15, int(self.base_opacity * 255))) 
        painter.drawPath(path)
        
        # 呼叫結構化繪圖邏輯
        self._draw_exp_content(painter, px, py, pw, ph, is_export=False)

        # 1.5 Title & Level
        pass
        
