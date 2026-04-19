import sys
import os
import shutil
import threading
import time
import logging
import subprocess
import psutil
import semver

import cv2
import pytesseract
from PyQt6 import sip
from PyQt6.QtWidgets import (QApplication, QWidget, QSystemTrayIcon, QMenu)
from PyQt6.QtCore import Qt, QPoint, QRect, QTimer, pyqtSignal, QRectF, QStandardPaths
from PyQt6.QtGui import QFont, QColor, QPainter, QPen, QPixmap, QIcon, QPainterPath, QAction, QImage, QLinearGradient

# Local imports
from .skill_timer import TimerManager
from .settings_window import SettingsWindow
from .platform import WindowManagerImpl, ScreenCaptureImpl

# 初始化日誌記錄器
logging.getLogger('pytesseract').setLevel(logging.WARNING)
logging.getLogger('PIL').setLevel(logging.WARNING)
logger = logging.getLogger(__name__)

# Using utils.py for VERSION, REPO_URL, resource_path, ConfigManager, EXP_TABLE
from .utils import VERSION, REPO_URL, resource_path, ConfigManager, EXP_TABLE, _project_root

# Tesseract Portable Setup (Cross-platform)
def get_tess_cmd():
    """Detect Tesseract-OCR path across platforms"""
    # 1. PyInstaller bundle
    if hasattr(sys, '_MEIPASS'):
        bundle_tess = os.path.join(sys._MEIPASS, "Tesseract-OCR", "tesseract.exe")
        if os.path.exists(bundle_tess):
            tess_dir = os.path.dirname(bundle_tess)
            if tess_dir not in os.environ.get("PATH", ""):
                os.environ["PATH"] = tess_dir + os.pathsep + os.environ.get("PATH", "")
            return bundle_tess

    # 2. Local vendor folder (Windows portable)
    if sys.platform == "win32":
        local_tess = os.path.join(_project_root(), "vendor", "Tesseract-OCR", "tesseract.exe")
        if os.path.exists(local_tess):
            tess_dir = os.path.dirname(local_tess)
            if tess_dir not in os.environ.get("PATH", ""):
                os.environ["PATH"] = tess_dir + os.pathsep + os.environ.get("PATH", "")
            return local_tess

    # 3. System-installed
    if sys.platform == "darwin":
        for p in ["/opt/homebrew/bin/tesseract", "/usr/local/bin/tesseract"]:
            if os.path.exists(p):
                return p
    elif sys.platform == "win32":
        common_p = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
        if os.path.exists(common_p):
            return common_p

    # 4. Fallback: system PATH
    return shutil.which("tesseract")

if pytesseract:
    pytesseract.pytesseract.tesseract_cmd = get_tess_cmd()


def _font_families():
    """Return preferred CJK font families for the current platform."""
    if sys.platform == "darwin":
        return ["PingFang TC", "Heiti TC"]
    return ["Microsoft JhengHei", "微軟正黑體"]


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
    exp_visual_request = pyqtSignal(ExpVisualData)    # 用於除錯影像: exp, lv, coin
    lv_update_request = pyqtSignal(LVUpdateData)
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
        self.click_zones = {}
        self.is_active = False # For timers compat
        self.show_preview = False
        self.active_profile_name = "F1"
        self._is_running = True
        self._wm = WindowManagerImpl()
        self._wm = WindowManagerImpl()
        
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
        self.current_exp_data = StatsData(
            text="---", value=0, percent=0.0, 
            gained_10m=0, percent_10m=0.0, time_to_level=-1,
            is_estimated=True, tracking_duration=0,
            money_10m=0, cumulative_money=0, cumulative_gain=0,
            cumulative_pct=0.0, max_10m_exp=0,
            exp_rate_history=[], money_rate_history=[]
        )
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
        # 必須重置後端的 Tracker 狀態，否則 UI 數值會立即被舊數據覆蓋
        if self.controller and self.controller.tracker:
            self.controller.tracker.reset_baseline()
            
        if not silent:
            self.show_notification("📊 經驗值統計已重置")

    def on_update_found(self, tag, url):
        self._latest_version_info = (tag, url)
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
                    latest_version = latest_tag[1:]
                    current_version = VERSION[1:]
                    if semver.compare(latest_version, current_version) == 1:
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
        # We no longer move the Overlay window; it stays full-screen on the virtual desktop.
        # This keeps the UI stable and independent of the game's movement.
        if not hasattr(self, '_wm'):
            return
        try:
            info = self._wm.find_game_window(self.target_window_title, "msw.exe")
        except Exception:
            self.game_hwnd = None
            return

        if info:
            try:
                dpr = self.screen().devicePixelRatio()
                sx, sy = self._wm.client_to_screen(info.window_id, 0, info.height)
                logical_pt = QPoint(int(sx / dpr), int(sy / dpr))
                local_bl = self.mapFromGlobal(logical_pt)
                self.bx, self.by = local_bl.x(), local_bl.y()
            except Exception as e:
                logger.debug(f"[Overlay] Window mapping failed: {e}")

        if not self.isVisible():
            self.show()
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
        self.cumulative_gain = stats.cumulative_gain
        self.cumulative_pct = stats.cumulative_pct
        self.cumulative_money = stats.cumulative_money
        self.max_10m_exp = stats.max_10m_exp
        self.exp_rate_history = stats.exp_rate_history
        self.money_rate_history = stats.money_rate_history
        self.update()

    def on_toggle_exp(self):
        self.show_exp_panel = not self.show_exp_panel
        status = "已啟用" if self.show_exp_panel else "已關閉"
        
        # Modular toggle logic via controller
        if self.controller:
            self.controller.toggle_tracking(self.show_exp_panel)
            
        # 使用者要求：開啟面板（F10）時重置數據，以便開始新的紀錄
        if self.show_exp_panel:
            self.reset_exp_stats(silent=True)
            
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

<<<<<<< HEAD
    # --- UI 輔助邏輯 ---
=======
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
        # We now use a fixed 3.0x scale instead of target_h
        scale = 3.0
        char_spacing = 30
        
        def build_pass_canvas(boxes):
            if not boxes: return None
            
            # Pre-calculate scaled dimensions to find max height
            scaled_dims = [(int(b[2] * scale), int(b[3] * scale)) for b in boxes]
            max_nh = max([d[1] for d in scaled_dims])
            
            # Increase horizontal padding: char_spacing + extra 180px (90px sides)
            total_w = sum([d[0] for d in scaled_dims]) + (len(boxes) + 1) * char_spacing + 180
            
            # Dynamic height: max character height + generous 80px (40px top/bot)
            canvas_h = max_nh + 80 
            canvas = np.ones((canvas_h, total_w), dtype=np.uint8) * 255
            curr_x = 90 # Start with 90px left padding
            
            # Baseline alignment: Ensure we have at least 40px at bottom
            baseline_y = max_nh + 40
            
            for idx, b in enumerate(boxes):
                char = thresh_img[b[1]:b[1]+b[3], b[0]:b[0]+b[2]]
                char_inv = cv2.bitwise_not(char)
                nw, nh = scaled_dims[idx]
                if nw > 0 and nh > 0:
                    resized = cv2.resize(char_inv, (nw, nh), interpolation=cv2.INTER_CUBIC)
                    
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
        current_capture = None

        def on_frame_arrived_callback(img):
            nonlocal last_processed, target_hwnd, current_capture
            try:
                # Process if panel is ON
                if not self._is_running: return
                if not getattr(self, "show_exp_panel", False):
                    # Hard stop when panel is hidden, even if debug is on (for performance)
                    logger.info("[ExpTracker] Stopping session (Panel Closed)")
                    self._exp_tracker_active = False # Manual break for the inner loop
                    if current_capture is not None:
                        current_capture.stop()
                    return

                now = time.time()
                if now - last_processed < 1.0: return

                # Window Search (Dynamic lookup)
                if not target_hwnd or not self._wm.is_valid(target_hwnd):
                    info_cand = None
                    for name in ["MapleStory Worlds-Artale (繁體中文版)", "MapleStory Worlds-Artale"]:
                        info_cand = self._wm.find_game_window(name, "msw.exe")
                        if info_cand: break
                    if not info_cand: return
                    target_hwnd = info_cand.window_id

                # Handle minimized state silently
                try:
                    if self._wm.is_minimized(target_hwnd): return
                except Exception:
                    target_hwnd = None; return

                # img is already a BGR numpy array from ScreenCaptureImpl
                h, w = img.shape[:2]

                # RESTART TRIGGER: Resolution change (Maximized/Restored)
                if not hasattr(self, 'last_cap_w'): self.last_cap_w, self.last_cap_h = w, h
                if abs(w - self.last_cap_w) > 10 or abs(h - self.last_cap_h) > 10:
                    logger.info(f"[ExpTracker] Resolution changed ({self.last_cap_w}x{self.last_cap_h} -> {w}x{h}), rebuilding session...")
                    self.last_cap_w, self.last_cap_h = w, h
                    self._exp_tracker_active = False # SIGNAL the supervisor loop to break
                    if current_capture is not None:
                        current_capture.stop()
                    return

                x_cr, y_cr, cw_ref, ch_ref = self._wm.get_client_rect(target_hwnd)
                scale = min(cw_ref / self.BASE_W, ch_ref / self.BASE_H)

                # Handle maximized vs windowed border logic
                if self._wm.is_maximized(target_hwnd):
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
                        # 2. Coin Recognition: Only scan icon if we HAVEN'T locked coordinates or if OCR is failing
                        needs_rescan = not hasattr(self, 'last_coin_pos') or self.last_coin_pos is None
                        # Check if last OCR was bad (we'll define self.last_money_ocr_conf)
                        if getattr(self, 'last_money_ocr_conf', 100) < 60:
                            needs_rescan = True

                        if needs_rescan:
                            img_roi = img
                            roi_x1, roi_y1 = 0, 0
                            
                            tpl_h, tpl_w = self.coin_tpl.shape[:2]
                            st_w = int(tpl_w * scale); st_h = int(tpl_h * scale)
                            if st_w > 5 and st_h > 5:
                                tpl_resized = cv2.resize(self.coin_tpl, (st_w, st_h))
                                res = cv2.matchTemplate(img_roi, tpl_resized, cv2.TM_CCOEFF_NORMED)
                                _, max_val, _, max_loc = cv2.minMaxLoc(res)
                                
                                if max_val > 0.7:
                                    real_x, real_y = max_loc[0] + roi_x1, max_loc[1] + roi_y1
                                    self.last_coin_pos = (real_x - off_x, real_y - off_y, st_w, st_h)
                                    self.last_coin_match_conf = max_val
                                    # Reset OCR conf to "Perfect" after finding icon
                                    self.last_money_ocr_conf = 100 

                        if hasattr(self, 'last_coin_pos') and self.last_coin_pos:
                            real_x = self.last_coin_pos[0] + off_x
                            real_y = self.last_coin_pos[1] + off_y
                            st_w, st_h = self.last_coin_pos[2], self.last_coin_pos[3]
                            
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
                                
                                # Emit debug image & confidence for UI
                                if self.show_debug and not sip.isdeleted(self):
                                    self.exp_visual_request.emit({
                                        "coin": ic_thresh.copy(),
                                        "coin_match_conf": self.last_coin_match_conf * 100,
                                        "money_ocr_conf": self.last_money_ocr_conf
                                    })
                                
                                # Perform OCR and capture confidence
                                if pytesseract and pytesseract.pytesseract.tesseract_cmd:
                                    try:
                                        ic_padded = cv2.copyMakeBorder(ic_thresh, 10, 10, 10, 10, cv2.BORDER_CONSTANT, value=0)
                                        ocr_data = pytesseract.image_to_data(ic_padded, output_type=pytesseract.Output.DICT, config='--psm 7 -c tessedit_char_whitelist=0123456789,')
                                        idx_list = [i for i, text in enumerate(ocr_data['text']) if text.strip()]
                                        if idx_list:
                                            # Average confidence of non-empty words
                                            self.last_money_ocr_conf = sum([float(ocr_data['conf'][i]) for i in idx_list]) / len(idx_list)
                                            if self.last_money_ocr_conf < 60:
                                                logger.warning(f"[MoneyTracker] Low OCR confidence ({self.last_money_ocr_conf:.1f}%), will rescan coin icon.")
                                            
                                            c_text = "".join([ocr_data['text'][i] for i in idx_list])
                                            # Rest of the logic...
                                            self.last_coin_ocr = c_text
                                            val_str = c_text.replace(',', '')
                                            if val_str.isdigit():
                                                m_val = int(val_str)
                                                if not sip.isdeleted(self): self.money_update_request.emit(m_val)
                                        else:
                                            self.last_money_ocr_conf = 0 # Force rescan
                                            logger.debug("[MoneyTracker] No text found in meso area, rescanning...")
                                    except Exception as e: 
                                        self.last_money_ocr_conf = 0
                                        logger.debug(f"[MoneyTracker] OCR Error: {e}")
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
                    
                    text, e_conf, e_processed = self._perform_enhanced_ocr(thresh, "exp", upscale=4, whitelist="0123456789.[]%")
                    
                    # Always emit visual update for debug panel, even if text parsing fails
                    if not sip.isdeleted(self):
                        self.exp_visual_request.emit({"thresh": e_processed, "conf": e_conf})
                        
                    if text:
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

            # Find game window via platform abstraction
            info = self._wm.find_game_window("MapleStory Worlds-Artale", "msw.exe")
            target_hwnd = info.window_id if info else None

            if target_hwnd:
                try:
                    logger.info(f"[ExpTracker] Starting session for window {target_hwnd}")
                    self.last_target_hwnd = target_hwnd # Store for paintEvent

                    self._exp_tracker_active = True
                    current_capture = ScreenCaptureImpl()
                    current_capture.start(info, on_frame_arrived_callback)

                    # Stay in this inner loop as long as session is healthy and panel is open
                    while self._exp_tracker_active and self._is_running and getattr(self, "show_exp_panel", False):
                        if not current_capture.is_active():
                            logger.info("[ExpTracker] Session closed by capture layer.")
                            self._exp_tracker_active = False
                            break
                        time.sleep(1.0)
                except Exception as e:
                    logger.debug(f"[ExpTracker] Session failed: {e}")
                    time.sleep(2.0)
                finally:
                    if current_capture is not None:
                        current_capture.stop()
                    current_capture = None
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
            
            # --- Smart Cross-Validation & Level Inference ---
            if pct_val > 0:
                is_valid = False
                # 1. Try current level first (Fast path)
                if self.current_lv:
                    try:
                        m_lv = re.search(r'\d+', self.current_lv)
                        if m_lv:
                            total = EXP_TABLE.get(int(m_lv.group()))
                            if total and abs(int(val * 10000 / total) / 100.0 - pct_val) <= 0.05:
                                is_valid = True
                    except: pass
                
                # 2. If invalid or no level, scan table to infer level
                if not is_valid:
                    for lvl, total in EXP_TABLE.items():
                        if abs(int(val * 10000 / total) / 100.0 - pct_val) <= 0.05:
                            is_valid = True
                            if not self.current_lv or f"LV.{lvl}" != self.current_lv:
                                logger.info(f"[ExpTracker] Level detected via EXP ratio: LV.{lvl}")
                                self.current_lv = f"LV.{lvl}"
                            break
                
                if not is_valid:
                    if self.show_debug:
                        logger.debug(f"[ExpTracker] Validation Failed: {val:,} and {pct_val:.2f}% don't match any known LEVEL (Skipped)")
                    return # Data is garbage, skip update
            
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
>>>>>>> 2bad57f (refactor: overlay.py uses platform abstraction layer)


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
                target_hwnd = None
                for name in ["MapleStory Worlds-Artale (繁體中文版)", "MapleStory Worlds-Artale"]:
                    info = self._wm.find_game_window(name, "msw.exe")
                    if info:
                        target_hwnd = info.window_id; break

                if target_hwnd:
                    # 1. Get Game Client Area size
                    _cx, _cy, client_w, client_h = self._wm.get_client_rect(target_hwnd)

                    # 2. Get Global Screen coord of Client BOTTOM-LEFT (Physical)
                    bl_x, bl_y = self._wm.client_to_screen(target_hwnd, 0, client_h)

                    # 3. DPI Scaled Map to Overlay Local coordinates
                    dpr = self.screen().devicePixelRatio()
                    logical_gl_pt = QPoint(int(bl_x / dpr), int(bl_y / dpr))
                    local_bl = self.mapFromGlobal(logical_gl_pt)
                    bx, by = local_bl.x(), local_bl.y()

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
            font.setFamilies(_font_families())
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
<<<<<<< HEAD
            font.setFamilies(["Microsoft JhengHei", "微軟正黑體"])
            font.setPointSize(15 if display_seconds >= 1000 else (22 if seconds > 3 else 26))
=======
            font.setFamilies(_font_families())
            font.setPointSize(22 if seconds > 3 else 26)
>>>>>>> 2bad57f (refactor: overlay.py uses platform abstraction layer)
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
                    if hwnd and self._wm.is_valid(hwnd):
                        cx, cy, cw, ch = self.last_coin_pos
                        pt_x, pt_y = self._wm.client_to_screen(hwnd, cx, cy)
                        dpr = self.screen().devicePixelRatio() if self.screen() else 1.0
                        logical_pt = QPoint(int(pt_x / dpr), int(pt_y / dpr))
                        local_pt = self.mapFromGlobal(logical_pt)

                        painter.setPen(QPen(QColor(255, 255, 0, 200), 2))
                        painter.drawRect(int(local_pt.x()), int(local_pt.y()), int(cw / dpr), int(ch / dpr))
                        conf_val = getattr(self, 'last_coin_match_conf', 0)
                        painter.setPen(QColor(255, 255, 0))
                        painter.drawText(local_pt.x(), local_pt.y() - 5, f"Coin ({int(conf_val*100)}%)")

                        if hasattr(self, 'last_coin_info_pos') and self.last_coin_info_pos:
                            ix, iy, iw, ih = self.last_coin_info_pos
                            ipt_x, ipt_y = self._wm.client_to_screen(hwnd, ix, iy)
                            ilogical_pt = QPoint(int(ipt_x / dpr), int(ipt_y / dpr))
                            ilocal_pt = self.mapFromGlobal(ilogical_pt)
                            painter.setPen(QPen(QColor(0, 255, 255, 180), 2))
                            painter.drawRect(int(ilocal_pt.x()), int(ilocal_pt.y()), int(iw / dpr), int(ih / dpr))
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
        
<<<<<<< HEAD
        # 加上浮水印與版權宣告
        painter.setPen(QPen(QColor(255, 255, 255, 80)))
        font = QFont("Microsoft JhengHei", 9)
=======
        # Add a more visible watermark
        painter.setPen(QPen(QColor(255, 255, 255, 80))) # Increased opacity
        font = QFont(_font_families()[0], 9)
>>>>>>> 2bad57f (refactor: overlay.py uses platform abstraction layer)
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
            try:
                if sys.platform == "darwin":
                    subprocess.Popen(["open", "-R", save_path])
                else:
                    subprocess.Popen(f'explorer /select,"{save_path}"')
            except: pass
        else:
            self.show_notification("❌ 產出失敗，請檢查權限")

    def _draw_exp_content(self, painter, px, py, pw, ph, is_export=False):
        now = time.time()
        if not self.current_exp_data: return
        # 1. 標題與標籤
        painter.setPen(QColor(255, 255, 255))
        font = QFont(_font_families()[0])
        font.setPointSize(12 if is_export else 11); font.setBold(True)
        painter.setFont(font)
        y = py + (30 if is_export else 25)
        painter.drawText(px + 15, y, f"📊 經驗值監測報告" if is_export else "📊 經驗值監測")
        
        # 2. 次要資訊 (紀錄時長與累計，整合在同一行)
        y += (28 if is_export else 25)
        duration_sec = self.current_exp_data.tracking_duration
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
        
        ttl_sec = self.current_exp_data.time_to_level
        val_ttl = f"{ttl_sec // 3600}小時 {(ttl_sec % 3600) // 60}分" if ttl_sec > 0 else "計算速率中..."
        painter.setPen(QColor(255, 255, 255))
        painter.drawText(QRect(px + 15, y - 18, pw - 30, 24), Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter, val_ttl)
        
        # 4. 十分鍾效率 (Sliding Window)
        y += (32 if is_export else 28)
        gain_val = self.current_exp_data.gained_10m
        gain_pct = self.current_exp_data.percent_10m
        is_est = self.current_exp_data.is_estimated
        
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
        
        duration_sec = self.current_exp_data.tracking_duration
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
            money_10m = self.current_exp_data.money_10m
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
        
