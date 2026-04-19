import time
import logging
import threading
import os
import cv2
import win32gui
import win32con
import win32process
import numpy as np
import psutil
from PyQt6.QtCore import QObject, pyqtSignal

try:
    from windows_capture import WindowsCapture
except ImportError:
    WindowsCapture = None

logger = logging.getLogger("ArtaleCapture")

class ArtaleCapture(QObject):
    """
    處理視窗尋找與螢幕截取的生命週期。
    透過訊號 (Signals) 將影格處理委派給其他組件。
    """
    frame_arrived = pyqtSignal(object, float, int, int, int, int) # img(ndarray), scale, off_x, off_y, cw, ch
    session_started = pyqtSignal(int) # hwnd
    session_closed = pyqtSignal()

    def __init__(self):
        super().__init__()
        self._is_running = True
        self._active = False
        self.stop_event = threading.Event()
        self.wake_event = threading.Event()
        self.last_cap_w = 0
        self.last_cap_h = 0
        self.target_hwnd = None
        self.target_window_title = "MapleStory Worlds-Artale (繁體中文版)"
        
        # 1080p 校準參考
        self.BASE_W, self.BASE_H = 1920, 1080
        self._session_start_maximized = False

    def start(self):
        self._is_running = True
        self.thread = threading.Thread(target=self._run_loop, daemon=True)
        self.thread.start()

    def stop(self):
        self._is_running = False
        self._active = False
        self.wake_event.set()

    def _get_window_metrics(self, target_hwnd, img):
        """計算視窗縮放比例以及相對於截取影格的偏移量"""
        try:
            h, w = img.shape[:2]
            
            # 1. 取得座標與當前狀態
            wr = win32gui.GetWindowRect(target_hwnd)
            client_tl = win32gui.ClientToScreen(target_hwnd, (0, 0))

            # 2. 取得內容區域 (Client Area) 的寬高
            crect = win32gui.GetClientRect(target_hwnd)
            cw_ref, ch_ref = crect[2], crect[3]
            if cw_ref <= 0 or ch_ref <= 0: return None
            
            scale = min(cw_ref / self.BASE_W, ch_ref / self.BASE_H)

            # 3. 核心邏輯：固定 Y 軸相對偏移 (Fixed Relative Offset Y)
            # WGC 影格與內容區的 Y 軸相對位置在 Session 開始時就固定了
            if not hasattr(self, '_session_fixed_off_y'):
                # 判定影格原點是在哪裡 (0 還是 wr[1])
                padding_top = 0
                for y in range(min(32, h)):
                    if np.any(img[y, w//2] != 0):
                        padding_top = y; break
                
                # 如果是最大化啟動且無黑邊，影格原點在螢幕 0，偏移量即為絕對座標
                if self._session_start_maximized and padding_top == 0:
                    self._session_fixed_off_y = client_tl[1]
                else:
                    # 否則影格原點在視窗頂部 (wr[1])，偏移量為相對座標 (標題列高度)
                    self._session_fixed_off_y = client_tl[1] - wr[1]
                
                logger.info(f"[Capture] Y-Offset Locked: {self._session_fixed_off_y} (Maximized: {self._session_start_maximized}, Padding: {padding_top})")

            # X 軸經測試不需要偏移 (WGC 影格左側即為內容區起點)
            return scale, 0, self._session_fixed_off_y, cw_ref, ch_ref
        except Exception as e:
            logger.error(f"Error getting window metrics: {e}")
            return None

    def _find_target_window(self):
        my_pid = os.getpid()
        found_hwnds = []
        
        def enum_handler(hwnd, lparam):
            if win32gui.IsWindowVisible(hwnd):
                _, pid = win32process.GetWindowThreadProcessId(hwnd)
                if pid != my_pid:
                    title = win32gui.GetWindowText(hwnd)
                    if "MapleStory Worlds-Artale" in title:
                        found_hwnds.append(hwnd)
        
        try:
            win32gui.EnumWindows(enum_handler, None)
            if found_hwnds: return found_hwnds[0]
            
            # 如果沒找到，回退到進程搜尋 (Process Search)
            for proc in psutil.process_iter(['pid', 'name']):
                if proc.info['name'] and proc.info['name'].lower() == 'msw.exe':
                    win_list = []
                    def cb(h, extra):
                        if win32gui.IsWindowVisible(h):
                            _, pid = win32process.GetWindowThreadProcessId(h)
                            if pid == proc.info['pid']: extra.append(h)
                        return True
                    win32gui.EnumWindows(cb, win_list)
                    if win_list:
                        return max(win_list, key=lambda h: len(win32gui.GetWindowText(h)))
        except: pass
        return None

    def _run_loop(self):
        last_processed_time = 0
        
        while self._is_running:
            if not self._active:
                self.wake_event.wait(timeout=1.0)
                self.wake_event.clear()
                if not self._active: continue

            self.target_hwnd = self.target_hwnd or self._find_target_window()
            if not self.target_hwnd:
                time.sleep(2.0); continue

            try:
                precise_name = win32gui.GetWindowText(self.target_hwnd)
                
                # 重置 Session 狀態
                if hasattr(self, '_session_start_maximized'): delattr(self, '_session_start_maximized')
                if hasattr(self, '_session_fixed_off_y'): delattr(self, '_session_fixed_off_y')
                
                placement = win32gui.GetWindowPlacement(self.target_hwnd)
                self._session_start_maximized = (placement[1] == win32con.SW_SHOWMAXIMIZED)
                logger.info(f"[Capture] Starting Session. Initial Maximized: {self._session_start_maximized}")
                
                cap_config = {
                    "window_name": precise_name,
                    "cursor_capture": False,
                    "minimum_update_interval": 1000
                }
                
                try:
                    capture = WindowsCapture(draw_border=False, **cap_config)
                except Exception as e:
                    if "Toggling the capture border" in str(e):
                        capture = WindowsCapture(draw_border=True, **cap_config)
                    else: raise e

                @capture.event
                def on_frame_arrived(frame, control):
                    nonlocal last_processed_time
                    if not self._active or not self._is_running:
                        control.stop(); return
                        
                    now = time.time()
                    if now - last_processed_time < 1.0: return
                    
                    img_orig = frame.frame_buffer
                    img = cv2.cvtColor(img_orig, cv2.COLOR_BGRA2BGR)
                    h, w = img.shape[:2]

                    self.last_cap_w, self.last_cap_h = w, h
                    
                    metrics = self._get_window_metrics(self.target_hwnd, img)
                    if metrics:
                        scale, off_x, off_y, cw, ch = metrics
                        # 只有當偏移量發生明顯變化時才記錄日誌，避免刷屏
                        if not hasattr(self, '_last_log_metrics') or self._last_log_metrics != (off_x, off_y, cw, ch):
                            logger.info(f"[Capture] Metrics: scale={scale:.2f}, offset=({off_x}, {off_y}), client={cw}x{ch}")
                            self._last_log_metrics = (off_x, off_y, cw, ch)

                        self.frame_arrived.emit(img, scale, off_x, off_y, cw, ch)
                        last_processed_time = now

                @capture.event
                def on_closed():
                    self.session_closed.emit()

                self.session_started.emit(self.target_hwnd)
                capture.start_free_threaded()
                
                # 健康狀況監控
                while self._active and self._is_running:
                    if not win32gui.IsWindow(self.target_hwnd): 
                        self.target_hwnd = None # 視窗消失，重置句柄
                        break
                    time.sleep(1.0)
                
                # 停止 Session：windows-capture 1.5.0+ 並不建議直接在 capture 物件呼叫 stop()
                # 而是透過 callback 中的 control.stop()。這裡我們只需等待 callback 收到下一張影格並自行停止。
                # 或者如果視窗已關閉，Session 會自動結束。
                pass
            except Exception as e:
                logger.error(f"[Capture] Session Error: {e}")
                time.sleep(2.0)

    def set_active(self, active):
        self._active = active
        if active: self.wake_event.set()
