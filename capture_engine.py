import time
import logging
import threading
import os
import cv2
import win32gui
import win32con
import win32process
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
    frame_arrived = pyqtSignal(object, float, int, int, int, int) # img, scale, off_x, off_y, cw, ch
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
        self.target_window_title = "MapleStory Worlds-Artale (繁體中文版)"
        
        # 1080p 校準參考
        self.BASE_W, self.BASE_H = 1920, 1080

    def start(self):
        self._is_running = True
        self.thread = threading.Thread(target=self._run_loop, daemon=True)
        self.thread.start()

    def stop(self):
        self._is_running = False
        self._active = False
        self.wake_event.set()

    def _get_window_metrics(self, target_hwnd, frame_w, frame_h):
        """計算視窗縮放比例以及相對於截取影格的偏移量"""
        try:
            crect = win32gui.GetClientRect(target_hwnd)
            cw_ref, ch_ref = crect[2], crect[3]
            scale = min(cw_ref / self.BASE_W, ch_ref / self.BASE_H)
            
            placement = win32gui.GetWindowPlacement(target_hwnd)
            if placement[1] == win32con.SW_SHOWMAXIMIZED:
                off_x = 0; off_y = frame_h - ch_ref
            else:
                border_w = max(0, (frame_w - cw_ref) // 2)
                off_x = border_w; off_y = max(0, frame_h - ch_ref - border_w)
            return scale, off_x, off_y, cw_ref, ch_ref
        except:
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

            target_hwnd = self._find_target_window()
            if not target_hwnd:
                time.sleep(2.0); continue

            try:
                precise_name = win32gui.GetWindowText(target_hwnd)
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
                    
                    # 偵測解析度變化
                    if self.last_cap_w != 0 and (abs(w - self.last_cap_w) > 10 or abs(h - self.last_cap_h) > 10):
                        logger.info(f"[Capture] 解析度發生變更，正在重啟...")
                        control.stop(); return

                    self.last_cap_w, self.last_cap_h = w, h
                    
                    metrics = self._get_window_metrics(target_hwnd, w, h)
                    if metrics:
                        scale, off_x, off_y, cw, ch = metrics
                        self.frame_arrived.emit(img, scale, off_x, off_y, cw, ch)
                        last_processed_time = now

                @capture.event
                def on_closed():
                    self.session_closed.emit()

                self.session_started.emit(target_hwnd)
                capture.start_free_threaded()
                
                # 健康狀況監控
                while self._active and self._is_running:
                    if not win32gui.IsWindow(target_hwnd): break
                    time.sleep(1.0)
                
                capture.stop()
            except Exception as e:
                logger.error(f"[Capture] Session Error: {e}")
                time.sleep(2.0)

    def set_active(self, active):
        self._active = active
        if active: self.wake_event.set()
