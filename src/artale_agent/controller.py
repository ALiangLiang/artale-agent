import logging
import threading
import time
import win32gui
from PyQt6.QtCore import QObject, QTimer
from artale_agent.capture_engine import ArtaleCapture
from artale_agent.ocr_engine import ArtaleOCR
from artale_agent.exp_tracker import ExpTracker
from artale_agent.data_types import LVUpdateData
from artale_agent.utils import resource_path, ConfigManager, REPO_URL, VERSION
from artale_agent.platform import SystemUtilsImpl
from artale_agent.report_manager import ReportManager
from artale_agent.video_recorder import VideoRecorder
import urllib.request
import json
from PyQt6.QtCore import QPoint, QRect, Qt
from PyQt6.QtGui import QPixmap, QImage
import numpy as np

logger = logging.getLogger(__name__)

class ArtaleController(QObject):
    """
    Artale Agent 的核心協調器。
    負責連接截圖 (Capture)、辨識 (OCR)、統計 (Tracker) 與介面 (Overlay/View)。
    """
    def __init__(self, overlay):
        super().__init__()
        self.overlay = overlay # 控制器對應的 View (介面)
        self.system_utils = SystemUtilsImpl()
        
        # 1. 初始化引擎與統計器
        self.capture_engine = ArtaleCapture()
        self.ocr_engine = ArtaleOCR(self)
        self.tracker = ExpTracker()
        self.report_manager = ReportManager(self)
        self.recorder = VideoRecorder(fps=30.0)
        self.record_hud = True
        self.video_save_path = ""
        self._last_ocr_time = 0
        
        self.ocr_engine.set_coin_template(resource_path("coin.png"))
        
        # 2. 連結 截圖引擎 -> OCR 處理 / 橋接
        self.capture_engine.frame_arrived.connect(self.on_frame_ready)
        self.capture_engine.session_started.connect(self.on_session_started)
        
        # 3. 連結 OCR 引擎 -> 統計器
        self.ocr_engine.money_update.connect(self.on_money_parsed)
        self.ocr_engine.exp_update.connect(self.on_exp_parsed)
        self.ocr_engine.lv_update.connect(self.on_lv_parsed)
        
        # 4. 連結 統計器 -> 介面更新
        self.tracker.stats_updated.connect(self.overlay.on_stats_updated)
        self.tracker.lv_inferred.connect(lambda lv: self.overlay.lv_update_request.emit(LVUpdateData(level=str(lv), conf=100.0)))
        
        # 5. 連結 OCR 視覺輔助 -> 介面
        self.ocr_engine.exp_visual_update.connect(self.overlay.exp_visual_request)
        
        # 6. 連結 介面訊號 -> 報表管理員動作 (從 SettingsWindow 解耦)
        sw = self.overlay.settings_window
        sw.export_report_requested.connect(self.report_manager.export_exp_report)
        self.overlay.export_report_request.connect(self.report_manager.export_exp_report)
        sw.export_csv_requested.connect(self.report_manager.export_csv_report)
        sw.import_csv_requested.connect(self.report_manager.import_csv_report)
        sw.open_dashboard_requested.connect(self.report_manager.open_analytics_dashboard)
        sw.notification_requested.connect(self.overlay.show_notification)
        sw.config_updated.connect(self.load_profile)
        
        # 7. 自動檢查更新
        QTimer.singleShot(3000, lambda: self.check_for_updates(auto=True))

    def start(self):
        """啟動核心引擎"""
        self.load_profile() # 啟動時讀取配置
        self.tracker.show_debug = self.overlay.show_debug
        
        self.capture_engine.start()
        if self.overlay.show_exp_panel:
            self.capture_engine.set_active(True)

    def on_session_started(self, hwnd):
        logger.info("[Controller] Capture session active for HWND %s", hwnd)
        self.overlay.last_target_hwnd = hwnd

    def on_frame_ready(self, img, scale, off_x, off_y, cw, ch):
        """截圖引擎與 OCR 引擎之間的橋接器"""
        if not self.overlay.isVisible(): return
        
        # 1. 錄影影格混合合成寫入背景 Queue (根據選定 FPS)
        if self.recorder.is_recording:
            try:
                # 只有當啟用 HUD 錄影且視窗句柄有效時才進行合成，否則直接寫入純淨底圖
                if self.record_hud:
                    hwnd = self.capture_engine.target_hwnd
                    if hwnd and win32gui.IsWindow(hwnd):
                        # 取得設備像素比 DPR 以進行高 DPI 顯示器對齊
                        dpr = self.overlay.devicePixelRatioF()
                        
                        # 取得客戶端左上角的螢幕絕對座標
                        x, y = win32gui.ClientToScreen(hwnd, (0, 0))
                        local_pt = self.overlay.mapFromGlobal(QPoint(x, y))
                        
                        # 計算邏輯寬高以在 PyQt 空間中進行裁剪
                        logical_cw = int(cw / dpr)
                        logical_ch = int(ch / dpr)
                        
                        # 在記憶體中渲染當前 Overlay 畫面
                        pixmap = QPixmap(self.overlay.size())
                        pixmap.fill(Qt.GlobalColor.transparent)
                        self.overlay.render(pixmap)
                        
                        # 裁剪出完美的 Client Area 影格
                        q_img = pixmap.toImage()
                        crop_rect = QRect(local_pt.x(), local_pt.y(), logical_cw, logical_ch)
                        cropped_img = q_img.copy(crop_rect)
                        
                        # 如果有縮放比例，將其拉伸至與 WGC 實體像素吻合的 (cw, ch)
                        if cropped_img.width() != cw or cropped_img.height() != ch:
                            cropped_img = cropped_img.scaled(
                                cw, ch, 
                                Qt.AspectRatioMode.IgnoreAspectRatio, 
                                Qt.TransformationMode.SmoothTransformation
                            )
                        
                        # 將裁剪後的 QImage 轉成 NumPy BGRA 矩陣
                        ptr = cropped_img.constBits()
                        ptr.setsize(ch * cw * 4)
                        hud_np = np.frombuffer(ptr, dtype=np.uint8).reshape((ch, cw, 4))
                        
                        # 向量化透明度混合 (Vectorized Alpha Blending) 印在 img 上
                        # WGC 回傳的 img 是 BGR 格式，對應 Client Area 切片是 img[off_y:off_y+ch, off_x:off_x+cw]
                        client_area = img[off_y:off_y+ch, off_x:off_x+cw]
                        
                        # 取得 HUD 的 alpha 通道 (0.0 ~ 1.0)
                        alpha = hud_np[:, :, 3:4] / 255.0
                        hud_rgb = hud_np[:, :, :3]
                        
                        # 進行混合：HUD * alpha + ClientArea * (1 - alpha)
                        blended = (hud_rgb * alpha + client_area * (1.0 - alpha)).astype(np.uint8)
                        img[off_y:off_y+ch, off_x:off_x+cw] = blended
                    
                self.recorder.write_frame(img)
            except Exception as e:
                logger.error("[Controller] Composite blending error: %s", e)
        
        # 2. OCR 引擎保持每 1.0 秒執行一次 (此處作為雙重保護)
        now = time.time()
        if now - self._last_ocr_time >= 1.0:
            self._last_ocr_time = now
            self.ocr_engine.show_money_log = self.overlay.show_money_log
            self.ocr_engine.show_debug = self.overlay.show_debug
            self.ocr_engine.exp_paused = self.overlay.exp_paused
            self.ocr_engine.process_frame(img, scale, off_x, off_y, cw, ch)
            
        # 3. 觸發介面重繪
        self.overlay.update()

    def on_exp_parsed(self, data):
        """將辨識出的經驗值數據傳遞給統計器"""
        self.tracker.update_exp(data.text, conf=data.conf)

    def on_money_parsed(self, data):
        self.tracker.update_money(data.text, conf=data.conf)

    def on_lv_parsed(self, data):
        """處理等級辨識結果"""
        lv_text = data.level
        conf = data.conf
        
        # 1. 只有當 OCR 真的抓到有效數字時，才更新統計器與暫存
        if lv_text and str(lv_text).isdigit() and len(str(lv_text)) <= 3:
            lv_val = int(lv_text)
            self.tracker.update_lv_ocr(lv_val, conf) # 更新輔助判定暫存
        
        # 2. 通知 UI 更新
        self.overlay.lv_update_request.emit(data)

    def toggle_tracking(self, active):
        """切換截圖引擎的活動狀態"""
        self.capture_engine.set_active(active)

    def toggle_video_recording(self):
        """開啟或停止錄影"""
        if self.recorder.is_recording:
            self.stop_recording()
        else:
            self.start_recording()

    def start_recording(self):
        if self.recorder.is_recording: return
        
        # A. 確保獲取遊戲視窗控制代碼 (優先使用擷取引擎已定位的視窗)
        hwnd = self.capture_engine.target_hwnd
        if not hwnd or not win32gui.IsWindow(hwnd):
            hwnd = self.capture_engine._find_target_window()
            self.capture_engine.target_hwnd = hwnd
            
        if not hwnd:
            self.overlay.show_notification("❌ 找不到遊戲視窗，無法錄影！")
            return
            
        # B. 啟動錄製狀態（傳入自定義儲存路徑與非同步完成回調）
        self.recorder.start(
            save_path=self.video_save_path,
            on_save_finished=self.on_video_saved
        )
        
        # C. 根據選定的 FPS 動態計算 WGC 重設擷取頻率
        interval_ms = int(1000.0 / self.recorder.fps)
        self.capture_engine.restart_session(interval_ms)
        self.overlay.show_notification(f"🎥 開始錄影 {int(self.recorder.fps)} FPS")
        
        # D. 同步 Control Center 按鈕狀態 (亮紅警示)
        sw = self.overlay.settings_window
        if sw and hasattr(sw, "record_video_btn"):
            sw.record_video_btn.setText("🛑 停止畫面錄製")
            sw.record_video_btn.setStyleSheet("background-color: #c62828; color: white; font-weight: bold; height: 32px;")

    def stop_recording(self):
        if not self.recorder.is_recording: return
        
        # A. 停止錄製（非阻塞立即返回）
        self.recorder.stop()
        
        # B. WGC 降頻回原先的 1 FPS (1000ms)
        self.capture_engine.restart_session(1000)
        
        # C. 如果平時沒開經驗面板，則關閉擷取
        if not self.overlay.show_exp_panel:
            self.capture_engine.set_active(False)
            
        # D. 同步 Control Center 按鈕狀態 (還原)
        sw = self.overlay.settings_window
        if sw and hasattr(sw, "record_video_btn"):
            sw.record_video_btn.setText("🎥 開始遊戲錄影")
            sw.record_video_btn.setStyleSheet(sw.btn_common_style)
            
        # 立即彈出通知提示正在非同步處理存檔中，體驗極致流暢！
        self.overlay.show_notification("💾 正在非同步儲存影片...")

    def on_video_saved(self, filepath):
        """影片寫入背景執行緒完成時的非同步完成回調 (執行緒安全)"""
        import os
        if filepath:
            filename = os.path.basename(filepath)
            # 透過 Qt 訊號發送以確保執行緒安全地彈出 Toast 提示
            self.overlay.notification_request.emit(f"💾 影片已儲存：{filename}")

    def load_profile(self):
        """核心配置載入邏輯：協調介面與引擎"""
        # 1. 載入檔案
        config = ConfigManager.load_config()
        active = config.get("active_profile", "F1")
        p_data = config["profiles"].get(active, {})
        nickname = p_data.get("name", active)
        
        # 2. 通知介面清理與更新 (僅在切換配置時清理計時器)
        if getattr(self.overlay, "active_profile_name", None) != active:
            self.overlay.clear_all_timers(show_msg=False)
        self.overlay.apply_profile_config(active, nickname, config)
        
        # 3. 同步至其他引擎 (若有需要)
        self.tracker.show_debug = config.get("show_debug", False)
        
        # 4. 同步錄影自定義設定
        fps_val = config.get("video_fps", 30)
        self.recorder.fps = float(fps_val)
        self.recorder.frame_interval = 1.0 / self.recorder.fps
        self.record_hud = config.get("record_hud", True)
        self.video_save_path = config.get("video_save_path", "")
        
        logger.info("[Controller] Profile '%s' loaded successfully. (Video Settings: FPS=%s, HUD=%s, Path=%s)", 
                    active, fps_val, self.record_hud, self.video_save_path)

    def check_for_updates(self, auto=False):
        """檢查 GitHub 上的新版本"""
        def _check():
            try:
                url = f"https://api.github.com/repos/{REPO_URL}/releases"
                req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
                with urllib.request.urlopen(req, timeout=5) as response:
                    releases = json.loads(response.read().decode())
                    if not isinstance(releases, list): return
                    
                    # 搜尋最新的非 alpha/beta 發行版本
                    latest_release = None
                    for r in releases:
                        tag = r.get("tag_name", "")
                        is_pre = r.get("prerelease", False)
                        if is_pre or "-alpha" in tag.lower() or "-beta" in tag.lower():
                            continue
                        latest_release = r
                        break
                    
                    if not latest_release: return
                    latest_tag = latest_release.get("tag_name", VERSION)
                    
                    if latest_tag != VERSION:
                        html_url = latest_release.get("html_url", f"https://github.com/{REPO_URL}/releases")
                        # 透過 Overlay 的訊號同步 UI 狀態
                        self.overlay.update_found.emit(latest_tag, html_url)
            except Exception as e:
                logger.debug("[Update] Check failed: %s", e)
        
        threading.Thread(target=_check, daemon=True).start()

