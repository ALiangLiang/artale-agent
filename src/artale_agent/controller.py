import logging
import threading
from PyQt6.QtCore import QObject, QTimer
from artale_agent.capture_engine import ArtaleCapture
from artale_agent.ocr_engine import ArtaleOCR
from artale_agent.exp_tracker import ExpTracker
from artale_agent.data_types import LVUpdateData
from artale_agent.utils import resource_path, ConfigManager, REPO_URL, VERSION
from artale_agent.platform import SystemUtilsImpl
from artale_agent.report_manager import ReportManager
import urllib.request
import json

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
        
        # 6. 連結 介面訊號 -> 報表管理員動作
        self.overlay.export_report_request.connect(self.report_manager.export_exp_report)
        self.overlay.export_csv_request.connect(self.report_manager.export_csv_report)
        self.overlay.import_csv_request.connect(self.report_manager.import_csv_report)
        self.overlay.open_dashboard_request.connect(self.report_manager.open_analytics_dashboard)
        self.overlay.profile_switch_request.connect(self.load_profile)
        self.overlay.settings_window.config_updated.connect(self.load_profile)
        
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
        
        # 將介面設定同步回辨識引擎
        self.ocr_engine.show_money_log = self.overlay.show_money_log
        self.ocr_engine.show_debug = self.overlay.show_debug
        self.ocr_engine.exp_paused = self.overlay.exp_paused
        
        # 分發至批次 OCR 任務
        self.ocr_engine.process_frame(img, scale, off_x, off_y, cw, ch)
        
        # 觸發介面重繪
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

    def load_profile(self):
        """核心配置載入邏輯：協調介面與引擎"""
        # 1. 載入檔案
        config = ConfigManager.load_config()
        active = config.get("active_profile", "F1")
        p_data = config["profiles"].get(active, {})
        nickname = p_data.get("name", active)
        
        # 2. 通知介面清理與更新
        self.overlay.clear_all_timers(show_msg=False)
        self.overlay.apply_profile_config(active, nickname, config)
        
        # 3. 同步至其他引擎 (若有需要)
        self.tracker.show_debug = config.get("show_debug", False)
        logger.info("[Controller] Profile '%s' loaded successfully.", active)

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

