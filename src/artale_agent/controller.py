import logging
import time
import os
import sys
import subprocess
import threading
import csv
from datetime import datetime
from PyQt6.QtCore import QObject, Qt, QStandardPaths, QTimer
from PyQt6.QtGui import QPixmap, QPainter
from PyQt6.QtWidgets import QApplication
from artale_agent.capture_engine import ArtaleCapture
from artale_agent.ocr_engine import ArtaleOCR
from artale_agent.exp_tracker import ExpTracker
from artale_agent.data_types import LVUpdateData
from artale_agent.utils import resource_path, _project_root
from artale_agent.platform import SystemUtilsImpl

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
        
        # 6. 連結 介面訊號 -> 控制器動作
        self.overlay.export_report_request.connect(self.export_exp_report)
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
        from artale_agent.utils import ConfigManager
        
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
        from artale_agent.utils import REPO_URL, VERSION
        
        def _check():
            try:
                import urllib.request, json, webbrowser
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
                        msg = f"✨ 發現新版本: {latest_tag}！請下載更新"
                        self.overlay.notification_request.emit(msg)
                        if not auto: webbrowser.open(html_url)
                    else:
                        if not auto: self.overlay.notification_request.emit("✅ 目前已是最新版本")
            except Exception as e:
                logger.debug("[Update] Check failed: %s", e)
                if not auto: self.overlay.notification_request.emit(f"❌ 檢查失敗: {e}")
        
        threading.Thread(target=_check, daemon=True).start()

    def export_exp_report(self):
        """
        產生成果報告圖並儲存。
        集中在此處處理以避免 UI 類別中出現文件 I/O 操作。
        """
        pw, ph = 330, 220
        pixmap = QPixmap(pw, ph)
        pixmap.fill(Qt.GlobalColor.transparent)
        
        painter = QPainter(pixmap)
        # 授權 Overlay 的專用方法進行實際繪製
        self.overlay._draw_exp_content(painter, 0, 0, pw, ph, is_export=True)
        painter.end()
        
        # 系統檔案與剪貼簿操作
        filename = f"Artale瑞士刀_{int(time.time())}.png"
        pictures_dir = QStandardPaths.writableLocation(QStandardPaths.StandardLocation.PicturesLocation)
        save_path = os.path.join(pictures_dir, filename)
        
        if pixmap.save(save_path, "PNG"):
            # 複製到剪貼簿
            QApplication.clipboard().setPixmap(pixmap)
            
            logger.info("[Report] Exported to %s", save_path)
            self.overlay.show_notification(f"✅ 成果圖已儲存並複製到剪貼簿！")
            try: subprocess.Popen(f'explorer /select,"{save_path}"')
            except: pass
        else:
            self.overlay.show_notification("❌ 產出失敗，請檢查權限")
    def export_csv_report(self):
        """
        將累積的歷史紀錄匯出為 CSV 檔案。
        儲存於執行檔 (.exe) 所在的目錄。
        """
        # 確定儲存目錄 (與 EXE 同目錄)
        if hasattr(sys, "_MEIPASS"):
            # 如果是 PyInstaller 打包環境，sys.executable 是 exe 路徑
            save_dir = os.path.dirname(sys.executable)
        else:
            # 開發環境
            save_dir = _project_root()
            
        # 建立 logs 子目錄
        logs_dir = os.path.join(save_dir, "logs")
        if not os.path.exists(logs_dir):
            try:
                os.makedirs(logs_dir)
            except Exception as e:
                logger.error("Failed to create logs directory: %s", e)
                # 如果建立失敗，就退回到原本的目錄
                logs_dir = save_dir
                
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"Artale紀錄_{timestamp}.csv"
        save_path = os.path.join(logs_dir, filename)
        
        history = self.tracker.csv_history
        if not history:
            self.overlay.show_notification("⚠️ 目前尚無紀錄資料可匯出")
            return
            
        headers = [
            "時間", "EXP數值", "EXP百分比", "取得EXP", "EXP/分", "預估10分", 
            "準確度", "統計時間", "升級預估剩餘時間", "累積經驗(10分)", 
            "累積經驗(60分)", "累積經驗(全部)", "預計60分經驗量", 
            "預計百分比(1|10|60分)", "辨識文字(前)", "辨識文字(後)", "OCR準確率", "等級"
        ]
        
        try:
            # 使用 utf-8-sig 以便 Excel 正確識別 BOM (中文字碼)
            with open(save_path, "w", newline="", encoding="utf-8-sig") as f:
                writer = csv.DictWriter(f, fieldnames=headers)
                writer.writeheader()
                writer.writerows(history)
            
            self.overlay.show_notification(f"✅ CSV 紀錄已儲存至執行檔目錄")
            logger.info("[Report] CSV Exported to %s", save_path)
            
            # 開啟檔案所在資料夾
            self.system_utils.open_file_manager(save_path, select=True)
        except Exception as e:
            logger.error("CSV Export failed: %s", e)
            self.overlay.show_notification(f"❌ 匯出失敗: {e}")
