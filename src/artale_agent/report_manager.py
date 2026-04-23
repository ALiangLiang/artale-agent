import csv
import logging
import os
import re
import shutil
import sys
import time
import webbrowser
from datetime import datetime
from PyQt6.QtCore import QObject, QStandardPaths, Qt, QRect, QRectF
from PyQt6.QtWidgets import QFileDialog, QApplication
from PyQt6.QtGui import QPixmap, QPainter, QColor, QFont, QPen, QPainterPath
from artale_agent.utils import resource_path, _project_root, get_version

logger = logging.getLogger(__name__)

class ReportManager(QObject):
    """
    負責經驗值報表的匯出與匯入邏輯，包含圖片生成與 CSV 處理。
    """
    def __init__(self, controller):
        super().__init__()
        self.controller = controller
        self.tracker = controller.tracker
        self.overlay = controller.overlay
        self.system_utils = controller.system_utils

    def export_exp_report(self):
        """
        產生報告圖並儲存。
        """
        import time
        
        pw, ph = 330, 220
        pixmap = QPixmap(pw, ph)
        pixmap.fill(Qt.GlobalColor.transparent)
        
        painter = QPainter(pixmap)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)

        # 1. 繪製報告背景 (較不透明)
        rect = QRect(0, 0, pw, ph)
        path = QPainterPath()
        path.addRoundedRect(QRectF(rect).adjusted(2, 2, -2, -2), 15, 15)
        painter.setPen(QPen(QColor(255, 215, 0), 2))
        painter.setBrush(QColor(10, 10, 15, 240))
        painter.drawPath(path)
        
        # 2. 加上浮水印與版權宣告
        painter.setPen(QPen(QColor(255, 255, 255, 80)))
        font = QFont("Microsoft JhengHei", 9)
        font.setItalic(True)
        painter.setFont(font)
        painter.drawText(
            rect.adjusted(0, 0, -15, -10),
            Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignBottom,
            "使用 Artale 瑞士刀記錄",
        )

        # 3. 呼叫共用的經驗值繪圖邏輯
        self.overlay._draw_exp_content(painter, 0, 0, pw, ph, is_export=True)
        painter.end()
        
        # 系統檔案與剪貼簿操作
        filename = f"Artale瑞士刀_{int(time.time())}.png"
        pictures_dir = QStandardPaths.writableLocation(QStandardPaths.StandardLocation.PicturesLocation)
        save_path = os.path.join(pictures_dir, filename)
        
        if pixmap.save(save_path, "PNG"):
            # 複製到剪貼簿
            QApplication.clipboard().setPixmap(pixmap)
            
            logger.info("[Report] Image exported to %s", save_path)
            self.overlay.show_notification(f"✅ 成果圖已儲存並複製到剪貼簿！")
            self.system_utils.open_file_manager(save_path, select=True)
        else:
            self.overlay.show_notification("❌ 產出失敗，請檢查權限")

    def open_analytics_dashboard(self):
        """開啟數據儀表板 HTML，若不存在則自動從內建資源產生"""
        if hasattr(sys, "_MEIPASS"):
            deploy_dir = os.path.dirname(sys.executable)
        else:
            deploy_dir = _project_root()

        target_path = os.path.join(deploy_dir, "analytics.html")

        # 每次開啟都嘗試部署最新版本的 HTML (確保 UI 更新能套用)
        source_path = resource_path("analytics.html")
        if os.path.exists(source_path):
            try:
                # 讀取並注入版本號
                with open(source_path, 'r', encoding='utf-8') as f:
                    content = f.read()
                
                version = get_version()
                content = content.replace("{{VERSION}}", version)
                
                with open(target_path, 'w', encoding='utf-8') as f:
                    f.write(content)
                
                logger.info("[Report] Analytics dashboard updated to %s at %s", version, target_path)
            except Exception as e:
                logger.error("[Report] Failed to update analytics.html: %s", e)
                # 如果失敗 (例如檔案被瀏覽器佔用)，則繼續開啟舊有的檔案而不中斷
        
        if not os.path.exists(target_path):
            self.overlay.show_notification("❌ 找不到儀表板檔案")
            return

        webbrowser.open(target_path)
        self.overlay.show_notification("📈 已開啟數據儀表板")

    def export_csv_report(self):
        """
        將累積的歷史紀錄匯出為 CSV 檔案。
        """
        # 確定儲存目錄 (與 EXE 同目錄)
        if hasattr(sys, "_MEIPASS"):
            save_dir = os.path.dirname(sys.executable)
        else:
            save_dir = _project_root()
            
        logs_dir = os.path.join(save_dir, "logs")
        if not os.path.exists(logs_dir):
            try:
                os.makedirs(logs_dir)
            except Exception as e:
                logger.error("Failed to create logs directory: %s", e)
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
            with open(save_path, "w", newline="", encoding="utf-8-sig") as f:
                writer = csv.DictWriter(f, fieldnames=headers)
                writer.writeheader()
                writer.writerows(history)
            
            self.overlay.show_notification(f"✅ CSV 紀錄已儲存至 logs 目錄")
            logger.info("[Report] CSV Exported to %s", save_path)
            self.system_utils.open_file_manager(save_path, select=True)
        except Exception as e:
            logger.error("CSV Export failed: %s", e)
            self.overlay.show_notification(f"❌ 匯出失敗: {e}")

    def import_csv_report(self):
        """
        匯入外部 CSV 紀錄，並進行資料標準化與等級補全。
        """
        file_path, _ = QFileDialog.getOpenFileName(
            None, "選取要匯入的 CSV 檔案", "", "CSV Files (*.csv)"
        )
        if not file_path:
            return
            
        try:
            imported_count = 0
            with open(file_path, "r", encoding="utf-8-sig") as f:
                reader = csv.DictReader(f)
                
                # 判定是否為外部工具格式：檢查標題最後一欄是否為「等級」
                fieldnames = reader.fieldnames if reader.fieldnames else []
                is_external = not (fieldnames and fieldnames[-1] == "等級")
                
                new_entries = []
                for row in reader:
                    if is_external:
                        # 外部格式才需要進行內容標準化轉換
                        row = self._normalize_row(row)
                        # 修正外部工具特性：第一筆資料的「取得EXP」歸零
                        if imported_count == 0:
                            row["取得EXP"] = "0"
                    
                    new_entries.append(row)
                    imported_count += 1
                
            if not new_entries:
                self.overlay.show_notification("⚠️ 選擇的檔案無有效數據")
                return

            if hasattr(sys, "_MEIPASS"):
                save_dir = os.path.dirname(sys.executable)
            else:
                save_dir = _project_root()
            
            logs_dir = os.path.join(save_dir, "logs")
            if not os.path.exists(logs_dir):
                os.makedirs(logs_dir, exist_ok=True)
                
            base_name = os.path.basename(file_path)
            name_part, ext = os.path.splitext(base_name)
            output_filename = f"{name_part}_標準化補全{ext}"
            output_path = os.path.join(logs_dir, output_filename)
            
            headers = [
                "時間", "EXP數值", "EXP百分比", "取得EXP", "EXP/分", "預估10分", 
                "準確度", "統計時間", "升級預估剩餘時間", "累積經驗(10分)", 
                "累積經驗(60分)", "累積經驗(全部)", "預計60分經驗量", 
                "預計百分比(1|10|60分)", "辨識文字(前)", "辨識文字(後)", "OCR準確率", "等級"
            ]
            
            with open(output_path, "w", newline="", encoding="utf-8-sig") as f:
                writer = csv.DictWriter(f, fieldnames=headers, extrasaction='ignore')
                writer.writeheader()
                writer.writerows(new_entries)
                    
            self.overlay.show_notification(f"✅ 已標準化並補全 {imported_count} 筆紀錄！")
            logger.info("[Report] Normalized and saved %d entries to %s", imported_count, output_path)
            self.system_utils.open_file_manager(output_path, select=True)
                
        except Exception as e:
            logger.error("CSV Import failed: %s", e)
            self.overlay.show_notification(f"❌ 匯入失敗: {e}")

    # --- 內部資料處理方法 ---

    def _normalize_row(self, row):
        """
        對單一列數據進行標準化處理，將外部格式轉換為本工具的統一格式。
        """
        # 1. 基礎清理 (移除前後空白與換行)
        for key in row:
            if row[key]: row[key] = str(row[key]).strip()

        # 2. 時間轉換: 處理 '2026/4/12 下午12:12:49' -> '2026-04-12 12:12:49'
        row["時間"] = self._parse_timestamp(row.get("時間", ""))

        # 3. 數值清理: 移除 ⚡, 逗號, 以及備註 (如 <10m)
        val_str = self._clean_numeric_str(row.get("EXP數值", "0"))
        pct_str = self._clean_numeric_str(row.get("EXP百分比", "0"))
        val = int(float(val_str)) if val_str else 0
        pct = float(pct_str) if pct_str else 0.0

        # 更新欄位內容為標準純淨格式
        row["EXP數值"] = val
        row["EXP百分比"] = f"{pct:.2f}%"
        row["取得EXP"] = self._clean_numeric_str(row.get("取得EXP", "0"))
        row["EXP/分"] = self._clean_numeric_str(row.get("EXP/分", "0"))
        row["預估10分"] = self._clean_numeric_str(row.get("預估10分", "0"))
        row["累積經驗(10分)"] = self._clean_numeric_str(row.get("累積經驗(10分)", "0"))
        row["累積經驗(60分)"] = self._clean_numeric_str(row.get("累積經驗(60分)", "0"))
        row["累積經驗(全部)"] = self._clean_numeric_str(row.get("累積經驗(全部)", "0"))
        row["預計60分經驗量"] = self._clean_numeric_str(row.get("預計60分經驗量", "0"))

        # 4. 文字修正: 縮寫單位
        if "升級預估剩餘時間" in row:
            row["升級預估剩餘時間"] = str(row["升級預估剩餘時間"]).replace("分鐘", "分")

        # 5. 標籤對應: 將 '高/中/低' 轉為百分比顯示
        if row.get("準確度") in ["低", "中", "高"]:
            row["準確度"] = self._map_accuracy_label(row["準確度"])

        # 6. 等級補全: 如果缺失等級則根據經驗值推算
        if not row.get("等級"):
            if val > 0 and pct > 0:
                inferred_lv = self.tracker.infer_level(val, pct)
                row["等級"] = str(inferred_lv) if inferred_lv else ""
            else:
                row["等級"] = ""
        
        return row

    def _clean_numeric_str(self, s):
        """清理數值字串中的雜訊，僅保留數字與小數點"""
        if not s: return "0"
        s = str(s).replace("⚡", "").replace(",", "")
        s = s.split(" ")[0] # 移除 "245 (<10m)" 後半部
        s = re.sub(r'[^\d.]', '', s)
        return s if s else "0"

    def _parse_timestamp(self, time_str):
        """解析各種格式的時間字串，統一輸出為 ISO 格式"""
        if not time_str: return ""
        time_str = str(time_str).strip()
        try:
            is_pm = "下午" in time_str
            is_am = "上午" in time_str
            clean_time = time_str.replace("下午", " ").replace("上午", " ").replace("/", "-")
            for fmt in ["%Y-%m-%d %H:%M:%S", "%Y-%m-%d %I:%M:%S"]:
                try:
                    dt = datetime.strptime(clean_time.strip(), fmt)
                    if is_pm and dt.hour < 12: dt = dt.replace(hour=dt.hour + 12)
                    elif is_am and dt.hour == 12: dt = dt.replace(hour=0)
                    return dt.strftime("%Y-%m-%d %H:%M:%S")
                except: continue
        except: pass
        return time_str

    def _map_accuracy_label(self, acc_str):
        """將外部工具的文字標籤映射為百分比字串"""
        mapping = {"高": "100.0%", "中": "50.0%", "低": "10.0%"}
        return mapping.get(str(acc_str).strip(), str(acc_str))
