import datetime
import logging
import os
from typing import override

import numpy as np
from PyQt6.QtCore import QSize, Qt, pyqtSignal
from PyQt6.QtGui import QIcon, QImage, QPixmap
from PyQt6.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QFrame,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QPushButton,
    QScrollArea,
    QSlider,
    QTabWidget,
    QVBoxLayout,
    QWidget,
)

try:
    from PyQt6 import sip
except ImportError:
    sip = None

from artale_agent.rjpq_tool import RJPQSyncClient, RJPQTabContent
from artale_agent.skill_timer import IconSelectorDialog, PositionHandle
from artale_agent.utils import VERSION, ConfigManager, platform_font_family, resource_path

logger = logging.getLogger(__name__)


class SettingsWindow(QWidget):
    config_updated = pyqtSignal()
    request_show = pyqtSignal()
    timer_request = pyqtSignal(str, int, str, bool)
    notification_request = pyqtSignal(str)

    def __init__(self, overlay=None):
        super().__init__()
        self.overlay = overlay
        self.setWindowTitle("Artale 瑞士刀 - Control Center")
        self.setFixedSize(400, 750)
        self.setWindowFlags(Qt.WindowType.WindowStaysOnTopHint | Qt.WindowType.Tool)
        self.is_recording = False
        self.trigger_data = {}
        self.recording_global_key = None # 目前正在錄製哪一個全域熱鍵
        self.global_hk_buttons = {} # 儲存按鈕引用
        self.handle = None
        self.exp_handle = None
        self.rjpq_handle = None

        self.init_ui()
        self.request_show.connect(self.safe_show)

    def update_debug_img(self, data):
        """更新統一的批次 OCR 監控影像"""
        if not data: return
        
        # 1. 更新大批次畫布 (Big Batch Canvas)
        img = data.exp # 在批次模式下，"exp" 包含完整的畫布
        if img is not None and img.size > 0:
            h, w = img.shape
            bytes_data = np.ascontiguousarray(img).tobytes()
            q_img = QImage(bytes_data, w, h, w, QImage.Format.Format_Grayscale8).copy()
            pixmap = QPixmap.fromImage(q_img)
            if not pixmap.isNull():
                self.debug_batch_img_lbl.setPixmap(pixmap.scaled(self.debug_batch_img_lbl.size(), Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation))
        
        # 1.5 更新等級裁切畫布
        lv_img = data.lv
        if lv_img is not None and lv_img.size > 0:
            lh, lw = lv_img.shape
            l_bytes = np.ascontiguousarray(lv_img).tobytes()
            ql_img = QImage(l_bytes, lw, lh, lw, QImage.Format.Format_Grayscale8).copy()
            l_pixmap = QPixmap.fromImage(ql_img)
            if not l_pixmap.isNull():
                self.debug_lv_img_lbl.setPixmap(l_pixmap.scaled(self.debug_lv_img_lbl.size(), Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation))
        
        # 1.6 更新楓幣裁切畫布
        coin_img = data.coin
        if coin_img is not None and coin_img.size > 0:
            ch, cw = coin_img.shape
            c_bytes = np.ascontiguousarray(coin_img).tobytes()
            qc_img = QImage(c_bytes, cw, ch, cw, QImage.Format.Format_Grayscale8).copy()
            c_pixmap = QPixmap.fromImage(qc_img)
            if not c_pixmap.isNull():
                self.debug_coin_img_lbl.setPixmap(c_pixmap.scaled(self.debug_coin_img_lbl.size(), Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation))
        
        # 2. 更新信心度指標
        conf = data.conf
        self.debug_global_conf_lbl.setText(f"OCR Confidence: {conf:.0f}%")
        color = "#51cf66" if conf >= 85 else ("#ffd700" if conf >= 60 else "#ff6b6b")
        self.debug_global_conf_lbl.setStyleSheet(f"color: {color}; font-family: Consolas; font-weight: bold; font-size: 12px;")

    def update_lv_debug_img(self, data):
        """僅更新等級數字資訊 (影像現在已成為批次的一部分)"""
        if not data: return
        lv_text = data.level
        lv_conf = data.conf
        
        self.debug_lv_stats_lbl.setText(f"LATEST LV: {lv_text or '--'}")

        if lv_text and lv_conf >= 100 and self.overlay: 
            try:
                # 處理等級更新與升級偵測
                lv_val = int(lv_text)
                self.overlay.current_lv = f"LV.{lv_val}"
                
                if self.overlay.last_confirmed_lv is not None:
                    level_diff = lv_val - self.overlay.last_confirmed_lv
                    if 0 < level_diff <= 2:
                        logger.info("[ExpTracker] 確認升級！ %s -> %s", self.overlay.last_confirmed_lv, lv_val)
                        self.overlay.exp_initial_val = None
                        self.overlay.cumulative_gain = 0
                        self.overlay.exp_history = []
                    elif level_diff != 0:
                        logger.debug("[ExpTracker] 已過濾掉不合理的等級跳變: %s -> %s", self.overlay.last_confirmed_lv, lv_val)
                
                self.overlay.last_confirmed_lv = lv_val
            except:
                pass

    def show_update_banner(self, tag, url):
        if hasattr(self, "update_banner"):
            self.update_banner.setText(
                f'✨ <a href="{url}" style="color: #ffdd00; text-decoration: underline;">發現新版本 {tag}！點此前往下載更新</a>'
            )
            self.update_banner.setVisible(True)

    @override
    def showEvent(self, event):
        super().showEvent(event)
        self.refresh_items()
        if self.overlay:
            if (
                hasattr(self.overlay, "_latest_version_info")
                and self.overlay._latest_version_info
            ):
                self.show_update_banner(*self.overlay._latest_version_info)

    def init_ui(self):
        self.setWindowTitle("Artale 瑞士刀")
        icon_path = resource_path("app_icon.png")
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))

        config = ConfigManager.load_config()
        self.layout = QVBoxLayout(self)
        self.setStyleSheet(
            f"background-color: #121212; color: #e0e0e0; font-family: {platform_font_family()};"
        )

        self.tabs = QTabWidget()
        self.tabs.setStyleSheet(f"""
            QTabWidget::pane {{ border: 1px solid #333; background: #121212; }}
            QTabBar::tab {{ background: #222; color: #888; padding: 10px 4px; min-width: 85px; font-size: 11px; font-family: {platform_font_family()}; }}
            QTabBar::tab:selected {{ background: #333; color: #ffd700; font-weight: bold; }}
        """)
        self.layout.addWidget(self.tabs)

        btn_common_style = f"""
            QPushButton {{
                background-color: #2a2a2a; color: #ccc; border: 1px solid #3d3d3d; border-radius: 4px;
                height: 32px; font-weight: bold; font-family: {platform_font_family()};
            }}
            QPushButton:hover {{ background-color: #333; border: 1px solid #555; }}
        """

        # 分頁 1: 計時器 (Timer)
        timer_tab = QWidget()
        timer_tab_layout = QVBoxLayout(timer_tab)

        profile_row = QHBoxLayout()
        self.profile_box = QComboBox()
        self.profile_box.setFixedHeight(30)
        self.profile_box.setStyleSheet("""
            QComboBox { background-color: #333; color: #ffd700; border-radius: 4px; padding: 5px; min-width: 80px; }
            QComboBox QAbstractItemView { background-color: #222; selection-background-color: #ffd700; selection-color: black; }
        """)

        self.nickname_inp = QLineEdit()
        self.nickname_inp.setPlaceholderText("在此輸入暱稱...")
        self.nickname_inp.setStyleSheet(
            "QLineEdit { background-color: #222; border: 1px solid #444; border-radius: 4px; padding: 5px; color: #fff; }"
        )
        self.nickname_inp.textChanged.connect(self.on_nickname_changed)

        profile_row.addWidget(self.profile_box, 1)
        profile_row.addWidget(self.nickname_inp, 2)
        timer_tab_layout.addLayout(profile_row)

        timer_tab_layout.addWidget(
            QLabel(
                "貼心提醒：雙擊 F1~F9 可快速切換配置組",
                styleSheet="color: #888; font-size: 11px; margin-bottom: 5px;",
            )
        )

        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setFrameShape(QFrame.Shape.NoFrame)
        self.scroll_content = QWidget()
        self.scroll_layout = QVBoxLayout(self.scroll_content)
        self.scroll_layout.setSpacing(2)
        self.scroll_layout.setContentsMargins(0, 0, 0, 0)
        self.scroll.setWidget(self.scroll_content)
        timer_tab_layout.addWidget(self.scroll)

        self.record_btn = QPushButton("➕ 新增按鍵 (點我後按鍵盤)")
        self.record_btn.setStyleSheet(btn_common_style)
        self.record_btn.clicked.connect(self.toggle_recording)
        timer_tab_layout.addWidget(self.record_btn)

        ship_group = QGroupBox("🚢 特殊提醒")
        ship_group.setStyleSheet(
            "QGroupBox { color: #aaa; font-weight: bold; border: 1px solid #333; border-radius: 8px; margin-top: 15px; padding-top: 10px; }"
        )
        ship_layout = QVBoxLayout()
        ship_btn = QPushButton("🚢 開始下班船班倒數")
        ship_btn.setStyleSheet(btn_common_style)
        ship_btn.clicked.connect(self.start_ship_timer)
        ship_layout.addWidget(ship_btn)

        elevator_row = QHBoxLayout()
        ed_btn = QPushButton("🔻 赫爾奧斯塔 (下樓)")
        ed_btn.setStyleSheet(btn_common_style)
        ed_btn.clicked.connect(lambda: self.start_elevator_timer("down"))
        eu_btn = QPushButton("🔺 赫爾奧斯塔 (上樓)")
        eu_btn.setStyleSheet(btn_common_style)
        eu_btn.clicked.connect(lambda: self.start_elevator_timer("up"))
        elevator_row.addWidget(ed_btn)
        elevator_row.addWidget(eu_btn)
        ship_layout.addLayout(elevator_row)
        ship_group.setLayout(ship_layout)
        timer_tab_layout.addWidget(ship_group)

        pos_btn = QPushButton("🔱 調整計時器位置")
        pos_btn.setStyleSheet(btn_common_style)
        pos_btn.clicked.connect(self.toggle_timer_handle)
        timer_tab_layout.addWidget(pos_btn)
        self.tabs.addTab(timer_tab, "⏲️ 計時器")

        # 分頁 2: 經驗值/楓幣
        exp_tab = QWidget()
        exp_tab_layout = QVBoxLayout(exp_tab)
        exp_info = QLabel("📊 經驗值/楓幣設定")
        exp_info.setStyleSheet("color: #ffd700; font-weight: bold; font-size: 14px; margin-top: 10px;")
        exp_tab_layout.addWidget(exp_info)
        
        self.exp_active_cb = QCheckBox("開啟經驗值監測面板 (Hotkey: F10)")
        self.exp_active_cb.setStyleSheet("color: #ccc; margin-top: 10px;")
        exp_tab_layout.addWidget(self.exp_active_cb)
        self.money_active_cb = QCheckBox("開啟楓幣記錄（實驗中）")
        self.money_active_cb.setStyleSheet("color: #ccc; margin-top: 5px;")
        exp_tab_layout.addWidget(self.money_active_cb)
        if self.overlay:
            self.exp_active_cb.setChecked(self.overlay.show_exp_panel)
            self.money_active_cb.setChecked(
                getattr(self.overlay, "show_money_log", True)
            )
        self.exp_active_cb.toggled.connect(self.on_exp_toggle_changed)
        self.money_active_cb.toggled.connect(self.on_money_toggle_changed)

        reset_exp_btn = QPushButton("🔄 重新開始紀錄 (歸零重算)")
        reset_exp_btn.setStyleSheet(btn_common_style)
        reset_exp_btn.clicked.connect(self.on_reset_exp_clicked)
        exp_tab_layout.addWidget(reset_exp_btn)

        opacity_info = QLabel("✨ 面板背景透明度")
        opacity_info.setStyleSheet("color: #aaa; font-size: 11px; margin-top: 15px;")
        exp_tab_layout.addWidget(opacity_info)
        opacity_row = QHBoxLayout()
        self.opacity_val_lbl = QLabel(f"{int(config.get('opacity', 0.5) * 100)}%")
        self.opacity_val_lbl.setStyleSheet("color: #ffd700; font-weight: bold;")
        self.opacity_slider = QSlider(Qt.Orientation.Horizontal)
        self.opacity_slider.setRange(10, 100)
        self.opacity_slider.setValue(int(config.get("opacity", 0.5) * 100))
        self.opacity_slider.valueChanged.connect(self.on_opacity_changed)
        opacity_row.addWidget(self.opacity_slider)
        opacity_row.addWidget(self.opacity_val_lbl)
        exp_tab_layout.addLayout(opacity_row)
        
        exp_pos_btn = QPushButton("📊 調整經驗值面板位置")
        exp_pos_btn.setStyleSheet(btn_common_style)
        exp_pos_btn.clicked.connect(self.toggle_exp_handle)
        exp_tab_layout.addWidget(exp_pos_btn)
        export_btn = QPushButton("📸 產出成果圖 (截圖分享)")
        export_btn.setStyleSheet(btn_common_style)
        export_btn.clicked.connect(self.overlay.export_exp_report if self.overlay else lambda: None)
        exp_tab_layout.addWidget(export_btn)
        
        self.debug_mode_cb = QCheckBox("顯示除錯訊息 (開發者模式)")
        self.debug_mode_cb.setStyleSheet("color: #888; font-size: 11px;")
        exp_tab_layout.addWidget(self.debug_mode_cb)
        
        # --- 批次 OCR 監控群組 ---
        self.debug_group = QWidget()
        self.debug_layout = QVBoxLayout(self.debug_group)
        self.debug_layout.setContentsMargins(0, 5, 0, 0)
        
        self.debug_info_lbl = QLabel("🔍 全域 OCR 監控 (LV | COIN | EXP)")
        self.debug_info_lbl.setStyleSheet("color: #888; font-size: 10px; font-weight: bold;")
        
        self.debug_lv_img_lbl = QLabel()
        self.debug_lv_img_lbl.setFixedHeight(40)
        self.debug_lv_img_lbl.setStyleSheet("border: 1px solid #444; background: #222; padding: 2px;")
        self.debug_lv_img_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        self.debug_coin_img_lbl = QLabel()
        self.debug_coin_img_lbl.setFixedHeight(40)
        self.debug_coin_img_lbl.setStyleSheet("border: 1px solid #444; background: #222; padding: 2px;")
        self.debug_coin_img_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        self.debug_batch_img_lbl = QLabel()
        self.debug_batch_img_lbl.setFixedSize(380, 110)
        self.debug_batch_img_lbl.setStyleSheet("border: 1px solid #444; background: #000; padding: 5px;")
        self.debug_batch_img_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        status_row = QHBoxLayout()
        self.debug_global_conf_lbl = QLabel("OCR Confidence: --%")
        self.debug_global_conf_lbl.setStyleSheet("color: #00ffff; font-family: Consolas; font-size: 11px;")
        self.debug_lv_stats_lbl = QLabel("LATEST LV: --")
        self.debug_lv_stats_lbl.setStyleSheet("color: #ffd700; font-family: Consolas; font-size: 11px;")
        
        status_row.addWidget(self.debug_global_conf_lbl)
        status_row.addStretch()
        status_row.addWidget(self.debug_lv_stats_lbl)
        
        self.debug_layout.addWidget(self.debug_info_lbl)
        self.debug_layout.addWidget(self.debug_lv_img_lbl)
        self.debug_layout.addWidget(self.debug_coin_img_lbl)
        self.debug_layout.addWidget(self.debug_batch_img_lbl)
        self.debug_layout.addLayout(status_row)
        
        exp_tab_layout.addWidget(self.debug_group)
        exp_tab_layout.addStretch()
        
        # 初始狀態設定
        if self.overlay: 
            self.debug_mode_cb.setChecked(self.overlay.show_debug)
            self.debug_group.setVisible(self.overlay.show_debug)
        self.debug_mode_cb.toggled.connect(self.on_debug_mode_changed)

        self.tabs.addTab(exp_tab, "📊 經驗值/楓幣")

        # 分頁 3: 羅茱 YzY 同步
        self.rjpq_client = RJPQSyncClient()
        self.rjpq_tab = RJPQTabContent(self.rjpq_client)
        if self.overlay:
            self.rjpq_tab.color_selected.connect(self.overlay.set_rjpq_color)
            self.rjpq_client.sync_received.connect(self.overlay.update_rjpq_data)
            self.rjpq_client.error_received.connect(
                lambda msg: self.overlay.show_notification(f"❌ YZY: {msg}")
            )
            self.rjpq_client.overlay_toggle_request.connect(
                self.overlay.set_rjpq_overlay_visible
            )
            self.overlay.rjpq_cell_clicked.connect(self.rjpq_tab.platform_clicked)

        move_rjpq_btn = QPushButton("🔱 調整羅茱面板位置")
        move_rjpq_btn.setStyleSheet(btn_common_style)
        move_rjpq_btn.clicked.connect(self.toggle_rjpq_handle)
        self.rjpq_tab.main_layout.insertWidget(2, move_rjpq_btn)
        self.tabs.addTab(self.rjpq_tab, "🎮 羅茱 YzY")

        # 分頁 4: 系統 / 熱鍵設定
        sys_tab = QWidget()
        sys_layout = QVBoxLayout(sys_tab)
        hk_grid = QGridLayout()
        hk_labels = {
            "exp_toggle": "📊 顯示/隱藏經驗面板",
            "exp_pause": "⏸ 暫停/恢復紀錄 (F11)",
            "reset": "🧹 重置清空所有計時器 (F9)",
            "exp_report": "📸 產出經驗成果圖 (F12)",
            "rjpq_1": "🎮 羅茱 - 標記位置 1",
            "rjpq_2": "🎮 羅茱 - 標記位置 2",
            "rjpq_3": "🎮 羅茱 - 標記位置 3",
            "rjpq_4": "🎮 羅茱 - 標記位置 4",
            "show_settings": "🍁 顯示/隱藏控制中心",
        }
        hotkeys = config.get("hotkeys", {})
        # 繪製全域熱鍵列表
        for idx, (hk_id, txt) in enumerate(hk_labels.items()):
            hk_grid.addWidget(QLabel(txt), idx, 0)
            raw_val = hotkeys.get(hk_id, "None").upper()
            btn = QPushButton("無" if raw_val == "NONE" else raw_val)
            btn.setFixedWidth(100)
            btn.setStyleSheet(btn_common_style)
            btn.clicked.connect(lambda checked, h=hk_id: self.start_recording_global(h))
            hk_grid.addWidget(btn, idx, 1)
            self.global_hk_buttons[hk_id] = btn
        sys_layout.addLayout(hk_grid)
        sys_layout.addStretch()
        
        # 底部資訊與檢查更新
        credit_lbl = QLabel("✨ 由 ALiangLiang 傾心製作 ❤️")
        credit_lbl.setOpenExternalLinks(True)
        credit_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        sys_layout.addWidget(credit_lbl)
        update_btn = QPushButton(f"🔍 檢查更新 (目前版本: {VERSION})")
        update_btn.setStyleSheet("background:transparent; color:#888;")
        update_btn.clicked.connect(lambda: self.overlay.controller.check_for_updates() if (self.overlay and self.overlay.controller) else None)
        sys_layout.addWidget(update_btn)
        
        self.update_banner = QLabel("")
        self.update_banner.setVisible(False)
        sys_layout.addWidget(self.update_banner)
        self.tabs.addTab(sys_tab, "⚙️ 設定")

        # 配置切換器 / 暱稱輸入
        self.update_profile_dropdown()

        # 儲存按鈕
        save_btn = QPushButton("💾 儲存並套用")
        save_btn.setStyleSheet(
            "background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #ffd54f, stop:1 #ffb300); color: #222; font-weight: bold; height: 40px;"
        )
        save_btn.clicked.connect(self.save_and_close)
        self.layout.addWidget(save_btn)

        exit_btn = QPushButton("🛑 關閉整個輔助")
        exit_btn.clicked.connect(QApplication.instance().quit)
        self.layout.addWidget(exit_btn)

    def start_recording_global(self, hk_id):
        self.is_recording = False
        self.recording_global_key = hk_id
        for btn in self.global_hk_buttons.values():
            btn.setText(btn.text().replace(" (錄製中...)", ""))
        self.global_hk_buttons[hk_id].setText("錄製中...")
        self.global_hk_buttons[hk_id].setStyleSheet(
            "background: #552222; color: #ff5555; border: 1px solid #ff0000;"
        )

    @override
    def keyPressEvent(self, event):
        key_code = event.key()
        is_escape = key_code == Qt.Key.Key_Escape
        if self.recording_global_key:
            key_name = "none" if is_escape else self.qt_key_to_name(event)
            if key_name:
                config = ConfigManager.load_config()
                config["hotkeys"][self.recording_global_key] = key_name
                ConfigManager.save_config(config)
                self.global_hk_buttons[self.recording_global_key].setText(
                    "無" if key_name == "none" else key_name.upper()
                )
                self.global_hk_buttons[self.recording_global_key].setStyleSheet("")
                self.recording_global_key = None
            return

        if self.is_recording:
            if is_escape:
                self.toggle_recording()
                return
            key_name = self.qt_key_to_name(event)
            if key_name:
                p_key = self.profile_box.currentData() or "F1"
                config = ConfigManager.load_config()
                triggers = self.capture_ui_data()
                config["profiles"][p_key]["triggers"] = triggers
                if key_name not in triggers:
                    config["profiles"][p_key]["triggers"][key_name] = {
                        "seconds": 300,
                        "icon": "",
                    }
                    ConfigManager.save_config(config)
                self.refresh_items()
                self.toggle_recording()

    def capture_ui_data(self):
        data = {}
        for k, ui in self.trigger_data.items():
            try:
                data[k] = {
                    "seconds": int(ui["inp"].text()),
                    "icon": ui["icon"],
                    "sound": ui["cb_sound"].isChecked(),
                }
            except:
                data[k] = {"seconds": 300, "icon": ui["icon"], "sound": True}
        return data

    def update_profile_dropdown(self):
        self.profile_box.blockSignals(True)
        self.profile_box.clear()
        config = ConfigManager.load_config()
        active = config.get("active_profile", "F1")
        for i in range(1, 10):
            key = f"F{i}"
            name = config["profiles"][key].get("name", f"Profile {i}")
            self.profile_box.addItem(f"{key}: {name}", key)
        for i in range(self.profile_box.count()):
            if self.profile_box.itemData(i) == active:
                self.profile_box.setCurrentIndex(i)
                self.nickname_inp.setText(config["profiles"][active].get("name", ""))
                break
        self.profile_box.blockSignals(False)
        self.profile_box.currentIndexChanged.connect(self.switch_profile_ui)

    def on_nickname_changed(self, text):
        key = self.profile_box.currentData()
        if not key:
            return
        config = ConfigManager.load_config()
        config["profiles"][key]["name"] = text
        ConfigManager.save_config(config)
        self.profile_box.blockSignals(True)
        self.profile_box.setItemText(self.profile_box.currentIndex(), f"{key}: {text}")
        self.profile_box.blockSignals(False)
        self.config_updated.emit()

    def switch_profile_ui(self, index):
        p_key = self.profile_box.itemData(index)
        config = ConfigManager.load_config()
        config["active_profile"] = p_key
        ConfigManager.save_config(config)
        self.nickname_inp.setText(config["profiles"][p_key].get("name", ""))
        self.refresh_items()
        self.config_updated.emit()

    def qt_key_to_name(self, event):
        code = event.key()
        is_numpad = bool(event.modifiers() & Qt.KeyboardModifier.KeypadModifier)
        special = {
            Qt.Key.Key_F1: "f1",
            Qt.Key.Key_F2: "f2",
            Qt.Key.Key_F3: "f3",
            Qt.Key.Key_F4: "f4",
            Qt.Key.Key_F5: "f5",
            Qt.Key.Key_F6: "f6",
            Qt.Key.Key_F7: "f7",
            Qt.Key.Key_F8: "f8",
            Qt.Key.Key_F9: "f9",
            Qt.Key.Key_F10: "f10",
            Qt.Key.Key_F11: "f11",
            Qt.Key.Key_F12: "f12",
            Qt.Key.Key_Shift: "shift",
            Qt.Key.Key_Control: "ctrl",
            Qt.Key.Key_Alt: "alt",
            Qt.Key.Key_Space: "space",
            Qt.Key.Key_Pause: "pause",
            Qt.Key.Key_Insert: "insert",
            Qt.Key.Key_Home: "home",
            Qt.Key.Key_End: "end",
            Qt.Key.Key_Delete: "delete",
            Qt.Key.Key_PageUp: "page_up",
            Qt.Key.Key_PageDown: "page_down",
        }
        if is_numpad:
            if Qt.Key.Key_0 <= code <= Qt.Key.Key_9:
                return f"num_{code - Qt.Key.Key_0}"
            if code == Qt.Key.Key_Period:
                return "num_dot"
        if code in special:
            return special[code]
        try:
            return chr(code).lower() if 32 <= code <= 126 else None
        except:
            return None

    def toggle_recording(self):
        self.is_recording = not self.is_recording
        if self.is_recording:
            self.record_btn.setText("🔴 請按鍵盤錄製中...")
            self.record_btn.setStyleSheet("background-color: #c62828; color: white;")
        else:
            self.record_btn.setText("➕ 新增按鍵 (點我後按鍵盤)")
            self.record_btn.setStyleSheet("")

    def safe_show(self):
        self.show()
        self.activateWindow()
        self.raise_()

    def toggle_timer_handle(self):
        if not self.handle:
            self.handle = PositionHandle()
            if self.overlay:
                self.handle.position_changed.connect(self.overlay.update_offset)
        if self.handle.isVisible():
            self.handle.hide()
        else:
            if self.overlay:
                geo = self.overlay.geometry()
                cx = geo.x() + geo.width() // 2 + self.overlay.x_offset
                cy = geo.y() + geo.height() // 2 + self.overlay.y_offset
                self.handle.move(cx - 30, cy - 30)
                self.overlay.show_preview = True
                self.overlay.update()
            self.handle.show()
            self.handle.raise_()

    def toggle_exp_handle(self):
        if not self.exp_handle:
            c = ConfigManager.load_config()
            ox, oy = c.get("exp_offset", [0, 0])
            self.exp_handle = PositionHandle()
            if self.overlay:
                self.exp_handle.move(
                    self.overlay.rect().center().x() + ox - 30,
                    self.overlay.rect().center().y() + oy - 150,
                )
                self.exp_handle.position_changed.connect(self.overlay.update_exp_offset)
            self.exp_handle.show()
        else:
            self.exp_handle.close()
            self.exp_handle = None
            self.save_and_close()

    def toggle_rjpq_handle(self):
        if not self.rjpq_handle:
            c = ConfigManager.load_config()
            ox, oy = c.get("rjpq_offset", [-200, 0])
            self.rjpq_handle = PositionHandle()
            center = self.overlay.mapToGlobal(self.overlay.rect().center())
            self.rjpq_handle.move(center.x() + ox - 30, center.y() + oy - 30)
            self.rjpq_handle.position_changed.connect(self.overlay.update_rjpq_offset)
            self.rjpq_handle.show()
            if self.overlay:
                self.overlay.show_rjpq_panel = True
                self.overlay.update()
        else:
            self.rjpq_handle.close()
            self.rjpq_handle = None
            self.save_and_close()

    def start_ship_timer(self):
        now = datetime.datetime.now()
        rem_min = 10 - (now.minute % 10)
        total_sec = (rem_min * 60) - now.second
        if total_sec <= 0:
            total_sec = 600
        if self.overlay:
            self.overlay.timer_request.emit(
                "Ship", total_sec, "buff_pngs/Others/ship_icon.png", True
            )

    def start_elevator_timer(self, dir):
        now = datetime.datetime.now()
        if dir == "down":
            total_sec = ((4 - (now.minute % 4)) * 60) - now.second
        else:
            total_sec = (120 - ((now.minute % 4) * 60 + now.second)) % 240
        if total_sec <= 0:
            total_sec = 240
        if self.overlay:
            icon = resource_path(f"buff_pngs/Others/elevator_{dir}.png")
            name = f"電梯({'下' if dir == 'down' else '上'})"
            self.overlay.timer_request.emit(name, total_sec, icon, True)

    def refresh_items(self):
        while self.scroll_layout.count():
            w = self.scroll_layout.takeAt(0).widget()
            if w:
                w.deleteLater()
        config = ConfigManager.load_config()
        p_key = self.profile_box.currentData() or "F1"
        p_data = config["profiles"].get(p_key, {"triggers": {}})
        self.trigger_data = {}
        for key, data in p_data["triggers"].items():
            if isinstance(data, int):
                data = {"seconds": data, "icon": ""}
            row = QFrame()
            row.setFixedHeight(50)
            row.setStyleSheet("background-color: #1e1e1e; border-radius: 4px;")
            layout = QHBoxLayout(row)
            lbl = QLabel(key.upper())
            lbl.setStyleSheet("color: #ffd700;")
            lbl.setFixedWidth(60)
            layout.addWidget(lbl)
            btn = QPushButton()
            btn.setFixedSize(32, 32)
            self.update_icon_button(btn, data.get("icon", ""))
            btn.clicked.connect(lambda checked, k=key, b=btn: self.pick_icon(k, b))
            layout.addWidget(btn)
            inp = QLineEdit(str(data.get("seconds", 300)))
            inp.setFixedWidth(35)
            layout.addWidget(inp)
            layout.addWidget(QLabel("秒"))
            cb = QCheckBox("20s音效")
            cb.setChecked(data.get("sound", True))
            layout.addWidget(cb)
            layout.addStretch()
            del_btn = QPushButton("🗑️")
            del_btn.setFixedWidth(30)
            del_btn.clicked.connect(lambda checked, k=key: self.delete_key(k))
            layout.addWidget(del_btn)
            self.trigger_data[key] = {
                "inp": inp,
                "icon": data.get("icon", ""),
                "cb_sound": cb,
            }
            self.scroll_layout.addWidget(row)
        self.scroll_layout.addStretch()

    def delete_key(self, key):
        config = ConfigManager.load_config()
        triggers = self.capture_ui_data()
        if key in triggers:
            del triggers[key]
            p_key = self.profile_box.currentData() or "F1"
            config["profiles"][p_key]["triggers"] = triggers
            ConfigManager.save_config(config)
            self.refresh_items()

    def update_icon_button(self, btn, path):
        real = resource_path(path) if path and not os.path.isabs(path) else path
        if real and os.path.exists(real):
            btn.setIcon(
                QIcon(
                    QPixmap(real).scaled(
                        24,
                        24,
                        Qt.AspectRatioMode.KeepAspectRatio,
                        Qt.TransformationMode.SmoothTransformation,
                    )
                )
            )
            btn.setIconSize(QSize(24, 24))
            btn.setText("")
        else:
            btn.setIcon(QIcon())
            btn.setText("🖼️")

    def pick_icon(self, key, btn):
        dlg = IconSelectorDialog(self)
        if dlg.exec():
            path = dlg.selected_icon
            self.trigger_data[key]["icon"] = path
            self.update_icon_button(btn, path)

    def on_money_toggle_changed(self, checked):
        if self.overlay:
            self.overlay.show_money_log = checked
            config = ConfigManager.load_config()
            config["show_money_log"] = checked
            ConfigManager.save_config(config)

    def on_exp_toggle_changed(self, checked):
        if self.overlay and self.overlay.show_exp_panel != checked:
            self.overlay.on_toggle_exp(checked)

    def on_reset_exp_clicked(self):
        if self.overlay: self.overlay.reset_exp_stats()

    def on_debug_mode_changed(self, checked):
        if self.overlay:
            self.overlay.show_debug = checked
            self.overlay.show_notification(f"除錯模式: {'開啟' if checked else '關閉'}")
        self.debug_group.setVisible(checked)

    def on_opacity_changed(self, v):
        self.opacity_val_lbl.setText(f"{v}%")
        if self.overlay:
            self.overlay.base_opacity = v / 100.0
            self.overlay.update()



    def save_and_close(self):
        config = ConfigManager.load_config()
        p_key = self.profile_box.currentData() or "F1"
        triggers = self.capture_ui_data()
        config["profiles"][p_key]["triggers"] = triggers
        config["profiles"][p_key]["name"] = self.nickname_inp.text()
        if self.overlay:
            config["offset"] = [self.overlay.x_offset, self.overlay.y_offset]
            config["exp_offset"] = [
                self.overlay.exp_x_offset,
                self.overlay.exp_y_offset,
            ]
            config["show_exp"] = self.overlay.show_exp_panel
            config["show_money_log"] = self.overlay.show_money_log
            config["show_debug"] = self.overlay.show_debug
            config["opacity"] = self.overlay.base_opacity
            config["rjpq_offset"] = [
                self.overlay.rjpq_x_offset,
                self.overlay.rjpq_y_offset,
            ]
        ConfigManager.save_config(config)
        self.config_updated.emit()
        self.hide()
        if self.handle:
            self.handle.hide()
        if self.exp_handle:
            self.exp_handle.hide()
        if self.rjpq_handle:
            self.rjpq_handle.hide()
        if self.overlay:
            self.overlay.show_preview = False
