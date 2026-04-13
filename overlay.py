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

# Initialize logger
logger = logging.getLogger(__name__)

def get_version():
    try:
        base_dir = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
        v_path = os.path.join(base_dir, "VERSION")
        with open(v_path, 'r', encoding='utf-8') as f:
            return f.read().strip()
    except:
        return "v?.?.?"

VERSION = get_version()
REPO_URL = "ALiangLiang/artale-agent"

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

CONFIG_FILE = "config.json"

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

class ConfigManager:
    @staticmethod
    def load_config():
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r", encoding='utf-8') as f:
                    config = json.load(f)
                    
                    # Migration: Old single profile -> Multi profile
                    if "profiles" not in config:
                        old_triggers = config.get("triggers", {"f1": {"seconds": 300, "icon": ""}})
                        old_offset = config.get("offset", [0, 0])
                        config = {
                            "active_profile": "F1",
                            "offset": old_offset,
                            "profiles": {
                                "F1": {"name": "預設配置", "triggers": old_triggers}
                            }
                        }
                    
                    # Migration: "Profile X" -> "FX"
                    new_profiles = {}
                    old_profiles = config.get("profiles", {})
                    for i in range(1, 9):
                        old_key = f"Profile {i}"
                        new_key = f"F{i}"
                        
                        if old_key in old_profiles:
                            data = old_profiles[old_key]
                            if "name" not in data or data["name"] == old_key:
                                data["name"] = f"切換組 {new_key}"
                            new_profiles[new_key] = data
                        elif new_key in old_profiles:
                            new_profiles[new_key] = old_profiles[new_key]
                        else:
                            new_profiles[new_key] = {"name": f"切換組 {new_key}", "triggers": {}}
                    
                    config["profiles"] = new_profiles
                    
                    # Sanitize active_profile
                    if config.get("active_profile", "").startswith("Profile "):
                        num = config["active_profile"].split(" ")[1]
                        config["active_profile"] = f"F{num}"
                    
                    if config.get("active_profile") not in config["profiles"]:
                        config["active_profile"] = "F1"

                    # Ensure multiple offsets exist
                    if "offset" not in config:
                        config["offset"] = [0, 0]
                    if "exp_offset" not in config:
                        config["exp_offset"] = [0, 0]

                    # Ensure migration for older trigger formats inside profiles
                    for p in config["profiles"].values():
                        if "name" not in p:
                            p["name"] = "未命名"
                        for k, v in p["triggers"].items():
                            if isinstance(v, (int, float)):
                                p["triggers"][k] = {"seconds": int(v), "icon": "", "sound": True}
                            if "sound" not in p["triggers"][k]:
                                p["triggers"][k]["sound"] = True
                    
                    if "opacity" not in config:
                        config["opacity"] = 0.5
                    
                    default_hks = {
                        "exp_toggle": "f10",
                        "exp_pause": "f11",
                        "reset": "f9",
                        "exp_report": "f12",
                        "rjpq_1": "num_1",
                        "rjpq_2": "num_2",
                        "rjpq_3": "num_3",
                        "rjpq_4": "num_4",
                        "show_settings": "pause"
                    }
                    if "hotkeys" not in config:
                        config["hotkeys"] = default_hks
                    else:
                        # Ensure missing keys are added
                        for k, v in default_hks.items():
                            if k not in config["hotkeys"]:
                                config["hotkeys"][k] = v

                    return config
            except Exception as e: 
                logger.error(f"Error loading config: {e}")
                pass
            
        # Default Multi-Profile Config
        default_profiles = {}
        for i in range(1, 10):
            default_profiles[f"F{i}"] = {"name": f"切換組 F{i}", "triggers": {}}
        return {
            "active_profile": "F1", 
            "offset": [0, 0], 
            "opacity": 0.5, 
            "profiles": default_profiles,
            "hotkeys": {
                "exp_toggle": "f10",
                "exp_pause": "f11",
                "reset": "f12",
                "rjpq_1": "1",
                "rjpq_2": "2",
                "rjpq_3": "3",
                "rjpq_4": "4"
            }
        }

    @staticmethod
    def save_config(config):
        with open(CONFIG_FILE, "w", encoding='utf-8') as f:
            json.dump(config, f, ensure_ascii=False, indent=2)

class IconSelectorDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("選擇技能圖示")
        self.setFixedSize(600, 500)
        self.selected_icon = None
        
        layout = QVBoxLayout(self)
        self.setStyleSheet("""
            QDialog { background-color: #121212; color: #e0e0e0; }
            QTabWidget::pane { border: 1px solid #333; }
            QTabBar::tab { background: #222; color: #888; padding: 8px 15px; }
            QTabBar::tab:selected { background: #333; color: #ffd700; }
            QScrollArea { border: none; background: transparent; }
        """)
        
        self.tabs = QTabWidget()
        layout.addWidget(self.tabs)
        
        priority = ["Warrior", "Magician", "Bowman", "Thief", "Pirate", "Common", "buff_items", "Others"]
        
        base_path = resource_path("buff_pngs")
        if os.path.exists(base_path):
            all_dirs = [d for d in os.listdir(base_path) if os.path.isdir(os.path.join(base_path, d)) and d != "Gray"]
            
            # Sort categories by priority, then alphabetically for unknown ones
            categories = []
            for p in priority:
                if p in all_dirs:
                    categories.append(p)
                    all_dirs.remove(p)
            categories.extend(sorted(all_dirs))
            
            cat_map = {
                "Warrior": "劍士", "Magician": "法師", "Bowman": "弓箭手",
                "Thief": "盜賊", "Pirate": "海盜", "Common": "共通",
                "buff_items": "消耗品", "Others": "其他"
            }
            
            for cat in categories:
                scroll = QScrollArea()
                container = QWidget()
                grid = QGridLayout(container)
                grid.setSpacing(5)
                
                cat_path = os.path.join(base_path, cat)
                icons = [f for f in os.listdir(cat_path) if f.endswith(".png")]
                
                col = 0
                row = 0
                for icon_file in sorted(icons):
                    icon_path = os.path.join(cat_path, icon_file)
                    btn = QPushButton()
                    btn.setFixedSize(50, 50)
                    pixmap = QPixmap(icon_path).scaled(40, 40, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation)
                    btn.setIcon(QIcon(pixmap))
                    btn.setIconSize(QSize(40, 40))
                    btn.setStyleSheet("QPushButton { background: #1e1e1e; border-radius: 4px; } QPushButton:hover { background: #333; border: 1px solid #ffd700; }")
                    
                    btn.clicked.connect(lambda checked, p=icon_path: self.select_icon(p))
                    
                    grid.addWidget(btn, row, col)
                    col += 1
                    if col >= 8:
                        col = 0
                        row += 1
                
                grid.setRowStretch(row + 1, 1)
                container.setLayout(grid)
                scroll.setWidget(container)
                scroll.setWidgetResizable(True)
                display_name = cat_map.get(cat, cat)
                self.tabs.addTab(scroll, display_name)
        
        cancel_btn = QPushButton("取消")
        cancel_btn.clicked.connect(self.reject)
        layout.addWidget(cancel_btn)

    def select_icon(self, path):
        abs_path = os.path.abspath(path)
        base_dir = os.path.abspath(".")
        
        try:
            if abs_path.lower().startswith(base_dir.lower()):
                rel = os.path.relpath(abs_path, base_dir)
                self.selected_icon = rel.replace("\\", "/")
            else:
                self.selected_icon = abs_path
        except:
            self.selected_icon = abs_path
        self.accept()

class PositionHandle(QWidget):
    position_changed = pyqtSignal(int, int)
    
    def __init__(self, icon_path="buff_pngs/arrow.png"):
        super().__init__()
        self.setWindowFlags(Qt.WindowType.WindowStaysOnTopHint | Qt.WindowType.FramelessWindowHint | Qt.WindowType.Tool)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.setFixedSize(60, 60)
        
        self.lbl = QLabel(self)
        self.lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        real_p = resource_path(icon_path)
        if os.path.exists(real_p):
            self.pixmap = QPixmap(real_p).scaled(50, 50, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation)
            self.lbl.setPixmap(self.pixmap)
        
        self.lbl.setFixedSize(60, 60)
        self.lbl.setStyleSheet("background: rgba(255, 255, 255, 50); border: 1px dashed white; border-radius: 5px;")
        
        self._dragging = False
        self._drag_start = QPoint()

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self._dragging = True
            self._drag_start = event.globalPosition().toPoint() - self.frameGeometry().topLeft()
            event.accept()

    def mouseMoveEvent(self, event):
        if self._dragging and event.buttons() & Qt.MouseButton.LeftButton:
            self.move(event.globalPosition().toPoint() - self._drag_start)
            self.emit_offset()
            event.accept()

    def mouseReleaseEvent(self, event):
        self._dragging = False
        self.emit_offset()

    def emit_offset(self):
        gp = self.mapToGlobal(self.rect().center())
        self.position_changed.emit(gp.x(), gp.y())

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
        self.overlay = overlay
        self.trigger_data = {}
        self.is_recording = False
        self.recording_global_key = None # Which global hotkey we are recording
        self.global_hk_buttons = {} # Store button refs
        
        # Connect to EXP updates for live preview
        if self.overlay:
            # Use visual request for UI images
            self.overlay.exp_visual_request.connect(self.update_debug_img)
            self.overlay.lv_update_request.connect(self.update_lv_debug_img)
            self.overlay.update_found.connect(self.show_update_banner)
            
        self.handle = None
        self.exp_handle = None
        self.rjpq_handle = None
        
        self.init_ui()
        self.request_show.connect(self.safe_show)

    def update_debug_img(self, data):
        if not data: return
        thresh = data.get("thresh") # This will now be the processed_img
        if thresh is not None:
            # We don't need to invert anymore because processed_img is already Black-on-White
            h, w = thresh.shape
            bytes_per_line = w
            q_img = QImage(thresh.data, w, h, bytes_per_line, QImage.Format.Format_Grayscale8)
            pixmap = QPixmap.fromImage(q_img)
            self.debug_img_lbl.setPixmap(pixmap.scaled(self.debug_img_lbl.size(), Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation))
        
        # Display simulated confidence for EXP
        exp_conf = data.get("conf", 0)
        self.debug_exp_conf_lbl.setText(f"Conf: {exp_conf:.0f}%")
        exp_conf_color = "#51cf66" if exp_conf >= 100 else ("#ffd700" if exp_conf >= 50 else "#ff6b6b")
        self.debug_exp_conf_lbl.setStyleSheet(f"color: {exp_conf_color}; font-family: Consolas; font-weight: bold; font-size: 13px;")

    def update_lv_debug_img(self, data):
        if not data: return
        thresh = data.get("thresh") # This will now be the lv_processed
        if thresh is not None:
            h, w = thresh.shape
            q_img = QImage(thresh.data, w, h, w, QImage.Format.Format_Grayscale8)
            pixmap = QPixmap.fromImage(q_img)
            self.debug_lv_img_lbl.setPixmap(pixmap.scaled(self.debug_lv_img_lbl.size(), Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation))
        
        lv_text = data.get("level")
        lv_conf = data.get("conf", 0)
        
        # Display simulated confidence based on consensus
        self.debug_lv_conf_lbl.setText(f"Conf: {lv_conf:.0f}%")
        conf_color = "#51cf66" if lv_conf >= 100 else ("#ffd700" if lv_conf >= 50 else "#ff6b6b")
        self.debug_lv_conf_lbl.setStyleSheet(f"color: {conf_color}; font-family: Consolas; font-weight: bold; font-size: 13px;")

        # Only act if consensus reached (e.g. 100% = 3 consecutive identical frames)
        if lv_text and lv_conf >= 100 and self.overlay: 
            try:
                lv_val = int(lv_text)
                self.overlay.current_lv = f"LV.{lv_val}"
                
                if self.overlay.last_confirmed_lv is not None:
                    level_diff = lv_val - self.overlay.last_confirmed_lv
                    # Reasonable Jump Check
                    if 0 < level_diff <= 2:
                        logger.info(f"[ExpTracker] Confirmed Level UP! {self.overlay.last_confirmed_lv} -> {lv_val}")
                        self.overlay.exp_initial_val = None
                        self.overlay.cumulative_gain = 0
                        self.overlay.exp_history = []
                    elif level_diff != 0:
                        logger.debug(f"[ExpTracker] Filtered out unreasonable jump: {self.overlay.last_confirmed_lv} -> {lv_val}")
                
                self.overlay.last_confirmed_lv = lv_val
            except:
                pass

    def show_update_banner(self, tag, url):
        if hasattr(self, 'update_banner'):
            self.update_banner.setText(f'✨ <a href="{url}" style="color: #ffdd00; text-decoration: underline;">發現新版本 {tag}！點此前往下載更新</a>')
            self.update_banner.setVisible(True)

    def showEvent(self, event):
        super().showEvent(event)
        self.refresh_items()
        if self.overlay:
            # Check cached version info
            if hasattr(self.overlay, '_latest_version_info') and self.overlay._latest_version_info:
                self.show_update_banner(*self.overlay._latest_version_info)
            else:
                # Auto check in silent/auto mode
                self.overlay.check_for_updates(auto=True)

    def init_ui(self):
        self.setWindowTitle("Artale 瑞士刀")
        # Set Window Icon
        icon_path = resource_path("app_icon.png")
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))
            
        config = ConfigManager.load_config()
        self.layout = QVBoxLayout(self)
        self.setStyleSheet("background-color: #121212; color: #e0e0e0; font-family: 'Microsoft JhengHei', '微軟正黑體', sans-serif;")
        
        # (Redundant body title removed as it matches window title-bar)
        
        self.tabs = QTabWidget()
        self.tabs.setStyleSheet("""
            QTabWidget::pane { border: 1px solid #333; background: #121212; }
            QTabBar::tab { background: #222; color: #888; padding: 10px 4px; min-width: 85px; font-size: 11px; font-family: 'Microsoft JhengHei', '微軟正黑體', sans-serif; }
            QTabBar::tab:selected { background: #333; color: #ffd700; font-weight: bold; }
        """)
        self.layout.addWidget(self.tabs)

        # --- Tab 1: Timer Settings ---
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
        self.nickname_inp.setStyleSheet("""
            QLineEdit { background-color: #222; border: 1px solid #444; border-radius: 4px; padding: 5px; color: #fff; }
            QLineEdit:focus { border: 1px solid #ffd700; }
        """)
        self.nickname_inp.textChanged.connect(self.on_nickname_changed)

        profile_row.addWidget(self.profile_box, 1)
        profile_row.addWidget(self.nickname_inp, 2)
        timer_tab_layout.addLayout(profile_row)

        subtitle = QLabel("貼心提醒：雙擊 F1~F9 可快速切換配置組")
        subtitle.setStyleSheet("color: #888; font-size: 11px; margin-bottom: 5px;")
        timer_tab_layout.addWidget(subtitle)

        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setFrameShape(QFrame.Shape.NoFrame)
        self.scroll_content = QWidget()
        self.scroll_layout = QVBoxLayout(self.scroll_content)
        self.scroll_layout.setSpacing(2)
        self.scroll_layout.setContentsMargins(0, 0, 0, 0)
        self.scroll.setWidget(self.scroll_content)
        timer_tab_layout.addWidget(self.scroll)

        btn_common_style = """
            QPushButton {
                background-color: #2a2a2a;
                color: #ccc;
                border: 1px solid #3d3d3d;
                border-radius: 4px;
                height: 32px;
                font-weight: bold;
                font-family: 'Microsoft JhengHei', '微軟正黑體', sans-serif;
            }
            QPushButton:hover {
                background-color: #333;
                border: 1px solid #555;
            }
        """

        self.record_btn = QPushButton("➕ 新增按鍵 (點我後按鍵盤)")
        self.record_btn.setStyleSheet(btn_common_style)
        self.record_btn.clicked.connect(self.toggle_recording)
        timer_tab_layout.addWidget(self.record_btn)
        
        # --- Ship Reminder Section ---
        ship_group = QGroupBox("🚢 特殊提醒")
        ship_group.setStyleSheet("QGroupBox { color: #aaa; font-weight: bold; border: 1px solid #333; border-radius: 8px; margin-top: 15px; padding-top: 10px; }")
        ship_layout = QVBoxLayout()
        
        ship_btn = QPushButton("🚢 開始下班船班倒數")
        ship_btn.setStyleSheet(btn_common_style)
        ship_btn.clicked.connect(self.start_ship_timer)
        ship_layout.addWidget(ship_btn)
        
        elevator_row = QHBoxLayout()
        elevator_down_btn = QPushButton("🔻 赫爾奧斯塔 (下樓)")
        elevator_down_btn.setStyleSheet(btn_common_style)
        elevator_down_btn.clicked.connect(lambda: self.start_elevator_timer("down"))
        
        elevator_up_btn = QPushButton("🔺 赫爾奧斯塔 (上樓)")
        elevator_up_btn.setStyleSheet(btn_common_style)
        elevator_up_btn.clicked.connect(lambda: self.start_elevator_timer("up"))
        
        elevator_row.addWidget(elevator_down_btn)
        elevator_row.addWidget(elevator_up_btn)
        ship_layout.addLayout(elevator_row)
        
        ship_group.setLayout(ship_layout)
        timer_tab_layout.addWidget(ship_group)
        
        # --- Timer Position Adjustment ---
        pos_btn = QPushButton("🔱 調整計時器位置")
        pos_btn.setStyleSheet(btn_common_style)
        pos_btn.clicked.connect(self.toggle_timer_handle)
        timer_tab_layout.addWidget(pos_btn)
        
        self.tabs.addTab(timer_tab, "⏲️ 計時器")

        # --- Tab 2: EXP Settings ---
        exp_tab = QWidget()
        exp_tab_layout = QVBoxLayout(exp_tab)
        
        exp_info = QLabel("📊 經驗值追蹤設定")
        exp_info.setStyleSheet("color: #ffd700; font-weight: bold; font-size: 14px; margin-top: 10px;")
        exp_tab_layout.addWidget(exp_info)
        
        self.exp_active_cb = QCheckBox("開啟經驗值監測面板 (Hotkey: F10)")
        self.exp_active_cb.setStyleSheet("color: #ccc; margin-top: 10px;")
        if self.overlay:
            self.exp_active_cb.setChecked(self.overlay.show_exp_panel)
        self.exp_active_cb.toggled.connect(self.on_exp_toggle_changed)
        exp_tab_layout.addWidget(self.exp_active_cb)
        
        self.money_active_cb = QCheckBox("開啟金錢記錄追蹤 (未開放)")
        self.money_active_cb.setStyleSheet("color: #666; margin-top: 5px;")
        self.money_active_cb.setChecked(False)
        self.money_active_cb.setEnabled(False)
        exp_tab_layout.addWidget(self.money_active_cb)
        
        reset_exp_btn = QPushButton("🔄 重新開始紀錄 (歸零重算)")
        reset_exp_btn.setStyleSheet(btn_common_style)
        reset_exp_btn.clicked.connect(self.on_reset_exp_clicked)
        exp_tab_layout.addWidget(reset_exp_btn)
        
        # Opacity Slider (General setting)
        opacity_info = QLabel("✨ 面板背景透明度")
        opacity_info.setStyleSheet("color: #aaa; font-size: 11px; margin-top: 15px;")
        exp_tab_layout.addWidget(opacity_info)
        
        opacity_row = QHBoxLayout()
        self.opacity_val_lbl = QLabel(f"{int(config.get('opacity', 0.5)*100)}%")
        self.opacity_val_lbl.setStyleSheet("color: #ffd700; font-weight: bold;")
        
        self.opacity_slider = QSlider(Qt.Orientation.Horizontal)
        self.opacity_slider.setRange(10, 100)
        self.opacity_slider.setValue(int(config.get("opacity", 0.5) * 100))
        self.opacity_slider.valueChanged.connect(self.on_opacity_changed)
        
        opacity_row.addWidget(self.opacity_slider)
        opacity_row.addWidget(self.opacity_val_lbl)
        exp_tab_layout.addLayout(opacity_row)
        
        # --- EXP Position Adjustment ---
        exp_pos_btn = QPushButton("📊 調整經驗值面板位置")
        exp_pos_btn.setStyleSheet(btn_common_style)
        exp_pos_btn.clicked.connect(self.toggle_exp_handle)
        exp_tab_layout.addWidget(exp_pos_btn)
        
        # Export Report Button
        self.export_btn = QPushButton("📸 產出成果圖 (截圖分享)")
        self.export_btn.setStyleSheet(btn_common_style)
        self.export_btn.clicked.connect(self.overlay.export_exp_report if self.overlay else lambda: None)
        exp_tab_layout.addWidget(self.export_btn)
        
        # Add Stretch to push debug info to the VERY BOTTOM
        exp_tab_layout.addStretch()

        # Show Debug Messages Checkbox (At bottom)
        self.debug_mode_cb = QCheckBox("顯示除錯訊息 (開發者模式)")
        self.debug_mode_cb.setStyleSheet("color: #888; font-size: 11px; margin-top: 10px;")
        if self.overlay:
            self.debug_mode_cb.setChecked(self.overlay.show_debug)
        self.debug_mode_cb.toggled.connect(self.on_debug_mode_changed)
        exp_tab_layout.addWidget(self.debug_mode_cb)

        # Debug OCR Image (Hidden by default, shown only when checkbox is on)
        self.debug_info_lbl = QLabel("🔍 OCR 監控 (白底黑字則為正常)")
        self.debug_info_lbl.setStyleSheet("color: #666; font-size: 10px; margin-top: 5px;")
        self.debug_info_lbl.setVisible(self.debug_mode_cb.isChecked())
        exp_tab_layout.addWidget(self.debug_info_lbl)
        
        debug_row = QHBoxLayout()
        self.debug_img_lbl = QLabel()
        self.debug_img_lbl.setFixedSize(250, 30)
        self.debug_img_lbl.setStyleSheet("border: 1px solid #444; background: #000;")
        self.debug_img_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.debug_img_lbl.setVisible(self.debug_mode_cb.isChecked())
        
        self.debug_exp_conf_lbl = QLabel("Conf: --%")
        self.debug_exp_conf_lbl.setStyleSheet("color: #00ffff; font-family: Consolas; font-weight: bold; font-size: 12px; margin-left: 10px;")
        self.debug_exp_conf_lbl.setVisible(self.debug_mode_cb.isChecked())
        
        debug_row.addWidget(self.debug_img_lbl)
        debug_row.addWidget(self.debug_exp_conf_lbl)
        debug_row.addStretch()
        exp_tab_layout.addLayout(debug_row)

        # LV OCR Monitor
        self.debug_lv_info_lbl = QLabel("🔍 LV 監控 (白底黑字則為正常)")
        self.debug_lv_info_lbl.setStyleSheet("color: #666; font-size: 10px; margin-top: 10px;")
        self.debug_lv_info_lbl.setVisible(self.debug_mode_cb.isChecked())
        exp_tab_layout.addWidget(self.debug_lv_info_lbl)
        
        lv_row = QHBoxLayout()
        self.debug_lv_img_lbl = QLabel()
        self.debug_lv_img_lbl.setFixedSize(150, 40)
        self.debug_lv_img_lbl.setStyleSheet("border: 1px solid #444; background: #000;")
        self.debug_lv_img_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.debug_lv_img_lbl.setVisible(self.debug_mode_cb.isChecked())
        
        self.debug_lv_conf_lbl = QLabel("Conf: --%")
        self.debug_lv_conf_lbl.setStyleSheet("color: #00ffff; font-family: Consolas; font-weight: bold; font-size: 12px; margin-left: 10px;")
        self.debug_lv_conf_lbl.setVisible(self.debug_mode_cb.isChecked())
        
        lv_row.addWidget(self.debug_lv_img_lbl)
        lv_row.addWidget(self.debug_lv_conf_lbl)
        lv_row.addStretch()
        exp_tab_layout.addLayout(lv_row)



        exp_tab_layout.addStretch()
        self.tabs.addTab(exp_tab, "📊 經驗值")

        # --- Tab 3: RJPQ YzY ---
        self.rjpq_client = RJPQSyncClient()
        self.rjpq_tab = RJPQTabContent(self.rjpq_client)
        self.rjpq_tab.color_selected.connect(self.overlay.set_rjpq_color)
        self.rjpq_client.sync_received.connect(self.overlay.update_rjpq_data)
        self.rjpq_client.overlay_toggle_request.connect(self.overlay.set_rjpq_overlay_visible)
        self.overlay.rjpq_cell_clicked.connect(self.rjpq_tab.platform_clicked)
        
        # Add Drag button to RJPQ tab
        move_rjpq_btn = QPushButton("🔱 調整羅茱面板位置")
        move_rjpq_btn.setStyleSheet(btn_common_style)
        move_rjpq_btn.clicked.connect(self.toggle_rjpq_handle)
        self.rjpq_tab.main_layout.insertWidget(2, move_rjpq_btn) # Insert after toggle

        # Sync checkbox state
        self.rjpq_tab.overlay_cb.setChecked(self.overlay.show_rjpq_panel)
        self.tabs.addTab(self.rjpq_tab, "🎮 羅茱 YzY")
        
        # --- Tab 4: System Settings ---
        sys_tab = QWidget()
        sys_tab_layout = QVBoxLayout(sys_tab)
        
        sys_info = QLabel("⚙️ 輔助全局快捷鍵設定")
        sys_info.setStyleSheet("color: #ffd700; font-weight: bold; font-size: 14px; margin-top: 10px;")
        sys_tab_layout.addWidget(sys_info)
        
        sys_desc = QLabel("點擊下方按鈕後，按下鍵盤上的按鍵即可變更快捷鍵。")
        sys_desc.setStyleSheet("color: #888; font-size: 11px;")
        sys_tab_layout.addWidget(sys_desc)
        
        # Hotkey Rows
        hk_grid = QGridLayout()
        hk_labels = {
            "exp_toggle": "📊 顯示/隱藏經驗面板",
            "exp_pause": "⏸ 暫停/恢復紀錄 (F11)",
            "reset": "🧹 重置清空所有計時器 (F9)",
            "exp_report": "📸 產出經驗成果圖 (F12)",
            "rjpq_1": "🎮 羅茱 - 標記位置 1 (Num1)",
            "rjpq_2": "🎮 羅茱 - 標記位置 2 (Num2)",
            "rjpq_3": "🎮 羅茱 - 標記位置 3 (Num3)",
            "rjpq_4": "🎮 羅茱 - 標記位置 4 (Num4)",
            "show_settings": "🍁 顯示/隱藏控制中心"
        }
        
        config = ConfigManager.load_config()
        hotkeys = config.get("hotkeys", {})
        
        for idx, (hk_id, label_text) in enumerate(hk_labels.items()):
            lbl = QLabel(label_text)
            lbl.setStyleSheet("color: #ccc; font-size: 12px;")
            hk_grid.addWidget(lbl, idx, 0)
            
            raw_val = hotkeys.get(hk_id, "None").upper()
            display_val = "無" if raw_val == "NONE" else raw_val
            btn = QPushButton(display_val)
            btn.setFixedWidth(100)
            btn.setStyleSheet("""
                QPushButton { background: #333; color: #fff; border: 1px solid #555; border-radius: 4px; padding: 5px; font-weight: bold; font-family: 'Microsoft JhengHei', '微軟正黑體', sans-serif; }
                QPushButton:hover { background: #444; border-color: #ffd700; }
            """)
            btn.clicked.connect(lambda checked, h=hk_id: self.start_recording_global(h))
            hk_grid.addWidget(btn, idx, 1)
            self.global_hk_buttons[hk_id] = btn
            
        sys_tab_layout.addLayout(hk_grid)
        sys_tab_layout.addStretch()
        
        credit_lbl = QLabel('✨ 由 <a href="https://github.com/ALiangLiang" style="color: #aaa; text-decoration: none;">ALiangLiang</a> 傾心製作 ❤️ | <a href="https://github.com/ALiangLiang/artale-agent" style="color: #88ccff; text-decoration: none;">🔗 原始碼</a> | <a href="https://buymeacoffee.com/aliangliang" style="color: #ffdd00; text-decoration: none;">🍵 請開發者喝杯茶</a>')
        credit_lbl.setStyleSheet("color: #666; font-size: 11px;")
        credit_lbl.setOpenExternalLinks(True)
        credit_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        sys_tab_layout.addWidget(credit_lbl)
        
        # Check Update Button
        update_btn = QPushButton(f"🔍 檢查更新 (目前版本: {VERSION})")
        update_btn.setStyleSheet("""
            QPushButton {
                background: transparent; color: #888; border: none; font-size: 11px; margin-top: 5px;
            }
            QPushButton:hover { color: #fff; text-decoration: underline; }
        """)
        update_btn.clicked.connect(self.overlay.check_for_updates if self.overlay else lambda: None)
        sys_tab_layout.addWidget(update_btn)
        
        self.update_banner = QLabel("")
        self.update_banner.setStyleSheet("color: #ffdd00; font-size: 11px; margin-top: 5px;")
        self.update_banner.setOpenExternalLinks(True)
        self.update_banner.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.update_banner.setVisible(False)
        sys_tab_layout.addWidget(self.update_banner)
        
        self.tabs.addTab(sys_tab, "⚙️ 設定")

        # --- Global Controls (Bottom) ---

        self.update_profile_dropdown()

        self.refresh_items()

        # --- Global Controls (Bottom) ---

        save_btn = QPushButton("💾 儲存並套用")
        save_btn.setStyleSheet("""
            QPushButton {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #ffd54f, stop:1 #ffb300);
                color: #222;
                border-radius: 6px;
                font-weight: bold;
                font-size: 14px;
                height: 40px;
            }
            QPushButton:hover {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #ffea00, stop:1 #ff8f00);
            }
        """)
        save_btn.clicked.connect(self.save_and_close)
        self.layout.addWidget(save_btn)

        exit_btn = QPushButton("🛑 關閉整個輔助")
        exit_btn.setStyleSheet("QPushButton { background-color: transparent; color: #666; border: none; font-size: 10px; margin-top: 20px; }")
        exit_btn.clicked.connect(QApplication.instance().quit)
        self.layout.addWidget(exit_btn)
        


    def start_recording_global(self, hk_id):
        # Reset any active recording
        self.is_recording = False 
        self.recording_global_key = hk_id
        for btn in self.global_hk_buttons.values(): btn.setText(btn.text().replace(" (錄製中...)", ""))
        self.global_hk_buttons[hk_id].setText("錄製中...")
        self.global_hk_buttons[hk_id].setStyleSheet("background: #552222; color: #ff5555; border: 1px solid #ff0000; border-radius: 4px; padding: 5px; font-weight: bold;")

    def keyPressEvent(self, event):
        key_code = event.key()
        # Handle ESC as "None" (unbind)
        is_escape = (key_code == Qt.Key.Key_Escape)
        
        if self.recording_global_key:
            key_name = "none" if is_escape else self.qt_key_to_name(event)
            if key_name:
                config = ConfigManager.load_config()
                config["hotkeys"][self.recording_global_key] = key_name
                ConfigManager.save_config(config)
                display_name = "無" if key_name == "none" else key_name.upper()
                self.global_hk_buttons[self.recording_global_key].setText(display_name)
                self.global_hk_buttons[self.recording_global_key].setStyleSheet("background: #333; color: #fff; border: 1px solid #555; border-radius: 4px; padding: 5px; font-weight: bold;")
                self.recording_global_key = None
                if self.overlay and is_escape: self.overlay.show_notification("已清除該全域快捷鍵綁定")
            return

        if self.is_recording:
            if is_escape:
                self.toggle_recording()
                if self.overlay: self.overlay.show_notification("已取消新增按鍵")
                return
            key_name = self.qt_key_to_name(event)
            if key_name:
                p_key = self.profile_box.itemData(self.profile_box.currentIndex())
                if not p_key: p_key = "F1"
                current_ui_triggers = self.capture_ui_data()
                config = ConfigManager.load_config()
                config["profiles"][p_key]["triggers"] = current_ui_triggers
                if key_name not in config["profiles"][p_key]["triggers"]:
                    config["profiles"][p_key]["triggers"][key_name] = {"seconds": 300, "icon": ""}
                    ConfigManager.save_config(config)
                    if self.overlay: self.overlay.show_notification(f"已新增按鍵: {key_name.upper()}")
                else:
                    if self.overlay: self.overlay.show_notification(f"按鍵 {key_name.upper()} 已存在")
                self.refresh_items()
                self.toggle_recording()

    def capture_ui_data(self):
        active_ui_data = {}
        for key, ui in self.trigger_data.items():
            try:
                active_ui_data[key] = {
                    "seconds": int(ui["inp"].text()), 
                    "icon": ui["icon"],
                    "sound": ui["cb_sound"].isChecked()
                }
            except Exception as e:
                logger.debug(f"[Settings] UI capture failed for {key}: {e}")
                active_ui_data[key] = {"seconds": 300, "icon": ui["icon"], "sound": True}
        return active_ui_data

    def update_profile_dropdown(self):
        self.profile_box.blockSignals(True)
        current_idx = self.profile_box.currentIndex()
        self.profile_box.clear()
        config = ConfigManager.load_config()
        active = config.get("active_profile", "F1")
        for i in range(1, 9):
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
        if not key: return
        config = ConfigManager.load_config()
        config["profiles"][key]["name"] = text
        ConfigManager.save_config(config)
        self.profile_box.blockSignals(True)
        idx = self.profile_box.currentIndex()
        self.profile_box.setItemText(idx, f"{key}: {text}")
        self.profile_box.blockSignals(False)
        self.config_updated.emit()

    def switch_profile_ui(self, index):
        p_key = self.profile_box.itemData(index)
        if not p_key: return
        config = ConfigManager.load_config()
        config["active_profile"] = p_key
        ConfigManager.save_config(config)
        self.nickname_inp.setText(config["profiles"][p_key].get("name", ""))
        self.refresh_items()
        if self.overlay: self.overlay.load_profile_immediately()
        self.config_updated.emit()

    def qt_key_to_name(self, event):
        code = event.key()
        is_numpad = bool(event.modifiers() & Qt.KeyboardModifier.KeypadModifier)
        
        special_map = {
            Qt.Key.Key_F1: "f1", Qt.Key.Key_F2: "f2", Qt.Key.Key_F3: "f3", Qt.Key.Key_F4: "f4",
            Qt.Key.Key_F5: "f5", Qt.Key.Key_F6: "f6", Qt.Key.Key_F7: "f7", Qt.Key.Key_F8: "f8",
            Qt.Key.Key_F9: "f9", Qt.Key.Key_F10: "f10", Qt.Key.Key_F11: "f11", 
            Qt.Key.Key_F12: "f12", Qt.Key.Key_Shift: "shift", 
            Qt.Key.Key_Control: "ctrl", Qt.Key.Key_Alt: "alt", Qt.Key.Key_Space: "space",
            Qt.Key.Key_Pause: "pause"
        }
        
        # Priority 1: Check for Numpad 0-9 and Dot
        if is_numpad:
            if Qt.Key.Key_0 <= code <= Qt.Key.Key_9:
                return f"num_{code - Qt.Key.Key_0}"
            if code == Qt.Key.Key_Period or code == Qt.Key.Key_Comma or code == 0x0100002c: # Dot or Comma on keypad
                return "num_dot"
            # Fallback for Qt's specific keypad range if above didn't catch it
            if 0x01000020 <= code <= 0x01000029:
                return f"num_{code - 0x01000020}"
        
        if code in special_map: return special_map[code]
        try:
            name = chr(code).lower() if 32 <= code <= 126 else None
            return name
        except Exception as e:
            logger.debug(f"[Settings] Key conversion failed: {e}")
            return None
            
        if code in special_map: return special_map[code]
        try:
            name = chr(code).lower() if 32 <= code <= 126 else None
            return name
        except Exception as e:
            logger.debug(f"[Settings] Key conversion failed: {e}")
            return None

    def toggle_recording(self):
        self.is_recording = not self.is_recording
        if self.is_recording:
            self.record_btn.setText("🔴 請按鍵盤錄製中...")
            self.record_btn.setStyleSheet("background-color: #c62828; color: white; border-radius: 5px; height: 35px;")
        else:
            self.record_btn.setText("➕ 新增按鍵 (點我後按鍵盤)")
            self.record_btn.setStyleSheet("""
                QPushButton {
                    background-color: #333;
                    color: #ddd;
                    border: 1px solid #444;
                    border-radius: 6px;
                    font-weight: bold;
                    height: 35px;
                }
                QPushButton:hover {
                    background-color: #3d3d3d;
                    border: 1px solid #555;
                }
            """)

    def safe_show(self):
        self.show(); self.activateWindow(); self.raise_()
        
    def toggle_timer_handle(self):
        if not self.handle:
            self.handle = PositionHandle()
            if self.overlay: self.handle.position_changed.connect(self.overlay.update_offset)
        if self.handle.isVisible():
            self.handle.hide()
            if self.overlay: self.overlay.show_preview = False
        else:
            if self.overlay:
                geo = self.overlay.geometry()
                cx = geo.x() + geo.width() // 2 + self.overlay.x_offset
                cy = geo.y() + geo.height() // 2 + self.overlay.y_offset
                self.handle.move(cx - 30, cy - 30)
                self.overlay.show_preview = True; self.overlay.update()
            self.handle.show(); self.handle.raise_()

    def toggle_exp_handle(self):
        if not self.exp_handle:
            config = ConfigManager.load_config()
            ox, oy = config.get("exp_offset", [0, 0])
            self.exp_handle = PositionHandle()
            self.exp_handle.move(self.overlay.rect().center().x() + ox - 30, self.overlay.rect().center().y() + oy - 150)
            self.exp_handle.position_changed.connect(self.overlay.update_exp_offset)
            self.exp_handle.show()
        else:
            self.exp_handle.close()
            self.exp_handle = None
            self.save_and_close()

    def toggle_rjpq_handle(self):
        if not self.rjpq_handle:
            # Re-read to get latest
            config = ConfigManager.load_config()
            ox, oy = config.get("rjpq_offset", [-200, 0])
            self.rjpq_handle = PositionHandle()
            
            # Map center to global, then add offset
            center_global = self.overlay.mapToGlobal(self.overlay.rect().center())
            self.rjpq_handle.move(center_global.x() + ox - 30, center_global.y() + oy - 30)
            
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
        import datetime
        now = datetime.datetime.now()
        minutes = now.minute
        seconds = now.second
        
        # Victoria/Orbis: Every 10 mins (0, 10, 20 ...)
        rem_min = 10 - (minutes % 10)
        total_seconds = (rem_min * 60) - seconds
        if total_seconds <= 0: total_seconds = 600
        
        if self.overlay:
             self.overlay.timer_request.emit("Ship", total_seconds, "buff_pngs/Others/ship_icon.png", True)

    def start_elevator_timer(self, direction):
        import datetime
        now = datetime.datetime.now()
        minutes = now.minute
        seconds = now.second
        
        if direction == "down": # 4n: 0, 4, 8...
            rem_min = 4 - (minutes % 4)
            total_seconds = (rem_min * 60) - seconds
        else: # 4n+2: 2, 6, 10...
            # This is complex but: ( (2 - minutes) % 4 ) minutes remaining?
            # Let's use total cycle logic
            rem_total = (120 - ((minutes % 4) * 60 + seconds)) % 240
            total_seconds = rem_total
            
        if total_seconds <= 0:
            total_seconds = 240
            
        if self.overlay:
             icon = "buff_pngs/Others/elevator_down.png" if direction == "down" else "buff_pngs/Others/elevator_up.png"
             name = "電梯(下)" if direction == "down" else "電梯(上)"
             self.overlay.timer_request.emit(name, total_seconds, icon, True)
        
    def refresh_items(self):
        while self.scroll_layout.count():
            item = self.scroll_layout.takeAt(0)
            if item.widget(): item.widget().deleteLater()
        config = ConfigManager.load_config()
        p_key = self.profile_box.itemData(self.profile_box.currentIndex())
        if not p_key: p_key = config.get("active_profile", "F1")
        p_data = config["profiles"].get(p_key, {"triggers": {}})
        self.trigger_data = {}
        for key, data in p_data["triggers"].items():
            if isinstance(data, int): data = {"seconds": data, "icon": ""}
            row_widget = QFrame()
            row_widget.setFixedHeight(50)
            row_widget.setStyleSheet("background-color: #1e1e1e; border-radius: 4px; margin: 2px;")
            row = QHBoxLayout(row_widget)
            lbl = QLabel(key.upper()); lbl.setStyleSheet("color: #ffd700;"); lbl.setFixedWidth(60)
            
            icon_btn = QPushButton(); icon_btn.setFixedSize(32, 32)
            self.update_icon_button(icon_btn, data.get("icon", ""))
            icon_btn.clicked.connect(lambda checked, k=key, btn=icon_btn: self.pick_icon(k, btn))
            
            inp = QLineEdit(str(data.get("seconds", 300))); inp.setFixedWidth(35)
            cb_sound = QCheckBox("20s音效")
            cb_sound.setChecked(data.get("sound", True))
            cb_sound.setStyleSheet("font-size: 10px; color: #888;")
            del_btn = QPushButton("🗑️"); del_btn.setFixedWidth(30)
            del_btn.clicked.connect(lambda checked, k=key: self.delete_key(k))
            row.addWidget(lbl); row.addWidget(icon_btn); row.addWidget(inp); row.addWidget(QLabel("秒")); row.addWidget(cb_sound); row.addStretch(); row.addWidget(del_btn)
            self.trigger_data[key] = {"inp": inp, "icon": data.get("icon", ""), "cb_sound": cb_sound}
            self.scroll_layout.addWidget(row_widget)
        self.scroll_layout.addStretch()

    def delete_key(self, key):
        p_key = self.profile_box.itemData(self.profile_box.currentIndex())
        if not p_key: p_key = "F1"
        current_ui_triggers = self.capture_ui_data()
        if key in current_ui_triggers: del current_ui_triggers[key]
        config = ConfigManager.load_config()
        config["profiles"][p_key]["triggers"] = current_ui_triggers
        ConfigManager.save_config(config)
        self.refresh_items()

    def update_icon_button(self, btn, path):
        real_path = path if path and os.path.exists(path) else resource_path(path) if path else None
        if real_path and os.path.exists(real_path):
            pixmap = QPixmap(real_path).scaled(24, 24, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation)
            btn.setIcon(QIcon(pixmap)); btn.setIconSize(QSize(24, 24)); btn.setText("")
        else:
            btn.setIcon(QIcon()); btn.setText("🖼️")

    def pick_icon(self, key, btn):
        dlg = IconSelectorDialog(self)
        if dlg.exec():
            path = dlg.selected_icon
            self.trigger_data[key]["icon"] = path
            self.update_icon_button(btn, path)

    def on_money_toggle_changed(self, checked):
        if self.overlay:
            self.overlay.show_money_log = checked
            status = "已啟用" if checked else "已關閉"
            logger.info(f"[Overlay] Money Log toggled: {status}")
            # Save preference
            try:
                config = ConfigManager.load_config()
                config["show_money_log"] = checked
                ConfigManager.save_config(config)
            except Exception as e:
                logger.debug(f"[Overlay] Money toggle save failed: {e}")

    def on_debug_mode_changed(self, checked):
        v = checked
        self.debug_info_lbl.setVisible(v)
        self.debug_img_lbl.setVisible(v)
        self.debug_exp_conf_lbl.setVisible(v)
        self.debug_lv_info_lbl.setVisible(v)
        self.debug_lv_img_lbl.setVisible(v)
        self.debug_lv_conf_lbl.setVisible(v)
        if self.overlay: 
            self.overlay.show_debug = v
            self.overlay.show_notification(f"除錯模式: {'開啟' if checked else '關閉'}")

    def on_opacity_changed(self, value):
        self.opacity_val_lbl.setText(f"{value}%")
        if self.overlay:
            self.overlay.base_opacity = value / 100.0
            self.overlay.update()

    def on_exp_toggle_changed(self, checked):
        if self.overlay and self.overlay.show_exp_panel != checked:
            self.overlay.on_toggle_exp()

    def on_reset_exp_clicked(self):
        if self.overlay:
            self.overlay.on_toggle_exp() # Turn off
            self.overlay.on_toggle_exp() # Turn back on (resets logic)
            self.refresh_items()

    def save_and_close(self):
        config = ConfigManager.load_config()
        p_key = self.profile_box.itemData(self.profile_box.currentIndex())
        if not p_key: p_key = "F1"
        base_dir = os.path.abspath(".")
        new_triggers = {}
        for key, data in self.trigger_data.items():
            try:
                icon_path = data["icon"]
                if icon_path and os.path.isabs(icon_path) and icon_path.lower().startswith(base_dir.lower()):
                    icon_path = os.path.relpath(icon_path, base_dir).replace("\\", "/")
                new_triggers[key] = {
                    "seconds": int(data["inp"].text()), 
                    "icon": icon_path,
                    "sound": data["cb_sound"].isChecked()
                }
            except Exception as e:
                logger.debug(f"[Settings] Save trigger failed for {key}: {e}")
        config["profiles"][p_key]["triggers"] = new_triggers
        config["profiles"][p_key]["name"] = self.nickname_inp.text()
        if self.overlay: 
            config["offset"] = [self.overlay.x_offset, self.overlay.y_offset]
            config["exp_offset"] = [self.overlay.exp_x_offset, self.overlay.exp_y_offset]
            config["show_exp"] = self.overlay.show_exp_panel
            config["show_money_log"] = getattr(self.overlay, 'show_money_log', True)
            config["show_debug"] = self.overlay.show_debug
            config["opacity"] = self.overlay.base_opacity
        config["rjpq_offset"] = [self.overlay.rjpq_x_offset, self.overlay.rjpq_y_offset]
        ConfigManager.save_config(config)
        if self.handle: self.handle.hide()
        if hasattr(self, 'exp_handle') and self.exp_handle: self.exp_handle.hide()
        if self.overlay: self.overlay.show_preview = False
        self.config_updated.emit(); self.hide()

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
    exp_visual_request = pyqtSignal(dict)    # For Debug Images
    lv_update_request = pyqtSignal(dict)
    toggle_exp_request = pyqtSignal()
    toggle_pause_request = pyqtSignal()
    toggle_rjpq_request = pyqtSignal()
    settings_show_request = pyqtSignal()
    rjpq_cell_clicked = pyqtSignal(int)
    export_report_request = pyqtSignal()
    update_found = pyqtSignal(str, str)
    
    def __init__(self, target_window_title="MapleStory Worlds-Artale (繁體中文版)"):
        super().__init__()
        self.target_window_title = target_window_title
        self.active_timers = {} 
        self.click_zones = {}  
        self.is_active = False
        self.show_preview = False
        self.active_profile_name = "F1"
        self._is_running = True # Start this early!
        self.show_debug = False # Start this early!
        
        # Load configs early
        config = ConfigManager.load_config()
        self.show_exp_panel = config.get("show_exp", False)
        self.show_money_log = False  # Disabled in v0.2.8
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
        self.exp_history = [] # List of (timestamp, value, percent)
        self.exp_initial_val = None # Initialize to avoid AttributeError
        self.selected_color = -1
        self.cumulative_gain = 0
        self.cumulative_pct = 0.0
        self.max_10m_exp = 0 # Track max 10m gain
        self.last_exp_pct = 0.0
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
        
        self.timer_request.connect(self.start_timer)
        self.clear_request.connect(self.clear_all_timers)
        self.notification_request.connect(self.show_notification)
        self.profile_switch_request.connect(self.load_profile_immediately)
        self.exp_update_request.connect(self.on_exp_update)
        self.toggle_exp_request.connect(self.on_toggle_exp)
        self.toggle_pause_request.connect(self.on_toggle_pause)
        self.toggle_rjpq_request.connect(self.on_toggle_rjpq)
        self.export_report_request.connect(self.export_exp_report)
        self.update_found.connect(self.on_update_found)
        
        # We'll use this signal for tray to talk to settings_window
        self.request_show_settings_signal = pyqtSignal()
        if hasattr(self, 'request_show_settings_signal'):
             # Handle signal in main.py later
             pass
        
        self.tracking_timer = QTimer(self); self.tracking_timer.timeout.connect(self.sync_with_game_window); self.tracking_timer.start(100)
        self.countdown_timer = QTimer(self); self.countdown_timer.timeout.connect(self.update_countdown)
        self.world_timers = {} # Keep empty or remove if not used
        
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

    def start_timer(self, key, seconds, icon_path=None, sound_enabled=True):
        pixmap = None
        if icon_path:
            # Check Absolute, Relative, and Resource paths
            real_path = icon_path
            if not os.path.exists(real_path):
                real_path = resource_path(icon_path)
            
            if os.path.exists(real_path):
                pixmap = QPixmap(real_path)
                if pixmap.isNull():
                    logger.error(f"[Timer] Failed to load icon (Malformed): {real_path}")
                    pixmap = None
            else:
                logger.warning(f"[Timer] Icon not found in any search path: {icon_path}")
        
        self.active_timers[key] = {"seconds": seconds, "pixmap": pixmap, "sound_enabled": sound_enabled}
        self.is_active = True
        if not self.countdown_timer.isActive(): self.countdown_timer.start(1000)
        self.update()

    def update_countdown(self):
        to_remove = []
        for key in list(self.active_timers.keys()):
            self.active_timers[key]["seconds"] -= 1
            rem = self.active_timers[key]["seconds"]
            sound_enabled = self.active_timers[key].get("sound_enabled", True)
            if rem == 20 and sound_enabled: self.play_sound(1)
            elif rem == 0: self.play_sound(2)
            elif -10 < rem < 0: self.play_sound(1)
            if rem <= -10: to_remove.append(key)
        for key in to_remove:
            if key in self.active_timers: del self.active_timers[key]

        if not self.active_timers:
            self.is_active = False; self.countdown_timer.stop()
        self.update()

    def play_sound(self, times=1):
        if not winsound: return
        def worker():
            for _ in range(times):
                try: winsound.Beep(800, 150); time.sleep(0.12)
                except Exception as e: logger.debug(f"[Sound] Beep failed: {e}")
        threading.Thread(target=worker, daemon=True).start()

    def sync_with_game_window(self):
        # We no longer move the Overlay window; it stays full-screen on the virtual desktop.
        # This keeps the UI stable and independent of the game's movement.
        if not win32gui: return
        hwnd = 0
        try:
            my_pid = os.getpid()
            def callback(h, extra):
                nonlocal hwnd
                if not win32gui.IsWindowVisible(h): return True
                _, pid = win32process.GetWindowThreadProcessId(h)
                if pid == my_pid: return True # Skip ourselves
                title = win32gui.GetWindowText(h).lower()
                if self.target_window_title.lower() in title:
                    hwnd = h; return False
                return True
            win32gui.EnumWindows(callback, None)
        except Exception as e:
            logger.debug(f"[Overlay] Window search failed: {e}")
            hwnd = 0
        
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
        self.active_timers = {}; self.click_zones = {}; self.is_active = False
        if self.countdown_timer.isActive(): self.countdown_timer.stop()
        if show_msg: self.show_notification("⚠️ 已強制關閉並重設計時器")
        self.update()

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
                if key in self.active_timers: del self.active_timers[key]
                self.update(); return True
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
        target_h = 60; char_spacing = 30
        
        def build_pass_canvas(boxes):
            if not boxes: return None
            # Find max height in this pass to calculate uniform scale
            max_h_pass = max([b[3] for b in boxes])
            scale = target_h / max_h_pass if max_h_pass > 0 else 1.0
            
            total_w = sum([int(b[2] * scale) for b in boxes]) + (len(boxes) + 1) * char_spacing
            canvas = np.ones((target_h + 40, total_w), dtype=np.uint8) * 255
            curr_x = char_spacing
            
            # Baseline alignment: put all characters on the same floor
            baseline_y = target_h + 20
            
            for b in boxes:
                char = thresh_img[b[1]:b[1]+b[3], b[0]:b[0]+b[2]]
                char_inv = cv2.bitwise_not(char)
                nw, nh = int(b[2] * scale), int(b[3] * scale)
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
            # If we split into parts, and this is the second part (Percentage)
            # We skip the identified brackets
            boxes_to_use = p['boxes']
            if found_brackets and i == 1:
                # Part B starts at bracket_idx. Let's slice relative to merged_boxes
                # Percentage content is between bracket_idx and bracket_end_idx
                start_in_part_b = 1 # Skip the '[' at 0
                end_in_part_b = len(boxes_to_use)
                if bracket_end_idx != -1 and bracket_end_idx > bracket_idx:
                    # If we found a trailing ], skip it
                    end_in_part_b = len(boxes_to_use) - 1
                
                boxes_to_use = boxes_to_use[start_in_part_b:end_in_part_b]

            cvs = build_pass_canvas(boxes_to_use)
            if cvs is not None:
                canvases.append(cvs)
                # Specific whitelist: No brackets in Tesseract input to avoid confusion
                wl = p['wl'].replace('[', '').replace(']', '')
                config = f'--psm 7 --oem 3 -c tessedit_char_whitelist={wl}'
                try:
                    data = pytesseract.image_to_data(cvs, config=config, output_type=pytesseract.Output.DICT)
                    txts = [t for t in data['text'] if t.strip()]
                    confs = [int(c) for c in data['conf'] if int(c) != -1]
                    
                    if i == 1 and found_brackets:
                        # Percentage part
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
            combined = np.ones(((target_h + 40) * len(canvases) + 10, max_w), dtype=np.uint8) * 255
            curr_y = 0
            for c in canvases:
                h_c, w_c = c.shape
                combined[curr_y:curr_y+h_c, :w_c] = c
                curr_y += h_c + 5 # 5px gap
            processed_img = combined
        elif canvases:
            processed_img = canvases[0]
        else:
            processed_img = cv2.bitwise_not(thresh_img)

        ocr_text = "".join(final_texts)
        native_conf = sum(conf_sums) / len(conf_sums) if conf_sums else 0
        if native_conf <= 0 and any(c.isdigit() for c in ocr_text):
            native_conf = 30.0
            
        print(f"[OCR-DEBUG-SPLIT] '{ocr_text}' | Conf: {native_conf:.1f}%")
        
        # Apply Stability
        curr_score_attr = f"_current_score_{key}"
        
        if not hasattr(self, stable_img_attr):
            setattr(self, stable_img_attr, None)
            setattr(self, curr_score_attr, 0)
            
        prev_img = getattr(self, stable_img_attr)
        prev_score = getattr(self, curr_score_attr)
        
        stability = 100 
        if prev_img is not None:
            if prev_img.shape == thresh_img.shape:
                # Anti-flicker: Use small blur to ignore single-pixel scaling noise from resizing
                b1 = cv2.GaussianBlur(thresh_img, (3, 3), 0)
                b2 = cv2.GaussianBlur(prev_img, (3, 3), 0)
                diff = cv2.absdiff(b1, b2)
                diff_pct = (np.count_nonzero(diff) / diff.size) * 100
                sens = 10 if key == "exp" else 5
                stability = max(0, 100 - (diff_pct * sens))
            else:
                stability = 100
        
        setattr(self, stable_img_attr, thresh_img.copy())
        
        # 4. Pure Neural Reporting (USER REQUEST: Raw Tesseract Scores)
        new_score = native_conf
        setattr(self, curr_score_attr, new_score)
        
        # 6. Logging Low Confidence (Updated to 90%)
        if new_score < 90 and self.show_debug:
            logger.warning(f"[OCR] Low Confidence ({key}): {new_score:.1f}% | Text: {ocr_text}")
            # Save frame to tmp for visual debugging
            try:
                if not os.path.exists("tmp"): os.makedirs("tmp")
                # Sanitize text for filename (remove symbols like [], %, etc.)
                clean_text = "".join(c for c in ocr_text if c.isalnum() or c in "._-")
                fname = f"tmp/{key}_{new_score:.1f}_{clean_text}.png"
                cv2.imwrite(fname, thresh_img)
            except Exception as e:
                logger.debug(f"[OCR] Failed to save debug image: {e}")
        
        return ocr_text, new_score, processed_img

    def run_exp_tracker(self):
        """Background thread to capture and recognize EXP with adaptive scaling"""
        last_processed = 0
        
        def on_frame_arrived_callback(frame: Frame, control):
            nonlocal last_processed
            try:
                # Process if panel is ON OR if we are in Debug Mode
                if not self._is_running: return
                if not getattr(self, "show_exp_panel", False):
                    # Hard stop when panel is hidden, even if debug is on (for performance)
                    logger.info("[ExpTracker] Stopping session (Panel Closed)")
                    self._exp_tracker_active = False # Manual break for the inner loop
                    control.stop()
                    return
                
                now = time.time()
                if now - last_processed < 1.0: return
                
                img_bgra = getattr(frame, "frame_buffer", None)
                if img_bgra is None: return
                img = cv2.cvtColor(img_bgra, cv2.COLOR_BGRA2BGR)
                h, w = img.shape[:2]
                
                # IMPORTANT: Find the window to get client rect
                target_hwnd = None
                for name in ["MapleStory Worlds-Artale (繁體中文版)", "MapleStory Worlds-Artale"]:
                    hwnd = win32gui.FindWindow(None, name)
                    if hwnd: target_hwnd = hwnd; break
                
                if target_hwnd:
                    # Monitor Placement change (Maximize/Restore)
                    placement = win32gui.GetWindowPlacement(target_hwnd)
                    current_state = placement[1]
                    if not hasattr(self, 'last_window_state'):
                        self.last_window_state = current_state
                    
                    if current_state != self.last_window_state:
                        logger.info(f"[ExpTracker] Window state changed to {current_state}. Restarting session...")
                        self.last_window_state = current_state
                        self._exp_tracker_active = False # Signal main loop to restart
                        control.stop() 
                        return
                    
                    crect = win32gui.GetClientRect(target_hwnd)
                    # Use client dimensions for scaling and anchoring
                    cw_ref, ch_ref = crect[2], crect[3]
                else:
                    cw_ref, ch_ref = w, h # Fallback
                
                # Calculate dynamic scale based on the smaller reference side
                scale = min(cw_ref / self.BASE_W, ch_ref / self.BASE_H)
                
                # Check window state for border logic
                placement = win32gui.GetWindowPlacement(target_hwnd)
                if placement[1] == win32con.SW_SHOWMAXIMIZED:
                    # Maximized: Top is title bar, sides are 0
                    off_x = 0
                    off_y = h - ch_ref
                else:
                    # Windowed: Borders are typically symmetrical on left/right/bottom
                    # Capture width (w) = Client width (cw_ref) + 2 * border_width
                    border_w = max(0, (w - cw_ref) // 2)
                    off_x = border_w
                    # Capture height (h) = Client height (ch_ref) + title_bar + bot_border
                    # Title bar = h - ch_ref - border_w (assuming bot_border == border_w)
                    off_y = max(0, h - ch_ref - border_w)
                
                cx = off_x + int(self.X_OFF_FROM_LEFT * scale)
                cy = off_y + (ch_ref - int(self.Y_OFF_FROM_BOTTOM * scale))
                
                cw = int(self.BASE_CW * scale)
                ch = int(self.BASE_CH * scale)
                
                # Exact crop
                crop = img[max(0, cy):min(h, cy+ch), max(0, cx):min(w, cx+cw)]
                if crop.size == 0: return
                
                # Update debug info for sync visualization
                self.last_crop_info = (cx, cy, cw, ch, w, h)
                
                # LV Crop calculation
                lv_cx = off_x + int(self.LV_X_OFF_FROM_LEFT * scale)
                lv_cy = off_y + (ch_ref - int(self.LV_Y_OFF_FROM_BOTTOM * scale))
                lv_cw = int(self.LV_BASE_CW * scale)
                lv_ch = int(self.LV_BASE_CH * scale)
                self.last_lv_crop_info = (lv_cx, lv_cy, lv_cw, lv_ch, w, h)
                
                # LV Binarization for debug
                lv_crop = img[max(0, lv_cy):min(h, lv_cy+lv_ch), max(0, lv_cx):min(w, lv_cx+lv_cw)]
                if lv_crop.size > 0:
                    lv_gray = cv2.cvtColor(lv_crop, cv2.COLOR_BGR2GRAY)
                    r_lv = 60 / lv_gray.shape[0] if lv_gray.shape[0] > 0 else 3
                    lv_gray = cv2.resize(lv_gray, None, fx=r_lv, fy=r_lv, interpolation=cv2.INTER_CUBIC)
                    # Use THRESH_BINARY (White text) to match the new OCR helper logic
                    _, lv_thresh = cv2.threshold(lv_gray, 180, 255, cv2.THRESH_BINARY)
                    
                    # OCR for Level using unified helper (Upscale 4 for precision)
                    lv_ocr_text, lv_conf, lv_processed = self._perform_enhanced_ocr(lv_thresh, "lv", upscale=4, whitelist="0123456789")
                    
                    # ALWAYS emit UI update so user can see what OCR sees
                    if not sip.isdeleted(self): 
                        self.lv_update_request.emit({"thresh": lv_processed, "level": lv_ocr_text, "conf": lv_conf})
                        
                    # DATA gate: Only process level logic if solid (>= 90%)
                    if lv_conf < 90:
                        if self.show_debug and lv_ocr_text:
                            logger.debug(f"[ExpTracker] LV change ignored (low confidence): {lv_conf:.1f}%")
                
                # --- Coin Recognition (Template Matching with Adaptive Scaling) ---
                if self.show_money_log and hasattr(self, 'coin_tpl') and self.coin_tpl is not None:
                    try:
                        tpl_h, tpl_w = self.coin_tpl.shape[:2]
                        scaled_tpl_w = int(tpl_w * scale); scaled_tpl_h = int(tpl_h * scale)
                        if scaled_tpl_w > 5 and scaled_tpl_h > 5:
                            tpl_resized = cv2.resize(self.coin_tpl, (scaled_tpl_w, scaled_tpl_h))
                            res = cv2.matchTemplate(img, tpl_resized, cv2.TM_CCOEFF_NORMED)
                            _, max_val, _, max_loc = cv2.minMaxLoc(res)
                            
                            if max_val > 0.8:
                                # Primary coin box
                                self.last_coin_pos = (max_loc[0] - off_x, max_loc[1] - off_y, scaled_tpl_w, scaled_tpl_h)
                                
                                # Secondary Info Area (Right 30px, 280x31)
                                info_w = int(280 * scale); info_h = int(31 * scale)
                                info_ix = max_loc[0] + scaled_tpl_w + int(30 * scale)
                                # Move down 1px total as requested
                                info_iy = max_loc[1] + (scaled_tpl_h // 2) - (info_h // 2) + int(1 * scale)
                                self.last_coin_info_pos = (info_ix - off_x, info_iy - off_y, info_w, info_h)
                                
                                # --- OCR for Coin Info Area ---
                                coin_info_text = ""
                                info_crop = img[max(0, info_iy):min(h, info_iy+info_h), max(0, info_ix):min(w, info_ix+info_w)]
                                if info_crop.size > 0:
                                    ic_gray = cv2.cvtColor(info_crop, cv2.COLOR_BGR2GRAY)
                                    ic_r = 60 / ic_gray.shape[0] if ic_gray.shape[0] > 0 else 3
                                    ic_gray = cv2.resize(ic_gray, None, fx=ic_r, fy=ic_r, interpolation=cv2.INTER_CUBIC)
                                    _, ic_thresh = cv2.threshold(ic_gray, 130, 255, cv2.THRESH_BINARY_INV)
                                    if pytesseract and pytesseract.pytesseract.tesseract_cmd:
                                        try:
                                            # Efficient single-call OCR
                                            ic_padded = cv2.copyMakeBorder(ic_thresh, 10, 10, 10, 10, cv2.BORDER_CONSTANT, value=255)
                                            coin_info_text = pytesseract.image_to_string(ic_padded, config='--psm 7 -c tessedit_char_whitelist=0123456789,').strip()
                                            
                                            if coin_info_text:
                                                # Standard log (High performance)
                                                logger.debug(f"[ExpTracker] Coin Match: {max_val:.2f} | OCR: {coin_info_text}")
                                                self.last_coin_ocr = coin_info_text
                                        except Exception as e:
                                            logger.debug(f"[Debug] OCR Failed: {e}")
                                self.last_coin_ocr = coin_info_text
                            else:
                                self.last_coin_pos = None
                                self.last_coin_info_pos = None
                                self.last_coin_ocr = ""
                    except: 
                        self.last_coin_pos = None
                        self.last_coin_info_pos = None
                        self.last_coin_ocr = ""
                
                if self.exp_paused:
                    last_processed = now
                    self.update()
                    return

                # OCR Processing
                gray = cv2.cvtColor(crop, cv2.COLOR_BGR2GRAY)
                target_ocr_h = 60
                r = target_ocr_h / gray.shape[0] if gray.shape[0] > 0 else 3
                gray = cv2.resize(gray, None, fx=r, fy=r, interpolation=cv2.INTER_CUBIC)
                # EXP OCR Processing (180 threshold as per EXACT Artale Efficiency logic)
                _, thresh = cv2.threshold(gray, 180, 255, cv2.THRESH_BINARY)
                
                # EXP OCR using unified helper (Upscale 4x, Digit-only conf)
                text, exp_conf, exp_processed = self._perform_enhanced_ocr(thresh, "exp", upscale=4, whitelist="0123456789.%")
                
                if text:
                    # Visual Signal: Send the processed_img as requested
                    if not sip.isdeleted(self):
                        self.exp_visual_request.emit({
                            "thresh": exp_processed, 
                            "crop": crop, 
                            "text": text, 
                            "conf": exp_conf
                        })
                        
                    # DATA processing: We trust the raw OCR result
                    self.parse_and_update_exp(text, thresh, crop, exp_conf)
                last_processed = now
                if not sip.isdeleted(self): self.update() # Ensure red box moves smoothly with window
            except: pass
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
            
            win32gui.EnumWindows(enum_handler, None)
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
                try:
                    precise_name = win32gui.GetWindowText(target_hwnd)
                    capture = WindowsCapture(
                        window_name=precise_name, 
                        cursor_capture=False, 
                        draw_border=False, 
                        minimum_update_interval=1000
                    )
                    
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
                        # Check if window still exists and is visible
                        if not win32gui.IsWindow(target_hwnd) or not win32gui.IsWindowVisible(target_hwnd):
                            break
                    
                    # No explicit stop() available on the capture object in this version
                    self._exp_tracker_active = False
                    
                except Exception as e:
                    err_msg = str(e)
                    if "0x80070490" in err_msg:
                        # Common API error when window is transitionary/hidden, retry silently
                        pass
                    else:
                        logger.error(f"[ExpTracker] Failed to start capture for HWND {target_hwnd}: {e}")
            
            # Adaptive retry delay
            time.sleep(2.0)

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
        self.msg_text = text; self.msg_opacity = 255
        if hasattr(self, 'fade_timer') and self.fade_timer.isActive(): self.fade_timer.stop()
        self.fade_timer = QTimer(self); self.fade_timer.timeout.connect(self.step_fade)
        QTimer.singleShot(3000, lambda: self.fade_timer.start(16)); self.update()

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
            
        if self.show_rjpq_panel:
            # Multi-profile check: Use current session offsets if not reloaded yet
            pw, ph = 180, 320
            # Anchor (ax, ay) is Top-Right
            ax = self.rect().center().x() + self.rjpq_x_offset
            ay = self.rect().center().y() + self.rjpq_y_offset
            # draw_rjpq_panel uses Top-Left as start, so start_x = ax - pw
            draw_rjpq_panel(painter, ax - pw, ay, 
                          pw, ph, self.base_opacity, self.rjpq_data, self.selected_color)
            
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

        if not self.active_timers and not self.show_preview: return

        timers_to_draw = []
        if self.active_timers:
            sorted_active = sorted(self.active_timers.items(), key=lambda x: x[1]["seconds"], reverse=True)
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
            if self.show_preview and not self.active_timers: color = QColor(255, 255, 255, 150)
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
        pw, ph = 330, 200 # Increased height for graph
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
        



