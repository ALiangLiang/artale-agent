import sys
import json
import os
import threading
import time
try:
    import win32gui
    import winsound
except ImportError:
    win32gui = None
    winsound = None

from PyQt6.QtWidgets import (QApplication, QWidget, QVBoxLayout, QHBoxLayout, 
                             QLabel, QLineEdit, QPushButton, QScrollArea, QFrame,
                             QGridLayout, QDialog, QTabWidget, QComboBox, QSlider, QCheckBox,
                             QSystemTrayIcon, QMenu)
try:
    from windows_capture import WindowsCapture, Frame
except ImportError:
    WindowsCapture = None

from PyQt6.QtCore import Qt, QPoint, QRect, QTimer, pyqtSignal, QSize, QRectF
from PyQt6.QtGui import QFont, QColor, QPainter, QPen, QPixmap, QIcon, QPainterPath, QAction
import cv2
import re
import pytesseract
import os

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
                    for i in range(1, 10):
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

                    # Ensure root offset exists
                    if "offset" not in config:
                        config["offset"] = [0, 0]

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

                    return config
            except Exception as e: 
                print(f"Error loading config: {e}")
                pass
            
        # Default Multi-Profile Config
        default_profiles = {}
        for i in range(1, 10):
            default_profiles[f"F{i}"] = {"name": f"切換組 F{i}", "triggers": {}}
        return {"active_profile": "F1", "offset": [0, 0], "opacity": 0.5, "profiles": default_profiles}

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
    
    def __init__(self, target_window_title="Artale"):
        super().__init__()
        self.target_window_title = target_window_title
        self.setWindowFlags(Qt.WindowType.WindowStaysOnTopHint | Qt.WindowType.FramelessWindowHint | Qt.WindowType.Tool)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.setFixedSize(60, 60)
        
        self.lbl = QLabel(self)
        self.lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        arrow_p = resource_path("buff_pngs/arrow.png")
        self.pixmap = QPixmap(arrow_p).scaled(50, 50, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation)
        self.lbl.setPixmap(self.pixmap)
        self.lbl.setFixedSize(60, 60)
        self.lbl.setToolTip("拖動我來調整計時器位置\n調整完畢後儲存即可")
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
    
    def __init__(self, overlay=None):
        super().__init__()
        self.overlay = overlay
        self.setWindowTitle("Artale Helper - Settings")
        self.setFixedSize(360, 580)
        self.setWindowFlags(Qt.WindowType.WindowStaysOnTopHint)
        self.is_recording = False
        self.handle = None
        self.trigger_data = {}
        
        # Connect to EXP updates for live preview
        if self.overlay:
            self.overlay.exp_update_request.connect(self.on_exp_data_received)
            
        self.init_ui()
        self.request_show.connect(self.safe_show)

    def on_exp_data_received(self, data):
        """Update live preview images using robust byte loading"""
        if not self.isVisible() or self.tabs.currentIndex() != 1:
            return
            
        try:
            from PyQt6.QtGui import QPixmap
            if data.get("debug_bytes"):
                pix = QPixmap()
                if pix.loadFromData(data["debug_bytes"]):
                    self.debug_img_lbl.setPixmap(pix.scaled(
                        self.debug_img_lbl.size(), 
                        Qt.AspectRatioMode.KeepAspectRatio, 
                        Qt.TransformationMode.SmoothTransformation
                    ))
        except Exception as e:
            pass

    def init_ui(self):
        self.setWindowTitle("Artale Helper ⚙️ 設定")
        # Set Window Icon
        icon_path = resource_path("app_icon.png")
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))
            
        config = ConfigManager.load_config()
        self.layout = QVBoxLayout(self)
        self.setStyleSheet("background-color: #121212; color: #e0e0e0; font-family: 'Segoe UI';")
        
        top_row = QHBoxLayout()
        title = QLabel("🎹 設定清單")
        title.setFont(QFont("Segoe UI Semibold", 18))
        title.setStyleSheet("color: #ffd700;")
        top_row.addWidget(title)
        
        self.tabs = QTabWidget()
        self.tabs.setStyleSheet("""
            QTabWidget::pane { border: 1px solid #333; background: #121212; }
            QTabBar::tab { background: #222; color: #888; padding: 10px; min-width: 80px; }
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
        self.scroll.setWidget(self.scroll_content)
        timer_tab_layout.addWidget(self.scroll)

        self.record_btn = QPushButton("➕ 新增按鍵 (點我後按鍵盤)")
        self.record_btn.setStyleSheet("QPushButton { background-color: #2e7d32; color: white; font-weight: bold; border-radius: 5px; height: 35px; }")
        self.record_btn.clicked.connect(self.toggle_recording)
        timer_tab_layout.addWidget(self.record_btn)
        
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
        
        reset_exp_btn = QPushButton("🔄 重新開始紀錄 (歸零重算)")
        reset_exp_btn.setStyleSheet("QPushButton { background-color: #444; color: #eee; border-radius: 4px; height: 30px; margin-top: 10px; } QPushButton:hover { background: #555; }")
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
        
        self.debug_img_lbl = QLabel()
        self.debug_img_lbl.setFixedSize(300, 30)
        self.debug_img_lbl.setStyleSheet("border: 1px solid #444; background: #000;")
        self.debug_img_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.debug_img_lbl.setVisible(self.debug_mode_cb.isChecked())
        exp_tab_layout.addWidget(self.debug_img_lbl)
        self.tabs.addTab(exp_tab, "📊 經驗值")

        # --- Global Controls (Bottom) ---

        self.update_profile_dropdown()

        self.refresh_items()

        # --- Global Controls (Bottom) ---

        save_btn = QPushButton("💾 儲存並套用")
        save_btn.setStyleSheet("QPushButton { background-color: #ffd700; color: black; font-weight: bold; border-radius: 5px; height: 35px; margin-top: 5px;}")
        save_btn.clicked.connect(self.save_and_close)
        self.layout.addWidget(save_btn)

        pos_btn = QPushButton("🔱 調整顯示位置")
        pos_btn.setStyleSheet("QPushButton { background-color: #5c6bc0; color: white; border-radius: 5px; height: 30px; margin-top: 5px;}")
        pos_btn.clicked.connect(self.toggle_handle)
        self.layout.addWidget(pos_btn)

        exit_btn = QPushButton("🛑 關閉整個輔助")
        exit_btn.setStyleSheet("QPushButton { background-color: transparent; color: #666; border: none; font-size: 10px; margin-top: 20px; }")
        exit_btn.clicked.connect(QApplication.instance().quit)
        self.layout.addWidget(exit_btn)
        


    def keyPressEvent(self, event):
        if self.is_recording:
            key_code = event.key()
            key_name = self.qt_key_to_name(key_code)
            if key_name:
                p_key = self.profile_box.itemData(self.profile_box.currentIndex())
                if not p_key: p_key = "F1"
                current_ui_triggers = self.capture_ui_data()
                config = ConfigManager.load_config()
                config["profiles"][p_key]["triggers"] = current_ui_triggers
                if key_name not in config["profiles"][p_key]["triggers"]:
                    config["profiles"][p_key]["triggers"][key_name] = {"seconds": 300, "icon": ""}
                    ConfigManager.save_config(config)
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
            except:
                active_ui_data[key] = {"seconds": 300, "icon": ui["icon"], "sound": True}
        return active_ui_data

    def update_profile_dropdown(self):
        self.profile_box.blockSignals(True)
        current_idx = self.profile_box.currentIndex()
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

    def qt_key_to_name(self, code):
        special_map = {
            Qt.Key.Key_F1: "f1", Qt.Key.Key_F2: "f2", Qt.Key.Key_F3: "f3", Qt.Key.Key_F4: "f4",
            Qt.Key.Key_F5: "f5", Qt.Key.Key_F6: "f6", Qt.Key.Key_F7: "f7", Qt.Key.Key_F8: "f8",
            Qt.Key.Key_F9: "f9", Qt.Key.Key_F12: "f12", Qt.Key.Key_Shift: "shift", 
            Qt.Key.Key_Control: "ctrl", Qt.Key.Key_Alt: "alt", Qt.Key.Key_Space: "space"
        }
        if code in special_map: return special_map[code]
        try:
            name = chr(code).lower() if 32 <= code <= 126 else None
            return name
        except: return None

    def toggle_recording(self):
        self.is_recording = not self.is_recording
        if self.is_recording:
            self.record_btn.setText("🔴 請按鍵盤錄製中...")
            self.record_btn.setStyleSheet("background-color: #c62828; color: white; border-radius: 5px; height: 35px;")
        else:
            self.record_btn.setText("➕ 新增按鍵 (點我後按鍵盤)")
            self.record_btn.setStyleSheet("background-color: #2e7d32; color: white; border-radius: 5px; height: 35px;")

    def safe_show(self):
        self.show(); self.activateWindow(); self.raise_()
        
    def toggle_handle(self):
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

    def on_debug_mode_changed(self, checked):
        if self.overlay:
            self.overlay.show_debug = checked
            self.debug_info_lbl.setVisible(checked)
            self.debug_img_lbl.setVisible(checked)
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
            except: pass
        config["profiles"][p_key]["triggers"] = new_triggers
        config["profiles"][p_key]["name"] = self.nickname_inp.text()
        if self.overlay: config["offset"] = [self.overlay.x_offset, self.overlay.y_offset]
        config["active_profile"] = p_key
        config["opacity"] = self.opacity_slider.value() / 100.0
        ConfigManager.save_config(config)
        if self.handle: self.handle.hide()
        if self.overlay: self.overlay.show_preview = False
        self.config_updated.emit(); self.hide()

class ArtaleOverlay(QWidget):
    timer_request = pyqtSignal(str, int, str, bool) 
    clear_request = pyqtSignal()
    notification_request = pyqtSignal(str)
    profile_switch_request = pyqtSignal()
    exp_update_request = pyqtSignal(dict)
    toggle_exp_request = pyqtSignal()
    settings_show_request = pyqtSignal()
    
    def __init__(self, target_window_title="Artale"):
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
        self.show_exp_panel = config.get("show_exp", True)
        self.show_debug = config.get("show_debug", False)
        self.base_opacity = config.get("opacity", 0.5)
        
        self.msg_text = ""; self.msg_opacity = 0
        self.x_offset = 0; self.y_offset = 0
        self.current_exp_data = {"text": "---", "value": 0, "percent": 0.0, "gained_10m": 0, "percent_10m": 0.0}
        self.exp_history = [] # List of (timestamp, value, percent)
        self.last_capture_time = 0
        self._tesseract_error_shown = False
        
        self.timer_request.connect(self.start_timer)
        self.clear_request.connect(self.clear_all_timers)
        self.notification_request.connect(self.show_notification)
        self.profile_switch_request.connect(self.load_profile_immediately)
        self.exp_update_request.connect(self.on_exp_update)
        self.toggle_exp_request.connect(self.on_toggle_exp)
        
        # We'll use this signal for tray to talk to settings_window
        self.request_show_settings_signal = pyqtSignal()
        if hasattr(self, 'request_show_settings_signal'):
             # Handle signal in main.py later
             pass
        
        self.tracking_timer = QTimer(self); self.tracking_timer.timeout.connect(self.sync_with_game_window); self.tracking_timer.start(100)
        self.countdown_timer = QTimer(self); self.countdown_timer.timeout.connect(self.update_countdown)
        
        # Initialize ExpTracker
        if WindowsCapture:
            self.exp_tracker_thread = threading.Thread(target=self.run_exp_tracker, daemon=True)
            self.exp_tracker_thread.start()
        else:
            print("[Warning] windows-capture is not installed. EXP tracking disabled.")
        
        frame_p = resource_path("buff_pngs/skill_frame.png")
        self.icon_frame = QPixmap(frame_p) if os.path.exists(frame_p) else None
        self.load_profile_immediately()
        self.init_tray()
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
        
        show_settings_action = QAction("⚙️ 開啟設定 (Pause)", self)
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
        self.tray_icon.setToolTip("Artale Helper")
        
        # Click to toggle settings
        self.tray_icon.activated.connect(self.on_tray_activated)
        self.tray_icon.show()

    def on_tray_activated(self, reason):
        if reason == QSystemTrayIcon.ActivationReason.Trigger:
            self.request_show_settings()

    def request_show_settings(self):
        self.settings_show_request.emit()

    def reset_exp_stats(self):
        """Reset EXP tracking baseline"""
        self.exp_history = []
        self.show_notification("📊 經驗值統計已重置")

    def init_ui(self):
        self.setWindowFlags(Qt.WindowType.WindowStaysOnTopHint | Qt.WindowType.FramelessWindowHint | Qt.WindowType.WindowTransparentForInput)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.setAttribute(Qt.WidgetAttribute.WA_AlwaysStackOnTop)
        virtual_geo = QApplication.primaryScreen().virtualGeometry()
        self.setGeometry(virtual_geo)
        self.show()

    def start_timer(self, key, seconds, icon_path=None, sound_enabled=True):
        pixmap = None
        if icon_path:
            real_path = icon_path if os.path.exists(icon_path) else resource_path(icon_path)
            if os.path.exists(real_path): pixmap = QPixmap(real_path)
        self.active_timers[key] = {"seconds": seconds, "pixmap": pixmap, "sound_enabled": sound_enabled}
        self.is_active = True
        if not self.countdown_timer.isActive(): self.countdown_timer.start(1000)
        # Limit UI update to avoid unnecessary repaints
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
                except: pass
        threading.Thread(target=worker, daemon=True).start()

    def sync_with_game_window(self):
        if not win32gui: return
        hwnd = 0
        try:
            def callback(h, extra):
                nonlocal hwnd
                if self.target_window_title.lower() in win32gui.GetWindowText(h).lower() and win32gui.IsWindowVisible(h):
                    hwnd = h; return False
                return True
            win32gui.EnumWindows(callback, None)
        except: hwnd = 0
        if hwnd:
            rect = win32gui.GetWindowRect(hwnd)
            x, y, x2, y2 = rect
            if self.geometry() != QRect(x, y, x2-x, y2-y): 
                self.setGeometry(x, y, x2-x, y2-y)
                self.update()
            # Absolute persistence
            if not self.isVisible(): self.show()
            self.raise_()

    def update_offset(self, gx, gy):
        local = self.mapFromGlobal(QPoint(gx, gy))
        self.x_offset = local.x() - self.rect().center().x()
        self.y_offset = local.y() - self.rect().center().y()
        self.click_zones = {}; self.update()

    def clear_all_timers(self, show_msg=True):
        self.active_timers = {}; self.click_zones = {}; self.is_active = False
        if self.countdown_timer.isActive(): self.countdown_timer.stop()
        if show_msg: self.show_notification("⚠️ 已強制關閉並重設 (F12)")
        self.update()

    def check_right_click(self, gx, gy):
        p = QPoint(gx, gy)
        for key, rect in list(self.click_zones.items()):
            if rect.contains(p):
                if key in self.active_timers: del self.active_timers[key]
                self.update(); return True
        return False

    # --- EXP Tracker Logic ---
    def closeEvent(self, event):
        self._is_running = False
        super().closeEvent(event)

    def run_exp_tracker(self):
        """Background thread to capture and recognize EXP with adaptive scaling"""
        BASE_W, BASE_H = 1920, 1080
        BASE_X, BASE_Y, BASE_CW, BASE_CH = 1022, 1017, 250, 19
        last_processed = 0
        
        def on_frame_arrived_callback(frame: Frame, _):
            nonlocal last_processed
            try:
                if not self._is_running or not self.show_exp_panel: return
                now = time.time()
                if now - last_processed < 1.0: return
                last_processed = now
                img_bgra = getattr(frame, "frame_buffer", None)
                if img_bgra is None: return
                img = cv2.cvtColor(img_bgra, cv2.COLOR_BGRA2BGR)
                h, w = img.shape[:2]
                sx, sy = w / BASE_W, h / BASE_H
                crop = img[int(BASE_Y*sy):int((BASE_Y+BASE_CH)*sy), int(BASE_X*sx):int((BASE_X+BASE_CW)*sx)]
                if crop.size == 0: return
                gray = cv2.cvtColor(crop, cv2.COLOR_BGR2GRAY)
                target_h = 60
                scale = target_h / gray.shape[0] if gray.shape[0] > 0 else 3
                gray = cv2.resize(gray, None, fx=scale, fy=scale, interpolation=cv2.INTER_CUBIC)
                _, thresh = cv2.threshold(gray, 180, 255, cv2.THRESH_BINARY_INV)
                text = ""
                if pytesseract and pytesseract.pytesseract.tesseract_cmd:
                    try:
                        padded = cv2.copyMakeBorder(thresh, 10, 10, 10, 10, cv2.BORDER_CONSTANT, value=255)
                        tess_config = '--psm 7 -c tessedit_char_whitelist=0123456789.[]%|'
                        text = pytesseract.image_to_string(padded, config=tess_config).strip()
                        self._tesseract_error_shown = False
                    except: pass
                if text:
                    if self.show_debug: print(f"[ExpTracker] Raw OCR: '{text}'")
                    self.parse_and_update_exp(text, thresh, crop)
                last_processed = now
            except: pass

        while self._is_running:
            capture = None
            target_name = None
            
            # 1. Check static candidates
            static_candidates = ["MapleStory Worlds-Artale (繁體中文版)", "Artale"]
            for name in static_candidates:
                try:
                    from windows_capture import WindowsCapture
                    # We just test if creating the object with this name works without error
                    # (In most versions, invalid names throw immediately or on start)
                    test = WindowsCapture(window_name=name)
                    target_name = name
                    break
                except: continue
            
            # 2. Check process fallback if not found
            if not target_name:
                try:
                    import psutil, win32gui, win32process
                    for proc in psutil.process_iter(['pid', 'name']):
                        if proc.info['name'] and proc.info['name'].lower() == 'msw.exe':
                            def callback(hwnd, extra):
                                if win32gui.IsWindowVisible(hwnd):
                                    _, pid = win32process.GetWindowThreadProcessId(hwnd)
                                    if pid == proc.info['pid']:
                                        t = win32gui.GetWindowText(hwnd); 
                                        if t: extra.append(t)
                                return True
                            titles = []
                            win32gui.EnumWindows(callback, titles)
                            if titles: target_name = titles[0]; break
                except: pass

            if target_name:
                try:
                    from windows_capture import WindowsCapture
                    temp_capture = WindowsCapture(window_name=target_name, cursor_capture=False, draw_border=False, minimum_update_interval=1000)
                    @temp_capture.event
                    def on_frame_arrived(frame, control):
                        on_frame_arrived_callback(frame, control)
                    @temp_capture.event
                    def on_closed():
                        print("[ExpTracker] Session closed."); self._exp_tracker_active = False
                    
                    print(f"[ExpTracker] Monitoring {target_name} for EXP...")
                    self._exp_tracker_active = True
                    self.capture_session = temp_capture.start_free_threaded()
                    capture = temp_capture
                except Exception as e:
                    print(f"[ExpTracker] Failed to start capture for {target_name}: {e}")
            
            if capture:
                while self._is_running and self._exp_tracker_active:
                    time.sleep(1)
                capture = None
            else:
                if self._is_running: time.sleep(5)

    def parse_and_update_exp(self, raw_text, debug_img=None, raw_img=None):
        try:
            # Prepare data
            data = {"text": "---", "value": 0, "percent": 0.0, "timestamp": time.time(), "debug_bytes": None, "raw_bytes": None}
            
            # Encode images to bytes (Bypasses QImage pointer issues)
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

            # Normalize and extract all numbers
            s = raw_text.replace(",", "")
            nums = re.findall(r"\d+\.\d+|\d+", s)
            if not nums: return
            
            val_str, pct_val = "", 0.0
            if len(nums) >= 2:
                # Normal case: [EXP, ..., PCT]
                pct_str = nums[-1]
                pct_val = float(pct_str)
                val_str = max(nums[:-1], key=len)
            elif len(nums) == 1 and "." in nums[0]:
                # Joined case: "522120511.46"
                full = nums[0]
                dot_idx = full.find(".")
                pct_str = full[dot_idx-2:] # Assume 2 digits before dot
                val_str = full[:dot_idx-2]
                pct_val = float(pct_str)
            else: return

            # Noise correction for brackets seen as '1'
            if pct_val > 100 and str(pct_str).startswith("1"):
                pct_val = float(str(pct_str)[1:])
            
            if not val_str: return
            val = int(val_str)
            if self.show_debug:
                print(f"[ExpTracker] Parse OK: {val:,} [{pct_val:.2f}%]")
            
            data["text"] = f"{val:,} [{pct_val:.2f}%]"
            data["value"] = val
            data["percent"] = pct_val
            
            self.exp_update_request.emit(data)
        except Exception as e:
            print(f"[ExpTracker] Parse Error: {e} | Raw: {raw_text}")

    def on_exp_update(self, data):
        # Initialize session variables if needed
        if not hasattr(self, 'exp_session_start_time'): self.exp_session_start_time = None
        if not hasattr(self, 'exp_initial_val') or self.exp_initial_val is None: 
            self.exp_initial_val = data["value"]
            print(f"[ExpTracker] Initial baseline: {self.exp_initial_val:,}")

        # Trigger session start on first actual EXP gain
        if self.exp_session_start_time is None and data["value"] > self.exp_initial_val:
            self.exp_session_start_time = data["timestamp"]
            print(f"[ExpTracker] Session triggered! First gain detected.")

        now = data["timestamp"]
        self.exp_history.append((now, data["value"], data["percent"]))
        
        # 1. 10min Sliding Window (for Efficiency)
        limit_10m = now - 600
        while len(self.exp_history) > 1 and self.exp_history[0][0] < limit_10m:
            self.exp_history.pop(0)
            
        # 2. Update UI Metadata
        self.current_exp_data["text"] = data["text"]
        self.current_exp_data["value"] = data["value"]
        self.current_exp_data["percent"] = data["percent"]

        if self.exp_session_start_time is not None:
            # Main recording duration (from first gain)
            self.current_exp_data["tracking_duration"] = int(now - self.exp_session_start_time)
            
            # Efficiency window (sliding, max 10m)
            ref_t, ref_v, ref_p = self.exp_history[0]
            win_elapsed = now - ref_t
            gain_val = data["value"] - ref_v
            gain_pct = data["percent"] - ref_p
            
            if win_elapsed > 1:
                if win_elapsed >= 595: # Real 10m window reached
                    self.current_exp_data["gained_10m"] = gain_val
                    self.current_exp_data["percent_10m"] = gain_pct
                    self.current_exp_data["is_estimated"] = False
                else:
                    # Estimate 10m amount
                    scale = 600 / win_elapsed
                    self.current_exp_data["gained_10m"] = int(gain_val * scale)
                    self.current_exp_data["percent_10m"] = gain_pct * scale
                    self.current_exp_data["is_estimated"] = True

                # Time to Level (based on current 10m rate)
                rate_per_sec = (self.current_exp_data["percent_10m"] / 600.0)
                if rate_per_sec > 0:
                    rem_pct = 100.0 - data["percent"]
                    self.current_exp_data["time_to_level"] = int(rem_pct / rate_per_sec)
                else:
                    self.current_exp_data["time_to_level"] = -1
        
        if self.show_exp_panel:
            print(f"[Overlay] EXP Data: {data['text']}")
        self.update()

    def load_profile_immediately(self):
        self.clear_all_timers(show_msg=False)
        config = ConfigManager.load_config()
        active = config.get("active_profile", "F1")
        nickname = config["profiles"].get(active, {}).get("name", active)
        self.active_profile_name = active
        self.x_offset, self.y_offset = config.get("offset", [0, 0])
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
        # Always draw EXP panel if we have data, even if not active
        painter = QPainter(self); painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        
        # 0. Draw EXP Statistics Panel (Top Left)
        if self.show_exp_panel:
            self.draw_exp_panel(painter)

        if not self.is_active and not self.show_preview and self.msg_opacity == 0: 
            return
        
        # Base coordinates
        base_x = self.rect().center().x() + self.x_offset
        base_y = self.rect().center().y() + self.y_offset

        # 1. Profile/Action Notification (Centered above anchor)
        if self.msg_opacity > 0:
            font = QFont("Segoe UI Bold", 18); painter.setFont(font)
            tw = painter.fontMetrics().horizontalAdvance(self.msg_text)
            # Draw notification clearly above the timer block
            bg_rect = QRect(base_x - (tw+40)//2, base_y - 70, tw+40, 45)
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
        elif self.show_preview:
            timers_to_draw.append(("preview", 300, QPixmap(resource_path("buff_pngs/arrow.png"))))

        new_click_zones = {}; spacing = 80; total_width = len(timers_to_draw) * spacing
        block_start_x = base_x - (total_width // 2)
        block_center_y = base_y + 60
        
        for idx, (key, seconds, pixmap) in enumerate(timers_to_draw):
            x_pos = block_start_x + idx * spacing + (spacing // 2); block_center = QPoint(x_pos, block_center_y)
            if pixmap:
                icon_size = 40; icon_rect = QRect(block_center.x() - 20, block_center.y() - 45, 40, 40)
                if key != "preview": new_click_zones[key] = QRect(self.mapToGlobal(icon_rect.topLeft()), QSize(40, 40))
                if self.icon_frame: painter.drawPixmap(icon_rect.adjusted(-2, -2, 2, 2), self.icon_frame)
                painter.drawPixmap(icon_rect, pixmap.scaled(40, 40, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation))
            display_seconds = max(0, seconds); text = str(display_seconds)
            color = QColor(100, 255, 100) if seconds > 30 else QColor(255, 50, 50)
            if self.show_preview and not self.active_timers: color = QColor(255, 255, 255, 150)
            font = QFont("Segoe UI Black", 22 if seconds > 3 else 26); painter.setFont(font)
            text_rect = QRect(block_center.x() - 50, block_center.y() - 13, 100, 50)
            painter.setPen(QPen(QColor(0,0,0,200), 4)); painter.drawText(text_rect, Qt.AlignmentFlag.AlignCenter, text)
            painter.setPen(QPen(color, 2)); painter.drawText(text_rect, Qt.AlignmentFlag.AlignCenter, text)
        self.click_zones = new_click_zones

    def on_toggle_exp(self):
        self.show_exp_panel = not self.show_exp_panel
        status = "已啟用" if self.show_exp_panel else "已關閉"
        print(f"[Overlay] EXP Panel toggled: {status}")
        self.show_notification(f"📊 經驗監測系統 {status} (F10)")
        if not self.show_exp_panel:
            # Reset all session variables
            self.current_exp_data = {
                "text": "---", "value": 0, "percent": 0.0, 
                "gained_10m": 0, "percent_10m": 0.0, 
                "is_estimated": True, "tracking_duration": 0, "time_to_level": -1
            }
            self.exp_history = []
            self.exp_session_start_time = None
            self.exp_initial_val = None
        self.update()

    def draw_exp_panel(self, painter):
        if not self.show_exp_panel:
            return
            
        # Positional logic
        bx = self.rect().center().x() + self.x_offset
        by = self.rect().center().y() + self.y_offset
        pw, ph = 330, 115 # Reduced height
        px = bx - pw // 2
        py = by + 140
        panel_rect = QRect(px, py, pw, ph)
        
        # 1. Background
        path = QPainterPath()
        path.addRoundedRect(QRectF(panel_rect), 12, 12)
        painter.setPen(QPen(QColor(255, 215, 0, 255), 2))
        painter.setBrush(QColor(10, 10, 15, int(self.base_opacity * 255))) 
        painter.drawPath(path)
        
        # 2. Recording Duration
        duration_sec = self.current_exp_data.get("tracking_duration", 0)
        h_dur = duration_sec // 3600
        m_dur = (duration_sec % 3600) // 60
        s_dur = duration_sec % 60
        duration_text = f"{h_dur:02d}:{m_dur:02d}:{s_dur:02d}" if h_dur > 0 else f"{m_dur:02d}:{s_dur:02d}"
            
        painter.setPen(QColor(200, 200, 200))
        painter.setFont(QFont("Segoe UI", 9))
        painter.drawText(px + 15, py + 32, f"紀錄時長: {duration_text}")
        
        # 3. Time to Level Up
        painter.setPen(QColor(255, 215, 0))
        painter.setFont(QFont("Segoe UI Bold", 13))
        ttl_sec = self.current_exp_data.get("time_to_level", -1)
        if ttl_sec > 0:
            h = ttl_sec // 3600
            m = (ttl_sec % 3600) // 60
            ttl_text = f"升級預計還需: {h}小時 {m}分"
        else:
            ttl_text = "升級預計還需: 計算速率中..."
        painter.drawText(px + 15, py + 62, ttl_text)
        
        # 4. 10min Efficiency
        gain_val = self.current_exp_data.get("gained_10m", 0)
        gain_pct = self.current_exp_data.get("percent_10m", 0.0)
        is_est = self.current_exp_data.get("is_estimated", True)
        label = "（預估）" if is_est else ""
        gain_text = f"{label}10分鐘效率: +{gain_val:,} ({gain_pct:+.2f}%)"
        
        painter.setPen(QColor(100, 255, 100) if gain_val >= 0 else QColor(255, 100, 100))
        painter.setFont(QFont("Segoe UI Semibold", 11))
        painter.drawText(px + 15, py + 95, gain_text)

        # 6. Progress Bar (Bottom)
        progress_pct = self.current_exp_data.get("percent", 0.0)
        bar_full_width = pw - 30
        bar_width = int(bar_full_width * (max(0, min(100, progress_pct)) / 100.0))
        
        painter.setBrush(QColor(255, 255, 255, 30))
        painter.setPen(Qt.PenStyle.NoPen)
        painter.drawRect(px + 15, py + 128, bar_full_width, 3) 
        
        if bar_width > 0:
            painter.setBrush(QColor(255, 215, 0, 180))
            painter.drawRect(px + 15, py + 128, bar_width, 3)
        



