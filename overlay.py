import sys
import json
import os
import threading
import time
try:
    import win32gui
    import win32con
    import winsound
except ImportError:
    win32gui = None
    winsound = None

from PyQt6.QtWidgets import (QApplication, QWidget, QVBoxLayout, QHBoxLayout, 
                             QLabel, QLineEdit, QPushButton, QScrollArea, QFrame,
                             QGridLayout, QDialog, QTabWidget, QComboBox, QSystemTrayIcon)
from PyQt6.QtCore import Qt, QPoint, QRect, QTimer, pyqtSignal, QSize, QDir, QFileInfo
from PyQt6.QtGui import QFont, QColor, QPainter, QBrush, QPen, QLinearGradient, QPixmap, QIcon

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
                                p["triggers"][k] = {"seconds": int(v), "icon": ""}
                    return config
            except Exception as e: 
                print(f"Error loading config: {e}")
                pass
            
        # Default Multi-Profile Config
        default_profiles = {}
        for i in range(1, 10):
            default_profiles[f"F{i}"] = {"name": f"切換組 F{i}", "triggers": {}}
        return {"active_profile": "F1", "offset": [0, 0], "profiles": default_profiles}

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
                self.tabs.addTab(scroll, cat)
        
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
        self.setFixedSize(350, 500)
        self.setWindowFlags(Qt.WindowType.WindowStaysOnTopHint)
        self.is_recording = False
        self.handle = None
        self.trigger_data = {}
        self.init_ui()
        self.request_show.connect(self.safe_show)

    def init_ui(self):
        self.layout = QVBoxLayout(self)
        self.setStyleSheet("background-color: #121212; color: #e0e0e0; font-family: 'Segoe UI';")
        
        top_row = QHBoxLayout()
        title = QLabel("🎹 設定清單")
        title.setFont(QFont("Segoe UI Semibold", 18))
        title.setStyleSheet("color: #ffd700;")
        top_row.addWidget(title)
        
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
        self.layout.addLayout(profile_row)

        subtitle = QLabel("貼心提醒：雙擊 F1~F9 可快速切換配置組")
        subtitle.setStyleSheet("color: #888; font-size: 11px; margin-bottom: 10px;")
        self.layout.addWidget(subtitle)

        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setFrameShape(QFrame.Shape.NoFrame)
        self.scroll_content = QWidget()
        self.scroll_layout = QVBoxLayout(self.scroll_content)
        self.scroll.setWidget(self.scroll_content)
        self.layout.addWidget(self.scroll)

        self.update_profile_dropdown()
        self.refresh_items()

        act_layout = QVBoxLayout()
        self.record_btn = QPushButton("➕ 新增按鍵 (點我後按鍵盤)")
        self.record_btn.setStyleSheet("QPushButton { background-color: #2e7d32; color: white; font-weight: bold; border-radius: 5px; height: 35px; }")
        self.record_btn.clicked.connect(self.toggle_recording)
        act_layout.addWidget(self.record_btn)

        save_btn = QPushButton("💾 儲存並套用")
        save_btn.setStyleSheet("QPushButton { background-color: #ffd700; color: black; font-weight: bold; border-radius: 5px; height: 35px; margin-top: 5px;}")
        save_btn.clicked.connect(self.save_and_close)
        act_layout.addWidget(save_btn)

        pos_btn = QPushButton("🔱 調整顯示位置")
        pos_btn.setStyleSheet("QPushButton { background-color: #5c6bc0; color: white; border-radius: 5px; height: 30px; margin-top: 5px;}")
        pos_btn.clicked.connect(self.toggle_handle)
        act_layout.addWidget(pos_btn)

        exit_btn = QPushButton("🛑 關閉整個輔助")
        exit_btn.setStyleSheet("QPushButton { background-color: transparent; color: #666; border: none; font-size: 10px; margin-top: 20px; }")
        exit_btn.clicked.connect(QApplication.instance().quit)
        act_layout.addWidget(exit_btn)
        
        self.layout.addLayout(act_layout)

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
                active_ui_data[key] = {"seconds": int(ui["inp"].text()), "icon": ui["icon"]}
            except:
                active_ui_data[key] = {"seconds": 300, "icon": ui["icon"]}
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
            inp = QLineEdit(str(data.get("seconds", 300))); inp.setFixedWidth(50)
            del_btn = QPushButton("🗑️"); del_btn.setFixedWidth(30)
            del_btn.clicked.connect(lambda checked, k=key: self.delete_key(k))
            row.addWidget(lbl); row.addWidget(icon_btn); row.addWidget(inp); row.addWidget(QLabel("秒")); row.addStretch(); row.addWidget(del_btn)
            self.trigger_data[key] = {"inp": inp, "icon": data.get("icon", "")}
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
                new_triggers[key] = {"seconds": int(data["inp"].text()), "icon": icon_path}
            except: pass
        config["profiles"][p_key]["triggers"] = new_triggers
        config["profiles"][p_key]["name"] = self.nickname_inp.text()
        if self.overlay: config["offset"] = [self.overlay.x_offset, self.overlay.y_offset]
        config["active_profile"] = p_key
        ConfigManager.save_config(config)
        if self.handle: self.handle.hide()
        if self.overlay: self.overlay.show_preview = False
        self.config_updated.emit(); self.hide()

class ArtaleOverlay(QWidget):
    timer_request = pyqtSignal(str, int, str) 
    clear_request = pyqtSignal()
    notification_request = pyqtSignal(str)
    profile_switch_request = pyqtSignal()
    
    def __init__(self, target_window_title="Artale"):
        super().__init__()
        self.target_window_title = target_window_title
        self.active_timers = {} 
        self.click_zones = {}  
        self.is_active = False
        self.show_preview = False
        self.active_profile_name = "F1"
        self.msg_text = ""; self.msg_opacity = 0
        self.x_offset = 0; self.y_offset = 0
        
        self.timer_request.connect(self.start_timer)
        self.clear_request.connect(self.clear_all_timers)
        self.notification_request.connect(self.show_notification)
        self.profile_switch_request.connect(self.load_profile_immediately)
        
        self.tracking_timer = QTimer(self); self.tracking_timer.timeout.connect(self.sync_with_game_window); self.tracking_timer.start(100)
        self.countdown_timer = QTimer(self); self.countdown_timer.timeout.connect(self.update_countdown)
        
        frame_p = resource_path("buff_pngs/skill_frame.png")
        self.icon_frame = QPixmap(frame_p) if os.path.exists(frame_p) else None
        self.init_ui()
        self.load_profile_immediately()

    def init_ui(self):
        self.setWindowFlags(Qt.WindowType.WindowStaysOnTopHint | Qt.WindowType.FramelessWindowHint | Qt.WindowType.WindowTransparentForInput | Qt.WindowType.Tool)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.setAttribute(Qt.WidgetAttribute.WA_ShowWithoutActivating)
        virtual_geo = QApplication.primaryScreen().virtualGeometry()
        self.setGeometry(virtual_geo); self.show()

    def start_timer(self, key, seconds, icon_path=None):
        pixmap = None
        if icon_path:
            real_path = icon_path if os.path.exists(icon_path) else resource_path(icon_path)
            if os.path.exists(real_path): pixmap = QPixmap(real_path)
        self.active_timers[key] = {"seconds": seconds, "pixmap": pixmap}
        self.is_active = True
        if not self.countdown_timer.isActive(): self.countdown_timer.start(1000)
        self.update()

    def update_countdown(self):
        to_remove = []
        for key in list(self.active_timers.keys()):
            self.active_timers[key]["seconds"] -= 1
            rem = self.active_timers[key]["seconds"]
            if rem == 20: self.play_sound(1)
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
            rect = win32gui.GetWindowRect(hwnd); x, y, x2, y2 = rect
            if self.geometry() != QRect(x, y, x2-x, y2-y): self.setGeometry(x, y, x2-x, y2-y); self.update()

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
        if not self.is_active and not self.show_preview and self.msg_opacity == 0: return
        painter = QPainter(self); painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        
        # Base coordinates
        base_x = self.rect().center().x() + self.x_offset
        base_y = self.rect().center().y() + self.y_offset

        # 1. Profile/Action Notification (Centered above anchor)
        if self.msg_opacity > 0:
            font = QFont("Segoe UI Bold", 18); painter.setFont(font)
            tw = painter.fontMetrics().horizontalAdvance(self.msg_text)
            # Draw notification clearly above the timer block
            bg_rect = QRect(base_x - (tw+40)//2, base_y - 70, tw+40, 45)
            painter.setBrush(QColor(0, 0, 0, min(200, self.msg_opacity)))
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
