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
                             QGridLayout, QDialog, QTabWidget, QComboBox)
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
                with open(CONFIG_FILE, "r") as f:
                    config = json.load(f)
                    
                    # Migration: Old single profile -> Multi profile
                    if "profiles" not in config:
                        old_triggers = config.get("triggers", {"f1": {"seconds": 300, "icon": ""}})
                        old_offset = config.get("offset", [0, 0])
                        config = {
                            "active_profile": "Profile 1",
                            "offset": old_offset,
                            "profiles": {
                                "Profile 1": {"triggers": old_triggers}
                            }
                        }
                        # Initialize empty profiles for F2-F8
                        for i in range(2, 9):
                            config["profiles"][f"Profile {i}"] = {"triggers": {}}
                    
                    # Ensure root offset exists
                    if "offset" not in config:
                        # Try to pick from any profile if missing at root
                        first_p = list(config["profiles"].values())[0] if config["profiles"] else {}
                        config["offset"] = first_p.get("offset", [0, 0])

                    # Ensure migration for older trigger formats inside profiles
                    for p in config["profiles"].values():
                        for k, v in p["triggers"].items():
                            if isinstance(v, (int, float)):
                                p["triggers"][k] = {"seconds": int(v), "icon": ""}
                    return config
            except: pass
            
        # Default Multi-Profile Config
        default_profiles = {}
        for i in range(1, 9):
            default_profiles[f"Profile {i}"] = {"triggers": {}}
        return {"active_profile": "Profile 1", "offset": [0, 0], "profiles": default_profiles}

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
            # First, add priority items that exist
            for p in priority:
                if p in all_dirs:
                    categories.append(p)
                    all_dirs.remove(p)
            # Add remaining items alphabetically
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
                    
                    # Store path in a lambda correctly
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
        self.selected_icon = path
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
        
        # Style hint for the draggable area
        self.lbl.setStyleSheet("background: rgba(255, 255, 255, 50); border: 1px dashed white; border-radius: 5px;")
        
        self._dragging = False
        self._drag_start = QPoint()
        
        # Periodic sync with the game window while visible
        self.sync_timer = QTimer(self)
        self.sync_timer.timeout.connect(self.keep_on_game_center)

    def showEvent(self, event):
        super().showEvent(event)
        self.sync_timer.start(100)

    def hideEvent(self, event):
        super().hideEvent(event)
        self.sync_timer.stop()

    def keep_on_game_center(self):
        if self._dragging: return

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
        # Emit the handle's global center position
        gp = self.mapToGlobal(self.rect().center())
        self.position_changed.emit(gp.x(), gp.y())

    def sync_with_offset(self, ox, oy):
        hwnd = win32gui.FindWindow(None, self.target_window_title)
        if hwnd:
            rect = win32gui.GetWindowRect(hwnd)
            x, y, x2, y2 = rect
            cx, cy = x + (x2-x) // 2, y + (y2-y) // 2
            self.move(cx + ox - 25, cy + oy - 25)

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
        self.init_ui()
        self.request_show.connect(self.safe_show)

    def init_ui(self):
        self.layout = QVBoxLayout(self)
        self.setStyleSheet("background-color: #121212; color: #e0e0e0; font-family: 'Segoe UI';")
        
        # Header
        top_row = QHBoxLayout()
        title = QLabel("🎹 設定清單")
        title.setFont(QFont("Segoe UI Semibold", 18))
        title.setStyleSheet("color: #ffd700;")
        top_row.addWidget(title)
        
        # Profile Selector
        self.profile_box = QComboBox()
        for i in range(1, 9):
            self.profile_box.addItem(f"Profile {i}")
        
        config = ConfigManager.load_config()
        idx = self.profile_box.findText(config.get("active_profile", "Profile 1"))
        self.profile_box.setCurrentIndex(idx if idx >= 0 else 0)
        self.profile_box.currentTextChanged.connect(self.switch_profile_ui)
        self.profile_box.setStyleSheet("""
            QComboBox { background-color: #333; color: #ffd700; border-radius: 4px; padding: 5px; min-width: 100px; }
            QComboBox QAbstractItemView { background-color: #222; selection-background-color: #ffd700; selection-color: black; }
        """)
        top_row.addWidget(self.profile_box)
        self.layout.addLayout(top_row)

        subtitle = QLabel("貼心提醒：雙擊 F1~F8 可快速切換配置組")
        subtitle.setStyleSheet("color: #888; font-size: 11px; margin-bottom: 10px;")
        self.layout.addWidget(subtitle)

        # Scroll Area
        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setFrameShape(QFrame.Shape.NoFrame)
        self.scroll_content = QWidget()
        self.scroll_layout = QVBoxLayout(self.scroll_content)
        self.scroll.setWidget(self.scroll_content)
        self.layout.addWidget(self.scroll)

        self.refresh_items()

        # Action Buttons
        act_layout = QVBoxLayout()
        
        self.record_btn = QPushButton("➕ 新增按鍵 (點我後按鍵盤)")
        self.record_btn.setStyleSheet("""
            QPushButton { background-color: #2e7d32; color: white; font-weight: bold; border-radius: 5px; height: 35px; }
            QPushButton:hover { background-color: #388e3c; }
        """)
        self.record_btn.clicked.connect(self.toggle_recording)
        act_layout.addWidget(self.record_btn)

        save_btn = QPushButton("💾 儲存並套用")
        save_btn.setStyleSheet("""
            QPushButton { background-color: #ffd700; color: black; font-weight: bold; border-radius: 5px; height: 35px; margin-top: 5px;}
            QPushButton:hover { background-color: #ffea00; }
        """)
        save_btn.clicked.connect(self.save_and_close)
        act_layout.addWidget(save_btn)

        pos_btn = QPushButton("🔱 調整顯示位置 (在遊戲中央會出現黃色圖示)")
        pos_btn.setStyleSheet("""
            QPushButton { background-color: #5c6bc0; color: white; border-radius: 5px; height: 30px; margin-top: 5px;}
            QPushButton:hover { background-color: #7986cb; }
        """)
        pos_btn.clicked.connect(self.toggle_handle)
        act_layout.addWidget(pos_btn)

        exit_btn = QPushButton("🛑 關閉整個輔助")
        exit_btn.setStyleSheet("""
            QPushButton { background-color: transparent; color: #666; border: none; font-size: 10px; margin-top: 20px; }
            QPushButton:hover { color: #ff5555; }
        """)
        exit_btn.clicked.connect(QApplication.instance().quit)
        act_layout.addWidget(exit_btn)
        
        self.layout.addLayout(act_layout)

    def keyPressEvent(self, event):
        if self.is_recording:
            key_code = event.key()
            key_name = self.qt_key_to_name(key_code)
            if key_name:
                config = ConfigManager.load_config()
                p_name = self.profile_box.currentText()
                if key_name not in config["profiles"][p_name]["triggers"]:
                    config["profiles"][p_name]["triggers"][key_name] = {"seconds": 300, "icon": ""}
                    ConfigManager.save_config(config)
                    self.refresh_items()
                self.toggle_recording()

    def switch_profile_ui(self, p_name):
        config = ConfigManager.load_config()
        config["active_profile"] = p_name
        ConfigManager.save_config(config)
        self.refresh_items()
        if self.overlay:
            self.overlay.load_profile_immediately()
        self.config_updated.emit()

    def qt_key_to_name(self, code):
        """Maps Qt key code to pynput compatible string names."""
        special_map = {
            Qt.Key.Key_F1: "f1", Qt.Key.Key_F2: "f2", Qt.Key.Key_F3: "f3", Qt.Key.Key_F4: "f4",
            Qt.Key.Key_F5: "f5", Qt.Key.Key_F6: "f6", Qt.Key.Key_F7: "f7", Qt.Key.Key_F8: "f8",
            Qt.Key.Key_F9: "f9", Qt.Key.Key_F10: "f10", Qt.Key.Key_F11: "f11", Qt.Key.Key_F12: "f12",
            Qt.Key.Key_Shift: "shift", Qt.Key.Key_Control: "ctrl", Qt.Key.Key_Alt: "alt",
            Qt.Key.Key_Space: "space", Qt.Key.Key_Return: "enter", Qt.Key.Key_Enter: "enter",
            Qt.Key.Key_PageUp: "page_up", Qt.Key.Key_PageDown: "page_down",
            Qt.Key.Key_Home: "home", Qt.Key.Key_End: "end",
            Qt.Key.Key_Insert: "insert", Qt.Key.Key_Delete: "delete"
        }
        if code in special_map:
            return special_map[code]
        # Handle alphanumeric
        name = chr(code) if 32 <= code <= 126 else None
        return name.lower() if name else None

    def toggle_recording(self):
        self.is_recording = not self.is_recording
        if self.is_recording:
            self.record_btn.setText("🔴 請按鍵盤錄製中...")
            self.record_btn.setStyleSheet("background-color: #c62828; color: white; border-radius: 5px; height: 35px;")
        else:
            self.record_btn.setText("➕ 新增按鍵 (點我後按鍵盤)")
            self.record_btn.setStyleSheet("background-color: #2e7d32; color: white; border-radius: 5px; height: 35px;")

    def safe_show(self):
        self.show()
        self.activateWindow()
        self.raise_()
        
    def toggle_handle(self):
        if not self.handle:
            self.handle = PositionHandle()
            if self.overlay:
                self.handle.position_changed.connect(self.overlay.update_offset)
        
        if self.handle.isVisible():
            self.handle.hide()
            if self.overlay: self.overlay.show_preview = False
        else:
            # Position handle at current timer location
            if self.overlay:
                geo = self.overlay.geometry()
                cx = geo.x() + geo.width() // 2 + self.overlay.x_offset
                cy = geo.y() + geo.height() // 2 + self.overlay.y_offset
                self.handle.move(cx - 30, cy - 30)
                self.overlay.show_preview = True
                self.overlay.update()
            self.handle.show()
            self.handle.raise_()

    def refresh_items(self):
        while self.scroll_layout.count():
            item = self.scroll_layout.takeAt(0)
            widget = item.widget()
            if widget: widget.deleteLater()
            
        config = ConfigManager.load_config()
        p_name = self.profile_box.currentText()
        p_data = config["profiles"].get(p_name, {"triggers": {}})
        
        self.trigger_data = {}
        for key, data in p_data["triggers"].items():
            # Handle potential migration issues here again just in case
            if isinstance(data, int):
                data = {"seconds": data, "icon": ""}
                
            row_widget = QFrame()
            row_widget.setFixedHeight(50) 
            row_widget.setStyleSheet("background-color: #1e1e1e; border-radius: 4px; margin: 2px;")
            
            row = QHBoxLayout(row_widget)
            row.setContentsMargins(10, 0, 10, 0)
            
            lbl = QLabel(key.upper())
            lbl.setFont(QFont("Segoe UI Bold", 10))
            lbl.setStyleSheet("color: #ffd700;")
            lbl.setFixedWidth(60)
            
            # Icon Preview
            icon_btn = QPushButton()
            icon_btn.setFixedSize(32, 32)
            self.update_icon_button(icon_btn, data.get("icon", ""))
            icon_btn.clicked.connect(lambda checked, k=key, btn=icon_btn: self.pick_icon(k, btn))
            
            inp = QLineEdit(str(data.get("seconds", 300)))
            inp.setFixedWidth(50)
            inp.setStyleSheet("background-color: #333; border: 1px solid #444; color: white;")
            
            unit = QLabel("秒")
            
            del_btn = QPushButton("🗑️")
            del_btn.setCursor(Qt.CursorShape.PointingHandCursor)
            del_btn.setFixedWidth(30)
            del_btn.setStyleSheet("background-color: transparent; color: #666; font-size: 14px;")
            del_btn.clicked.connect(lambda checked, k=key: self.delete_key(k))
            
            row.addWidget(lbl)
            row.addWidget(icon_btn)
            row.addWidget(inp)
            row.addWidget(unit)
            row.addStretch()
            row.addWidget(del_btn)
            
            # Store data for saving
            self.trigger_data[key] = {
                "inp": inp,
                "icon": data.get("icon", "")
            }
            self.scroll_layout.addWidget(row_widget)
            
        # Add a single final stretch to keep items at the top
        self.scroll_layout.addStretch()

    def delete_key(self, key):
        config = ConfigManager.load_config()
        if key in config["triggers"]:
            del config["triggers"][key]
            ConfigManager.save_config(config)
            self.refresh_items()

    def update_icon_button(self, btn, path):
        if path and os.path.exists(path):
            pixmap = QPixmap(path).scaled(24, 24, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation)
            btn.setIcon(QIcon(pixmap))
            btn.setIconSize(QSize(24, 24))
            btn.setText("")
        else:
            btn.setIcon(QIcon())
            btn.setText("🖼️")
            btn.setStyleSheet("QPushButton { border: 1px dashed #444; color: #666; }")

    def pick_icon(self, key, btn):
        dlg = IconSelectorDialog(self)
        if dlg.exec():
            path = dlg.selected_icon
            self.trigger_data[key]["icon"] = path
            self.update_icon_button(btn, path)

    def save_and_close(self):
        config = ConfigManager.load_config()
        p_name = self.profile_box.currentText()
        
        # Update current profile in the master config
        new_triggers = {}
        for key, data in self.trigger_data.items():
            try:
                new_triggers[key] = {
                    "seconds": int(data["inp"].text()),
                    "icon": data["icon"]
                }
            except: pass
            
        config["profiles"][p_name]["triggers"] = new_triggers
        if self.overlay:
            config["offset"] = [self.overlay.x_offset, self.overlay.y_offset]
        
        config["active_profile"] = p_name
        ConfigManager.save_config(config)
        
        if self.handle: self.handle.hide()
        if self.overlay: self.overlay.show_preview = False
        self.config_updated.emit()
        self.hide()

class ArtaleOverlay(QWidget):
    timer_request = pyqtSignal(str, int, str) # key, seconds, icon_path
    
    def __init__(self, target_window_title="Artale"):
        super().__init__()
        self.target_window_title = target_window_title
        self.active_timers = {} 
        self.click_zones = {}  
        self.is_active = False
        self.show_preview = False
        
        # Profile & Notification UI
        self.active_profile_name = "Profile 1"
        self.msg_text = ""
        self.msg_opacity = 0
        
        self.load_profile_immediately()
        
        frame_p = resource_path("buff_pngs/skill_frame.png")
        self.icon_frame = QPixmap(frame_p) if os.path.exists(frame_p) else None
        self.init_ui()
        
        # Connect internal signal to ensure thread safety
        self.timer_request.connect(self.start_timer)
        
        # Tracking & Logic
        self.tracking_timer = QTimer(self)
        self.tracking_timer.timeout.connect(self.sync_with_game_window)
        self.tracking_timer.start(100)

        self.countdown_timer = QTimer(self)
        self.countdown_timer.timeout.connect(self.update_countdown)

    def init_ui(self):
        self.setWindowFlags(
            Qt.WindowType.WindowStaysOnTopHint |
            Qt.WindowType.FramelessWindowHint |
            Qt.WindowType.WindowTransparentForInput |
            Qt.WindowType.Tool 
        )
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.setAttribute(Qt.WidgetAttribute.WA_ShowWithoutActivating)
        
        # Use virtual geometry to span ALL screens
        virtual_geo = QApplication.primaryScreen().virtualGeometry()
        self.setGeometry(virtual_geo)
        self.show()

    def start_timer(self, key, seconds, icon_path=None):
        if icon_path:
            # Check if current path exists, otherwise try relative to resource_path
            real_path = icon_path if os.path.exists(icon_path) else resource_path(icon_path)
            pixmap = QPixmap(real_path) if os.path.exists(real_path) else None
        else:
            pixmap = None
            
        self.active_timers[key] = {
            "seconds": seconds,
            "pixmap": pixmap
        }
        self.is_active = True
        if not self.countdown_timer.isActive():
            self.countdown_timer.start(1000)
        self.update()

    def update_countdown(self):
        to_remove = []
        for key in list(self.active_timers.keys()):
            self.active_timers[key]["seconds"] -= 1
            
            # Sound triggers
            rem = self.active_timers[key]["seconds"]
            if rem == 20:
                self.play_sound(1)
            elif rem == 0:
                self.play_sound(2)
            elif rem < 0 and rem > -10:
                # Continuous beep during linger phase
                self.play_sound(1)
            
            # Linger on screen for 10 more seconds after hitting 0
            if rem <= -10:
                to_remove.append(key)
        
        for key in to_remove:
            if key in self.active_timers:
                del self.active_timers[key]
            
        if not self.active_timers:
            self.is_active = False
            self.countdown_timer.stop()
        
        self.update()

    def play_sound(self, times=1):
        if not winsound: return
        def worker():
            for _ in range(times):
                try:
                    winsound.Beep(800, 150)
                    time.sleep(0.12)
                except:
                    pass
        threading.Thread(target=worker, daemon=True).start()

    def sync_with_game_window(self):
        if not win32gui: return
        
        # Consistent window search
        hwnd = 0
        try:
            def callback(h, extra):
                nonlocal hwnd
                title = win32gui.GetWindowText(h)
                if self.target_window_title.lower() in title.lower() and win32gui.IsWindowVisible(h):
                    hwnd = h
                    return False
                return True
            win32gui.EnumWindows(callback, None)
        except:
            hwnd = 0

        if hwnd:
            rect = win32gui.GetWindowRect(hwnd)
            x, y, x2, y2 = rect
            if self.geometry() != QRect(x, y, x2-x, y2-y):
                self.setGeometry(x, y, x2-x, y2-y)
                self.update()

    def update_offset(self, global_x, global_y):
        # Convert global screen position to local offset
        local = self.mapFromGlobal(QPoint(global_x, global_y))
        self.x_offset = local.x() - self.rect().center().x()
        self.y_offset = local.y() - self.rect().center().y()
        self.click_zones = {} # Clear zones on drag to avoid phantom clicks
        self.update()

    def clear_all_timers(self):
        """Immediately stops all timers and clears the screen."""
        self.active_timers = {}
        self.click_zones = {}
        self.is_active = False
        if self.countdown_timer.isActive():
            self.countdown_timer.stop()
        self.show_notification("⚠️ 已強制關閉並重設 (F9)")
        self.update()

    def check_right_click(self, gx, gy):
        """Called from global mouse listener to cancel a timer via right click."""
        p = QPoint(gx, gy)
        for key, rect in list(self.click_zones.items()):
            if rect.contains(p):
                if key in self.active_timers:
                    del self.active_timers[key]
                self.update()
                return True
        return False

    def load_profile_immediately(self):
        config = ConfigManager.load_config()
        self.active_profile_name = config.get("active_profile", "Profile 1")
        self.x_offset, self.y_offset = config.get("offset", [0, 0])
        self.show_notification(f"切換至 {self.active_profile_name}")
        self.update()

    def show_notification(self, text):
        self.msg_text = text
        self.msg_opacity = 255
        QTimer.singleShot(2000, self.fade_notification)
        self.update()

    def fade_notification(self):
        self.msg_opacity = 0
        self.update()

    def paintEvent(self, event):
        if not self.is_active and not self.show_preview and self.msg_opacity == 0: return
        
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        
        # 1. Profile Switch Notification
        if self.msg_opacity > 0:
            painter.setPen(QColor(255, 215, 0, self.msg_opacity))
            painter.setFont(QFont("Segoe UI Bold", 16))
            msg_rect = QRect(self.rect().width()//2 - 100, 50, 200, 40)
            painter.drawText(msg_rect, Qt.AlignmentFlag.AlignCenter, self.msg_text)

        if not self.active_timers and not self.show_preview: return
        timers_to_draw = [] # list of (key, seconds, pixmap)
        if self.active_timers:
            # Sort by remaining seconds (Descending: Longest first on the left)
            sorted_active = sorted(self.active_timers.items(), key=lambda x: x[1]["seconds"], reverse=True)
            for key, data in sorted_active:
                timers_to_draw.append((key, data["seconds"], data["pixmap"]))
        elif self.show_preview:
            # Show a dummy preview timer
            preview_pixmap = QPixmap(resource_path("buff_pngs/arrow.png"))
            timers_to_draw.append(("preview", 300, preview_pixmap))

        # Reset zones for this frame
        new_click_zones = {}
        
        spacing = 80 
        base_x = self.rect().center().x() + self.x_offset
        base_y = self.rect().center().y() + self.y_offset
        block_center_y = base_y + 60
        
        for idx, (key, seconds, pixmap) in enumerate(timers_to_draw):
            x_pos = base_x + idx * spacing + 20
            block_center = QPoint(x_pos, block_center_y)
            
            # Draw Icon
            if pixmap:
                icon_size = 40
                icon_rect = QRect(block_center.x() - icon_size // 2, block_center.y() - 45, icon_size, icon_size)
                
                # Store Click Zone (Global)
                if key != "preview":
                    global_top_left = self.mapToGlobal(icon_rect.topLeft())
                    new_click_zones[key] = QRect(global_top_left, QSize(icon_size, icon_size))

                if self.icon_frame:
                    painter.drawPixmap(icon_rect.adjusted(-2, -2, 2, 2), self.icon_frame)
                
                painter.drawPixmap(icon_rect, pixmap.scaled(icon_size, icon_size, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation))

            # Draw Number (show "0" during linger phase)
            display_seconds = max(0, seconds)
            text = str(display_seconds)
            color = QColor(100, 255, 100) if seconds > 30 else QColor(255, 50, 50)
            if self.show_preview and not self.active_timers:
                color = QColor(255, 255, 255, 150) # Grayish for preview
            
            font_size = 22 if seconds > 3 else 26
            font = QFont("Segoe UI Black", font_size)
            painter.setFont(font)
            
            text_y_offset = 12 if pixmap else 0
            text_rect = QRect(block_center.x() - 50, block_center.y() + text_y_offset - 25, 100, 50)
            
            # Outline
            painter.setPen(QPen(QColor(0,0,0,200), 4))
            painter.drawText(text_rect, Qt.AlignmentFlag.AlignCenter, text)
            
            # Main
            painter.setPen(QPen(color, 2))
            painter.drawText(text_rect, Qt.AlignmentFlag.AlignCenter, text)

        self.click_zones = new_click_zones

        self.click_zones = new_click_zones
