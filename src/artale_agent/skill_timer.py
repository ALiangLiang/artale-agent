import os
import time
import logging
import threading
from PyQt6.QtCore import Qt, QPoint, QTimer, pyqtSignal, QSize, QObject
from PyQt6.QtGui import QPixmap, QIcon
from PyQt6.QtWidgets import (QWidget, QVBoxLayout, QLabel, 
                             QPushButton, QScrollArea, QGridLayout, 
                             QDialog, QTabWidget)

from .platform import AudioPlayerImpl
from .utils import resource_path

logger = logging.getLogger(__name__)

# resource_path 已移除 (現位於 utils.py)
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
                rel = rel.replace("\\", "/")
                # Normalize: strip "assets/" prefix so stored paths match PyInstaller bundle layout
                if rel.startswith("assets/"):
                    rel = rel[len("assets/"):]
                self.selected_icon = rel
            else:
                self.selected_icon = abs_path
        except:
            self.selected_icon = abs_path
        self.accept()

class PositionHandle(QWidget):
    """用於拖曳調整 UI 位置的把手介面"""
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

class TimerManager(QObject):
    """管理所有活動計時器的核心邏輯及其生命週期"""
    updated = pyqtSignal()
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.active_timers = {}
        self.countdown_timer = QTimer(self)
        self.countdown_timer.timeout.connect(self.update_countdown)
        self.is_active = False

    def start_timer(self, key, seconds, icon_path=None, sound_enabled=True):
        pixmap = None
        if icon_path:
            real_path = icon_path
            if not os.path.exists(real_path):
                real_path = resource_path(icon_path)
            
            if os.path.exists(real_path):
                pixmap = QPixmap(real_path)
                if pixmap.isNull():
                    logger.error(f"[Timer] Failed to load icon: {real_path}")
                    pixmap = None
            else:
                logger.warning(f"[Timer] Icon not found: {icon_path}")
        
        self.active_timers[key] = {"seconds": seconds, "pixmap": pixmap, "sound_enabled": sound_enabled}
        self.is_active = True
        # 如果計時器尚未啟動，則立即啟動
        if not self.countdown_timer.isActive():
            self.countdown_timer.start(1000)
        self.updated.emit()

    def update_countdown(self):
        to_remove = []
        for key in list(self.active_timers.keys()):
            self.active_timers[key]["seconds"] -= 1
            rem = self.active_timers[key]["seconds"]
            sound_enabled = self.active_timers[key].get("sound_enabled", True)
            
            if rem == 20 and sound_enabled: self.play_sound(1)
            elif rem == 0: self.play_sound(2) # 倒數結束提示
            elif -10 < rem < 0: self.play_sound(1) # 負數超時提示
            
            if rem <= -10:
                to_remove.append(key)
        
        for key in to_remove:
            if key in self.active_timers:
                del self.active_timers[key]

        if not self.active_timers:
            self.is_active = False
            self.countdown_timer.stop()
        
        self.updated.emit()

    def clear_all(self):
        self.active_timers = {}
        self.is_active = False
        self.countdown_timer.stop()
        self.updated.emit()

    def play_sound(self, times=1):
        player = AudioPlayerImpl()
        def worker():
            for _ in range(times):
                try:
                    player.beep(800, 150)
                    time.sleep(0.12)
                except Exception as e:
                    logger.debug(f"[Sound] Beep failed: {e}")
        threading.Thread(target=worker, daemon=True).start()
