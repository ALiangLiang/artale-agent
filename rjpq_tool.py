import json
import logging
import urllib.request
from PyQt6.QtCore import Qt, QPoint, QTimer, pyqtSignal, QObject, QUrl, QSize
from PyQt6.QtGui import QFont, QColor, QPainter, QPen, QPainterPath, QBrush
from PyQt6.QtWidgets import QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, QPushButton, QFrame, QGridLayout, QMessageBox

try:
    from PyQt6.QtWebSockets import QWebSocket
except ImportError:
    QWebSocket = None
from PyQt6.QtNetwork import QAbstractSocket

# --- RJPQ Sync Client ---
class RJPQSyncClient(QObject):
    sync_received = pyqtSignal(list)
    char_counts_received = pyqtSignal(list)
    status_changed = pyqtSignal(bool)
    error_received = pyqtSignal(str)
    room_created = pyqtSignal(str, str)
    overlay_toggle_request = pyqtSignal(bool) # New signal to talk to overlay

    def __init__(self):
        super().__init__()
        self.ws = None
        if QWebSocket is not None:
            self.ws = QWebSocket()
            self.ws.connected.connect(self.on_connected)
            self.ws.disconnected.connect(self.on_disconnected)
            self.ws.textMessageReceived.connect(self.on_message)
            self.ws.errorOccurred.connect(self.on_error)
        self.room_code = ""
        self.room_pwd = ""
        self.is_connected = False
        self.reconnect_enabled = False # Becomes True after user manually connects
        
        # Reconnect timer
        self.reconnect_timer = QTimer(self)
        self.reconnect_timer.setSingleShot(True)
        self.reconnect_timer.timeout.connect(self.perform_reconnect)

    def connect_to_room(self, code, pwd):
        if self.ws is None:
            logging.error("[RJPQ Sync] Error: QWebSocket is not installed. Run 'pip install PyQt6-WebSockets'")
            return
        self.room_code = code
        self.room_pwd = pwd
        self.reconnect_enabled = True # Enable auto-reconnect
        url = "wss://rjpq.juanwang.cc"
        self.ws.open(QUrl(url))

    def disconnect_from_room(self):
        self.reconnect_enabled = False # Disable auto-reconnect when user clicks disconnect
        if self.ws:
            self.ws.close()

    def on_connected(self):
        self.is_connected = True
        self.status_changed.emit(True)
        self.reconnect_timer.stop()
        join_msg = {"type": "join", "code": self.room_code, "password": self.room_pwd}
        if self.ws:
            self.ws.sendTextMessage(json.dumps(join_msg))

    def on_disconnected(self):
        self.is_connected = False
        self.status_changed.emit(False)
        
        # Trigger auto-reconnect if enabled
        if self.reconnect_enabled:
            logging.info("[RJPQ Sync] Unexpectedly disconnected. Reconnecting in 3s...")
            self.reconnect_timer.start(3000)
            
    def perform_reconnect(self):
        if self.reconnect_enabled and not self.is_connected:
            self.connect_to_room(self.room_code, self.room_pwd)

    def on_message(self, message):
        try:
            msg = json.loads(message)
            if msg["type"] == "sync" and "data" in msg:
                self.sync_received.emit(msg["data"])
            elif msg["type"] == "charCounts" and "counts" in msg:
                self.char_counts_received.emit(msg["counts"])
            elif msg["type"] == "created":
                self.room_created.emit(msg["code"], msg.get("password", ""))
            elif msg["type"] == "error":
                self.error_received.emit(msg["error"])
            elif msg["type"] == "pong":
                # Silently ignore pong
                pass
            else:
                # Silently ignore other unknown types to avoid terminal spam/errors
                pass
        except Exception as e:
            logging.error(f"[RJPQ Sync] Message error: {e}")

    def on_error(self, error):
        logging.error(f"[RJPQ Sync] Connection error")
        self.status_changed.emit(False)

    def create_room(self, pwd):
        try:
            url = f"https://rjpq.juanwang.cc/api/room?action=create&pwd={pwd}"
            with urllib.request.urlopen(url) as response:
                if response.getcode() == 200:
                    data = json.loads(response.read().decode())
                    if "error" in data:
                        self.error_received.emit(data["error"])
                    else:
                        self.room_created.emit(data["code"], data.get("password", ""))
                else:
                    self.error_received.emit("伺服器請求失敗")
        except Exception as e:
            self.error_received.emit(f"建立房間失敗: {str(e)}")

    def send_action(self, action):
        if self.ws and self.ws.state() == QAbstractSocket.SocketState.ConnectedState:
            self.ws.sendTextMessage(json.dumps(action))

# --- RJPQ UI Component ---
class RJPQTabContent(QWidget):
    color_selected = pyqtSignal(int)
    
    def __init__(self, client):
        super().__init__()
        self.client = client
        self.selected_color = -1
        self.current_data = [4] * 40
        self.init_ui()
        
        self.client.status_changed.connect(self.update_status)
        self.client.sync_received.connect(self.update_grid)
        self.client.error_received.connect(self.on_error_message)
        self.client.room_created.connect(self.on_room_created)

    def on_error_message(self, error):
        logging.error(f"[RJPQ Sync Error] {error}")
        # Only show critical popups for real failures, not protocol warnings
        if "失敗" in error or "連線" in error:
            QMessageBox.critical(self, "錯誤", f"YZY 伺服器錯誤：\n{error}")
        self.update_status(False)

    def init_ui(self):
        self.main_layout = QVBoxLayout(self)
        
        # 1. Room Widget (Always Visible)
        self.room_widget = QWidget()
        room_row = QHBoxLayout(self.room_widget)
        self.code_inp = QLineEdit()
        self.code_inp.setPlaceholderText("房間")
        self.code_inp.setMaxLength(6)
        self.code_inp.setFixedWidth(70)
        self.code_inp.setStyleSheet("background: #222; color: #ffd700; border: 1px solid #444; border-radius: 4px; padding: 4px;")
        
        self.pwd_inp = QLineEdit()
        self.pwd_inp.setPlaceholderText("密碼")
        self.pwd_inp.setMaxLength(4)
        self.pwd_inp.setFixedWidth(70)
        self.pwd_inp.setStyleSheet("background: #222; color: #ffd700; border: 1px solid #444; border-radius: 4px; padding: 4px;")
        
        self.conn_btn = QPushButton("連線")
        self.conn_btn.setFixedWidth(60)
        self.conn_btn.setStyleSheet("QPushButton { background: #333; color: #fff; font-weight: bold; border-radius: 4px; height: 26px; }")
        self.conn_btn.clicked.connect(self.on_connect_clicked)
        
        self.create_btn = QPushButton("創建")
        self.create_btn.setFixedWidth(60)
        self.create_btn.setStyleSheet("QPushButton { background: #333; color: #88ccff; font-weight: bold; border-radius: 4px; height: 26px; }")
        self.create_btn.clicked.connect(self.on_create_clicked)
        
        self.status_dot = QFrame()
        self.status_dot.setFixedSize(12, 12)
        self.status_dot.setStyleSheet("background: #555; border-radius: 6px;")
        
        room_row.addWidget(QLabel("房:"))
        room_row.addWidget(self.code_inp)
        room_row.addWidget(QLabel("密:"))
        room_row.addWidget(self.pwd_inp)
        room_row.addWidget(self.conn_btn)
        room_row.addWidget(self.create_btn)
        room_row.addWidget(self.status_dot)
        room_row.addStretch()
        self.main_layout.addWidget(self.room_widget)

        # 1.1 Overlay Visibility Toggle
        self.overlay_ctrl = QWidget()
        ctrl_layout = QHBoxLayout(self.overlay_ctrl)
        ctrl_layout.setContentsMargins(10, 0, 10, 5)
        
        from PyQt6.QtWidgets import QCheckBox
        self.overlay_cb = QCheckBox("在遊戲畫面顯示路徑面板")
        self.overlay_cb.setStyleSheet("color: #00ffff; font-size: 11px; font-weight: bold;")
        self.overlay_cb.toggled.connect(self.client.overlay_toggle_request.emit)
        ctrl_layout.addWidget(self.overlay_cb)
        self.main_layout.addWidget(self.overlay_ctrl)

        # 2. Char Widget (Visible after Connected)
        self.char_widget = QWidget()
        char_row = QHBoxLayout(self.char_widget)
        self.char_btns = []
        char_colors = ["#ff6b6b", "#51cf66", "#339af0", "#cc5de8"]
        for i in range(4):
            btn = QPushButton(f"10{i+1}")
            btn.setCheckable(True)
            btn.setFixedSize(65, 30)
            btn.setStyleSheet(f"QPushButton {{ background: #222; color: {char_colors[i]}; border: 1px solid #444; border-radius: 4px; font-family: 'Microsoft JhengHei', '微軟正黑體', sans-serif; }} "
                             f"QPushButton:checked {{ background: {char_colors[i]}; color: #fff; font-weight: bold; font-family: 'Microsoft JhengHei', '微軟正黑體', sans-serif; }}")
            btn.clicked.connect(lambda checked, idx=i: self.select_char(idx))
            self.char_btns.append(btn)
            char_row.addWidget(btn)
        self.char_widget.setVisible(False)
        self.main_layout.addWidget(self.char_widget)

        # 3. Grid Widget (Visible after Char Selected)
        self.grid_widget = QFrame()
        self.grid_widget.setObjectName("Dashboard")
        self.grid_widget.setStyleSheet("""
            QFrame#Dashboard { 
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #1a1a1a, stop:1 #121212);
                border: 1px solid #333; 
                border-radius: 12px;
                padding: 10px;
            }
        """)
        grid_vbox = QVBoxLayout(self.grid_widget)
        grid_vbox.setContentsMargins(10, 15, 10, 10)
        
        dashboard_title = QLabel("📡 YzY 團隊路徑中控台")
        dashboard_title.setStyleSheet("color: #888; font-size: 10px; font-weight: bold; margin-bottom: 5px;")
        dashboard_title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        grid_vbox.addWidget(dashboard_title)
        
        self.platform_btns = []
        grid_layout = QGridLayout()
        grid_layout.setSpacing(6)
        for row in range(10):
            row_label = QLabel(str(10 - row))
            row_label.setFixedWidth(20)
            row_label.setStyleSheet("color: #444; font-weight: bold; font-family: 'Microsoft JhengHei', '微軟正黑體', sans-serif;")
            row_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
            grid_layout.addWidget(row_label, row, 0)
            for col in range(4):
                idx = row * 4 + col
                btn = QPushButton(str(col + 1))
                btn.setFixedSize(58, 28)
                btn.setCursor(Qt.CursorShape.PointingHandCursor)
                # Standard glass style
                btn.setStyleSheet("QPushButton { background: rgba(255,255,255,0.05); color: #666; border: 1px solid #222; border-radius: 4px; font-weight: bold; } "
                                 "QPushButton:hover { background: rgba(255,255,255,0.1); border-color: #444; }")
                btn.clicked.connect(lambda checked, i=idx: self.platform_clicked(i))
                self.platform_btns.append(btn)
                grid_layout.addWidget(btn, row, col + 1)
        grid_vbox.addLayout(grid_layout)

        # Reset button in grid widget
        reset_btn = QPushButton("🔄 重置所有人的標記 (全隊歸零)")
        reset_btn.setStyleSheet("""
            QPushButton { 
                background: #331111; color: #ff8888; border: 1px solid #552222; 
                border-radius: 6px; height: 32px; margin-top: 10px; font-weight: bold;
            }
            QPushButton:hover { background: #442222; }
        """)
        reset_btn.clicked.connect(self.on_reset_clicked)
        grid_vbox.addWidget(reset_btn)
        
        self.grid_widget.setVisible(False)
        self.main_layout.addWidget(self.grid_widget)
        
        self.main_layout.addStretch()
        
        yzy_credit = QLabel("💫 感謝 YzY 公會提供優秀補助工具")
        yzy_credit.setStyleSheet("color: #666; font-size: 10px; margin-top: 5px;")
        yzy_credit.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.main_layout.addWidget(yzy_credit)

    def find_target_row(self):
        if self.selected_color == -1: return -1
        # Loop from row 1 (bottom) to 10 (top). Grid indices 0-3 is row 10, 36-39 is row 1
        for row in range(9, -1, -1):
            row_marked_by_me = False
            for col in range(4):
                idx = row * 4 + col
                if self.current_data[idx] == self.selected_color:
                    row_marked_by_me = True
                    break
            if not row_marked_by_me:
                return row
        return -1

    def mark_by_hotkey(self, col_index):
        if not self.client.is_connected or self.selected_color == -1:
            return False
        target_row = self.find_target_row()
        if target_row != -1:
            idx = target_row * 4 + col_index
            self.platform_clicked(idx)
            return True
        return False

    def on_create_clicked(self):
        pwd = self.pwd_inp.text()
        if not pwd:
            QMessageBox.warning(self, "提醒", "請先輸入你想設定的房間密碼")
            return
        self.client.create_room(pwd)
        self.create_btn.setText("創建中")
        self.create_btn.setEnabled(False)

    def on_room_created(self, code, pwd):
        # Auto fill the room code and notify
        self.code_inp.setText(code)
        self.create_btn.setText("創建")
        self.create_btn.setEnabled(True)
        self.create_btn.setVisible(False) # Hide after success
        # Removed success popup as per user request
        # Trigger actual join process
        self.on_connect_clicked()

    def on_connect_clicked(self):
        if self.client.is_connected:
            self.client.disconnect_from_room()
            return

        code = self.code_inp.text()
        pwd = self.pwd_inp.text()
        if len(code) == 6 and pwd:
            self.client.connect_to_room(code, pwd)
            self.conn_btn.setText("連線中")
            self.conn_btn.setEnabled(False)

    def update_status(self, connected):
        if not hasattr(self, 'disconnect_timer'):
            self.disconnect_timer = QTimer(self)
            self.disconnect_timer.setSingleShot(True)
            self.disconnect_timer.timeout.connect(self.hide_ui_on_disconnect)
            
        color = "#51cf66" if connected else "#ff6b6b"
        self.status_dot.setStyleSheet(f"background: {color}; border-radius: 6px;")
        
        if connected:
            self.disconnect_timer.stop()
            self.conn_btn.setText("中斷")
            self.conn_btn.setStyleSheet("QPushButton { background: #c62828; color: #fff; font-weight: bold; border-radius: 4px; height: 26px; }")
            self.conn_btn.setEnabled(True)
            
            self.char_widget.setVisible(True)
            if self.selected_color != -1:
                self.grid_widget.setVisible(True)
                # Auto-reselect character if we had one
                QTimer.singleShot(500, lambda: self.select_char(self.selected_color))
        else:
            self.conn_btn.setText("連線")
            self.conn_btn.setStyleSheet("QPushButton { background: #333; color: #fff; font-weight: bold; border-radius: 4px; height: 26px; }")
            self.conn_btn.setEnabled(True)
            
            # Start 3s grace timer
            if not self.disconnect_timer.isActive():
                logging.info("[RJPQ Sync] Grace period started (3s)...")
                self.disconnect_timer.start(3000)
                # Keep status dot red but don't hide widgets yet
        
        self.create_btn.setVisible(not connected)

    def hide_ui_on_disconnect(self):
        # Only hide if we are still actually disconnected
        if not self.client.is_connected:
            logging.info("[RJPQ Sync] Grace period expired. Hiding UI.")
            self.char_widget.setVisible(False)
            self.grid_widget.setVisible(False)
            for btn in self.char_btns: btn.setChecked(False)

    def select_char(self, idx):
        self.selected_color = idx
        for i, btn in enumerate(self.char_btns):
            btn.setChecked(i == idx)
        self.client.send_action({"type": "selectChar", "color": idx})
        self.color_selected.emit(idx)
        
        # Show grid
        self.grid_widget.setVisible(True)
        self.update_grid(self.current_data)

    def platform_clicked(self, index):
        if not self.client.is_connected:
            QMessageBox.warning(self, "提醒", "請先連線到同步房間")
            return
        if self.selected_color == -1:
            QMessageBox.warning(self, "提醒", "請先選擇你的角色 (101-104)")
            return
        current_val = self.current_data[index]
        if current_val == self.selected_color:
            self.client.send_action({"type": "unmark", "index": index, "color": self.selected_color})
        elif current_val == 4:
            self.client.send_action({"type": "mark", "index": index, "color": self.selected_color})

    def update_grid(self, data):
        self.current_data = data
        char_colors = ["#ff6b6b", "#51cf66", "#339af0", "#cc5de8"]
        target_row = self.find_target_row()
        
        for i in range(40):
            val = data[i]
            btn = self.platform_btns[i]
            row_i = i // 4
            
            is_target = (row_i == target_row)
            border_style = "2px solid #ffd700" if is_target else "1px solid #333"
            
            if val < 4:
                # Active state: Glow background with white text
                btn.setStyleSheet(f"QPushButton {{ background: {char_colors[val]}; color: #fff; border: {border_style}; border-radius: 4px; font-weight: bold; }}")
                if self.selected_color != -1 and val != self.selected_color:
                    btn.setStyleSheet(f"QPushButton {{ background: {char_colors[val]}; color: #fff; border: {border_style}; border-radius: 4px; font-weight: bold; opacity: 0.6; }}")
            else:
                # Idle state: Darkened glass
                btn.setStyleSheet(f"QPushButton {{ background: rgba(255,255,255,0.03); color: #444; border: {border_style}; border-radius: 4px; }}")

    def on_reset_clicked(self):
        reply = QMessageBox.question(self, "確認重置", "確定要重置所有標記嗎？\n這將清空全隊目前的路徑紀錄。", 
                                     QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if reply == QMessageBox.StandardButton.Yes:
            self.client.send_action({"type": "reset"})

# --- Overlay Drawer ---
def draw_rjpq_panel(painter, px, py, pw, ph, opacity, data, selected_color):
    painter.setRenderHint(QPainter.RenderHint.Antialiasing)
    path = QPainterPath()
    path.addRoundedRect(px, py, pw, ph, 12, 12)
    painter.setPen(QPen(QColor(0, 255, 255, 200), 2))
    painter.setBrush(QColor(10, 15, 20, int(opacity * 255)))
    painter.drawPath(path)
    
    painter.setPen(QColor(0, 255, 255))
    font = QFont()
    font.setFamilies(["Microsoft JhengHei", "微軟正黑體"])
    font.setPointSize(11)
    font.setBold(True)
    painter.setFont(font)
    painter.drawText(px + 10, py + 25, "羅茱平台標記")
    
    cell_w, cell_h = 32, 22
    start_x = px + 35
    start_y = py + 45
    char_colors = [QColor("#ff6b6b"), QColor("#51cf66"), QColor("#339af0"), QColor("#cc5de8")]
    
    for row in range(10):
        painter.setPen(QColor(150, 150, 150))
        font = QFont()
        font.setFamilies(["Microsoft JhengHei", "微軟正黑體"])
        font.setPointSize(8)
        painter.setFont(font)
        painter.drawText(px + 10, start_y + row*25 + 16, str(10-row))
        
        for col in range(4):
            idx = row * 4 + col
            cx = start_x + col * 35
            cy = start_y + row * 25
            val = data[idx]
            
            painter.setPen(QPen(QColor(255, 255, 255, 40), 1))
            painter.setBrush(Qt.BrushStyle.NoBrush)
            painter.drawRoundedRect(cx, cy, cell_w, cell_h, 2, 2)
            
            if val < 4:
                color = QColor(char_colors[val])
                alpha = 255 if (selected_color == -1 or val == selected_color) else 100
                color.setAlpha(alpha)
                painter.setPen(Qt.PenStyle.NoPen)
                painter.setBrush(color)
                painter.drawEllipse(QPoint(int(cx + cell_w//2), int(cy + cell_h//2)), 7, 7)
                if alpha == 255:
                    painter.setBrush(QColor(255, 255, 255, 150))
                    painter.drawEllipse(QPoint(int(cx + cell_w//2), int(cy + cell_h//2)), 3, 3)

    # Draw Target Row Highlight (Yellow Border)
    if selected_color != -1:
        # Find target row using duplicate logic for drawer
        target_row = -1
        for row in range(9, -1, -1):
            row_marked_by_me = False
            for col in range(4):
                if data[row * 4 + col] == selected_color:
                    row_marked_by_me = True
                    break
            if not row_marked_by_me:
                target_row = row
                break
        
        if target_row != -1:
            row_y = start_y + target_row * 25
            painter.setPen(QPen(QColor(255, 215, 0, 180), 2))
            painter.setBrush(Qt.BrushStyle.NoBrush)
            painter.drawRoundedRect(start_x - 5, row_y - 2, cell_w * 4 + 15, cell_h + 4, 4, 4)
