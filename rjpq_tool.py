import json
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

    def connect_to_room(self, code, pwd):
        if self.ws is None:
            print("[RJPQ Sync] Error: QWebSocket is not installed. Run 'pip install PyQt6-WebSockets'")
            return
        self.room_code = code
        self.room_pwd = pwd
        url = "wss://rjpq.juanwang.cc"
        self.ws.open(QUrl(url))

    def disconnect_from_room(self):
        if self.ws:
            self.ws.close()

    def on_connected(self):
        self.is_connected = True
        self.status_changed.emit(True)
        join_msg = {"type": "join", "code": self.room_code, "password": self.room_pwd}
        if self.ws:
            self.ws.sendTextMessage(json.dumps(join_msg))

    def on_disconnected(self):
        self.is_connected = False
        self.status_changed.emit(False)

    def on_message(self, message):
        try:
            msg = json.loads(message)
            if msg["type"] == "sync" and "data" in msg:
                self.sync_received.emit(msg["data"])
            elif msg["type"] == "charCounts" and "counts" in msg:
                self.char_counts_received.emit(msg["counts"])
            elif msg["type"] == "error":
                self.error_received.emit(msg["error"])
        except Exception as e:
            print(f"[RJPQ Sync] Message error: {e}")

    def on_error(self, error):
        print(f"[RJPQ Sync] Connection error")
        self.status_changed.emit(False)

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

    def on_error_message(self, error):
        QMessageBox.critical(self, "錯誤", f"YZY 伺服器回傳錯誤：\n{error}")
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
        self.pwd_inp.setFixedWidth(70)
        self.pwd_inp.setStyleSheet("background: #222; color: #ffd700; border: 1px solid #444; border-radius: 4px; padding: 4px;")
        
        self.conn_btn = QPushButton("連線")
        self.conn_btn.setFixedWidth(60)
        self.conn_btn.setStyleSheet("QPushButton { background: #333; color: #fff; font-weight: bold; border-radius: 4px; height: 26px; }")
        self.conn_btn.clicked.connect(self.on_connect_clicked)
        
        self.status_dot = QFrame()
        self.status_dot.setFixedSize(12, 12)
        self.status_dot.setStyleSheet("background: #555; border-radius: 6px;")
        
        room_row.addWidget(QLabel("房:"))
        room_row.addWidget(self.code_inp)
        room_row.addWidget(QLabel("密:"))
        room_row.addWidget(self.pwd_inp)
        room_row.addWidget(self.conn_btn)
        room_row.addWidget(self.status_dot)
        room_row.addStretch()
        self.main_layout.addWidget(self.room_widget)

        # 2. Char Widget (Visible after Connected)
        self.char_widget = QWidget()
        char_row = QHBoxLayout(self.char_widget)
        self.char_btns = []
        char_colors = ["#ff6b6b", "#51cf66", "#339af0", "#cc5de8"]
        for i in range(4):
            btn = QPushButton(f"10{i+1}")
            btn.setCheckable(True)
            btn.setFixedSize(65, 30)
            btn.setStyleSheet(f"QPushButton {{ background: #222; color: {char_colors[i]}; border: 1px solid #444; border-radius: 4px; }} "
                             f"QPushButton:checked {{ background: {char_colors[i]}; color: #fff; font-weight: bold; }}")
            btn.clicked.connect(lambda checked, idx=i: self.select_char(idx))
            self.char_btns.append(btn)
            char_row.addWidget(btn)
        self.char_widget.setVisible(False)
        self.main_layout.addWidget(self.char_widget)

        # 3. Grid Widget (Visible after Char Selected)
        self.grid_widget = QWidget()
        grid_vbox = QVBoxLayout(self.grid_widget)
        grid_vbox.setContentsMargins(0, 5, 0, 0)
        
        self.platform_btns = []
        grid_layout = QGridLayout()
        grid_layout.setSpacing(4)
        for row in range(10):
            row_label = QLabel(str(10 - row))
            row_label.setFixedWidth(20)
            row_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
            grid_layout.addWidget(row_label, row, 0)
            for col in range(4):
                idx = row * 4 + col
                btn = QPushButton(str(col + 1))
                btn.setFixedSize(60, 28)
                btn.setStyleSheet("QPushButton { background: #222; color: #888; border: 1px solid #333; border-radius: 2px; }")
                btn.clicked.connect(lambda checked, i=idx: self.platform_clicked(i))
                self.platform_btns.append(btn)
                grid_layout.addWidget(btn, row, col + 1)
        grid_vbox.addLayout(grid_layout)

        # Reset button in grid widget
        reset_btn = QPushButton("🔄 重置所有標記")
        reset_btn.setStyleSheet("QPushButton { background: #444; color: #eee; border-radius: 4px; height: 30px; margin-top: 5px; }")
        reset_btn.clicked.connect(self.on_reset_clicked)
        grid_vbox.addWidget(reset_btn)
        
        self.grid_widget.setVisible(False)
        self.main_layout.addWidget(self.grid_widget)
        
        self.main_layout.addStretch()

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
        color = "#51cf66" if connected else "#ff6b6b"
        self.status_dot.setStyleSheet(f"background: {color}; border-radius: 6px;")
        
        if connected:
            self.conn_btn.setText("中斷")
            self.conn_btn.setStyleSheet("QPushButton { background: #c62828; color: #fff; font-weight: bold; border-radius: 4px; height: 26px; }")
            self.conn_btn.setEnabled(True)
        else:
            self.conn_btn.setText("連線")
            self.conn_btn.setStyleSheet("QPushButton { background: #333; color: #fff; font-weight: bold; border-radius: 4px; height: 26px; }")
            self.conn_btn.setEnabled(True)
        
        # Dynamic Visibility
        self.char_widget.setVisible(connected)
        if not connected:
            self.grid_widget.setVisible(False)
            self.selected_color = -1
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
        for i in range(40):
            val = data[i]
            btn = self.platform_btns[i]
            if val < 4:
                btn.setStyleSheet(f"background: {char_colors[val]}; color: #fff; border: 1px solid #444; border-radius: 2px;")
                if self.selected_color != -1 and val != self.selected_color:
                    btn.setStyleSheet(f"background: {char_colors[val]}; color: #fff; border: 1px solid #444; border-radius: 2px; opacity: 0.5;")
            else:
                btn.setStyleSheet("QPushButton { background: #222; color: #888; border: 1px solid #333; border-radius: 2px; }")

    def on_reset_clicked(self):
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
    painter.setFont(QFont("Segoe UI Bold", 11))
    painter.drawText(px + 10, py + 25, "YzY 羅茱同步路徑")
    
    cell_w, cell_h = 32, 22
    start_x = px + 35
    start_y = py + 45
    char_colors = [QColor("#ff6b6b"), QColor("#51cf66"), QColor("#339af0"), QColor("#cc5de8")]
    
    for row in range(10):
        painter.setPen(QColor(150, 150, 150))
        painter.setFont(QFont("Segoe UI", 8))
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
