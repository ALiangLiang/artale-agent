import json
import logging
import certifi
import os
logger = logging.getLogger(__name__)
import urllib.request
import urllib.parse
from PyQt6.QtCore import Qt, QPoint, QTimer, pyqtSignal, QObject, QUrl
from PyQt6.QtGui import QFont, QColor, QPainter, QPen, QPainterPath
from PyQt6.QtWidgets import QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, QPushButton, QFrame, QGridLayout, QMessageBox, QCheckBox

from PyQt6.QtWebSockets import QWebSocket
from PyQt6.QtNetwork import QAbstractSocket, QSslSocket, QSslConfiguration

# 在模組層級檢查 SSL 支援
try:
    ssl_ok = QSslSocket.supportsSsl()
    import logging
    temp_logger = logging.getLogger("Artale")
    temp_logger.info(f"[RJPQ] Qt6 SSL Support: {ssl_ok}")
    if ssl_ok:
        temp_logger.info(f"[RJPQ] SSL Library Build: {QSslSocket.sslLibraryBuildVersionString()}")
        temp_logger.info(f"[RJPQ] SSL Library Runtime: {QSslSocket.sslLibraryVersionString()}")
except Exception as e:
    pass

from .utils import platform_font_family, platform_font_families

# --- RJPQ Sync Client ---
class RJPQSyncClient(QObject):
    sync_received = pyqtSignal(list)
    char_counts_received = pyqtSignal(list)
    status_changed = pyqtSignal(bool)
    error_received = pyqtSignal(str)
    room_created = pyqtSignal(str, str)
    overlay_toggle_request = pyqtSignal(bool) # 用於與 overlay 通訊的新訊號

    def __init__(self):
        super().__init__()
        self.ws = QWebSocket()
        self.ws.connected.connect(self.on_connected)
        self.ws.disconnected.connect(self.on_disconnected)
        self.ws.textMessageReceived.connect(self.on_message)
        self.ws.errorOccurred.connect(self.on_error)
        self.room_code = ""
        self.room_pwd = ""
        self.is_connected = False
        self.reconnect_enabled = False # 在使用者手動連線後變更為 True
        
        # 重新連線計時器
        self.reconnect_timer = QTimer(self)
        self.reconnect_timer.setSingleShot(True)
        self.reconnect_timer.timeout.connect(self.perform_reconnect)

    def connect_to_room(self, code, pwd):
        try:
            logger.info(f"[RJPQ] Connecting to room: {code}")
            self.room_code = code
            self.room_pwd = pwd
            self.reconnect_enabled = True 
            url = "wss://rjpq.juanwang.cc"
            
            # 優先使用預設配置，必要時使用 certifi 補充
            from PyQt6.QtNetwork import QSslConfiguration
            conf = QSslConfiguration.defaultConfiguration()
            
            # 若處於 EXE 封裝後或系統缺乏憑證，手動從 certifi 載入
            import sys
            if getattr(sys, 'frozen', False) or not conf.caCertificates():
                if os.path.exists(certifi.where()):
                    from PyQt6.QtNetwork import QSslCertificate
                    certs = QSslCertificate.fromPath(certifi.where())
                    conf.setCaCertificates(certs)
                    # 為提升封裝版穩定性，明確允許標準連線交握
                    logger.info(f"[RJPQ] 已手動載入 CA 憑證 (數量: {len(certs)})")

            self.ws.setSslConfiguration(conf)
            self.ws.open(QUrl(url))
        except Exception as e:
            logger.error(f"[RJPQ] Connection failure: {e}")

    def disconnect_from_room(self):
        self.reconnect_enabled = False # 當使用者點擊斷開時，停用自動重新連線
        if self.ws:
            self.ws.close()

    def on_connected(self):
        logger.info("[RJPQ] Connected to server!")
        self.is_connected = True
        self.status_changed.emit(True)
        self.reconnect_timer.stop()
        join_msg = {"type": "join", "code": self.room_code, "password": self.room_pwd}
        if self.ws:
            self.ws.sendTextMessage(json.dumps(join_msg))

    def on_disconnected(self):
        logger.info("[RJPQ] Unexpectedly disconnected from server.")
        self.is_connected = False
        self.status_changed.emit(False)
        
        # 若已啟用則觸發自動重新連線
        if self.reconnect_enabled:
            logger.info("[RJPQ] Reconnecting in 3s...")
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
                # 安靜地忽略 pong 回應
                pass
            else:
                # 安靜地忽略其他未知類型，避免造成終端機負擔或報錯
                pass
        except Exception as e:
            logger.error(f"[RJPQ] Message runtime error: {e}")
            import traceback
            logger.error(traceback.format_exc())

    def on_error(self, error):
        err_str = self.ws.errorString()
        logger.error(f"[RJPQ] WebSocket Internal Error: {err_str} (Code: {error})")
        self.error_received.emit(f"連線錯誤: {err_str}")
        self.status_changed.emit(False)

    def create_room(self, pwd):
        try:
            quoted_pwd = urllib.parse.quote(pwd)
            url = f"https://rjpq.juanwang.cc/api/room?action=create&pwd={quoted_pwd}"
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

# --- RJPQ UI 元件 ---
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
        logging.error(f"[RJPQ 同步錯誤] {error}")
        # 在彈出對話框前立即重置按鈕狀態
        self.create_btn.setText("創建")
        self.create_btn.setEnabled(True)
        self.create_btn.setVisible(True)
        
        # 針對真正的失敗 (包含密碼錯誤) 顯示警告視窗
        if any(kw in error for kw in ["失敗", "連線", "密碼錯誤", "密码错误"]):
            # 如果是密碼錯誤，則停用自動重連以停止無限迴圈
            if "密碼" in error or "密码" in error:
                self.client.reconnect_enabled = False
                if hasattr(self, 'disconnect_timer'):
                    self.disconnect_timer.stop()
            QMessageBox.critical(self, "YZY 伺服器錯誤", f"同步失敗：\n{error}")
        self.update_status(False)

    def init_ui(self):
        self.main_layout = QVBoxLayout(self)
        
        # 1. 房間小工具 (始終顯示)
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

        # 1.1 Overlay 顯示切換控制
        self.overlay_ctrl = QWidget()
        ctrl_layout = QHBoxLayout(self.overlay_ctrl)
        ctrl_layout.setContentsMargins(10, 0, 10, 5)
        
        self.overlay_cb = QCheckBox("在遊戲畫面顯示路徑面板")
        self.overlay_cb.setStyleSheet("color: #00ffff; font-size: 11px; font-weight: bold;")
        self.overlay_cb.toggled.connect(self.client.overlay_toggle_request.emit)
        ctrl_layout.addWidget(self.overlay_cb)
        self.main_layout.addWidget(self.overlay_ctrl)

        # 2. 角色選擇小工具 (連線後顯示)
        self.char_widget = QWidget()
        char_row = QHBoxLayout(self.char_widget)
        self.char_btns = []
        char_colors = ["#ff6b6b", "#51cf66", "#339af0", "#cc5de8"]
        for i in range(4):
            btn = QPushButton(f"10{i+1}")
            btn.setCheckable(True)
            btn.setFixedSize(65, 30)
            btn.setStyleSheet(f"QPushButton {{ background: #222; color: {char_colors[i]}; border: 1px solid #444; border-radius: 4px; font-family: {platform_font_family()}; }} "
                             f"QPushButton:checked {{ background: {char_colors[i]}; color: #fff; font-weight: bold; font-family: {platform_font_family()}; }}")
            btn.clicked.connect(lambda checked, idx=i: self.select_char(idx))
            self.char_btns.append(btn)
            char_row.addWidget(btn)
        self.char_widget.setVisible(False)
        self.main_layout.addWidget(self.char_widget)

        # 3. 網格儀表板 (選擇角色後顯示)
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
            row_label.setStyleSheet(f"color: #444; font-weight: bold; font-family: {platform_font_family()};")
            row_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
            grid_layout.addWidget(row_label, row, 0)
            for col in range(4):
                idx = row * 4 + col
                btn = QPushButton(str(col + 1))
                btn.setFixedSize(58, 28)
                btn.setCursor(Qt.CursorShape.PointingHandCursor)
                # 標準毛玻璃風格樣式
                btn.setStyleSheet("QPushButton { background: rgba(255,255,255,0.05); color: #666; border: 1px solid #222; border-radius: 4px; font-weight: bold; } "
                                 "QPushButton:hover { background: rgba(255,255,255,0.1); border-color: #444; }")
                btn.clicked.connect(lambda checked, i=idx: self.platform_clicked(i))
                self.platform_btns.append(btn)
                grid_layout.addWidget(btn, row, col + 1)
        grid_vbox.addLayout(grid_layout)

        # 網格小工具中的重置按鈕
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
        # 從第 1 排 (底部) 掃描至第 10 排 (頂部)。網格索引 0-3 為第 10 排，36-39 為第 1 排
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
        # 自動填入房間代碼並通知
        self.code_inp.setText(code)
        self.create_btn.setText("創建")
        self.create_btn.setEnabled(True)
        self.create_btn.setVisible(False) # 成功後隱藏
        # 根據使用者要求移除成功彈窗
        # 觸發實際的加入流程
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
        try:
            if not hasattr(self, 'disconnect_timer'):
                self.disconnect_timer = QTimer(self)
                self.disconnect_timer.setSingleShot(True)
                self.disconnect_timer.timeout.connect(self.hide_ui_on_disconnect)
                
            self.create_btn.setVisible(not connected)
            if not connected:
                self.create_btn.setText("創建")
                self.create_btn.setEnabled(True)
                
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
                    # 若原本已有選擇角色，則延遲 500ms 後自動重新選擇
                    QTimer.singleShot(500, lambda: self.select_char(self.selected_color))
            else:
                self.conn_btn.setText("連線")
                self.conn_btn.setStyleSheet("QPushButton { background: #333; color: #fff; font-weight: bold; border-radius: 4px; height: 26px; }")
                self.conn_btn.setEnabled(True)
                
                if not self.client.reconnect_enabled:
                    self.hide_ui_on_disconnect()
                    return

                if not self.disconnect_timer.isActive():
                    logger.info("[RJPQ] 非預期中斷，進入寬限期 (3秒)...")
                    self.disconnect_timer.start(3000)
        except Exception as e:
            logger.error(f"[RJPQ] update_status error: {e}")

    def hide_ui_on_disconnect(self):
        # 僅在目前仍處於斷線狀態時隱藏
        if not self.client.is_connected:
            logging.info("[RJPQ 同步] 寬限期結束，隱藏介面。")
            self.char_widget.setVisible(False)
            self.grid_widget.setVisible(False)
            for btn in self.char_btns: btn.setChecked(False)

    def select_char(self, idx):
        self.selected_color = idx
        for i, btn in enumerate(self.char_btns):
            btn.setChecked(i == idx)
        self.client.send_action({"type": "selectChar", "color": idx})
        self.color_selected.emit(idx)
        
        # 顯示網格儀表板
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
        try:
            if not data or len(data) < 40:
                logger.warning(f"[RJPQ] Malformed grid data: {data}")
                return
                
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
                    # 啟用狀態
                    color = char_colors[val]
                    # 若是其他人標記的，則設定為半透明
                    opacity = "1.0" if (self.selected_color == -1 or val == self.selected_color) else "0.6"
                    btn.setStyleSheet(f"QPushButton {{ background: {color}; color: #fff; border: {border_style}; border-radius: 4px; font-weight: bold; opacity: {opacity}; }}")
                else:
                    # 閒置狀態 (未標記)
                    btn.setStyleSheet(f"QPushButton {{ background: rgba(255,255,255,0.03); color: #444; border: {border_style}; border-radius: 4px; }}")
        except Exception as e:
            logger.error(f"[RJPQ] update_grid error: {e}")

    def on_reset_clicked(self):
        reply = QMessageBox.question(self, "確認重置", "確定要重置所有標記嗎？\n這將清空全隊目前的路徑紀錄。", 
                                     QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if reply == QMessageBox.StandardButton.Yes:
            self.client.send_action({"type": "reset"})

# --- Overlay 影像繪製邏輯 ---
def draw_rjpq_panel(painter, px, py, pw, ph, opacity, data, selected_color):
    painter.setRenderHint(QPainter.RenderHint.Antialiasing)
    path = QPainterPath()
    path.addRoundedRect(px, py, pw, ph, 12, 12)
    painter.setPen(QPen(QColor(0, 255, 255, 200), 2))
    painter.setBrush(QColor(10, 15, 20, int(opacity * 255)))
    painter.drawPath(path)
    
    painter.setPen(QColor(0, 255, 255))
    font = QFont()
    font.setFamilies(platform_font_families())
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
        font.setFamilies(platform_font_families())
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

    # 繪製目標排的高亮框 (金色外框)
    if selected_color != -1:
        # 針對繪製邏輯尋找目標排 (複寫邏輯)
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