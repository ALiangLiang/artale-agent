import time
import re
import logging
from PyQt6.QtCore import QObject, pyqtSignal
from utils import EXP_TABLE

logger = logging.getLogger("ExpTracker")

class ExpTracker(QObject):
    """
    負責經驗值與楓幣的統計運算邏輯，與 UI 渲染完全解耦。
    """
    lv_inferred = pyqtSignal(str) # 當推算出新等級時發送
    updated = pyqtSignal(dict)

    def __init__(self):
        super().__init__()
        # 經驗值相關狀態
        self.exp_initial_val = None 
        self.last_exp_val = 0
        self.last_exp_pct = 0.0
        self.cumulative_gain = 0
        self.cumulative_pct = 0.0
        self.exp_history = []  # [(timestamp, value, percent)]
        self.exp_session_start_time = None
        self.max_10m_exp = 0
        self.exp_rate_history = [] # 趨勢圖數據
        
        # 楓幣相關狀態
        self.money_initial_val = None
        self.last_total_money = 0
        self.cumulative_money = 0
        self.money_history = [] # [(timestamp, gain)]
        self.money_rate_history = []
        
        # 系統狀態
        self.current_lv = None
        self.exp_session_start_pct = 0.0
        self.last_graph_sample_time = 0
        self.show_debug = False
        self.is_paused = False
        self.pause_start_time = 0

        # 最後輸出的 UI 數據包
        self.stats_data = {
            "text": "---",
            "value": 0,
            "percent": 0.0,
            "gained_10m": 0,
            "percent_10m": 0.0,
            "time_to_level": -1,
            "is_estimated": True,
            "tracking_duration": 0,
            "money_10m": 0,
            "cumulative_money": 0
        }

    def parse_exp_text(self, raw_text):
        """
        解析 OCR 原始文字並提取數值。
        支援格式如: "45013389 [98.85%]"
        """
        try:
            cleaned = raw_text.replace(' ', '')
            match = re.search(r'(\d+)\[(\d+\.?\d*)%\]', cleaned)
            if match:
                val = int(match.group(1))
                pct = float(match.group(2))
                return val, pct
        except Exception as e:
            if self.show_debug:
                logger.debug(f"解析錯誤: {e} | 原始文字: {raw_text}")
        return None, None

    def infer_level(self, current_exp, current_pct):
        """
        透過經驗值與百分比推算目前等級。
        """
        best_lv = None
        min_diff = 999999999
        
        for lv, next_exp in EXP_TABLE.items():
            calc_exp = int(next_exp * (current_pct / 100.0))
            diff = abs(calc_exp - current_exp)
            if diff < min_diff:
                min_diff = diff
                best_lv = lv
                
        if min_diff < 10000: # 容錯門檻
            return best_lv
        return None

    def update_exp(self, raw_text, timestamp=None):
        """處理經驗值數據更新"""
        now = timestamp or time.time()
        val, pct = self.parse_exp_text(raw_text)
        if val is None: return

        # 1. 初始化基準值
        if self.exp_initial_val is None:
            self.exp_initial_val = val
            self.last_exp_val = val
            self.last_exp_pct = pct
            self.exp_history = [(now, val, pct)]
            self.exp_session_start_time = None
            self.exp_session_start_pct = pct
            
            # 嘗試推測等級
            inf_lv = self.infer_level(val, pct)
            if inf_lv:
                self.current_lv = f"LV.{inf_lv}"
                self.lv_inferred.emit(self.current_lv)
                
            logger.info(f"建立初始基準值: {val:,} ({pct}%)")
            self._broadcast(raw_text, val, pct, now)
            return

        # 2. 偵測等級提升
        if val < self.last_exp_val - 1000:
            logger.info("偵測到等級變更，重置統計基準。")
            self.exp_initial_val = val
            self.last_exp_val = val
            self.last_exp_pct = pct
            self.cumulative_gain = 0
            self.cumulative_pct = 0.0
            self.exp_session_start_time = None
            self.exp_history = [(now, val, pct)]
            self._broadcast(raw_text, val, pct, now)
            return

        # 3. 計算增量與計時啟動
        v_diff = val - self.last_exp_val
        if v_diff > 0:
            if self.exp_session_start_time is None:
                self.exp_session_start_time = now
                logger.info("偵測到經驗值增加，正式啟動計時。")
            self.cumulative_gain += v_diff
            
        # 4. 更新歷史紀錄 (Sliding Window: 1小時)
        self.exp_history.append((now, val, pct))
        self.exp_history = [h for h in self.exp_history if h[0] >= now - 3600]
        self.last_exp_val = val
        self.last_exp_pct = pct

        self._broadcast(raw_text, val, pct, now)

    def update_money(self, total_val, timestamp=None):
        """處理楓幣數據更新"""
        now = timestamp or time.time()
        
        if self.money_initial_val is None:
            self.money_initial_val = total_val
            self.last_total_money = total_val
            logger.info(f"建立楓幣基準值: {total_val:,}")
            return

        gain = total_val - self.last_total_money
        if gain != 0:
            self.cumulative_money += gain
            self.money_history.append((now, gain))
        
        self.last_total_money = total_val
        self.money_history = [h for h in self.money_history if h[0] >= now - 3600]
        
        # 楓幣不需要頻繁廣播，隨經驗值更新一起發送即可

    def _broadcast(self, raw_text, val, pct, now):
        """計算統計結果並發送給 UI"""
        # A. 效率計算 (10分鐘滑動視窗)
        h_ago_10m = now - 600
        recent = [h for h in self.exp_history if h[0] >= h_ago_10m]
        
        gain_10m = 0; pct_10m = 0.0; time_to_lv = -1; is_est = True
        if len(recent) >= 2:
            dt = recent[-1][0] - recent[0][0]
            if dt > 5:
                dv = recent[-1][1] - recent[0][1]
                dp = recent[-1][2] - recent[0][2]
                gain_10m = int(dv * 600 / dt)
                pct_10m = dp * 600 / dt
                is_est = (dt < 580)
                
                # 預計升級時間
                rem_p = 100.0 - pct
                p_per_sec = dp / dt
                if p_per_sec > 0:
                    time_to_lv = int(rem_p / p_per_sec)

        # B. 楓幣效率計算
        m_history_10m = [h for h in self.money_history if h[0] >= h_ago_10m]
        m_gain_10m = 0
        if len(m_history_10m) >= 1:
            m_sum = sum(h[1] for h in m_history_10m)
            m_dt = now - m_history_10m[0][0]
            if m_dt > 10:
                m_gain_10m = int(m_sum * 600 / m_dt)

        # C. 更新歷史最高與圖表數據
        if not is_est and gain_10m > self.max_10m_exp:
            self.max_10m_exp = gain_10m

        if now - self.last_graph_sample_time >= 10:
            self.exp_rate_history.append(max(0, gain_10m))
            self.money_rate_history.append(max(0, m_gain_10m))
            if len(self.exp_rate_history) > 60: self.exp_rate_history.pop(0)
            if len(self.money_rate_history) > 60: self.money_rate_history.pop(0)
            self.last_graph_sample_time = now

        # D. 組裝數據包
        duration = now - self.exp_session_start_time if self.exp_session_start_time else 0
        
        # 修正累積百分比邏輯 (如果是開程式後的總累積)
        if self.exp_session_start_time is None:
             self.exp_session_start_pct = pct
             
        self.stats_data = {
            "text": raw_text,
            "value": val,
            "percent": pct,
            "gained_10m": max(0, gain_10m),
            "percent_10m": max(0.0, pct_10m),
            "time_to_level": time_to_lv,
            "is_estimated": is_est,
            "tracking_duration": int(duration),
            "money_10m": max(0, m_gain_10m),
            "cumulative_money": self.cumulative_money,
            "cumulative_gain": self.cumulative_gain,
            "cumulative_pct": pct - self.exp_session_start_pct,
            "max_10m_exp": self.max_10m_exp,
            "exp_rate_history": self.exp_rate_history,
            "money_rate_history": self.money_rate_history
        }

        self.updated.emit(self.stats_data)

    def toggle_pause(self):
        """切換暫停狀態並補償時間戳"""
        now = time.time()
        if not self.is_paused:
            self.is_paused = True
            self.pause_start_time = now
            logger.info("統計已暫停")
        else:
            self.is_paused = False
            shift = now - self.pause_start_time
            # 補償計時起點
            if self.exp_session_start_time:
                self.exp_session_start_time += shift
            # 補償所有歷史紀錄時間戳，避免效率被稀釋
            self.exp_history = [(t + shift, v, p) for t, v, p in self.exp_history]
            self.money_history = [(t + shift, g) for t, g in self.money_history]
            logger.info(f"統計已恢復，補償時長: {shift:.1f}秒")
        self.updated.emit(self.stats_data)

    def reset_baseline(self):
        """手動重置統計"""
        self.exp_initial_val = None
        self.money_initial_val = None
        self.cumulative_gain = 0
        self.cumulative_money = 0
        self.exp_history = []
        self.money_history = []
        self.exp_session_start_time = None
        self.updated.emit(self.stats_data)

    def set_lv(self, data):
        """外部同步等級訊息"""
        self.current_lv = data.get("level")
