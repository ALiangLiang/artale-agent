import time
import re
import logging
from PyQt6.QtCore import QObject, pyqtSignal
from utils import EXP_TABLE
from typing import List, Tuple, Optional, Dict, Any
from data_types import StatsData

logger = logging.getLogger("ExpTracker")

class ExpTracker(QObject):
    """
    負責經驗值與楓幣的統計運算邏輯，與 UI 渲染完全解耦。
    """
    lv_inferred = pyqtSignal(int) # 當推算出新等級時發送
    stats_updated = pyqtSignal(StatsData) # 修改訊號名稱與類型以匹配 Overlay 預期

    def __init__(self):
        super().__init__()
        # 經驗值相關狀態
        self.exp_initial_val: Optional[int] = None 
        self.last_exp_val: int = 0
        self.last_exp_pct: float = 0.0
        self.cumulative_gain: int = 0
        self.cumulative_pct: float = 0.0
        self.exp_history: List[Tuple[float, int, float]] = []  # [(timestamp, value, percent)]
        self.exp_session_start_time: Optional[float] = None
        self.max_10m_exp: int = 0
        self.exp_rate_history: List[int] = [] # 趨勢圖數據
        
        # 楓幣相關狀態
        self.money_initial_val: Optional[int] = None
        self.last_total_money: int = 0
        self.cumulative_money: int = 0
        self.money_history: List[Tuple[float, int]] = [] # [(timestamp, gain)]
        self.money_rate_history: List[int] = []
        
        # 系統狀態
        self.current_lv: Optional[int] = None
        self.exp_session_start_pct: float = 0.0
        self.last_graph_sample_time: float = 0
        self.show_debug: bool = False
        self.is_paused: bool = False
        self.pause_start_time: float = 0
        self.last_known_lv: Optional[int] = None # 用於更精準的等級提升判斷
        self.last_exp_val_time: float = 0 # 記錄最後一次成功的 OCR 時間點

        # 最後輸出的 UI 數據包
        self.stats_data: StatsData = {
            "text": "---",
            "value": 0,
            "percent": 0.0,
            "gained_10m": 0,
            "percent_10m": 0.0,
            "time_to_level": -1,
            "is_estimated": True,
            "tracking_duration": 0,
            "money_10m": 0,
            "cumulative_money": 0,
            "cumulative_gain": 0,
            "cumulative_pct": 0.0,
            "max_10m_exp": 0,
            "exp_rate_history": [],
            "money_rate_history": []
        }

    def parse_exp_text(self, raw_text):
        """
        解析 OCR 原始文字並提取數值。
        支援格式如: "45013389 [98.85%]"
        """
        try:
            match = re.search(r'(\d+)[\s]*(\d+\.?\d*)%', raw_text)
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
        possible_levels = []
        best_lv = None
        min_diff = 999999999
        
        for lv, next_exp in EXP_TABLE.items():
            calc_exp = int(next_exp * (current_pct / 100.0))
            diff = abs(calc_exp - current_exp)
            
            # 使用動態容錯門檻
            tolerance = max(100, next_exp * 0.0001)
            
            if diff < tolerance:
                possible_levels.append(lv)
                if diff < min_diff:
                    min_diff = diff
                    best_lv = lv
        
        # 判斷推斷品質
        if len(possible_levels) > 1:
            # 如果可能等級太多 (例如超過 3 個)，代表數據極度模糊 (通常是 0.00%)
            logger.debug(f"等級推斷存在模糊性: 共有 {len(possible_levels)} 個可能等級 {possible_levels[:5]}... 目前選擇偏差最小的 LV.{best_lv}")
        elif len(possible_levels) == 0:
            logger.debug(f"等級推斷失敗: 經驗值 {current_exp} 與百分比 {current_pct}% 在所有等級中皆不符合一致性檢查。")

        return best_lv

    def update_exp(self, raw_text, timestamp=None):
        """處理經驗值數據更新"""
        now = timestamp or time.time()
        val, pct = self.parse_exp_text(raw_text)
        if val is None: return

        # --- 數據一致性校驗 (Data Consistency Check) ---
        # 每次都推斷一次等級，確保這組 [經驗值 + 百分比] 是合法的
        inf_lv = self.infer_level(val, pct)
        if inf_lv is None:
            if self.show_debug:
                logger.debug(f"跳過不一致數據: 數值 {val:,} 與百分比 {pct}% 無法匹配任何等級。")
            # 廣播舊的經驗值，維持介面更新
            self._broadcast(raw_text, self.last_exp_val, self.last_exp_pct, now)
            return

        # 1. 初始化基準值
        if self.exp_initial_val is None:
            self.exp_initial_val = val
            self.last_exp_val = val
            self.last_exp_pct = pct
            self.exp_history = [(now, val, pct)]
            self.exp_session_start_time = None
            self.exp_session_start_pct = pct
            self.current_lv = inf_lv
            self.lv_inferred.emit(self.current_lv)
            self.last_exp_val_time = now
                
            logger.info(f"建立初始基準值: {val:,} ({pct}%) -> 識別等級: LV.{inf_lv}")
            self._broadcast(raw_text, val, pct, now)
            return

        # 2. 偵測等級變動 (由推斷結果決定)
        level_up_triggered = False
        if inf_lv != self.current_lv:
            # 只有等級上升1等才觸發升級邏輯 (避免 OCR 偶爾跳回低等級造成的誤判)
            if inf_lv == (self.current_lv or 0) + 1:
                logger.info(f"偵測到等級提升 (自動推斷): {self.current_lv} -> {inf_lv}")
                level_up_triggered = True
                self.current_lv = inf_lv
                self.lv_inferred.emit(self.current_lv)
            else:
                # 如果推斷等級變小了，可能是極其嚴重的誤判，暫時忽略這筆
                if self.show_debug:
                    logger.debug(f"忽略異常等級跳變: {self.current_lv} -> {inf_lv}")
                # 維持介面更新
                self._broadcast(raw_text, self.last_exp_val, self.last_exp_pct, now)
                return

        if level_up_triggered:
            logger.info("重置統計基準 (升級)。")
            # 升級時重置累積值
            self.exp_initial_val = val
            self.last_exp_val = val
            self.last_exp_pct = pct
            self.exp_session_start_time = None
            self.cumulative_gain = 0
            self.exp_history = [(now, val, pct)]
            self.money_history = []
            self.money_initial_val = None 
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
        self.last_exp_val_time = now
        
        self._broadcast(raw_text, val, pct, now)

    def update_tick(self):
        """僅更新時間與效率廣播 (用於辨識失敗時維持 UI 時鐘運作)"""
        # 只要計時已經開始，就持續廣播當前狀態以更新 Duration
        if self.exp_session_start_time:
            self._broadcast(None, self.last_exp_val, self.last_exp_pct, time.time())

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
            if dt > 3:
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
             
        self.stats_data = StatsData(
            text=raw_text,
            value=val,
            percent=pct,
            gained_10m=max(0, gain_10m),
            percent_10m=max(0.0, pct_10m),
            time_to_level=time_to_lv,
            is_estimated=is_est,
            tracking_duration=int(duration),
            money_10m=max(0, m_gain_10m),
            cumulative_money=self.cumulative_money,
            cumulative_gain=self.cumulative_gain,
            cumulative_pct=pct - self.exp_session_start_pct,
            max_10m_exp=self.max_10m_exp,
            exp_rate_history=self.exp_rate_history,
            money_rate_history=self.money_rate_history
        )

        self.stats_updated.emit(self.stats_data)

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
