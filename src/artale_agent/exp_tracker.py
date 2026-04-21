import time
import re
import logging
from PyQt6.QtCore import QObject, pyqtSignal
from artale_agent.utils import EXP_TABLE
from typing import List, Tuple, Optional, Dict, Any
from artale_agent.data_types import StatsData

logger = logging.getLogger(__name__)

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
        self.cumulative_exp_gain: int = 0
        self.cumulative_exp_pct: float = 0.0
        
        # 經驗值歷史紀錄：用於計算滑動視窗效率 (10分鐘)
        # 格式: [(時間戳記, 當前等級經驗值, 當前等級百分比, 總累積獲得經驗值)]
        # 註：儲存「總累積獲得經驗值」是為了在跨等級(Level Up)時仍能精準計算增量。
        self.exp_history: List[Tuple[float, int, float, int]] = []
        self.exp_session_start_time: Optional[float] = None
        self.max_10m_exp: int = 0
        self.exp_rate_history: List[int] = [] # 趨勢圖數據
        
        # 楓幣相關狀態
        self.money_initial_val: Optional[int] = None
        self.last_money_val: int = 0
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
        self.needs_baseline_correction: bool = False # 恢復後是否需要修正基準值
        self.last_known_lv: Optional[int] = None # 用於更精準的等級提升判斷
        self.last_exp_val_time: float = 0 # 記錄最後一次成功的 OCR 時間點
        
        # 等級 OCR 暫存 (用於輔助推論)
        self.last_lv_ocr_val: Optional[int] = None
        self.last_lv_ocr_conf: float = 0.0

        # 最後輸出的 UI 數據包
        self.stats_data = StatsData(
            text="---",
            value=0,
            percent=0.0,
            gained_10m=0,
            percent_10m=0.0,
            time_to_level=-1,
            is_estimated=True,
            tracking_duration=0,
            money_10m=0,
            cumulative_money=0,
            cumulative_exp_gain=0,
            cumulative_exp_pct=0.0,
            max_10m_exp=0,
            exp_rate_history=[],
            money_rate_history=[],
            is_paused=False
        )

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
                logger.debug("解析錯誤: %s | 原始文字: %s", e, raw_text)
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
        
        # 判斷推斷品質：必須唯一才能採納
        if len(possible_levels) == 1:
            return possible_levels[0]
        
        if len(possible_levels) > 1:
            logger.debug("等級推斷存在模糊性: 共有 %s 個可能等級 %s，視為推論失敗。", len(possible_levels), possible_levels[:5])
        elif len(possible_levels) == 0:
            logger.debug("等級推斷失敗: 經驗值 %s 與百分比 %s%% 無法匹配任何等級。", current_exp, current_pct)

        return None

    def update_lv_ocr(self, level: int, conf: float):
        """更新最後一次看到的等級 OCR 結果"""
        self.last_lv_ocr_val = level
        self.last_lv_ocr_conf = conf

    def validate_exp(self, raw_text, conf):
        """
        驗證經驗值 OCR 結果是否可採納。
        回傳: (val, pct, inf_lv) 或 (None, None, None)
        """
        # 1. 信心度過濾 (僅採納 0 或 >= 90)
        if 0 < conf < 90:
            logger.debug("[OCR] 經驗值信心度不足: %s", conf)
            return None, None, None

        # 2. 解析文字
        val, pct = self.parse_exp_text(raw_text)
        if val is None:
            logger.debug("[OCR] 經驗值解析失敗: %s", raw_text)
            return None, None, None

        # 3. 數據一致性校驗
        inf_lv = self.infer_level(val, pct)
        
        # 若推論無結果 (通常是 0.00%)，但等級 OCR 信心度高，則採納等級 OCR 結果
        if inf_lv is None and self.last_lv_ocr_conf >= 90:
            inf_lv = self.last_lv_ocr_val
            logger.debug("[OCR] 經驗值推論無效，使用等級 OCR 結果: %s", inf_lv)
            
        if inf_lv is None:
            logger.debug("[OCR] 經驗值與百分比無法匹配任何等級: %s", raw_text)
            return None, None, None

        # 4. 等級變動合理性檢查 (僅在正式計時開始後執行)
        if self.exp_session_start_time is not None:
            # 允許等級不變，或剛好 +1
            if inf_lv != self.current_lv and inf_lv != (self.current_lv or 0) + 1:
                logger.debug("[OCR] 等級變動合理性檢查失敗: %s -> %s", self.current_lv, inf_lv)
                return None, None, None

        return val, pct, inf_lv

    def update_exp(self, raw_text, conf=100, timestamp=None):
        """處理經驗值數據更新"""
        now = timestamp or time.time()

        # 執行獨立校驗函式
        val, pct, inf_lv = self.validate_exp(raw_text, conf)

        # 處理不採納的情況：維持舊數據，但時間與統計照常計算
        if val is None:
            self._broadcast(raw_text, self.last_exp_val, self.last_exp_pct, now)
            return

        # 1. 初始化基準值
        if self.exp_initial_val is None:
            self.exp_initial_val = val
            self.last_exp_val = val
            self.last_exp_pct = pct
            self.exp_history = [(now, val, pct, 0)] # (time, val, pct, cumulative_exp_gain)
            self.exp_session_start_time = None
            self.exp_session_start_pct = pct
            self.current_lv = inf_lv
            self.lv_inferred.emit(self.current_lv)
            self.last_exp_val_time = now
                
            logger.info("建立初始基準值: %s (%s%%) -> 識別等級: LV.%s", val, pct, inf_lv)
            self._broadcast(raw_text, val, pct, now)
            return

        # 1.5 暫停恢復後的基準值修正
        if self.needs_baseline_correction:
            logger.info("暫停恢復：更新基準值為 %s，忽略期間增量。", val)
            self.last_exp_val = val
            self.last_exp_pct = pct
            self.current_lv = inf_lv
            self.needs_baseline_correction = False
            # 這裡不 return，讓它繼續執行後面的歷史紀錄與廣播

        # 2. 偵測等級變動
        level_up_triggered = False
        if inf_lv != self.current_lv:
            # A. 正常升級
            if inf_lv == (self.current_lv or 0) + 1:
                logger.info("偵測到等級提升 (自動推斷): %s -> %s", self.current_lv, inf_lv)
                level_up_triggered = True
                self.current_lv = inf_lv
                self.lv_inferred.emit(self.current_lv)
            # B. 初始修正
            elif self.exp_session_start_time is None:
                logger.info("修正初始等級辨識: %s -> %s", self.current_lv, inf_lv)
                self.current_lv = inf_lv
                self.lv_inferred.emit(self.current_lv)

        if level_up_triggered:
            # 獲取前一等級所需的總經驗值以計算跨級增量
            prev_lv = inf_lv - 1
            max_exp_prev = EXP_TABLE.get(prev_lv, self.last_exp_val)
            
            # 計算增量：(前一級剩餘) + (這一級現有)
            v_diff = (max_exp_prev - self.last_exp_val) + val
            if v_diff > 0:
                if self.exp_session_start_time is None:
                    self.exp_session_start_time = now
                self.cumulative_exp_gain += v_diff
                # 基於使用者要求「baseline 改成 0」，調整百分比起點以持續累加
                self.exp_session_start_pct -= 100.0
                
            self.exp_history.append((now, val, pct, self.cumulative_exp_gain))
            self.last_exp_val = val
            self.last_exp_pct = pct
            self.last_exp_val_time = now
            
            logger.info("升級持續累計: 跨級增益 %s exp", v_diff)
            self._broadcast(raw_text, val, pct, now)
            return

        # 3. 計算增量與計時啟動
        v_diff = val - self.last_exp_val
        if v_diff > 0:
            if self.exp_session_start_time is None:
                self.exp_session_start_time = now
                logger.info("偵測到經驗值增加，正式啟動計時。")
            self.cumulative_exp_gain += v_diff
            
        # 4. 更新歷史紀錄 (Sliding Window: 1小時)
        self.exp_history.append((now, val, pct, self.cumulative_exp_gain))
        self.exp_history = [h for h in self.exp_history if h[0] >= now - 3600]
        self.last_exp_val = val
        self.last_exp_pct = pct
        self.last_exp_val_time = now
        
        logger.debug("經驗值持續累計: %s exp", v_diff)
        self._broadcast(raw_text, val, pct, now)

    def update_tick(self, timestamp=None):
        """僅更新時間與效率廣播 (用於辨識失敗時維持 UI 時鐘運作)"""
        now = timestamp or time.time()
        # 只要計時已經開始，就持續廣播當前狀態以更新 Duration
        if self.exp_session_start_time:
            self._broadcast(None, self.last_exp_val, self.last_exp_pct, now)

    def update_money(self, text, conf=100, timestamp=None):
        """處理楓幣數據更新"""
        now = timestamp or time.time()

        # 信心度過濾 (僅採納 0 或 >= 90)
        if 0 < conf < 90:
            logger.debug("[OCR] 楓幣信心度不足: %s", conf)
            return
        
        if self.money_initial_val is None:
            self.money_initial_val = int(text)
            self.last_money_val = int(text)
            logger.info("建立楓幣基準值: %s", text)
            return

        gain = int(text) - self.last_money_val
        if gain != 0:
            self.cumulative_money += gain
            self.last_money_val = int(text) # 更新基準值
            
        # 紀錄歷史總量 (Sliding Window: 1小時)
        self.money_history.append((now, int(text)))
        self.money_history = [h for h in self.money_history if h[0] >= now - 3600]
        
        logger.debug("楓幣持續累計: %s 楓幣", gain)

    def _broadcast(self, raw_text, val, pct, now):
        """計算統計結果並發送給 UI"""
        # A. 效率計算 (10分鐘滑動視窗)
        h_ago_10m = now - 600
        recent = [h for h in self.exp_history if h[0] >= h_ago_10m]
        
        gain_10m = 0; pct_10m = 0.0; time_to_lv = -1; is_est = True
        if len(recent) >= 2:
            dt = recent[-1][0] - recent[0][0]
            if dt > 3:
                # 使用第 4 個元素 (cumulative_exp_gain) 計算增量
                dv = recent[-1][3] - recent[0][3]
                
                # 計算百分比增量 (透過總獲得量推算，以支援跨級)
                # 如果有推斷等級，用當前等級的總量來估算百分比增益
                total_for_cur_lv = EXP_TABLE.get(self.current_lv, 1)
                dp = (dv / total_for_cur_lv) * 100.0
                
                gain_10m = int(dv * 600 / dt)
                pct_10m = dp * 600 / dt
                is_est = (dt < 580)
                
                # 預計升級時間
                rem_p = 100.0 - pct
                p_per_sec = dp / dt
                if p_per_sec > 0:
                    time_to_lv = int(rem_p / p_per_sec)

        # B. 楓幣效率計算 (參考經驗值邏輯：取基準點差值)
        m_history_10m = [h for h in self.money_history if h[0] >= h_ago_10m]
        m_gain_10m = 0
        if len(m_history_10m) >= 2:
            m_h_base = m_history_10m[0]
            m_dt = now - m_h_base[0]
            if m_dt > 10:
                # 使用總量差值計算
                m_dv = self.last_money_val - m_h_base[1]
                m_gain_10m = int(m_dv * 600 / m_dt)

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
            cumulative_exp_gain=self.cumulative_exp_gain,
            cumulative_exp_pct=pct - self.exp_session_start_pct,
            max_10m_exp=self.max_10m_exp,
            exp_rate_history=self.exp_rate_history,
            money_rate_history=self.money_rate_history,
            is_paused=self.is_paused
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
            # 補償所有歷史紀錄時間戳，避免效率被稀釋 (注意: exp_history 現在是 4 元組)
            self.exp_history = [(t + shift, v, p, c) for t, v, p, c in self.exp_history]
            self.money_history = [(t + shift, g) for t, g in self.money_history]
            self.needs_baseline_correction = True # 標記下次更新需修正基準
            logger.info("統計已恢復，補償時長: %.1f秒", shift)
            
        # 立即廣播最新狀態 (包含 is_paused 旗標)
        self._broadcast(None, self.last_exp_val, self.last_exp_pct, now)

    def reset_baseline(self):
        """手動重置統計"""
        self.exp_initial_val = None
        self.money_initial_val = None
        self.last_exp_val = None
        self.last_exp_pct = 0.0
        self.current_lv = None
        self.cumulative_exp_gain = 0
        self.cumulative_money = 0
        self.exp_history = []
        self.money_history = []
        self.exp_rate_history = []
        self.money_rate_history = []
        self.exp_session_start_time = None
        self.max_10m_exp = 0
        
        # 重置最後輸出的數據包
        self.stats_data = StatsData(
            text="---", value=0, percent=0.0, gained_10m=0, percent_10m=0.0,
            time_to_level=-1, is_estimated=True, tracking_duration=0,
            money_10m=0, cumulative_money=0, cumulative_exp_gain=0,
            cumulative_exp_pct=0.0, max_10m_exp=0,
            exp_rate_history=[], money_rate_history=[]
        )
        self.stats_updated.emit(self.stats_data)
