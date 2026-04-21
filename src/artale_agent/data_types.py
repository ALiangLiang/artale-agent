from dataclasses import dataclass
from typing import Optional, List
import numpy as np

@dataclass
class UpdateData:
    text: str
    conf: float

@dataclass
class LVUpdateData:
    level: str
    conf: float

@dataclass
class MoneyUpdateData(UpdateData):
    pass

@dataclass
class ExpUpdateData(UpdateData):
    pass

@dataclass
class ExpVisualData:
    exp: Optional[np.ndarray] = None
    lv: Optional[np.ndarray] = None
    coin: Optional[np.ndarray] = None
    conf: float = 0.0

@dataclass
class StatsData:
    text: Optional[str]
    value: int
    percent: float
    gained_10m: int
    percent_10m: float
    time_to_level: int
    is_estimated: bool
    tracking_duration: int
    money_10m: int
    cumulative_money: int
    cumulative_exp_gain: int
    cumulative_exp_pct: float
    max_10m_exp: int
    exp_rate_history: List[int]
    money_rate_history: List[int]
    is_paused: bool = False
