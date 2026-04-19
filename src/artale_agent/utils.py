import os
import sys
import json
import logging

logger = logging.getLogger(__name__)

CONFIG_FILE = "config.json"

def _project_root():
    """Get the project root directory (one level up from src/)"""
    return os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

def get_version():
    try:
        base_dir = getattr(sys, '_MEIPASS', _project_root())
        v_path = os.path.join(base_dir, "VERSION")
        with open(v_path, 'r', encoding='utf-8') as f:
            return f.read().strip()
    except:
        return "v?.?.?"

VERSION = get_version()
REPO_URL = "ALiangLiang/artale-agent"

def resource_path(relative_path):
    """ 獲取資源的絕對路徑，支援開發環境與 PyInstaller 打包環境 """
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = _project_root()
    if sys.platform == "win32":
        relative_path = relative_path.replace("/", "\\")

    return os.path.join(base_path, "assets", relative_path)

def platform_font_family():
    """Return the preferred font family CSS string for the current platform."""
    if sys.platform == "darwin":
        return "'PingFang TC', 'Heiti TC', sans-serif"
    return "'Microsoft JhengHei', '\u5fae\u8edf\u6b63\u9ed1\u9ad4', sans-serif"

def platform_font_families():
    """Return font families list for QFont.setFamilies()."""
    if sys.platform == "darwin":
        return ["PingFang TC", "Heiti TC"]
    return ["Microsoft JhengHei", "\u5fae\u8edf\u6b63\u9ed1\u9ad4"]

class ConfigManager:
    @staticmethod
    def load_config():
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r", encoding='utf-8') as f:
                    config = json.load(f)
                    
                    # 版本遷移：舊版單一配置 -> 新版多重配置
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
                    
                    new_profiles = {}
                    old_profiles = config.get("profiles", {})
                    for i in range(1, 10): # 目前支援 F1 到 F9
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
                    
                    if config.get("active_profile", "").startswith("Profile "):
                        num = config["active_profile"].split(" ")[1]
                        config["active_profile"] = f"F{num}"
                    
                    if config.get("active_profile") not in config["profiles"]:
                        config["active_profile"] = "F1"

                    if "offset" not in config: config["offset"] = [0, 0]
                    if "exp_offset" not in config: config["exp_offset"] = [0, 0]

                    for p in config["profiles"].values():
                        if "name" not in p: p["name"] = "未命名"
                        for k, v in p["triggers"].items():
                            if isinstance(v, (int, float)):
                                p["triggers"][k] = {"seconds": int(v), "icon": "", "sound": True}
                            if "sound" not in p["triggers"][k]:
                                p["triggers"][k]["sound"] = True
                    
                    if "opacity" not in config: config["opacity"] = 0.5
                    
                    default_hks = {
                        "exp_toggle": "f10", "exp_pause": "f11", "reset": "f9",
                        "exp_report": "f12", "rjpq_1": "num_1", "rjpq_2": "num_2",
                        "rjpq_3": "num_3", "rjpq_4": "num_4", "show_settings": "pause"
                    }
                    if "hotkeys" not in config:
                        config["hotkeys"] = default_hks
                    else:
                        for k, v in default_hks.items():
                            if k not in config["hotkeys"]: config["hotkeys"][k] = v

                    return config
            except Exception as e: 
                logger.error(f"Error loading config: {e}")
                pass
            
        default_profiles = {}
        for i in range(1, 10):
            default_profiles[f"F{i}"] = {"name": f"切換組 F{i}", "triggers": {}}
        return {
            "active_profile": "F1", "offset": [0, 0], "opacity": 0.5, 
            "profiles": default_profiles,
            "hotkeys": {
                "exp_toggle": "f10", "exp_pause": "f11", "reset": "f12",
                "rjpq_1": "1", "rjpq_2": "2", "rjpq_3": "3", "rjpq_4": "4"
            }
        }

    @staticmethod
    def save_config(config):
        with open(CONFIG_FILE, "w", encoding='utf-8') as f:
            json.dump(config, f, ensure_ascii=False, indent=2)

# Artale 經驗值表 (Level 1 - 200)
# 格式: { 當前等級: 升至下一級所需的經驗值 }
EXP_TABLE = {
    1: 15, 2: 34, 3: 57, 4: 92, 5: 135, 6: 372, 7: 560, 8: 840, 9: 1242, 10: 1716,
    11: 2360, 12: 3216, 13: 4200, 14: 5460, 15: 7050, 16: 8840, 17: 11040, 18: 13716, 19: 16680, 20: 20216,
    21: 24402, 22: 28980, 23: 34320, 24: 40512, 25: 47216, 26: 54900, 27: 63666, 28: 73080, 29: 83720, 30: 95700,
    31: 108480, 32: 122760, 33: 138666, 34: 155540, 35: 174216, 36: 194832, 37: 216600, 38: 240500, 39: 266682, 40: 294216,
    41: 324240, 42: 356916, 43: 391160, 44: 428280, 45: 468450, 46: 510420, 47: 555680, 48: 604416, 49: 655200, 50: 709716,
    51: 748608, 52: 789631, 53: 832902, 54: 878545, 55: 926689, 56: 977471, 57: 1031036, 58: 1087536, 59: 1147132, 60: 1209994,
    61: 1276301, 62: 1346242, 63: 1420016, 64: 1497832, 65: 1579913, 66: 1666492, 67: 1757815, 68: 1854143, 69: 1955750, 70: 2062925,
    71: 2175973, 72: 2295216, 73: 2420993, 74: 2553663, 75: 2693603, 76: 2841212, 77: 2996910, 78: 3161140, 79: 3334370, 80: 3517093,
    81: 3709829, 82: 3913127, 83: 4127566, 84: 4353756, 85: 4592341, 86: 4844001, 87: 5109452, 88: 5389449, 89: 5684790, 90: 5996316,
    91: 6324914, 92: 6671519, 93: 7037118, 94: 7422752, 95: 7829518, 96: 8258575, 97: 8711144, 98: 9188514, 99: 9692044, 100: 10223168,
    101: 10783397, 102: 11374327, 103: 11997640, 104: 12655110, 105: 13348610, 106: 14080113, 107: 14851703, 108: 15665576, 109: 16524049, 110: 17429566,
    111: 18384706, 112: 19392187, 113: 20454878, 114: 21575805, 115: 22758159, 116: 24005306, 117: 25320796, 118: 26708375, 119: 28171993, 120: 29715818,
    121: 31344244, 122: 33061908, 123: 34873700, 124: 36784778, 125: 38800583, 126: 40926854, 127: 43169645, 128: 45535341, 129: 48030677, 130: 50662758,
    131: 53439077, 132: 56367538, 133: 59456479, 134: 62714694, 135: 66151459, 136: 69776558, 137: 73600313, 138: 77633610, 139: 81887931, 140: 86375389,
    141: 91108760, 142: 96101520, 143: 101367883, 144: 106922842, 145: 112782213, 146: 118962678, 147: 125481832, 148: 132358236, 149: 139611467, 150: 147262175,
    151: 155332142, 152: 163844343, 153: 172823012, 154: 182293713, 155: 192283408, 156: 202820538, 157: 213935103, 158: 225658746, 159: 238024845, 160: 251068606,
    161: 264827165, 162: 279339693, 163: 294647508, 164: 310794191, 165: 327825712, 166: 345790561, 167: 364739883, 168: 384727628, 169: 405810702, 170: 428049128,
    171: 451506220, 172: 476248760, 173: 502347192, 174: 529875818, 175: 558913012, 176: 589541445, 177: 621848316, 178: 655925603, 179: 691870326, 180: 729784819,
    181: 769777027, 182: 811960808, 183: 856456260, 184: 903390063, 185: 952895838, 186: 1005114529, 187: 1060194805, 188: 1118293480, 189: 1179575962, 190: 1244216724,
    191: 1312399800, 192: 1384319309, 193: 1460180007, 194: 1540197871, 195: 1624600714, 196: 1713628833, 197: 1807535693, 198: 1906588648, 199: 2011069705, 200: 2121276324
}
