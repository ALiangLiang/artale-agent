import os
import sys
import json
import logging

logger = logging.getLogger(__name__)

CONFIG_FILE = "config.json"

def get_version():
    try:
        base_dir = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
        v_path = os.path.join(base_dir, "VERSION")
        with open(v_path, 'r', encoding='utf-8') as f:
            return f.read().strip()
    except:
        return "v?.?.?"

VERSION = get_version()
REPO_URL = "ALiangLiang/artale-agent"

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

class ConfigManager:
    @staticmethod
    def load_config():
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r", encoding='utf-8') as f:
                    config = json.load(f)
                    
                    # Migration: Old single profile -> Multi profile
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
                    for i in range(1, 10): # Supporting up to F9 now
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
