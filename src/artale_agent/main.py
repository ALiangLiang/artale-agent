import logging
import os
import platform as platform_mod
import sys
import time

import sentry_sdk
from pynput import keyboard, mouse
from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import QApplication

# 強制匯入以確保 PyInstaller 的可見性與執行時的執行緒安全 (sip.isdeleted)
try:
    from PyQt6 import QtWebSockets, QtNetwork
    import sip
except ImportError:
    pass
from sentry_sdk.integrations.logging import LoggingIntegration

# Local imports
from .overlay import ArtaleOverlay
from .settings_window import SettingsWindow
from .utils import ConfigManager, get_version
from .platform import FocusTrackerImpl

# 初始化日誌記錄器
logger = logging.getLogger(__name__)


def check_dynamic_console():
    """Enable console window if --debug or --console argument is present"""
    if sys.platform != "win32":
        return
    if "--debug" in sys.argv or "--console" in sys.argv:
        import ctypes
        import ctypes.wintypes

        try:
            # 附加到父進程控制台或分配一個新的控制台
            if not ctypes.windll.kernel32.AttachConsole(-1):
                ctypes.windll.kernel32.AllocConsole()
            # 將標準 I/O 重新映射到新控制台
            sys.stdout = open("CONOUT$", "w", encoding="utf-8")
            sys.stderr = open("CONOUT$", "w", encoding="utf-8")

            # 確保控制台字元集支援特殊字元
            os.system("chcp 65001 > nul")

            print("\n" + "=" * 50)
            print("Artale Swiss Knife - Debug Console Enabled")
            print("=" * 50 + "\n")
        except Exception as e:
            logger.error("[Main] Failed to allocate console: %s", e)


check_dynamic_console()

# Platform-specific tracker implementation is now in .platform module
def start_keyboard_listener(overlay, settings_window, focus_tracker):
    # --- 用於取消右鍵點擊的滑鼠監聽器 ---
    def on_click(x, y, button, pressed):
        if not pressed:
            return
        if not focus_tracker.is_game_active:
            return

        if button == mouse.Button.right:
            overlay.check_right_click(x, y)
        elif button == mouse.Button.left:
            overlay.check_left_click(x, y)

    mouse_l = mouse.Listener(on_click=on_click)
    mouse_l.start()

    # --- 鍵盤監聽器 ---
    current_config = ConfigManager.load_config()

    # 連按偵測狀態 (用於 F1-F9 配置切換)
    last_key = None
    last_time = 0
    double_press_delay = 0.35  # 秒
    is_globally_enabled = True

    def update_local_config():
        nonlocal current_config, is_globally_enabled
        current_config = ConfigManager.load_config()
        active = current_config.get("active_profile", "F1")
        p_data = current_config["profiles"].get(active, {"triggers": {}})
        triggers = p_data.get("triggers", {})
        is_globally_enabled = True  # 切換配置時重新啟用
        logger.info(
            "[Config] Switched to %s. Triggers: %s", active, list(triggers.keys())
        )

    # 將設定視窗的訊號與監聽器更新連結
    settings_window.config_updated.connect(update_local_config)

    def on_press(key):
        nonlocal last_key, last_time, current_config, is_globally_enabled
        try:
            k_name = None
            # 1. 優先處理數字小鍵盤 (Numpad) 的 VK 碼 (96-105, 110) 以區分主鍵盤數字
            vk = getattr(key, "vk", None)
            if vk is None:
                s_key = str(key)
                if s_key.startswith("<") and s_key.endswith(">"):
                    try:
                        vk = int(s_key[1:-1])
                    except Exception as e:
                        logger.debug("[Input] Failed to parse VK: %s", e)

            if vk is not None:
                if 96 <= vk <= 105:
                    k_name = f"num_{vk - 96}"
                elif vk == 110:  # 數字鍵盤的點 (.)
                    k_name = "num_dot"

            # 2. 如果不是小鍵盤，則嘗試獲取按鍵名稱或字元
            if k_name is None:
                if hasattr(key, "name"):
                    k_name = key.name
                elif hasattr(key, "char") and key.char:
                    k_name = key.char.lower()

                pass

            if k_name:
                for base in ["alt", "shift", "ctrl"]:
                    if k_name.startswith(base):
                        k_name = base
                        break

            # 3. 配置切換 (連按兩下 F1-F8) 或 停用鍵 (F12)
            now = time.time()

            # 使用可自定義的熱鍵
            hks = current_config.get("hotkeys", {})

            # 2. 全域始終開啟的控制鍵
            if k_name == hks.get("show_settings", "pause"):
                logger.info("[Input] %s pressed. Emitting show signal.", k_name.upper())
                settings_window.request_show.emit()
                return

            # --- 如果正在錄製熱鍵，則攔截觸發器 ---
            if settings_window.is_recording or settings_window.recording_global_key:
                return

            if k_name == hks.get("reset", "f12"):
                logger.info("[Input] %s Reset Triggered.", k_name.upper())
                is_globally_enabled = False
                overlay.clear_request.emit()
                return

            if k_name == hks.get("exp_toggle", "f10"):
                overlay.toggle_exp_request.emit()
                return

            if k_name == hks.get("exp_pause", "f11"):
                overlay.toggle_pause_request.emit()
                return

            if k_name == hks.get("exp_report", "f12"):
                overlay.export_report_request.emit()
                return

            # --- RJPQ 遠程同步快捷鍵 ---
            rjpq_keys = {
                hks.get("rjpq_1", "num_1"): 0,
                hks.get("rjpq_2", "num_2"): 1,
                hks.get("rjpq_3", "num_3"): 2,
                hks.get("rjpq_4", "num_4"): 3,
            }

            if k_name in rjpq_keys:
                col_idx = rjpq_keys[k_name]
                # 檢查 RJPQ 是否啟動並已連線
                if (
                    hasattr(settings_window, "rjpq_tab")
                    and settings_window.rjpq_tab.client.is_connected
                ):
                    if settings_window.rjpq_tab.mark_by_hotkey(col_idx):
                        return  # 如果按鍵用於 RJPQ，則攔截它

            # 3. 配置切換 (連按兩下 F1-F8)
            now = time.time()
            if k_name and k_name.startswith("f") and len(k_name) <= 3:
                f_num = k_name[1:]
                if f_num.isdigit() and 1 <= int(f_num) <= 8:
                    if last_key == k_name and (now - last_time) < double_press_delay:
                        # 成功！切換配置
                        p_key = f"F{f_num}"
                        config = ConfigManager.load_config()
                        config["active_profile"] = p_key
                        ConfigManager.save_config(config)
                        update_local_config()
                        overlay.profile_switch_request.emit()
                        last_key = None  # 重置
                        return
                    last_key = k_name
                    last_time = now

            # 4. 僅在啟用且 msw.exe 處於焦點時觸發計時器
            if not is_globally_enabled or not focus_tracker.is_game_active:
                return

            # 獲取當前配置的觸發器
            active_p = current_config.get("active_profile", "F1")
            prof_data = current_config["profiles"].get(active_p, {"triggers": {}})
            triggers = prof_data.get("triggers", {})

            if k_name and k_name in triggers:
                trigger_data = triggers[k_name]
                if isinstance(trigger_data, dict):
                    seconds = trigger_data.get("seconds", 10)
                    icon = trigger_data.get("icon", "")
                    sound = trigger_data.get("sound", True)
                else:
                    seconds = trigger_data
                    icon = ""
                    sound = True
                overlay.timer_request.emit(k_name, seconds, icon if icon else "", sound)

        except Exception as e:
            logger.error("[Error] Listener: %s", e)

    listener = keyboard.Listener(on_press=on_press)
    listener.start()
    logger.info("[Input] Listener active. Press 'Pause Break' for Control Center (🍁).")

def check_network_drive():
    if sys.platform != "win32":
        return
    try:
        import win32file
        from PyQt6.QtWidgets import QMessageBox
        # 獲取執行檔路徑
        app_path = os.path.abspath(sys.argv[0])
        drive = os.path.splitdrive(app_path)[0]
        if drive:
            drive_type = win32file.GetDriveType(drive + "\\")
            if drive_type == win32file.DRIVE_REMOTE:
                msg = QMessageBox()
                msg.setIcon(QMessageBox.Icon.Warning)
                msg.setWindowTitle("環境建議 - Artale 瑞士刀")
                msg.setText("偵測到程式正在網路磁碟機 (Samba) 上執行，這可能導致 OCR 效能嚴重下降或辨識失敗。請考慮將程式移動到本機硬碟執行以獲得最佳效能。")
                msg.setInformativeText("在網路硬碟上執行可能會導致視窗捕捉失敗 (0x80070490)。\n\n建議將程式複製到「本機磁碟」(如桌面或 C 槽) 以獲得最佳穩定性。")
                msg.setStandardButtons(QMessageBox.StandardButton.Ok)
                msg.exec()
    except Exception as e:
        logger.debug("[Main] Network drive check skipped or failed: %s", e)
def run_app():
    # 設定儲存日誌
    log_file = os.path.join(os.getcwd(), "artale_agent.log")
    logging.basicConfig(
        level=logging.DEBUG,
        format="%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%H:%M:%S",
        handlers=[
            logging.StreamHandler(),
            logging.FileHandler(log_file, encoding="utf-8", mode="w"),
        ],
    )
    app_logger = logging.getLogger("Artale")
    app_logger.info("--- Artale Agent Initializing (Log: %s) ---", log_file)
    app_logger.info("[System] OS: %s", platform_mod.platform())

    # --- 初始化 Sentry (僅在打包後的正式環境啟用) ---
    if getattr(sys, "frozen", False):
        sentry_sdk.init(
            dsn="https://b120418a69ec5d8ccd74a0bb4d2acacf@o4511210222452736.ingest.us.sentry.io/4511210254565376",
            integrations=[
                LoggingIntegration(level=logging.INFO, event_level=logging.ERROR),
            ],
            traces_sample_rate=1.0,
            enable_logs=True,
            release=f"artale-agent@{get_version()}",
        )
    else:
        logger.info("[Main] Dev mode: Sentry disabled.")

    # --- 啟用高 DPI 感知 ---
    # Qt 6 預設為 PerMonitorAwareV2，因此手動呼叫 ctypes 是多餘的，有時會導致「拒絕訪問」錯誤。
    try:
        QApplication.setHighDpiScaleFactorRoundingPolicy(
            Qt.HighDpiScaleFactorRoundingPolicy.PassThrough
        )
    except Exception as e:
        logger.debug("[Main] Rounding policy set failed: %s", e)

    app = QApplication(sys.argv)

    from PyQt6.QtGui import QFont

    font = QFont()
    if sys.platform == "darwin":
        font.setFamilies(["PingFang TC", "Heiti TC"])
    else:
        font.setFamilies(["Microsoft JhengHei", "微軟正黑體"])
    app.setFont(font)
    app.setQuitOnLastWindowClosed(False)
    # 檢查是否在 Samba/網路磁碟機上執行 (目前註解掉)
    # check_network_drive()
    main_overlay = ArtaleOverlay()
    settings_window = main_overlay.settings_window
    focus_tracker = FocusTrackerImpl()
    focus_tracker.start("msw.exe")
    # 2026/04 模組化架構：實例化控制器以管理引擎與橋接邏輯
    from .controller import ArtaleController
    app_controller = ArtaleController(main_overlay)
    main_overlay.controller = app_controller
    app_controller.start()
    
    # 系統匣圖示連結
    main_overlay.settings_show_request.connect(settings_window.safe_show)

    start_keyboard_listener(main_overlay, settings_window, focus_tracker)
    # 啟動時自動顯示控制中心
    settings_window.safe_show()

    logger.info("[Main] Artale 瑞士刀 initialized. Waiting for input...")
    sys.exit(app.exec())


if __name__ == "__main__":
    run_app()
