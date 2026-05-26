import os
import cv2
import queue
import logging
import threading
import time
from datetime import datetime
from PyQt6.QtCore import QStandardPaths

logger = logging.getLogger(__name__)

class VideoRecorder:
    def __init__(self, fps=30.0):
        self.fps = fps
        self.is_recording = False
        self.frame_queue = queue.Queue()
        self.write_thread = None
        self.video_writer = None
        self.filepath = None
        self.on_save_finished = None  # 新增完成存檔時的非同步回調
        
        # 動態獲取 Windows 系統 Videos (影片) 目錄，完美適配您的網路磁碟機路徑
        system_video_dir = QStandardPaths.writableLocation(QStandardPaths.StandardLocation.MoviesLocation)
        self.base_dir = os.path.join(system_video_dir, "ArtaleAgent")
        
        # 30 FPS 節流參數：確保影片以 1.0x 實時速度播放
        self.last_write_time = 0
        self.frame_interval = 1.0 / fps
        
    def start(self, save_path=None, on_save_finished=None):
        """開始錄影，可傳入自定義儲存路徑與存檔完成回調"""
        if self.is_recording:
            return
            
        try:
            # 支援自定義儲存目錄
            if save_path:
                self.base_dir = save_path
            else:
                system_video_dir = QStandardPaths.writableLocation(QStandardPaths.StandardLocation.MoviesLocation)
                self.base_dir = os.path.join(system_video_dir, "ArtaleAgent")
                
            if not os.path.exists(self.base_dir):
                os.makedirs(self.base_dir, exist_ok=True)
                
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"Artale_Record_{timestamp}.mp4"
            self.filepath = os.path.join(self.base_dir, filename)
            
            self.video_writer = None
            self.last_write_time = 0
            self.is_recording = True
            self.on_save_finished = on_save_finished
            
            # 清空隊列
            while not self.frame_queue.empty():
                try:
                    self.frame_queue.get_nowait()
                except queue.Empty:
                    break
                    
            # 啟動非同步寫入背景執行緒
            self.write_thread = threading.Thread(target=self._write_loop, daemon=True)
            self.write_thread.start()
            logger.info("[Recorder] Recording session initialized. Path: %s", self.filepath)
        except Exception as e:
            logger.error("[Recorder] Error starting recorder: %s", e)

    def write_frame(self, frame):
        """外部影格輸入端，完全非同步並進行 30 FPS 節流以確保 1.0x 實時播放速度"""
        if self.is_recording and frame is not None:
            now = time.time()
            # 容許 5ms 的執行緒抖動誤差，確保影片能穩定錄製滿 30 FPS
            if now - self.last_write_time >= self.frame_interval - 0.005:
                self.last_write_time = now
                # 拷貝影格以防主執行緒在寫入硬碟前修改影像
                self.frame_queue.put(frame.copy())
            
    def stop(self):
        """停止錄影，放入結束標記後立即返回（非阻塞）"""
        if not self.is_recording:
            return None
            
        self.is_recording = False
        # 塞入結束標記 None，背景執行緒會自動默默寫完剩下影格並安全釋放資源
        self.frame_queue.put(None)
        
        # 立即返回當前錄影路徑，不調用 join() 阻塞主執行緒，實現 0 毫秒極速響應！
        ret_path = self.filepath
        self.filepath = None
        self.write_thread = None
        return ret_path
        
    def _write_loop(self):
        """背景寫入背景執行緒循環"""
        while True:
            try:
                frame = self.frame_queue.get(timeout=1.0)
                if frame is None:  # 收到結束標記
                    break
                if frame.size > 0:
                    # 延遲初始化 VideoWriter (以第一幀的大小為準)
                    if self.video_writer is None:
                        h, w = frame.shape[:2]
                        logger.info("[Recorder] Creating VideoWriter with resolved dimensions: %sx%s", w, h)
                        fourcc = cv2.VideoWriter_fourcc(*'mp4v')
                        self.video_writer = cv2.VideoWriter(self.filepath, fourcc, self.fps, (w, h))
                        if not self.video_writer.isOpened():
                            logger.error("[Recorder] Failed to open VideoWriter!")
                            self.video_writer = None
                            
                    if self.video_writer:
                        self.video_writer.write(frame)
                self.frame_queue.task_done()
            except queue.Empty:
                if not self.is_recording:
                    break
            except Exception as e:
                logger.error("[Recorder] Write loop error: %s", e)
                
        # 迴圈退出時，代表所有影格都已經寫入完畢，安全釋放 VideoWriter 資源
        if self.video_writer:
            self.video_writer.release()
            self.video_writer = None
            
        logger.info("[Recorder] Thread finished writing all frames asynchronously.")
        
        # 執行完成回調
        if self.on_save_finished:
            try:
                self.on_save_finished(self.filepath)
            except Exception as e:
                logger.error("[Recorder] Save finished callback error: %s", e)
