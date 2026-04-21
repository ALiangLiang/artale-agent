import cv2
import numpy as np
import logging
import os
from typing import Optional, Tuple
from PyQt6.QtCore import QObject, pyqtSignal
from dataclasses import dataclass

# Optional Tesseract
try:
    import pytesseract
except ImportError:
    pytesseract = None

logger = logging.getLogger(__name__)


from artale_agent.data_types import LVUpdateData, ExpParsedData, ExpVisualData

class ArtaleOCR(QObject):
    """
    處理 Artale 的所有 OCR (光學字元辨識) 與模板比對邏輯。
    與主要的 Overlay 介面邏輯解耦。
    """
    # 用於 UI 更新的訊號 (由 ArtaleOverlay 連接)
    lv_update = pyqtSignal(LVUpdateData)
    money_update = pyqtSignal(int)
    exp_visual_update = pyqtSignal(ExpVisualData)
    exp_parsed_update = pyqtSignal(ExpParsedData) # 用於呼叫 overlay 中的 parse_and_update_exp

    def __init__(self, parent: Optional[QObject] = None) -> None:
        super().__init__(parent)
        # 設計常數設定
        # 1080p 校準參考 (與 ArtaleOverlay 保持同步)
        self.BASE_W, self.BASE_H = 1920, 1080
        self.LV_X_OFF_FROM_LEFT, self.LV_Y_OFF_FROM_BOTTOM = 92, 46
        self.LV_BASE_CW, self.LV_BASE_CH = 84, 26
        self.X_OFF_FROM_LEFT, self.Y_OFF_FROM_BOTTOM = 1084, 69
        self.BASE_CW, self.BASE_CH = 240, 26
        
        # 內部狀態
        self.coin_tpl = None
        self.last_coin_pos = None
        self.last_coin_match_conf = 0.0
        self.last_money_ocr_conf = 0.0
        self.last_crop_info = None
        self.exp_paused = False
        self.show_money_log = True
        self.show_debug = False

    def set_coin_template(self, tpl_path: str) -> None:
        if os.path.exists(tpl_path):
            self.coin_tpl = cv2.imread(tpl_path)
            if self.coin_tpl is not None:
                logger.info("[OCR] Coin template loaded from %s", tpl_path)
        else:
            logger.warning("[OCR] Coin template NOT found at %s", tpl_path)

    def perform_enhanced_ocr(self, thresh_img, key, upscale=3.0, whitelist=None):
        """
        從舊版 overlay.py 恢復的雙階段分割模式邏輯。
        將括號與數值分開辨識，以提升 Tesseract 的準確度。
        """
        if not pytesseract or not pytesseract.pytesseract.tesseract_cmd:
            if self.show_debug: logger.warning("[OCR] Tesseract path missing!")
            return "", 0, thresh_img

        # 1. 填充與分割 (尋找個別字元)
        # 使用黑色背景，因為之後會進行反相處理
        thresh_img = cv2.copyMakeBorder(thresh_img, 10, 10, 50, 50, cv2.BORDER_CONSTANT, value=0)
        contours, _ = cv2.findContours(thresh_img, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        
        raw_boxes = []
        for cnt in contours:
            x, y, w, h = cv2.boundingRect(cnt)
            # 字元高度過濾
            if h >= 2: raw_boxes.append([x, y, w, h])
        raw_boxes.sort(key=lambda b: b[0])
        
        # 合併重疊或相鄰的區塊
        merged_boxes = []
        if raw_boxes:
            curr = raw_boxes[0]
            for i in range(1, len(raw_boxes)):
                nxt = raw_boxes[i]
                if nxt[0] <= (curr[0] + curr[2] + 2): 
                    x1 = min(curr[0], nxt[0]); y1 = min(curr[1], nxt[1])
                    x2 = max(curr[0] + curr[2], nxt[0] + nxt[2])
                    y2 = max(curr[1] + curr[3], nxt[1] + nxt[3])
                    curr = [x1, y1, x2 - x1, y2 - y1]
                else:
                    merged_boxes.append(curr); curr = nxt
            merged_boxes.append(curr)
            
        if not merged_boxes:
            return "", 0, cv2.bitwise_not(thresh_img)
            
        # 2. 邏輯分割：依高度識別括號 (括號通常最高)
        bracket_idx = -1
        bracket_end_idx = -1
        
        if key == "exp" and len(merged_boxes) > 4:
            heights = [b[3] for b in merged_boxes]
            max_h = max(heights)
            min_h = min(heights)
            
            # 高度差異顯著則代表存在括號
            if max_h >= min_h + 2:
                for i, b in enumerate(merged_boxes):
                    if b[3] >= max_h - 1:
                        if bracket_idx == -1: bracket_idx = i
                        bracket_end_idx = i
        
        # 3. 畫布建構
        char_spacing = 30
        
        def build_pass_canvas(boxes):
            if not boxes: return None
            scaled_dims = [(int(b[2] * upscale), int(b[3] * upscale)) for b in boxes]
            max_nh = max([d[1] for d in scaled_dims])
            total_w = sum([d[0] for d in scaled_dims]) + (len(boxes) + 1) * char_spacing + 180
            canvas_h = max_nh + 80 
            canvas = np.ones((canvas_h, total_w), dtype=np.uint8) * 255 # 白色背景 (黑色文字)
            curr_x = 90 
            baseline_y = max_nh + 40
            
            for idx, b in enumerate(boxes):
                char = thresh_img[b[1]:b[1]+b[3], b[0]:b[0]+b[2]]
                char_inv = cv2.bitwise_not(char) # 反相後為黑色文字
                nw, nh = scaled_dims[idx]
                if nw > 0 and nh > 0:
                    resized = cv2.resize(char_inv, (nw, nh), interpolation=cv2.INTER_CUBIC)
                    y_off = baseline_y - nh
                    canvas[y_off:y_off+nh, curr_x:curr_x+nw] = resized
                    curr_x += nw + char_spacing
            return canvas

        parts = []
        if bracket_idx != -1:
            parts.append({'boxes': merged_boxes[:bracket_idx], 'wl': '0123456789'})
            parts.append({'boxes': merged_boxes[bracket_idx:], 'wl': '0123456789.%'})
        else:
            parts.append({'boxes': merged_boxes, 'wl': whitelist if whitelist else '0123456789'})

        # 4. OCR 處理與除錯影像拼合
        final_texts = []
        conf_sums = []
        canvases = []
        found_brackets = (bracket_idx != -1)

        for i, p in enumerate(parts):
            boxes_to_use = p['boxes']
            if found_brackets and i == 1:
                start_in_part_b = 1
                end_in_part_b = len(boxes_to_use)
                if bracket_end_idx != -1 and bracket_end_idx > bracket_idx:
                    end_in_part_b = len(boxes_to_use) - 1
                boxes_to_use = boxes_to_use[start_in_part_b:end_in_part_b]

            cvs = build_pass_canvas(boxes_to_use)
            if cvs is not None:
                canvases.append(cvs)
                wl = p['wl']
                config = f'--psm 7 --oem 3 -c tessedit_char_whitelist={wl}'
                try:
                    data = pytesseract.image_to_data(cvs, config=config, output_type=pytesseract.Output.DICT)
                    txts = [t for t in data['text'] if t.strip()]
                    confs = [int(c) for c in data['conf'] if int(c) != -1]
                    
                    if i == 1 and found_brackets:
                        res = "".join(txts)
                        final_texts.append(f"[{res}]")
                    else:
                        final_texts.append("".join(txts))
                    if confs: conf_sums.extend(confs)
                except Exception as e:
                    logger.debug("[OCR] Pass Error: %s", e)
        
        # 為除錯介面拼合畫布影像
        if len(canvases) > 1:
            max_w = max(c.shape[1] for c in canvases)
            total_h = sum(c.shape[0] for c in canvases) + (len(canvases) - 1) * 10
            combined = np.ones((total_h, max_w), dtype=np.uint8) * 255
            curr_y = 0
            for c in canvases:
                h_c, w_c = c.shape
                combined[curr_y:curr_y+h_c, :w_c] = c
                curr_y += h_c + 10 
            processed_img = combined
        elif canvases:
            processed_img = canvases[0]
        else:
            processed_img = cv2.bitwise_not(thresh_img)

        ocr_text = "".join(final_texts)
        score = sum(conf_sums) / len(conf_sums) if conf_sums else 0
        
        return ocr_text, score, processed_img

    def _get_lv_crop(self, img, scale, off_x, off_y, ch_ref):
        """計算並裁切等級區域"""
        h, w = img.shape[:2]
        lv_cx, lv_cy = off_x + int(self.LV_X_OFF_FROM_LEFT * scale), off_y + (ch_ref - int(self.LV_Y_OFF_FROM_BOTTOM * scale))
        lv_cw, lv_ch = int(self.LV_BASE_CW * scale), int(self.LV_BASE_CH * scale)
        return img[max(0, lv_cy):min(h, lv_cy+lv_ch), max(0, lv_cx):min(w, lv_cx+lv_cw)]

    def _get_money_crop(self, img, scale):
        """透過模板匹配動態定位並裁切楓幣區域"""
        if self.coin_tpl is None: return None
        h, w = img.shape[:2]
        tpl_h, tpl_w = self.coin_tpl.shape[:2]
        st_w, st_h = int(tpl_w * scale), int(tpl_h * scale)
        if st_w <= 5 or st_h <= 5: return None
        
        tpl_resized = cv2.resize(self.coin_tpl, (st_w, st_h))
        res = cv2.matchTemplate(img, tpl_resized, cv2.TM_CCOEFF_NORMED)
        _, max_val, _, max_loc = cv2.minMaxLoc(res)
        
        if max_val > 0.7:
            info_w, info_h = int(280 * scale), int(31 * scale)
            m_ix = max_loc[0] + st_w + int(20 * scale)
            m_iy = max_loc[1] + (st_h // 2) - (info_h // 2) + int(1 * scale)
            return img[max(0, m_iy):min(h, m_iy+info_h), max(0, m_ix):min(w, m_ix+info_w)]
        return None

    def _get_exp_crop(self, img, scale, off_x, off_y, ch_ref):
        """計算並裁切經驗值區域"""
        h, w = img.shape[:2]
        exp_lx, exp_ly = off_x + int(self.X_OFF_FROM_LEFT * scale), off_y + (ch_ref - int(self.Y_OFF_FROM_BOTTOM * scale)) + 2
        exp_cw, exp_ch = int(self.BASE_CW * scale), int(self.BASE_CH * scale) - 2
        return img[max(0, exp_ly):min(h, exp_ly+exp_ch), max(0, exp_lx):min(w, exp_lx+exp_cw)]

    def process_frame(self, img: np.ndarray, scale: float, off_x: int, off_y: int, cw_ref: int, ch_ref: int) -> None:
        """核心處理引擎：將影像分區並執行獨立的 OCR 辨識"""
        if self.exp_paused and not self.show_debug: return
        try:
            results = {"exp_conf": 0, "lv_conf": 0, "money_conf": 0}
            lv_thresh = None; m_thresh = None; exp_stack = None
            
            # --- 1. 等級辨識 ---
            lv_crop = self._get_lv_crop(img, scale, off_x, off_y, ch_ref)
            if lv_crop.size > 0:
                if self.show_debug: cv2.imwrite("./tmp/debug_lv_crop.png", lv_crop)
                # 預處理 (加入 scale 以還原解析度並放大)
                lv_thresh = self.preprocess_for_ocr(lv_crop, scale, threshold=180)
                # 新增 Padding 提升 Tesseract 辨識品質
                lv_thresh = cv2.copyMakeBorder(lv_thresh, 20, 20, 20, 20, cv2.BORDER_CONSTANT, value=(255,255,255))
                lv_txt, lv_conf = self._do_single_ocr(lv_thresh, "0123456789", psm=7)
                self.lv_update.emit(LVUpdateData(level=lv_txt, conf=lv_conf))
                results["lv_conf"] = lv_conf

            # --- 2. 楓幣辨識 ---
            if self.show_money_log:
                m_crop = self._get_money_crop(img, scale)
                if m_crop is not None and m_crop.size > 0:
                    # 預處理
                    m_thresh = self.preprocess_for_ocr(m_crop, scale, threshold=120)
                    m_txt, m_conf = self._do_single_ocr(m_thresh, "0123456789,", psm=7)
                    m_val_clean = "".join(filter(lambda c: c.isdigit(), m_txt))
                    if m_val_clean: self.money_update.emit(int(m_val_clean))
                    results["money_conf"] = m_conf

            # --- 3. 經驗值辨識 ---
            if not self.exp_paused:
                exp_crop = self._get_exp_crop(img, scale, off_x, off_y, ch_ref)
                if exp_crop.size > 0:
                    # 預處理
                    full_thresh = self.preprocess_for_ocr(exp_crop, scale, threshold=150)
                    if self.show_debug: cv2.imwrite("./tmp/debug_exp_processed.png", full_thresh)

                    # 切分經驗值與百分比，有可能拆壞
                    ev, ep = self.split_already_threshed(full_thresh, scale)
                    if ev is not None and ep is not None:
                        h_v, w_v = ev.shape; h_p, w_p = ep.shape
                        max_w = max(w_v, w_p)
                        spacing = 20; pad = 20
                        canvas_h = h_v + h_p + spacing + (pad * 2)
                        canvas_w = max_w + (pad * 2)
                        
                        exp_stack = np.ones((canvas_h, canvas_w), dtype=np.uint8) * 255
                        exp_stack[pad : pad+h_v, pad : pad+w_v] = ev
                        exp_stack[pad+h_v+spacing : pad+h_v+spacing+h_p, pad : pad+w_p] = ep
                        
                        exp_txt, exp_conf = self._do_single_ocr(exp_stack, "0123456789.%", psm=6)
                        self.exp_parsed_update.emit(ExpParsedData(
                            text=exp_txt, 
                            e_conf=exp_conf,
                            thresh=exp_stack,
                            scale=scale
                        ))
                        results["exp_conf"] = exp_conf

            # --- 最終除錯影像發送 ---
            if self.show_debug:
                self.exp_visual_update.emit(ExpVisualData(
                    exp=exp_stack,
                    lv=lv_thresh,
                    coin=m_thresh,
                    conf=sum(results.values()) / 3
                ))

        except Exception as e:
            logger.debug("[OCR] Process Frame Error: %s", e)

    def _do_single_ocr(self, img: np.ndarray, whitelist: str, psm: int = 7) -> Tuple[str, float]:
        """執行單一影像區域的 OCR"""
        if img is None: return "", 0
        try:
            config = f'--psm {psm} -c tessedit_char_whitelist={whitelist}'
            data = pytesseract.image_to_data(img, config=config, output_type=pytesseract.Output.DICT)
            txts = [t for t in data['text'] if t.strip()]
            confs = [c for c in data['conf'] if c >= 0]
            avg_conf = sum(confs)/len(confs) if confs else 0
            return " ".join(txts), avg_conf
        except:
            return "", 0

    def split_already_threshed(self, threshed_exp: np.ndarray, scale: float) -> Tuple[Optional[np.ndarray], Optional[np.ndarray]]:
        """抹除括號並保留原始字元細節 (包含孔洞)"""
        try:
            # 1. 偵測括號輪廓 (需在黑底白字上進行)
            inv = cv2.bitwise_not(threshed_exp)
            contours, _ = cv2.findContours(inv, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
            if len(contours) < 3: return None, None
            
            objs = []
            for c in contours:
                x, y, w, h = cv2.boundingRect(c)
                objs.append({'x': x, 'y': y, 'w': w, 'h': h, 'c': c})
            
            # --- 新增驗證邏輯：最右側的物件必須是最高的兩個之一 (右括號判定) ---
            # 1. 找到最右側的物件
            objs_by_x = sorted(objs, key=lambda o: o['x'])
            last_obj = objs_by_x[-1]
            
            # 2. 找到最高的兩個物件
            objs_by_h = sorted(objs, key=lambda o: o['h'], reverse=True)
            top2_h = objs_by_h[:2]
            
            # 3. 驗證：如果最右側的物件不在最高的前二中，代表右括號抓錯或被遮擋
            if last_obj not in top2_h:
                return None, None
            
            # 確認括號位置
            brackets = sorted(top2_h, key=lambda o: o['x'])
            b_left = brackets[0]; b_right = brackets[1]
            
            # 2. 建立「抹除遮罩」並套用至原圖
            # 我們只想要抹除 brackets，其餘字元維持原圖像素以保留孔洞
            clean_img = threshed_exp.copy()
            mask = np.zeros_like(threshed_exp)
            cv2.drawContours(mask, [b_left['c'], b_right['c']], -1, 255, thickness=-1)
            # 將遮罩範圍內的像素塗白 (背景色)
            clean_img[mask == 255] = 255
            
            # 3. 依據括號位置分割已抹除括號的原圖
            # 數值部分：第一個括號左側 (加上小補償避免切到邊緣)
            val_img = clean_img[:, :max(0, b_left['x'] - 2)]
            
            # 百分比部分：第一個括號與第二個括號之間
            pct_start = b_left['x'] + b_left['w'] + 2
            pct_end = b_right['x'] - 2
            if pct_end > pct_start:
                pct_img = clean_img[:, pct_start:pct_end]
            else:
                pct_img = None
                
            return val_img, pct_img
        except Exception as e:
            logger.debug("[OCR] Masked Split failed: %s", e)
        return None, None

    def preprocess_for_ocr(self, img: np.ndarray, scale: float, threshold: Optional[int] = 150) -> Optional[np.ndarray]:
        """二值化處理並加入平滑化邏輯"""
        if img is None or img.size == 0: return None
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        
        # 1. 放大影格：先根據 scale 還原回 1080p 基準，再放大 3 倍以利辨識
        upscale = (1.0 / scale) * 3.0
        gray = cv2.resize(gray, None, fx=upscale, fy=upscale, interpolation=cv2.INTER_CUBIC)
        
        # 2. 邊緣平滑化：在二值化前進行高斯模糊
        gray = cv2.GaussianBlur(gray, (3, 3), 0)
        
        # 3. 二值化
        _, thresh = cv2.threshold(gray, threshold, 255, cv2.THRESH_BINARY)
        
        # 4. 中值濾波：去除噪點並稍微平滑二值化後的邊緣
        thresh = cv2.medianBlur(thresh, 3)
        
        return cv2.bitwise_not(thresh)
