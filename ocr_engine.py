import cv2
import numpy as np
import logging
import os
from PyQt6.QtCore import QObject, pyqtSignal

# Optional Tesseract
try:
    import pytesseract
except ImportError:
    pytesseract = None

logger = logging.getLogger("ArtaleOCR")

class ArtaleOCR(QObject):
    """
    處理 Artale 的所有 OCR (光學字元辨識) 與模板比對邏輯。
    與主要的 Overlay 介面邏輯解耦。
    """
    # 用於 UI 更新的訊號 (由 ArtaleOverlay 連接)
    lv_update = pyqtSignal(dict)
    lv_update = pyqtSignal(dict)
    money_update = pyqtSignal(int)
    exp_visual_update = pyqtSignal(dict)
    exp_parsed_update = pyqtSignal(dict) # 用於呼叫 overlay 中的 parse_and_update_exp

    def __init__(self, parent=None):
        super().__init__(parent)
        # 設計常數設定
        # 1080p 校準參考 (與 ArtaleOverlay 保持同步)
        self.BASE_W, self.BASE_H = 1920, 1080
        self.LV_X_OFF_FROM_LEFT, self.LV_Y_OFF_FROM_BOTTOM = 96, 46
        self.LV_BASE_CW, self.LV_BASE_CH = 75, 26
        self.X_OFF_FROM_LEFT, self.Y_OFF_FROM_BOTTOM = 1084, 66
        self.BASE_CW, self.BASE_CH = 240, 22
        
        # 內部狀態
        self.coin_tpl = None
        self.last_coin_pos = None
        self.last_coin_match_conf = 0.0
        self.last_money_ocr_conf = 0.0
        self.last_crop_info = None
        self.exp_paused = False
        self.show_money_log = True
        self.show_debug = False

    def set_coin_template(self, tpl_path):
        if os.path.exists(tpl_path):
            self.coin_tpl = cv2.imread(tpl_path)
            if self.coin_tpl is not None:
                logger.info(f"[OCR] Coin template loaded from {tpl_path}")
        else:
            logger.warning(f"[OCR] Coin template NOT found at {tpl_path}")

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
                    logger.debug(f"[OCR] Pass Error: {e}")
        
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

    def process_frame(self, img, scale, off_x, off_y, cw_ref, ch_ref):
        """單次批次 OCR 處理，兼顧清晰度與效能"""
        if self.exp_paused and not self.show_debug: return
        try:
            h, w = img.shape[:2]
            parts = [] 
            
            # --- 1. 等級區域 (LEVEL) ---
            lv_cx, lv_cy = off_x + int(self.LV_X_OFF_FROM_LEFT * scale), off_y + (ch_ref - int(self.LV_Y_OFF_FROM_BOTTOM * scale))
            lv_cw, lv_ch = int(self.LV_BASE_CW * scale), int(self.LV_BASE_CH * scale)
            lv_crop = img[max(0, lv_cy):min(h, lv_cy+lv_ch), max(0, lv_cx):min(w, lv_cx+lv_cw)]
            if lv_crop.size > 0:
                parts.append(("LV", self.preprocess_for_ocr(lv_crop)))

            # --- 2. 楓幣區域 (MONEY) ---
            if self.show_money_log and self.coin_tpl is not None:
                tpl_h, tpl_w = self.coin_tpl.shape[:2]
                st_w, st_h = int(tpl_w * scale), int(tpl_h * scale)
                if st_w > 5 and st_h > 5:
                    tpl_resized = cv2.resize(self.coin_tpl, (st_w, st_h))
                    res = cv2.matchTemplate(img, tpl_resized, cv2.TM_CCOEFF_NORMED)
                    _, max_val, _, max_loc = cv2.minMaxLoc(res)
                    if max_val > 0.7:
                        info_w, info_h = int(280 * scale), int(31 * scale)
                        m_ix = max_loc[0] + st_w + int(30 * scale)
                        m_iy = max_loc[1] + (st_h // 2) - (info_h // 2) + int(1 * scale)
                        m_crop = img[max(0, m_iy):min(h, m_iy+info_h), max(0, m_ix):min(w, m_ix+info_w)]
                        if m_crop.size > 0:
                            parts.append(("MONEY", self.preprocess_for_ocr(m_crop, threshold=120)))

            # --- 3. 經驗值區域 (EXP - 優化後的單次分割) ---
            if not self.exp_paused:
                # 垂直方向微調以避開 UI 邊框
                exp_lx, exp_ly = off_x + int(self.X_OFF_FROM_LEFT * scale), off_y + (ch_ref - int(self.Y_OFF_FROM_BOTTOM * scale)) + 2
                exp_cw, exp_ch = int(self.BASE_CW * scale), int(self.BASE_CH * scale) - 2
                exp_crop = img[max(0, exp_ly):min(h, exp_ly+exp_ch), max(0, exp_lx):min(w, exp_lx+exp_cw)]
                if exp_crop.size > 0:
                    # 僅執行一次預處理
                    full_thresh = self.preprocess_for_ocr(exp_crop)
                    # 嘗試基於已二值化的影像進行分割
                    ev, ep = self.split_already_threshed(full_thresh, scale)
                    if ev is not None and ep is not None:
                        parts.append(("EXP_V", ev))
                        parts.append(("EXP_P", ep))
                    else:
                        parts.append(("EXP_FULL", full_thresh))

            if not parts: return
            
            # --- 2. 垂直拼接影像 (Stitch vertically) ---
            # 我們將小圖拼成一張大圖，減少 Tesseract 的啟動次數，提高效率。
            spacing = 45
            total_h = sum(p[1].shape[0] for p in parts) + spacing * (len(parts) - 1)
            max_w = max(p[1].shape[1] for p in parts)
            canvas = np.ones((total_h + 80, max_w + 80), dtype=np.uint8) * 255
            cur_y = 40
            for i, (p_type, p_img) in enumerate(parts):
                ph, pw = p_img.shape[:2]
                canvas[cur_y:cur_y+ph, 40:40+pw] = p_img
                if i < len(parts) - 1:
                    line_y = cur_y + ph + (spacing // 2)
                    cv2.line(canvas, (20, line_y), (max_w + 60, line_y), 200, 1)
                cur_y += ph + spacing

            # --- 3. 呼叫 Tesseract 字典模式 (用於獲取信心度) ---
            config = '--psm 6 -c tessedit_char_whitelist=0123456789.%,LV[] '
            data = pytesseract.image_to_data(canvas, config=config, output_type=pytesseract.Output.DICT)
            
            lines_data = [] 
            current_line_text, current_line_conf = [], []
            last_line_num = -1
            
            for i in range(len(data['text'])):
                txt = data['text'][i].strip()
                if not txt: continue
                line_n = data['line_num'][i]
                conf_val = data['conf'][i]
                
                if line_n != last_line_num and last_line_num != -1:
                    # 真實 Tesseract 信心度：僅對有效分數 (>0) 進行平均
                    valid_scores = [c for c in current_line_conf if c >= 0]
                    avg_conf = sum(valid_scores)/len(valid_scores) if valid_scores else 0
                    lines_data.append({"text": " ".join(current_line_text), "conf": avg_conf})
                    current_line_text, current_line_conf = [], []
                
                current_line_text.append(txt)
                current_line_conf.append(conf_val)
                last_line_num = line_n
                if self.show_debug:
                    logger.debug(f"[OCR] Token: '{txt}' | Conf: {conf_val}")
                
            if current_line_text:
                valid_scores = [c for c in current_line_conf if c >= 0]
                avg_conf = sum(valid_scores)/len(valid_scores) if valid_scores else 0
                lines_data.append({"text": " ".join(current_line_text), "conf": avg_conf})

            if self.show_debug:
                logger.debug(f"[OCR] Raw Tesseract Confidences: {data['conf']}")

            # --- 4. 最終數據對齊與分發 (Final Mapping) ---
            results = {} 
            line_idx = 0
            for p_type, _ in parts:
                if line_idx >= len(lines_data): break
                row = lines_data[line_idx]
                
                if p_type == "LV":
                    self.lv_update.emit({"level": row['text'], "conf": row['conf']})
                    results['lv_conf'] = row['conf']
                elif p_type == "MONEY":
                    m_val = "".join(filter(str.isdigit, row['text']))
                    if m_val: self.money_update.emit(int(m_val))
                    results['money_conf'] = row['conf']
                elif p_type == "EXP_V":
                    nxt = lines_data[line_idx+1] if line_idx+1 < len(lines_data) else {"text": "0", "conf": 0}
                    combined_conf = (row['conf'] + nxt['conf']) / 2
                    self.exp_parsed_update.emit({
                        "text": f"{row['text']} {nxt['text']}", 
                        "e_conf": combined_conf,
                        "thresh": canvas,
                        "crop": None
                    })
                    results['exp_conf'] = combined_conf
                    line_idx += 1
                elif p_type == "EXP_FULL":
                    self.exp_parsed_update.emit({
                        "text": row['text'], 
                        "e_conf": row['conf'],
                        "thresh": canvas,
                        "crop": None
                    })
                    results['exp_conf'] = row['conf']
                
                line_idx += 1
                
            if self.show_debug:
                results.update({"exp": canvas})
                self.exp_visual_update.emit(results)

        except Exception as e:
            logger.debug(f"[OCR] Batch process error: {e}")

    def split_already_threshed(self, threshed_exp, scale):
        """將已二值化並放大過的影像分割為 [數值] 與 [百分比]"""
        try:
            # 用於尋找括號的比例必須計入預處理中的 4 倍放大
            upscale_f = 4.0
            contours, _ = cv2.findContours(threshed_exp, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
            brackets = []
            for c in contours:
                x, y, w, h = cv2.boundingRect(c)
                # 調整放縮影像的判定基準
                if h > 20 * scale and w < 15 * scale:
                    brackets.append((x, y, w, h))
            
            if len(brackets) >= 2:
                brackets.sort(key=lambda b: b[0])
                bx1 = brackets[0][0]
                # 從第一個括號前幾個像素處切割
                val_img = threshed_exp[:, :bx1 - 2]
                pct_img = threshed_exp[:, bx1 - 2:]
                return val_img, pct_img
        except: pass
        return None, None

    def preprocess_for_ocr(self, img, threshold=None):
        """單一二值化處理點，支援 Otsu 演算法"""
        if img is None or img.size == 0: return None
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        gray = cv2.bitwise_not(gray) # 反相處理
        gray = cv2.resize(gray, None, fx=4.0, fy=4.0, interpolation=cv2.INTER_CUBIC)
        # 預設使用 Otsu 以避免手動調整閾值的麻煩
        _, thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
        return thresh

    def ocr_level(self, *args): pass
    def ocr_money(self, *args): pass
    def ocr_exp(self, *args): pass
