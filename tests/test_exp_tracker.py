import unittest
from artale_agent.exp_tracker import ExpTracker
from artale_agent.utils import EXP_TABLE

class TestExpTrackerInferLevel(unittest.TestCase):
    def setUp(self):
        """每個測試執行前初始化 ExpTracker"""
        self.tracker = ExpTracker()

    def test_exact_match_low_level(self):
        """測試低等級 (LV.30) 的精準匹配"""
        # LV.30 升級所需經驗值為 95700
        # 假設當前經驗值百分比為 50.0%
        next_exp = EXP_TABLE[30]  # 95700
        current_exp = int(next_exp * 0.5)  # 47850
        current_pct = 50.0
        
        inferred = self.tracker.infer_level(current_exp, current_pct)
        self.assertEqual(inferred, 30)

    def test_exact_match_high_level(self):
        """測試高等級 (LV.120) 的精準匹配"""
        # LV.120 升級所需經驗值為 29715818
        # 假設當前經驗值百分比為 25.12%
        next_exp = EXP_TABLE[120]  # 29715818
        current_exp = int(next_exp * 0.2512)  # 7464613
        current_pct = 25.12
        
        inferred = self.tracker.infer_level(current_exp, current_pct)
        self.assertEqual(inferred, 120)

    def test_ocr_error_within_tolerance(self):
        """測試當 OCR 識別存在微小誤差時，在容錯門檻內仍能正確推算"""
        # LV.30 的容錯門檻為 max(100, 95700 * 0.0001) = 100
        # 精準 50% 經驗值為 47850
        # 假設 OCR 識別為 47870 (誤差為 +20，在 100 的容錯門檻內)
        current_pct = 50.0
        current_exp = 47870
        
        inferred = self.tracker.infer_level(current_exp, current_pct)
        self.assertEqual(inferred, 30)

    def test_ocr_error_outside_tolerance(self):
        """測試當 OCR 識別誤差超出容錯門檻時，推算失敗並回傳 None"""
        # LV.30 的容錯門檻為 100
        # 精準 50% 經驗值為 47850
        # 假設 OCR 識別為 48000 (誤差為 +150，大於 100 的容錯門檻)
        current_pct = 50.0
        current_exp = 48000
        
        inferred = self.tracker.infer_level(current_exp, current_pct)
        self.assertIsNone(inferred)

    def test_ambiguous_match_returns_none(self):
        """測試當存在多個可能等級（模糊性）時，推算失敗並回傳 None"""
        # 當經驗值與百分比皆為 0 時，對所有等級來說都是 0 經驗值 (0% 佔比)，
        # 導致多個等級都符合條件，此時應回傳 None
        inferred = self.tracker.infer_level(0, 0.0)
        self.assertIsNone(inferred)

    def test_no_match_returns_none(self):
        """測試當完全不匹配任何等級時，回傳 None"""
        # 給定一個完全不合理的極端值组合
        inferred = self.tracker.infer_level(999999999, 0.01)
        self.assertIsNone(inferred)

    def test_all_levels_exact_match(self):
        """遍歷所有等級，驗證在其 50% 經驗值時皆能唯一且正確地推導出該等級"""
        # 排除 1 等級，因為 1 等級所需經驗值極小 (15)，且 50% 只有 7 點，
        # 在容錯門檻 tolerance=100 下，容易與其他極低等級混淆。
        for lv in EXP_TABLE.keys():
            if lv <= 9:
                continue
            next_exp = EXP_TABLE[lv]
            current_pct = 50.0
            current_exp = int(next_exp * 0.5)
            
            inferred = self.tracker.infer_level(current_exp, current_pct)
            self.assertEqual(
                inferred, 
                lv, 
                f"無法正確推導等級 {lv} (EXP: {current_exp}, PCT: {current_pct}%)"
            )

    def test_infer_level_highest_boundary_level_200(self):
        """測試最高等級 LV.200 (極大數值與動態容錯門檻校驗)"""
        # LV.200 的總經驗值為 2121276324
        next_exp = EXP_TABLE[200]
        
        # 1. 精準 50% 匹配
        current_exp = int(next_exp * 0.5)
        current_pct = 50.0
        inferred = self.tracker.infer_level(current_exp, current_pct)
        self.assertEqual(inferred, 200)

        # LV.200 的動態容錯為 2121276324 * 0.0001 = 212127.6324 -> 212127
        # 2. 誤差在容錯內 (+50,000 EXP)，應仍能正確推導出 LV.200
        inferred_within = self.tracker.infer_level(current_exp + 50000, current_pct)
        self.assertEqual(inferred_within, 200)

        # 3. 誤差超出容錯外 (+500,000 EXP)，應推導失敗並傳回 None
        inferred_outside = self.tracker.infer_level(current_exp + 500000, current_pct)
        self.assertIsNone(inferred_outside)

    def test_infer_level_lowest_density_ambiguity(self):
        """測試極低等級 (LV.1 ~ LV.9) 的高密度模糊性校驗，避免誤判"""
        # 當經驗值極小且百分比非 0 時，因基本容錯門檻為 100，
        # 會同時匹配多個低等級（如 LV.1 的 15、LV.2 的 34、LV.3 的 60、LV.4 的 95 等皆在其容錯內）。
        # 此時應觸發模糊性保護，返回 None。
        current_exp = 7
        current_pct = 50.0
        inferred = self.tracker.infer_level(current_exp, current_pct)
        self.assertIsNone(inferred)

    def test_infer_level_anomalous_inputs(self):
        """測試異常與超出界限的極端輸入校驗"""
        # 1. 負數經驗值
        self.assertIsNone(self.tracker.infer_level(-50, 50.0))
        # 2. 負數百分比
        self.assertIsNone(self.tracker.infer_level(100, -10.0))
        # 3. 超出 100% 的百分比
        self.assertIsNone(self.tracker.infer_level(100, 150.0))
        # 4. 極大百分比
        self.assertIsNone(self.tracker.infer_level(100, 99999.0))
        # 5. 精確 100% 的邊界
        next_exp = EXP_TABLE[100]
        self.assertEqual(self.tracker.infer_level(next_exp, 100.0), 100)


class TestExpTrackerValidation(unittest.TestCase):
    def setUp(self):
        """每個測試執行前初始化 ExpTracker"""
        self.tracker = ExpTracker()

    def test_parse_exp_text(self):
        """測試經驗值 OCR 原始文字解析"""
        # 支援中括號格式
        val, pct = self.tracker.parse_exp_text("45013389 [98.85%]")
        self.assertEqual(val, 45013389)
        self.assertEqual(pct, 98.85)

        # 支援無中括號空格格式
        val, pct = self.tracker.parse_exp_text("45013389 98.85%")
        self.assertEqual(val, 45013389)
        self.assertEqual(pct, 98.85)

        # 測試解析失敗
        val, pct = self.tracker.parse_exp_text("invalid ocr text")
        self.assertIsNone(val)
        self.assertIsNone(pct)

    def test_validate_exp_confidence_filtering(self):
        """測試 validate_exp 的 OCR 信心度過濾"""
        # 信心度小於 90 且大於 0 時應拒絕
        val, pct, lv = self.tracker.validate_exp("47850 [50.00%]", 85)
        self.assertIsNone(val)
        
        # 信心度大於等於 90 時應採納
        val, pct, lv = self.tracker.validate_exp("47850 [50.00%]", 90)
        self.assertEqual(val, 47850)
        self.assertEqual(pct, 50.00)
        self.assertEqual(lv, 30)

        # 信心度等於 0 時應採納 (模擬手動觸發/強制採納場景)
        val, pct, lv = self.tracker.validate_exp("47850 [50.00%]", 0)
        self.assertEqual(val, 47850)

    def test_validate_exp_parsing_failure(self):
        """測試 validate_exp 解析文字失敗"""
        val, pct, lv = self.tracker.validate_exp("broken-text", 95)
        self.assertIsNone(val)
        self.assertIsNone(pct)
        self.assertIsNone(lv)

    def test_validate_exp_fallback_to_level_ocr(self):
        """測試當等級推論失敗時，若等級 OCR 信心度高則進行強制校驗"""
        # 當經驗值與百分比皆為 0 時，正常推論 infer_level 會因為模糊而傳回 None。
        # 但若此時等級 OCR 結果信心度高 (>= 90)，validate_exp 應結合此等級進行驗證。
        self.tracker.update_lv_ocr(30, 95.0) # 設定最後看到的等級 OCR 為 LV.30，信心度 95%
        
        # 0 EXP 符合 LV.30 (總量 95700)，且 0/95700 = 0% 與 OCR 百分比 0% 差距小於 2%
        val, pct, lv = self.tracker.validate_exp("0 [0.00%]", 95)
        self.assertEqual(val, 0)
        self.assertEqual(pct, 0.0)
        self.assertEqual(lv, 30)

        # 若百分比差距過大 (超過 2%)，仍應拒絕
        val, pct, lv = self.tracker.validate_exp("0 [5.00%]", 95)
        self.assertIsNone(lv)

    def test_validate_exp_user_case_level_160(self):
        """測試使用者提供之真實數據：88263884 EXP at 34.74% 應被判定為等級 160"""
        self.tracker.update_lv_ocr(160, 95.0) # 設定最後看到的等級 OCR 為 LV.160，信心度 95%
        val, pct, lv = self.tracker.validate_exp("88263884 [34.74%]", 95)
        self.assertEqual(val, 88263884)
        self.assertEqual(pct, 34.74)
        self.assertEqual(lv, 160)

    def test_validate_exp_level_sanity_check(self):
        """測試 validate_exp 的等級變動合理性校驗"""
        # 1. 初始化狀態
        self.tracker.current_lv = 30
        self.tracker.exp_session_start_time = 1000.0  # 模擬計時已啟動
        
        # 2. 等級相同：應通過
        # LV.30, 50% => 47850 exp
        val, pct, lv = self.tracker.validate_exp("47850 [50.00%]", 95)
        self.assertEqual(lv, 30)

        # 3. 等級剛好 +1 (升級)：應通過
        # LV.31, 50% => 108480 * 0.5 = 54240 exp
        val, pct, lv = self.tracker.validate_exp("54240 [50.00%]", 95)
        self.assertEqual(lv, 31)

        # 4. 等級異常跳變 (例如從 30 變成 35)：
        # LV.35, 50% => 174216 * 0.5 = 87108 exp
        # 前 9 次跳變應被過濾（傳回 None），以避免單次 OCR 誤判造成升級統計出錯
        for i in range(1, 10):
            val, pct, lv = self.tracker.validate_exp("87108 [50.00%]", 95)
            self.assertIsNone(lv, f"第 {i} 次異常跳變應該被攔截並返回 None")
            self.assertEqual(self.tracker.lv_mismatch_counter, i)
            self.assertEqual(self.tracker.pending_lv, 35)

        # 第 10 次跳變時應執行強制修正並採納新等級
        val, pct, lv = self.tracker.validate_exp("87108 [50.00%]", 95)
        self.assertEqual(lv, 35)
        self.assertEqual(self.tracker.lv_mismatch_counter, 0)
        self.assertIsNone(self.tracker.pending_lv)

    def test_validate_exp_dynamic_ratio_cases(self):
        """測試動態比例檢查：自動產生多筆不同等級的測資，經驗值百分比為當前經驗量除以總量並計算至小數以下第二位"""
        # 挑選不同量級的等級進行測試（包含極低等級、中等級、高等級）
        test_levels = [30, 70, 120, 160, 190]
        # 測試不同的百分比分佈
        test_percentages = [12.34, 34.74, 50.00, 75.89, 99.12]
        
        for lv in test_levels:
            next_exp = EXP_TABLE[lv]
            self.tracker.update_lv_ocr(lv, 95.0) # 設定最後看到的等級 OCR 為當前等級，信心度 95%
            
            for base_pct in test_percentages:
                # 1. 根據百分比計算對應的經驗值數值
                val = int(next_exp * (base_pct / 100.0))
                if val > next_exp:
                    val = next_exp
                
                # 2. 計算精準百分比，依據使用者要求：為當前經驗量除以總量計算到百分比小數點以下第二位
                pct = round((val / next_exp) * 100.0, 2)
                
                # 3. 構造 OCR 原始文字，測試是否能成功通過比例檢查
                raw_text = f"{val} [{pct:.2f}%]"
                res_val, res_pct, res_lv = self.tracker.validate_exp(raw_text, 95)
                
                # 4. 驗證結果
                self.assertEqual(res_lv, lv, f"等級 {lv} 在經驗值 {val} ({pct}%) 時應通過校驗")
                self.assertEqual(res_val, val)
                self.assertEqual(res_pct, pct)
                
                # 5. 測試不合法的百分比（超出容錯範圍）
                val_tolerance = max(100, next_exp * 0.0001)
                dynamic_tolerance = (val_tolerance / next_exp) * 100.0
                
                # 偏移量設定為大於動態容錯門檻，且至少有 0.05% 的安全波動
                offset = (dynamic_tolerance * 2) + 0.05
                invalid_pct = round(pct + offset, 2)
                
                # 構造不合理的 OCR 原始文字，應被比例檢查拒絕並回傳 None
                raw_text_invalid = f"{val} [{invalid_pct:.2f}%]"
                res_val_inv, res_pct_inv, res_lv_inv = self.tracker.validate_exp(raw_text_invalid, 95)
                self.assertIsNone(res_lv_inv, f"等級 {lv} 的不合理百分比 {invalid_pct}% 應該要被比例檢查拒絕")


if __name__ == '__main__':
    unittest.main()
