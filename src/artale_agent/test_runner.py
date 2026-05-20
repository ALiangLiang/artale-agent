import os
import sys
import unittest

def run_tests():
    """專案單元測試執行入口點"""
    # 取得專案根目錄路徑
    current_dir = os.path.dirname(os.path.abspath(__file__))
    project_root = os.path.dirname(os.path.dirname(current_dir))
    tests_dir = os.path.join(project_root, 'tests')
    src_dir = os.path.join(project_root, 'src')
    
    # 確保 src 目錄在 sys.path 中
    if src_dir not in sys.path:
        sys.path.insert(0, src_dir)
        
    # 自動尋找與載入 tests 目錄底下的測試
    loader = unittest.TestLoader()
    suite = loader.discover(tests_dir)
    
    # 執行測試並回傳結束代碼
    runner = unittest.TextTestRunner(verbosity=2)
    result = runner.run(suite)
    sys.exit(0 if result.wasSuccessful() else 1)

if __name__ == '__main__':
    run_tests()
