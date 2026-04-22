import logging
import threading
import urllib.request
import re
from PyQt6.QtCore import Qt, QTimer, pyqtSignal
from PyQt6.QtWidgets import QWidget, QVBoxLayout, QTextBrowser
from artale_agent.utils import platform_font_family

logger = logging.getLogger(__name__)

class AwesomeTabContent(QWidget):
    content_ready = pyqtSignal(str)

    def __init__(self):
        super().__init__()
        self._loaded = False
        self.init_ui()
        self.content_ready.connect(self._update_ui)

    def init_ui(self):
        self.layout = QVBoxLayout(self)
        self.browser = QTextBrowser()
        self.browser.setOpenExternalLinks(True)
        self.browser.setHtml("<p style='color: #888; text-align: center; margin-top: 50px;'>正在載入工具清單...</p>")
        self.browser.setStyleSheet("border: none; background: transparent;")
        self.layout.addWidget(self.browser)

    def trigger_load(self):
        if not self._loaded:
            self._fetch_content()

    def _fetch_content(self):
        url = "https://raw.githubusercontent.com/vongola12324/awesome-artale/refs/heads/main/README.md"
        def _run():
            try:
                req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
                with urllib.request.urlopen(req, timeout=10) as response:
                    content = response.read().decode('utf-8')
                    html = self._md_to_html(content)
                    self.content_ready.emit(html)
            except Exception as e:
                logger.error(f"[Awesome] Fetch failed: {e}")
                error_html = f"<p style='color: #ff6b6b; text-align: center;'>載入失敗: {e}<br>請檢查網路連線後重試。</p>"
                self.content_ready.emit(error_html)
        
        threading.Thread(target=_run, daemon=True).start()

    def _update_ui(self, html):
        self.browser.setHtml(html)
        self._loaded = True

    def _md_to_html(self, md_text):
        import markdown2
        
        # 移除不支援的 HTML 標籤
        md_text = re.sub(r'</?details>', '', md_text)
        md_text = re.sub(r'</?summary>', '', md_text)
        
        # 處理 GitHub 特有的 [!TIP]
        md_text = re.sub(r'> \[!TIP\]', '> 💡 提示：', md_text)

        # 使用 markdown2 轉換，開啟一些常見擴展功能
        body_html = markdown2.markdown(md_text, extras=[
            "fenced-code-blocks", 
            "tables", 
            "task_list", 
            "break-on-newline",
            "header-ids"
        ])
        
        style = f"""
        <style>
            body {{ color: #e0e0e0; font-family: {platform_font_family()}; line-height: 1.6; font-size: 13px; padding: 10px; }}
            a {{ color: #4dabf7; text-decoration: none; }}
            a:hover {{ text-decoration: underline; }}
            h1, h2, h3 {{ color: #ffd700; margin-top: 20px; border-bottom: 1px solid #333; padding-bottom: 5px; }}
            li {{ margin-bottom: 8px; }}
            ul {{ padding-left: 20px; }}
            blockquote {{ border-left: 4px solid #ffd700; background: #1e1e1e; padding: 10px; color: #aaa; margin: 15px 0; }}
            code {{ background: #222; padding: 2px 4px; border-radius: 3px; font-family: Consolas, monospace; }}
            pre {{ background: #222; padding: 10px; border-radius: 5px; }}
            table {{ border-collapse: collapse; width: 100%; margin: 15px 0; }}
            th, td {{ border: 1px solid #444; padding: 8px; text-align: left; }}
            th {{ background: #2a2a2a; color: #ffd700; }}
        </style>
        """
        return f"<html><head>{style}</head><body>{body_html}</body></html>"
