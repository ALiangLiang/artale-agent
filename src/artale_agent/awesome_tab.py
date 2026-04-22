import logging
import threading
import urllib.request
import re
from PyQt6.QtCore import Qt, QTimer
from PyQt6.QtWidgets import QWidget, QVBoxLayout, QTextBrowser
from artale_agent.utils import platform_font_family

logger = logging.getLogger(__name__)

class AwesomeTabContent(QWidget):
    def __init__(self):
        super().__init__()
        self._loaded = False
        self.init_ui()

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
                    QTimer.singleShot(0, lambda: self._update_ui(html))
            except Exception as e:
                logger.error(f"[Awesome] Fetch failed: {e}")
                QTimer.singleShot(0, lambda: self._update_ui(f"<p style='color: #ff6b6b; text-align: center;'>載入失敗: {e}<br>請檢查網路連線後重試。</p>"))
        
        threading.Thread(target=_run, daemon=True).start()

    def _update_ui(self, html):
        self.browser.setHtml(html)
        self._loaded = True

    def _md_to_html(self, md_text):
        # 移除不支援的 HTML 標籤
        html = re.sub(r'</?details>', '', md_text)
        html = re.sub(r'</?summary>', '', html)
        
        # 處理 GitHub 特有的 [!TIP]
        html = re.sub(r'> \[!TIP\]', '> 💡 提示：', html)

        # Headers
        html = re.sub(r'^### (.*)$', r'<h3 style="color: #ffd700; border-bottom: 1px solid #333; padding-bottom: 5px;">\1</h3>', html, flags=re.M)
        html = re.sub(r'^## (.*)$', r'<h2 style="color: #ffd700; border-bottom: 1px solid #444;">\1</h2>', html, flags=re.M)
        
        # Bold
        html = re.sub(r'\*\*(.*?)\*\*', r'<b>\1</b>', html)
        
        # Links
        html = re.sub(r'\[(.*?)\]\((.*?)\)', r'<a href="\2" style="color: #4dabf7; text-decoration: none;">\1</a>', html)
        
        # Lists (簡單處理)
        html = re.sub(r'^- (.*)$', r'<li>\1</li>', html, flags=re.M)
        
        # Blockquotes
        html = re.sub(r'^> (.*)$', r'<blockquote style="border-left: 4px solid #ffd700; background: #1e1e1e; padding: 10px; color: #aaa; margin: 10px 0;">\1</blockquote>', html, flags=re.M)
        
        style = f"""
        <style>
            body {{ color: #e0e0e0; font-family: {platform_font_family()}; line-height: 1.5; font-size: 13px; }}
            a {{ color: #4dabf7; }}
            li {{ margin-bottom: 5px; }}
            ul {{ padding-left: 20px; }}
            h3 {{ margin-top: 20px; }}
        </style>
        """
        # 將換行轉為 <br>，但避免在已經有標籤的地方重複
        html = html.replace('\n', '<br>')
        return f"<html><head>{style}</head><body>{html}</body></html>"
