
import re
import os

path = r'z:\Github\artale-agent\overlay.py'
with open(path, 'r', encoding='utf-8') as f:
    content = f.read()

# Target area is the fuzzy search logic
# We'll replace it with a more robust fallback system
# Looking for everything between '# Start capture loop' and 'on_frame_arrived(frame, capture_control):'

pattern = r'(# Start capture loop\s+try:)(.*?)(\s+@capture.event)'
replacement = r"""\1
            # 1. Try common window names in sequence to find Artale
            capture = None
            for name in ["Artale", "artale", "ARTALE", "msw.exe"]:
                try:
                    from windows_capture import WindowsCapture
                    capture = WindowsCapture(
                        window_name=name,
                        cursor_capture=False,
                        draw_border=False,
                        minimum_update_interval=1000
                    )
                    break
                except:
                    continue

            if not capture:
                print(f"[ExpTracker] Error: Artale window not found. Using default 'Artale' and hoping for the best.")
                capture = WindowsCapture(window_name="Artale", minimum_update_interval=1000)
\3"""

new_content = re.sub(pattern, replacement, content, flags=re.DOTALL)

with open(path, 'w', encoding='utf-8') as f:
    f.write(new_content)

print("Successfully patched overlay.py")
