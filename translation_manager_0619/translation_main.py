# translation_main.py (수정 후)

import tkinter as tk
import sys
import os

# --- 프로젝트 환경 설정 ---
project_root = os.path.dirname(os.path.abspath(__file__))
if project_root not in sys.path:
    sys.path.append(project_root)

# --- 메인 UI 임포트 ---
# 'from tools.translate_tool_main'으로 경로 변경
from tools.translate_tool_main import TranslationAutomationTool

if __name__ == "__main__":
    root = tk.Tk()
    app = TranslationAutomationTool(root)
    root.mainloop()