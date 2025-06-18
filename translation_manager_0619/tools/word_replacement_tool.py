# tools/word_replacement_tool.py (수정 후)

import tkinter as tk
from tkinter import messagebox, ttk
import os
import sys # sys 모듈 추가

# --- 경로 문제 해결을 위한 코드 ---
project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if project_root not in sys.path:
    sys.path.append(project_root)
# ---------------------------------

# 'tools.translate' 경로에서 'tools'로 변경
from tools.word_replacement_manager import WordReplacementManager

class WordReplacementTool(tk.Frame):
    # ... (이하 코드는 변경 없음) ...
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent

        # WordReplacementManager가 UI와 로직을 모두 포함하므로,
        # 이를 직접 생성하여 사용합니다.
        # WordReplacementManager의 __init__ 인자를 (self, parent)로 전달하여
        # 올바른 부모-자식 관계를 형성합니다.
        try:
            self.manager = WordReplacementManager(self, parent)
            self.manager.pack(fill="both", expand=True)
        except Exception as e:
            # 오류 발생 시 메시지 표시
            ttk.Label(self, text=f"단어 치환 모듈 로드 오류:\n{e}").pack(pady=20)
            print(f"Error loading WordReplacementManager: {e}")