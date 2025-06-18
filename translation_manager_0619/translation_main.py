import tkinter as tk
import sys
import os

# --- 프로젝트 환경 설정 ---
# 현재 파일의 위치를 기반으로 프로젝트의 루트 폴더 경로를 찾습니다.
project_root = os.path.dirname(os.path.abspath(__file__))
# 파이썬이 다른 폴더에 있는 모듈(예: tools, ui, utils)을 찾을 수 있도록 경로를 추가합니다.
if project_root not in sys.path:
    sys.path.append(project_root)

# --- 메인 UI 임포트 ---
# 실제 클래스 이름인 'TranslationAutomationTool'을 임포트합니다.
from tools.translate_tool_main import TranslationAutomationTool

# TranslationAutomationTool 클래스는 스스로 창을 설정하고 전체 UI를 구성합니다.
# 따라서 별도의 Application 클래스 없이 바로 사용합니다.

if __name__ == "__main__":
    # 1. Tkinter 루트 윈도우를 생성합니다.
    root = tk.Tk()
    
    # 2. TranslationAutomationTool 클래스의 인스턴스를 생성하고,
    #    root 윈도우를 인자로 전달합니다.
    #    이 클래스가 창 제목 설정, UI 구성 등을 모두 처리합니다.
    app = TranslationAutomationTool(root)
    
    # 3. Tkinter 이벤트 루프를 시작하여 프로그램을 실행합니다.
    root.mainloop()