# tools/translate_tool_main.py (수정 후)

import tkinter as tk
import os
import logging
from tkinter import ttk
import sys # sys 모듈 추가

# --- 경로 문제 해결을 위한 코드 ---
# 현재 파일(translate_tool_main.py)의 위치를 기준으로 프로젝트 루트 디렉토리를 계산합니다.
# tools/translate_tool_main.py -> ../ (한 단계 위)
project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# 프로젝트 루트를 Python 경로에 추가합니다.
if project_root not in sys.path:
    sys.path.append(project_root)
# ---------------------------------


# --- 분리된 도구 모듈들 임포트 (경로에서 'translate' 제거) ---
from tools.db_compare_tool import DBCompareTool
from tools.translation_db_tool import TranslationDBTool
from tools.translation_apply_tool import TranslationApplyTool
from tools.string_sync_manager import StringSyncManager
from tools.excel_split_tool import ExcelSplitTool
from tools.word_replacement_tool import WordReplacementTool
from tools.translation_request_extractor import TranslationRequestExtractor
from tools.translation_workflow_tool import TranslationWorkflowTool


class TranslationAutomationTool(tk.Frame):
    def __init__(self, root):
        """
        메인 애플리케이션 창을 초기화하고 모든 도구 탭을 구성합니다.
        """
        super().__init__(root)
        self.root = root
        self.root.title("번역 자동화 툴")
        self.root.geometry("1400x800")
        
        logging.basicConfig(
            filename='translation_tool.log',
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            filemode='w'
        )

        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill="both", expand=True, padx=10, pady=10)
        
        self.workflow_frame = ttk.Frame(self.notebook)
        self.db_compare_frame = ttk.Frame(self.notebook)
        self.translation_db_frame = ttk.Frame(self.notebook)
        self.translation_apply_frame = ttk.Frame(self.notebook)        
        self.translation_request_frame = ttk.Frame(self.notebook)
        self.string_sync_frame = ttk.Frame(self.notebook)
        self.excel_split_frame = ttk.Frame(self.notebook)
        self.word_replacement_frame = ttk.Frame(self.notebook)
    
        self.notebook.add(self.workflow_frame, text="통합 번역 워크플로우")                
        self.notebook.add(self.translation_request_frame, text="번역 요청 추출")
        self.notebook.add(self.translation_db_frame, text="번역 DB 구축")
        self.notebook.add(self.translation_apply_frame, text="번역 적용")
        self.notebook.add(self.db_compare_frame, text="DB 비교 추출")
        self.notebook.add(self.string_sync_frame, text="STRING 동기화")
        self.notebook.add(self.excel_split_frame, text="엑셀 시트 분리")
        self.notebook.add(self.word_replacement_frame, text="단어 치환")
        
        self.excluded_files = self.load_excluded_files()
        
        self.setup_tabs()

    def load_excluded_files(self):
        """제외 파일 목록 로드"""
        try:
            excluded_file_path = os.path.join(project_root, "제외 파일 목록.txt")
            with open(excluded_file_path, "r", encoding="utf-8") as f:
                return [line.strip() for line in f.readlines() if line.strip()]
        except Exception:
            return []

    def setup_tabs(self):
        """
        각 탭 프레임에 해당하는 도구 모듈을 로드하고 배치합니다.
        """
        # [추가] 워크플로우 탭 설정
        workflow_tool = TranslationWorkflowTool(self.workflow_frame, self.root)
        workflow_tool.pack(fill="both", expand=True)
        
        # DB 비교 추출 탭
        db_compare_tool = DBCompareTool(self.db_compare_frame)
        db_compare_tool.pack(fill="both", expand=True)

        # 번역 DB 구축 탭
        translation_db_tool = TranslationDBTool(self.translation_db_frame, self.excluded_files)
        translation_db_tool.pack(fill="both", expand=True)
        
        # 번역 적용 탭
        translation_apply_tool = TranslationApplyTool(self.translation_apply_frame, self.excluded_files)
        translation_apply_tool.pack(fill="both", expand=True)

        # STRING 동기화 탭
        try:
            string_sync_manager = StringSyncManager(self.string_sync_frame, self.root)
            string_sync_manager.pack(fill="both", expand=True)
        except ImportError as e:
            ttk.Label(self.string_sync_frame, text=f"STRING 동기화 모듈 로드 오류:\n{e}").pack(pady=20)

        # 엑셀 시트 분리 탭
        excel_split_tool = ExcelSplitTool(self.excel_split_frame)
        excel_split_tool.pack(fill="both", expand=True)
        
        # 단어 치환 탭
        word_replacement_tool = WordReplacementTool(self.word_replacement_frame)
        word_replacement_tool.pack(fill="both", expand=True)
        
        # 번역 요청 추출 탭
        try:
            translation_request_extractor = TranslationRequestExtractor(self.translation_request_frame)
            translation_request_extractor.pack(fill="both", expand=True)
        except ImportError as e:
            ttk.Label(self.translation_request_frame, text=f"번역 요청 추출 모듈 로드 오류:\n{e}").pack(pady=20)


if __name__ == "__main__":
    root = tk.Tk()
    app = TranslationAutomationTool(root)
    root.mainloop()