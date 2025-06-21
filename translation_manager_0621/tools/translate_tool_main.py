# tools/translate_tool_main.py (개선된 도구들로 업데이트)

import tkinter as tk
import os
import logging
from tkinter import ttk
import sys # sys 모듈 추가
import ctypes  # [신규] Windows API 사용을 위한 모듈
from tkinter import font as tkfont # [신규] 폰트 모듈

def resource_path(relative_path):
    """ 개발 환경과 PyInstaller 배포 환경 모두에서 리소스 경로를 올바르게 가져옵니다. """
    try:
        # PyInstaller는 임시 폴더를 만들고 _MEIPASS에 경로를 저장합니다.
        base_path = sys._MEIPASS
    except Exception:
        # 개발 환경에서는 현재 파일의 상위 폴더를 기준으로 합니다.
        base_path = os.path.abspath(os.path.dirname(__file__))
    return os.path.join(base_path, relative_path)

def load_custom_font(font_path):
    """ 주어진 경로의 폰트 파일을 시스템에 로드합니다. (Windows 전용) """
    try:
        # AddFontResourceW 함수는 유니코드 파일 경로를 처리합니다.
        result = ctypes.windll.gdi32.AddFontResourceW(font_path)
        if result == 0:
            logging.error(f"폰트 로드 실패: {font_path}")
        else:
            logging.info(f"폰트 로드 성공 ({result}개): {font_path}")
    except Exception as e:
        logging.error(f"폰트 로드 중 예외 발생: {e}")
        
# --- 경로 문제 해결을 위한 코드 ---
# 현재 파일(translate_tool_main.py)의 위치를 기준으로 프로젝트 루트 디렉토리를 계산합니다.
# tools/translate_tool_main.py -> ../ (한 단계 위)
project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# 프로젝트 루트를 Python 경로에 추가합니다.
if project_root not in sys.path:
    sys.path.append(project_root)
# ---------------------------------

# --- 분리된 도구 모듈들 임포트 ---
from tools.db_compare_tool import DBCompareTool
from tools.enhanced_integrated_translation_tool import EnhancedIntegratedTranslationTool
from tools.enhanced_translation_apply_tool import EnhancedTranslationApplyTool
from tools.string_sync_manager import StringSyncManager
from tools.excel_split_tool import ExcelSplitTool
from tools.word_replacement_tool import WordReplacementTool
from tools.translation_request_extractor import TranslationRequestExtractor
from tools.translation_workflow_tool import TranslationWorkflowTool
from tools.advanced_excel_diff_tool import AdvancedExcelDiffTool
from tools.resolution_manager_tool import ResolutionManagerTool
        
class TranslationAutomationTool(tk.Frame):
    def __init__(self, root):
        """
        메인 애플리케이션 창을 초기화하고 모든 도구 탭을 구성합니다.
        """
        super().__init__(root)
        self.root = root
        
        # --- [신규] 프로그램 시작 시 커스텀 폰트 로드 ---
        font_name = "GmarketSansTTFMedium.ttf"  # 사용할 폰트 파일 이름
        custom_font_path = resource_path(os.path.join('fonts', font_name))
        if os.path.exists(custom_font_path):
            load_custom_font(custom_font_path)
        else:
            logging.warning(f"커스텀 폰트 파일을 찾을 수 없습니다: {custom_font_path}")
        # ---------------------------------------------

        self.root.title("번역 자동화 툴 (Enhanced)")  # [수정] 타이틀에 Enhanced 표시
        self.root.geometry("1400x800")
        
        logging.basicConfig(
            filename='translation_tool.log',
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            filemode='w'
        )

        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill="both", expand=True, padx=10, pady=10)
        
        # 탭 프레임들 생성
        self.workflow_frame = ttk.Frame(self.notebook)
        self.db_compare_frame = ttk.Frame(self.notebook)
        self.enhanced_integrated_frame = ttk.Frame(self.notebook)  # [신규] 개선된 통합 도구
        self.enhanced_apply_frame = ttk.Frame(self.notebook)       # [신규] 개선된 적용 도구
        self.translation_request_frame = ttk.Frame(self.notebook)
        self.string_sync_frame = ttk.Frame(self.notebook)
        self.excel_split_frame = ttk.Frame(self.notebook)
        self.word_replacement_frame = ttk.Frame(self.notebook)
        self.resolution_manager_frame = ttk.Frame(self.notebook)
    
        # 탭 추가 (순서 조정: 개선된 도구들을 상단에 배치)        
        self.notebook.add(self.translation_request_frame, text="번역 요청 추출")
        self.notebook.add(self.enhanced_apply_frame, text="🎯 번역 적용")
        self.notebook.add(self.resolution_manager_frame, text="⚙️ 해결 내역 관리")
        self.notebook.add(self.excel_split_frame, text="엑셀 시트 분리")        
        self.notebook.add(self.enhanced_integrated_frame, text="🚀엑셀 비교")        
        #self.notebook.add(advanced_excel_diff_tab, text="[기존] 고급 엑셀 비교")
        #self.notebook.add(self.db_compare_frame, text="DB 비교 추출")
        #self.notebook.add(self.workflow_frame, text="통합 번역 워크플로우")                
        #self.notebook.add(self.string_sync_frame, text="STRING 동기화")        
        #self.notebook.add(self.word_replacement_frame, text="단어 치환")
        
        self.excluded_files = self.load_excluded_files()
        self.shared_manager = EnhancedTranslationApplyTool(self.enhanced_apply_frame, self.excluded_files).translation_apply_manager
        
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
        # [신규] 개선된 번역 적용 탭
        try:
            enhanced_apply_tool = EnhancedTranslationApplyTool(self.enhanced_apply_frame, self.excluded_files)
            enhanced_apply_tool.pack(fill="both", expand=True)
            
            # 성공 메시지 로그
            logging.info("Enhanced Translation Apply Tool loaded successfully")
            
        except ImportError as e:
            error_message = f"개선된 번역 적용 도구 로드 오류:\n{e}\n\n필요한 파일:\n- enhanced_translation_apply_tool.py\n- enhanced_translation_apply_manager.py"
            ttk.Label(self.enhanced_apply_frame, text=error_message, foreground="red").pack(pady=20)
            logging.error(f"Enhanced Translation Apply Tool import error: {e}")
        except Exception as e:
            error_message = f"개선된 번역 적용 도구 초기화 오류:\n{e}"
            ttk.Label(self.enhanced_apply_frame, text=error_message, foreground="red").pack(pady=20)
            logging.error(f"Enhanced Translation Apply Tool initialization error: {e}")
        
        # [신규] 해결 내역 관리 탭 설정
        self.shared_manager = enhanced_apply_tool.translation_apply_manager        
        try:
            # 위에서 생성한 shared_manager를 전달
            resolution_tool = ResolutionManagerTool(self.resolution_manager_frame, self.shared_manager)
            resolution_tool.pack(fill="both", expand=True)
        except Exception as e:
            error_message = f"해결 내역 관리 도구 로드 오류:\n{e}"
            ttk.Label(self.resolution_manager_frame, text=error_message, foreground="red").pack(pady=20)
   
        
        # [신규] 개선된 통합 번역 도구 탭
        try:
            enhanced_integrated_tool = EnhancedIntegratedTranslationTool(self.enhanced_integrated_frame, self.excluded_files)
            enhanced_integrated_tool.pack(fill="both", expand=True)
            
            # 성공 메시지 로그
            logging.info("Enhanced Integrated Translation Tool loaded successfully")
            
        except ImportError as e:
            error_message = f"개선된 통합 번역 도구 로드 오류:\n{e}\n\n필요한 파일:\n- enhanced_integrated_translation_tool.py\n- enhanced_integrated_translation_manager.py"
            ttk.Label(self.enhanced_integrated_frame, text=error_message, foreground="red").pack(pady=20)
            logging.error(f"Enhanced Integrated Translation Tool import error: {e}")
        except Exception as e:
            error_message = f"개선된 통합 번역 도구 초기화 오류:\n{e}"
            ttk.Label(self.enhanced_integrated_frame, text=error_message, foreground="red").pack(pady=20)
            logging.error(f"Enhanced Integrated Translation Tool initialization error: {e}")
            
        
        # [기존] 워크플로우 탭 설정
        try:
            workflow_tool = TranslationWorkflowTool(self.workflow_frame, self.root)
            workflow_tool.pack(fill="both", expand=True)
        except ImportError as e:
            ttk.Label(self.workflow_frame, text=f"워크플로우 도구 로드 오류:\n{e}").pack(pady=20)
        
        # [기존] DB 비교 추출 탭
        try:
            db_compare_tool = DBCompareTool(self.db_compare_frame)
            db_compare_tool.pack(fill="both", expand=True)
        except ImportError as e:
            ttk.Label(self.db_compare_frame, text=f"DB 비교 도구 로드 오류:\n{e}").pack(pady=20)

        # [기존] STRING 동기화 탭
        try:
            string_sync_manager = StringSyncManager(self.string_sync_frame, self.root)
            string_sync_manager.pack(fill="both", expand=True)
        except ImportError as e:
            ttk.Label(self.string_sync_frame, text=f"STRING 동기화 모듈 로드 오류:\n{e}").pack(pady=20)

        # [기존] 엑셀 시트 분리 탭
        try:
            excel_split_tool = ExcelSplitTool(self.excel_split_frame)
            excel_split_tool.pack(fill="both", expand=True)
        except ImportError as e:
            ttk.Label(self.excel_split_frame, text=f"엑셀 분리 도구 로드 오류:\n{e}").pack(pady=20)
        
        # [기존] 단어 치환 탭
        try:
            word_replacement_tool = WordReplacementTool(self.word_replacement_frame)
            word_replacement_tool.pack(fill="both", expand=True)
        except ImportError as e:
            ttk.Label(self.word_replacement_frame, text=f"단어 치환 도구 로드 오류:\n{e}").pack(pady=20)
        
        # [기존] 번역 요청 추출 탭
        try:
            translation_request_extractor = TranslationRequestExtractor(self.translation_request_frame)
            translation_request_extractor.pack(fill="both", expand=True)
        except ImportError as e:
            ttk.Label(self.translation_request_frame, text=f"번역 요청 추출 모듈 로드 오류:\n{e}").pack(pady=20)
        
        # # [기존] 고급 엑셀 비교 탭
        # try:
        #     advanced_excel_diff_tab = AdvancedExcelDiffTool(self.notebook)
        # except ImportError as e:
        #     advanced_diff_frame = ttk.Frame(self.notebook)            
        #     self.notebook.add(advanced_diff_frame, text="[기존] 고급 엑셀 비교")
        #     ttk.Label(advanced_diff_frame, text=f"고급 엑셀 비교 도구 로드 오류:\n{e}").pack(pady=20)

if __name__ == "__main__":
    root = tk.Tk()
    app = TranslationAutomationTool(root)
    root.mainloop()