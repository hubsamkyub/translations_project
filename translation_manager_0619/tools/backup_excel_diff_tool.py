# tools/excel_diff_tool.py 파일 생성
import tkinter as tk
import os
import sys

# 현재 파일 경로 추가
current_dir = os.path.dirname(os.path.abspath(__file__))
# 상위 디렉토리 경로 (기본 프로젝트 루트)
root_dir = os.path.dirname(current_dir)

# 루트 디렉토리를 모듈 검색 경로에 추가
if root_dir not in sys.path:
    sys.path.append(root_dir)

# 원본 Excel_Data_Diff_Tool.py에서 ExcelDiffTool 클래스와 관련 코드 임포트
from tools.Excel_Data_Diff_Tool import ExcelDiffTool, LoadingPopup, ImprovedFileList

# 외부에서 쉽게 호출할 수 있는 함수 추가
def open_excel_diff_tool(parent, data_path=None):
    """Excel Diff 툴을 새 창에서 열기"""
    diff_window = tk.Toplevel(parent)
    diff_app = ExcelDiffTool(diff_window)
    
    # 데이터 경로가 제공되면 설정
    if data_path and os.path.exists(data_path):
        diff_app.source_path.set(data_path)
        diff_app.target_path.set(data_path)
        diff_app.update_file_list(data_path)
    
    return diff_app