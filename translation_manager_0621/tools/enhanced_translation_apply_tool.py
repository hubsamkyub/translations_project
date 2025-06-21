import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import threading
import time
import sys
import pandas as pd
from datetime import datetime
from collections import defaultdict

# --- 경로 문제 해결을 위한 코드 ---
project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if project_root not in sys.path:
    sys.path.append(project_root)
# ---------------------------------

from ui.common_components import ScrollableCheckList, LoadingPopup
from tools.enhanced_translation_apply_manager import EnhancedTranslationApplyManager
import openpyxl

class EnhancedTranslationApplyTool(tk.Frame):
    ### __init__: [변경 없음]
    def __init__(self, parent, excluded_files):
        super().__init__(parent)
        self.parent = parent
        self.translation_apply_manager = EnhancedTranslationApplyManager(self)
        
        # --- UI 변수 선언 ---
        # 소스 선택 관련
        self.translation_db_var = tk.StringVar()
        self.excel_source_path_var = tk.StringVar()
        self.original_folder_var = tk.StringVar()
        
        # 다중 시트 선택 관련
        self.selected_sheets_display_var = tk.StringVar(value="선택된 시트 없음")
        self.selected_sheets = []

        # [수정] 언어 옵션 (KR, CN, TW만)
        self.available_languages = ["KR", "CN", "TW"]
        self.apply_lang_vars = {}
        
        # 기본 번역 적용 옵션
        self.record_date_var = tk.BooleanVar(value=True)
        self.kr_match_check_var = tk.BooleanVar(value=True)
        self.kr_mismatch_delete_var = tk.BooleanVar(value=False)
        self.apply_smart_lookup_var = tk.BooleanVar(value=True)
        
        # [개선] 특수 컬럼 필터링 옵션 - 직접 입력 방식
        self.use_filtered_data_var = tk.BooleanVar(value=False)
        self.special_column_input_var = tk.StringVar(value="#번역요청")  # 기본값 설정
        self.special_condition_var = tk.StringVar()
        self.filter_status_var = tk.StringVar(value="필터링 데이터 없음")
        
        # 조건부 적용 옵션
        self.apply_on_new_var = tk.BooleanVar(value=True)
        self.apply_on_change_var = tk.BooleanVar(value=True)
        self.apply_on_transferred_var = tk.BooleanVar(value=False)
    
        # --- 내부 데이터 ---
        self.view_data_button = None
        self.view_filtered_button = None  
        self.original_files = []
        self.excluded_files = excluded_files
        self.cached_excel_path = None
        self.cached_sheet_names = [] # 이제 DB에서 가져온 시트 목록을 저장
        self.detected_special_columns = {}  
        
        self.setup_ui()

    ### setup_ui: [변경]
    def setup_ui(self):
        """[변경] DB 캐시 구축 시 진행률을 표시할 Progressbar를 추가합니다."""

        # 메인 좌우 분할 프레임
        main_paned = ttk.PanedWindow(self, orient="horizontal")
        main_paned.pack(fill="both", expand=True, padx=5, pady=5)

        # --- 좌측 패널: 설정 영역 ---
        left_frame = ttk.Frame(main_paned, width=800)
        left_canvas = tk.Canvas(left_frame)
        left_scrollbar = ttk.Scrollbar(left_frame, orient="vertical", command=left_canvas.yview)
        left_scrollable_frame = ttk.Frame(left_canvas)
        
        left_scrollable_frame.bind(
            "<Configure>",
            lambda e: left_canvas.configure(scrollregion=left_canvas.bbox("all"))
        )
        
        left_canvas.create_window((0, 0), window=left_scrollable_frame, anchor="nw")
        left_canvas.configure(yscrollcommand=left_scrollbar.set)

        # --- 우측 패널: 파일 및 실행 영역 ---
        right_frame = ttk.Frame(main_paned, width=600)
        right_canvas = tk.Canvas(right_frame)
        right_scrollbar = ttk.Scrollbar(right_frame, orient="vertical", command=right_canvas.yview)
        right_scrollable_frame = ttk.Frame(right_canvas)
        
        right_scrollable_frame.bind(
            "<Configure>",
            lambda e: right_canvas.configure(scrollregion=right_canvas.bbox("all"))
        )
        
        right_canvas.create_window((0, 0), window=right_scrollable_frame, anchor="nw")
        right_canvas.configure(yscrollcommand=right_scrollbar.set)

        # 패널 추가
        main_paned.add(left_frame, weight=4)  # 좌측 패널에 가중치 2 부여
        main_paned.add(right_frame, weight=3) # 우측 패널에 가중치 1 부여

        # ==================== 좌측 패널 구성 ====================

        # --- 1. 소스 선택 프레임 ---
        source_selection_frame = ttk.LabelFrame(left_scrollable_frame, text="🔧 번역 데이터 소스 선택")
        source_selection_frame.pack(fill="x", padx=5, pady=5)

        # DB 선택
        db_frame = ttk.Frame(source_selection_frame)
        db_frame.pack(fill="x", padx=5, pady=3)
        ttk.Label(db_frame, text="번역 DB:").grid(row=0, column=0, padx=5, pady=2, sticky="w")
        db_entry = ttk.Entry(db_frame, textvariable=self.translation_db_var)
        db_entry.grid(row=0, column=1, padx=5, pady=2, sticky="ew")
        ttk.Button(db_frame, text="찾아보기", command=self.select_translation_db_file).grid(row=0, column=2, padx=5, pady=2)
        db_frame.columnconfigure(1, weight=1)

        # 구분선
        ttk.Separator(source_selection_frame, orient="horizontal").pack(fill="x", padx=5, pady=5)

        # 엑셀 선택
        excel_frame = ttk.Frame(source_selection_frame)
        excel_frame.pack(fill="x", padx=5, pady=3)
        ttk.Label(excel_frame, text="엑셀 파일:").grid(row=0, column=0, padx=5, pady=2, sticky="w")
        excel_entry = ttk.Entry(excel_frame, textvariable=self.excel_source_path_var, state="readonly") # 직접 수정 방지
        excel_entry.grid(row=0, column=1, padx=5, pady=2, sticky="ew")
        ttk.Button(excel_frame, text="찾아보기", command=self.select_excel_source_file).grid(row=0, column=2, padx=5, pady=2)
        excel_frame.columnconfigure(1, weight=1)

        # 시트 선택
        sheet_frame = ttk.Frame(source_selection_frame)
        sheet_frame.pack(fill="x", padx=5, pady=3)
        ttk.Label(sheet_frame, text="시트 선택:").grid(row=0, column=0, padx=5, pady=2, sticky="w")
        selected_sheets_entry = ttk.Entry(sheet_frame, textvariable=self.selected_sheets_display_var, state="readonly")
        selected_sheets_entry.grid(row=0, column=1, padx=5, pady=2, sticky="ew")
        self.sheet_select_button = ttk.Button(sheet_frame, text="시트 선택", command=self.open_sheet_selection_popup, state="disabled")
        self.sheet_select_button.grid(row=0, column=2, padx=5, pady=2)

        # ▼▼▼ '캐시 재구축' 버튼 추가 ▼▼▼
        self.rebuild_cache_button = ttk.Button(sheet_frame, text="캐시 재구축", command=self.force_rebuild_cache, state="disabled")
        self.rebuild_cache_button.grid(row=0, column=3, padx=5, pady=2) 

        sheet_frame.columnconfigure(1, weight=1)
        

        # --- 2. [개선] 특수 컬럼 필터링 설정 ---
        filter_frame = ttk.LabelFrame(left_scrollable_frame, text="🔍 특수 컬럼 필터링 (고급 옵션)")
        filter_frame.pack(fill="x", padx=5, pady=5)
        
        # 필터링 활성화
        filter_enable_frame = ttk.Frame(filter_frame)
        filter_enable_frame.pack(fill="x", padx=5, pady=3)
        ttk.Checkbutton(filter_enable_frame, text="특수 컬럼 필터링된 데이터만 적용", 
                        variable=self.use_filtered_data_var,
                        command=self.toggle_filter_options).pack(side="left")

        # 필터링 설정 프레임
        self.filter_config_frame = ttk.Frame(filter_frame)
        self.filter_config_frame.pack(fill="x", padx=15, pady=5)
        
        # [개선] 컬럼명 직접 입력
        col_frame = ttk.Frame(self.filter_config_frame)
        col_frame.pack(fill="x", pady=2)
        ttk.Label(col_frame, text="컬럼명:").grid(row=0, column=0, padx=5, pady=2, sticky="w")
        ttk.Entry(col_frame, textvariable=self.special_column_input_var, width=25).grid(row=0, column=1, padx=5, pady=2, sticky="ew")
        self.detect_button = ttk.Button(col_frame, text="감지", command=self.detect_special_column, state="disabled") # 초기 비활성화
        self.detect_button.grid(row=0, column=2, padx=5, pady=2)
        col_frame.columnconfigure(1, weight=1)
        
        # 조건값 입력
        condition_frame = ttk.Frame(self.filter_config_frame)
        condition_frame.pack(fill="x", pady=2)
        ttk.Label(condition_frame, text="조건값:").grid(row=0, column=0, padx=5, pady=2, sticky="w")
        ttk.Entry(condition_frame, textvariable=self.special_condition_var, width=25).grid(row=0, column=1, padx=5, pady=2, sticky="ew")
        condition_frame.columnconfigure(1, weight=1)
        
        # 필터링 상태 표시
        status_frame = ttk.Frame(self.filter_config_frame)
        status_frame.pack(fill="x", pady=2)
        ttk.Label(status_frame, text="상태:").grid(row=0, column=0, padx=5, pady=2, sticky="w")
        filter_status_label = ttk.Label(status_frame, textvariable=self.filter_status_var, foreground="blue")
        filter_status_label.grid(row=0, column=1, padx=5, pady=2, sticky="w")
        
        # 초기에는 비활성화
        self.toggle_filter_options()

        # --- 3. 적용 옵션 ---
        options_frame = ttk.LabelFrame(left_scrollable_frame, text="⚙️ 적용 옵션")
        options_frame.pack(fill="x", padx=5, pady=5)
        # (이하 옵션 프레임 UI 구성은 변경 없음)
        self.apply_mode_var = tk.StringVar(value="id")
        self.apply_mode_var.trace_add("write", self.toggle_options_by_mode)
        mode_frame = ttk.Frame(options_frame)
        mode_frame.pack(fill="x", padx=5, pady=2)
        ttk.Label(mode_frame, text="적용 기준:").pack(side="left", padx=5)
        ttk.Radiobutton(mode_frame, text="ID 기반", variable=self.apply_mode_var, value="id").pack(side="left", padx=5)
        ttk.Radiobutton(mode_frame, text="KR 기반", variable=self.apply_mode_var, value="kr").pack(side="left", padx=5)
        self.id_based_options_frame = ttk.Frame(options_frame)
        self.id_based_options_frame.pack(fill="x", padx=15, pady=5)
        self.kr_match_check_var = tk.BooleanVar(value=True)
        self.kr_mismatch_delete_var = tk.BooleanVar(value=False)
        self.kr_overwrite_var = tk.BooleanVar(value=False)
        id_opt1 = ttk.Checkbutton(self.id_based_options_frame, text="KR 일치 검사", variable=self.kr_match_check_var, command=self.toggle_kr_options)
        id_opt1.pack(anchor="w", padx=5, pady=1)
        self.id_mismatch_delete_cb = ttk.Checkbutton(self.id_based_options_frame, text="└ KR 불일치 시 다국어 제거", variable=self.kr_mismatch_delete_var)
        self.id_mismatch_delete_cb.pack(anchor="w", padx=20, pady=1)
        self.id_overwrite_cb = ttk.Checkbutton(self.id_based_options_frame, text="└ 선택 언어 덮어쓰기", variable=self.kr_overwrite_var)
        self.id_overwrite_cb.pack(anchor="w", padx=20, pady=1)
        self.kr_based_options_frame = ttk.Frame(options_frame)
        self.kr_overwrite_on_kr_mode_var = tk.BooleanVar(value=False)
        kr_opt1 = ttk.Checkbutton(self.kr_based_options_frame, text="선택 언어 덮어쓰기", variable=self.kr_overwrite_on_kr_mode_var)
        kr_opt1.pack(anchor="w", padx=5, pady=1)
        lang_frame = ttk.LabelFrame(options_frame, text="적용 언어")
        lang_frame.pack(fill="x", padx=5, pady=5)
        lang_inner_frame = ttk.Frame(lang_frame)
        lang_inner_frame.pack(fill="x", padx=5, pady=2)
        for i, lang in enumerate(self.available_languages):
            var = tk.BooleanVar(value=True if lang in ["CN", "TW"] else False)
            self.apply_lang_vars[lang] = var
            ttk.Checkbutton(lang_inner_frame, text=lang, variable=var).pack(side="left", padx=10)
        conditional_frame = ttk.LabelFrame(options_frame, text="조건부 적용 (#번역요청 컬럼)")
        conditional_frame.pack(fill="x", padx=5, pady=5)
        cond_inner_frame = ttk.Frame(conditional_frame)
        cond_inner_frame.pack(fill="x", padx=5, pady=2)
        ttk.Checkbutton(cond_inner_frame, text="신규", variable=self.apply_on_new_var).pack(side="left", padx=5)
        ttk.Checkbutton(cond_inner_frame, text="change", variable=self.apply_on_change_var).pack(side="left", padx=5)
        ttk.Checkbutton(cond_inner_frame, text="전달", variable=self.apply_on_transferred_var).pack(side="left", padx=5)
        other_frame = ttk.Frame(options_frame)
        other_frame.pack(fill="x", padx=5, pady=2)
        ttk.Checkbutton(other_frame, text="번역 적용 표시", variable=self.record_date_var).pack(anchor="w", padx=5)


        # ==================== 우측 패널 구성 ====================

        # --- 1. 원본 파일 관련 ---
        original_files_frame = ttk.LabelFrame(right_scrollable_frame, text="📁 번역을 적용할 원본 파일")
        original_files_frame.pack(fill="x", padx=5, pady=5)
        
        # 원본 폴더 선택
        folder_frame = ttk.Frame(original_files_frame)
        folder_frame.pack(fill="x", padx=5, pady=3)
        ttk.Label(folder_frame, text="원본 폴더:").grid(row=0, column=0, padx=5, pady=2, sticky="w")
        ttk.Entry(folder_frame, textvariable=self.original_folder_var).grid(row=0, column=1, padx=5, pady=2, sticky="ew")
        ttk.Button(folder_frame, text="찾아보기", command=self.select_original_folder).grid(row=0, column=2, padx=5, pady=2)
        ttk.Button(folder_frame, text="파일 검색", command=self.search_original_files).grid(row=0, column=3, padx=5, pady=2)
        folder_frame.columnconfigure(1, weight=1)
        
        # 파일 목록
        files_list_frame = ttk.LabelFrame(right_scrollable_frame, text="원본 파일 목록")
        files_list_frame.pack(fill="both", expand=True, padx=5, pady=5)
        self.original_files_list = ScrollableCheckList(files_list_frame, height=200)
        self.original_files_list.pack(fill="both", expand=True, padx=5, pady=5)

        # --- 2. 액션 버튼들 ---
        action_frame = ttk.Frame(right_scrollable_frame)
        action_frame.pack(fill="x", padx=5, pady=5)
        
        # 상단 버튼 행 (데이터 관련)
        data_button_frame = ttk.Frame(action_frame)
        data_button_frame.pack(fill="x", pady=2)
        
        self.load_data_button = ttk.Button(data_button_frame, text="번역 데이터 로드", command=self.load_translation_data, state="disabled") # 초기 비활성화
        self.load_data_button.pack(side="left", padx=5)
        self.view_data_button = ttk.Button(data_button_frame, text="전체 데이터 보기", command=self.show_loaded_data_viewer, state="disabled")
        self.view_data_button.pack(side="left", padx=5)
        self.view_filtered_button = ttk.Button(data_button_frame, text="필터링 데이터 보기", command=self.show_filtered_data_viewer, state="disabled")
        self.view_filtered_button.pack(side="left", padx=5)
        
        # ▼▼▼ [요청 3] '덮어쓴 데이터 보기' 버튼 추가 ▼▼▼
        self.view_overwritten_button = ttk.Button(data_button_frame, text="덮어쓴 데이터 보기", command=self.show_overwritten_data_viewer, state="disabled")
        self.view_overwritten_button.pack(side="left", padx=5)
        
        # 하단 버튼 행 (실행)
        exec_button_frame = ttk.Frame(action_frame)
        exec_button_frame.pack(fill="x", pady=2)
        
        self.apply_button = ttk.Button(exec_button_frame, text="🚀 번역 적용", command=self.apply_translation, state="disabled") # 초기 비활성화
        self.apply_button.pack(side="right", padx=5)

        # --- 3. 로그 영역 ---
        log_frame = ttk.LabelFrame(right_scrollable_frame, text="📋 작업 로그")
        log_frame.pack(fill="both", expand=True, padx=5, pady=5)
        self.log_text = tk.Text(log_frame, wrap="word", height=15)
        log_scrollbar_v = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_scrollbar_v.set)
        log_scrollbar_v.pack(side="right", fill="y")
        self.log_text.pack(fill="both", expand=True)
        
        # --- 4. 상태바 ---
        status_frame_bottom = ttk.Frame(right_scrollable_frame)
        status_frame_bottom.pack(fill="x", padx=5, pady=5)
        self.status_label_apply = ttk.Label(status_frame_bottom, text="대기 중... 번역 엑셀 파일을 선택하세요.")
        self.status_label_apply.pack(side="left", padx=5)
        
        # [신규] 진행률 표시줄
        self.progress_bar = ttk.Progressbar(status_frame_bottom, orient="horizontal", mode="determinate")
        # self.progress_bar.pack(side="left", fill="x", expand=True, padx=5) # pack 대신 grid로 관리하여 숨기기/보이기 용이하게
        
        # 스크롤바 패킹
        left_canvas.pack(side="left", fill="both", expand=True)
        left_scrollbar.pack(side="right", fill="y")
        right_canvas.pack(side="left", fill="both", expand=True)
        right_scrollbar.pack(side="right", fill="y")
        
        # 초기 설정
        self.toggle_options_by_mode()
        self.toggle_kr_options()

    ### toggle_filter_options: [변경 없음]
    def toggle_filter_options(self):
        if self.use_filtered_data_var.get():
            self.filter_config_frame.pack(fill="x", padx=15, pady=5)
        else:
            self.filter_config_frame.pack_forget()

    ### detect_special_column: [변경]
    def detect_special_column(self):
        """[변경] Manager의 변경된 DB기반 감지 메서드를 호출합니다."""
        excel_path = self.excel_source_path_var.get()
        # excel_path가 캐시 생성의 기준이므로, 이 경로가 유효한지 확인하는 것은 필수
        if not excel_path or not self.cached_excel_path:
            messagebox.showwarning("파일 선택 필요", "먼저 번역 엑셀 파일을 선택하고 캐시를 생성해주세요.", parent=self)
            return
        
        column_name = self.special_column_input_var.get().strip()
        if not column_name:
            messagebox.showwarning("컬럼명 입력 필요", "감지할 특수 컬럼명을 입력하세요.", parent=self)
            return
        
        self.log_text.insert(tk.END, f"DB 캐시에서 특수 컬럼 '{column_name}' 감지 중...\n")
        
        loading_popup = LoadingPopup(self, "특수 컬럼 감지 중", f"'{column_name}' 컬럼을 분석하고 있습니다...")
        
        def detect_thread():
            try:
                # [변경] 이제 sheet_names 인자가 필요 없음
                result = self.translation_apply_manager.detect_special_column_in_excel(
                    excel_path, column_name
                )
                self.after(0, lambda: self.process_special_detection_result(result, loading_popup, column_name))
            except Exception as e:
                self.after(0, lambda: [
                    loading_popup.close(),
                    self.log_text.insert(tk.END, f"특수 컬럼 감지 중 오류: {str(e)}\n"),
                    messagebox.showerror("오류", f"특수 컬럼 감지 중 오류 발생: {str(e)}", parent=self)
                ])
        
        threading.Thread(target=detect_thread, daemon=True).start()

    ### process_special_detection_result: [변경 없음]
    def process_special_detection_result(self, result, loading_popup, column_name):
        loading_popup.close()
        
        if result["status"] == "error":
            messagebox.showerror("오류", f"특수 컬럼 감지 실패: {result['message']}", parent=self)
            return
        
        detected_info = result.get("detected_info", {})
        
        if detected_info:
            self.detected_special_columns = {column_name: detected_info}
            
            if detected_info.get('unique_values'):
                most_common = detected_info.get('most_common', [])
                if most_common:
                    self.special_condition_var.set(most_common[0][0])
            
            found_sheets_count = len(detected_info.get('found_in_sheets', []))
            self.filter_status_var.set(f"'{column_name}' 컬럼 발견 ({found_sheets_count}개 시트)")
            
            self.log_text.insert(tk.END, f"✅ 특수 컬럼 '{column_name}' 감지 완료:\n")
            self.log_text.insert(tk.END, f"  • 발견된 시트: {found_sheets_count}개\n")
            self.log_text.insert(tk.END, f"  • 데이터 항목: {detected_info['non_empty_rows']}개\n")
            
            suggested_values = [item[0] for item in detected_info.get('most_common', [])[:3]]
            if suggested_values:
                self.log_text.insert(tk.END, f"  • 추천값: {', '.join(suggested_values)}\n")
            
            messagebox.showinfo("완료", f"특수 컬럼 '{column_name}' 감지 완료!\n\n발견된 시트: {found_sheets_count}개\n데이터 항목: {detected_info['non_empty_rows']}개\n\n조건값을 확인하고 '번역 데이터 로드'를 실행하세요.", parent=self)
        else:
            self.filter_status_var.set(f"'{column_name}' 컬럼 없음")
            self.log_text.insert(tk.END, f"⚠️ 특수 컬럼 '{column_name}'을(를) 찾을 수 없습니다.\n")
            messagebox.showinfo("알림", f"'{column_name}' 컬럼을 찾을 수 없습니다.\n\n컬럼명을 확인하고 다시 시도해주세요.", parent=self)

    ### toggle_options_by_mode: [변경 없음]
    def toggle_options_by_mode(self, *args):
        mode = self.apply_mode_var.get()
        if mode == "id":
            self.kr_based_options_frame.pack_forget()
            self.id_based_options_frame.pack(fill="x", padx=15, pady=5)
        elif mode == "kr":
            self.id_based_options_frame.pack_forget()
            self.kr_based_options_frame.pack(fill="x", padx=15, pady=5)
        self.toggle_kr_options()

    ### toggle_kr_options: [변경 없음]
    def toggle_kr_options(self):
        mode = self.apply_mode_var.get()
        if mode == "id":
            is_kr_check_enabled = self.kr_match_check_var.get()
            state = "normal" if is_kr_check_enabled else "disabled"
            
            self.id_mismatch_delete_cb.config(state=state)
            self.id_overwrite_cb.config(state=state)
            
            if not is_kr_check_enabled:
                self.kr_mismatch_delete_var.set(False)
                self.kr_overwrite_var.set(False)
        else:
            self.id_mismatch_delete_cb.config(state="disabled")
            self.id_overwrite_cb.config(state="disabled")

    def select_excel_source_file(self):
        """[수정] 파일 선택 시 재구축 버튼 및 상태 플래그를 초기화합니다."""
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel 파일", "*.xlsx"), ("모든 파일", "*.*")],
            title="번역 엑셀 파일 선택", parent=self
        )
        if not file_path:
            return

        self.excel_source_path_var.set(file_path)
        self.translation_db_var.set("")
        self.selected_sheets = []
        self.selected_sheets_display_var.set("DB 캐시 생성 필요")
        self.cached_sheet_names = []
        self.filter_status_var.set("필터링 데이터 없음")
        
        # ▼▼▼ 상태 초기화 추가 ▼▼▼
        self.initial_info_shown = False # 첫 정보 메시지 표시 플래그 초기화
        self.rebuild_cache_button.config(state="disabled") 
        # ▲▲▲ 여기까지 추가 ▲▲▲
        
        self.sheet_select_button.config(state="disabled")
        self.detect_button.config(state="disabled")
        self.load_data_button.config(state="disabled")
        self.apply_button.config(state="disabled")
        self.view_data_button.config(state="disabled")
        self.view_filtered_button.config(state="disabled")
        
        self._start_db_caching_thread(file_path)

    ### _start_db_caching_thread: [신규]
    def _start_db_caching_thread(self, excel_path, force=False):
        """[수정] force 플래그를 받아 Manager에게 전달합니다."""
        
        self.progress_bar.pack(side="left", fill="x", expand=True, padx=5)
        self.progress_bar["value"] = 0
        self.status_label_apply.config(text="DB 캐시 구축 준비 중...")
        
        # 다른 버튼들 비활성화
        self.sheet_select_button.config(state="disabled")
        self.rebuild_cache_button.config(state="disabled")
        self.load_data_button.config(state="disabled")

        def progress_callback(percentage, message):
            self.after(0, lambda: [
                self.progress_bar.config(value=percentage),
                self.status_label_apply.config(text=message)
            ])

        def task():
            # ▼▼▼ force_rebuild 파라미터 전달 ▼▼▼
            result = self.translation_apply_manager.initiate_excel_caching(
                excel_path, force_rebuild=force, progress_callback=progress_callback
            )
            self.after(0, lambda: self._process_caching_result(result, excel_path))

        threading.Thread(target=task, daemon=True).start()

    def _process_caching_result(self, result, excel_path):
        """[수정] 작업 완료 후 '캐시 재구축' 버튼을 활성화합니다."""
        self.progress_bar.pack_forget() 
        
        if result["status"] == "error":
            self.status_label_apply.config(text="DB 캐시 구축 실패.")
            messagebox.showerror("캐시 생성 오류", result["message"], parent=self)
            return

        self.cached_excel_path = excel_path
        self.cached_sheet_names = result.get("sheets", [])
        self.status_label_apply.config(text=f"DB 캐시 준비 완료. 시트를 선택하세요. (총 {len(self.cached_sheet_names)}개)")
        self.selected_sheets_display_var.set(f"{len(self.cached_sheet_names)}개 시트 발견됨")
        
        # ▼▼▼ 버튼 활성화 로직 수정 ▼▼▼
        self.sheet_select_button.config(state="normal")
        self.rebuild_cache_button.config(state="normal") # 재구축 버튼 활성화
        self.detect_button.config(state="normal")
        
        if not hasattr(self, 'initial_info_shown') or self.initial_info_shown is False:
                messagebox.showinfo("캐시 준비 완료", f"'{os.path.basename(excel_path)}'에 대한 DB 캐시가 준비되었습니다.\n'시트 선택' 버튼으로 작업할 시트를 골라주세요.", parent=self)
                self.initial_info_shown = True # 첫 정보 메시지 표시 후 플래그 설정
    
    def load_translation_data(self):
        """[변경] Manager의 DB기반 메서드를 호출하여 메모리 캐시를 구성합니다."""
        db_path = self.translation_db_var.get()
        excel_path = self.excel_source_path_var.get()

        if db_path: # 사용자가 직접 선택한 DB 파일 처리 (기존 로직 유지)
            self.load_from_db(db_path)
            self.apply_button.config(state="normal")
        elif excel_path: # 자동 생성된 DB 캐시로부터 데이터 로드
            if not self.selected_sheets:
                messagebox.showwarning("경고", "'시트 선택' 버튼을 눌러 데이터를 읽어올 시트를 선택하세요.", parent=self)
                return
            
            special_filter = None
            if self.use_filtered_data_var.get():
                column_name = self.special_column_input_var.get().strip()
                condition_value = self.special_condition_var.get().strip()
                if not column_name or not condition_value:
                    messagebox.showwarning("경고", "특수 컬럼 필터링을 사용하려면 컬럼명과 조건값을 모두 설정하세요.", parent=self)
                    return
                special_filter = {"column_name": column_name, "condition_value": condition_value}
            
            # 이 작업은 DB 쿼리라 매우 빠르지만, 일관성을 위해 로딩 팝업 유지
            loading_popup = LoadingPopup(self, "메모리 캐싱 중", "번역 데이터 로딩 중...")
            
            def task():
                # [변경] Manager의 새 메서드 호출
                result = self.translation_apply_manager.load_translation_cache_from_excel_with_filter(
                    excel_path, self.selected_sheets, special_filter
                )
                self.after(0, lambda: [
                    self.process_cache_load_result(result, loading_popup),
                    self.apply_button.config(state="normal") if result["status"] == "success" else None
                ])
            
            threading.Thread(target=task, daemon=True).start()
        else:
            messagebox.showwarning("경고", "번역 DB 또는 엑셀 파일을 선택하세요.", parent=self)

    ### open_sheet_selection_popup: [변경]
    def open_sheet_selection_popup(self):
        """[변경] 더 이상 엑셀을 읽지 않고, 캐시된 시트 목록을 사용합니다."""
        if not self.cached_sheet_names:
            messagebox.showwarning("오류", "사용 가능한 시트 목록이 없습니다. 먼저 엑셀 파일을 선택하여 캐시를 생성하세요.", parent=self)
            return

        all_sheets = self.cached_sheet_names
        
        popup = tk.Toplevel(self)
        popup.title("시트 선택")
        popup.geometry("400x500")
        popup.transient(self)
        popup.grab_set()

        checklist = ScrollableCheckList(popup, height=350)
        checklist.pack(fill="both", expand=True, padx=10, pady=10)

        selected_set = set(self.selected_sheets)
        for sheet in all_sheets:
            checklist.add_item(sheet, checked=(sheet in selected_set))

        def on_confirm():
            self.selected_sheets = checklist.get_checked_items()
            if self.selected_sheets:
                if len(self.selected_sheets) > 3:
                    display_text = f"{len(self.selected_sheets)}개 시트 선택됨"
                else:
                    display_text = ", ".join(self.selected_sheets)
                self.selected_sheets_display_var.set(display_text)
                self.load_data_button.config(state="normal") # 시트 선택 후 로드 버튼 활성화
            else:
                self.selected_sheets_display_var.set("선택된 시트 없음")
                self.load_data_button.config(state="disabled")
            popup.destroy()

        confirm_button = ttk.Button(popup, text="확인", command=on_confirm)
        confirm_button.pack(pady=10)

    ### load_from_db: [변경 없음]
    def load_from_db(self, db_path):
        if not os.path.isfile(db_path):
            messagebox.showwarning("경고", "유효한 번역 DB 파일을 선택하세요.", parent=self)
            return
            
        self.log_text.insert(tk.END, "번역 DB 캐싱 중...\n")
        loading_popup = LoadingPopup(self, "DB 캐싱 중", "번역 데이터 캐싱 중...")
        
        def task():
            result = self.translation_apply_manager.load_translation_cache_from_db(db_path)
            self.after(0, lambda: self.process_cache_load_result(result, loading_popup))
        
        threading.Thread(target=task, daemon=True).start()

    ### load_from_excel: [삭제]
    # 이 함수의 기능은 `load_translation_data`로 통합되었습니다.

    ### process_cache_load_result: [변경]
    def process_cache_load_result(self, result, loading_popup):
        """[변경] 로드 결과 처리 로직을 새 반환값 구조에 맞게 일부 수정합니다."""
        loading_popup.close()
        
        if result["status"] == "error":
            messagebox.showerror("오류", f"캐싱 중 오류 발생: {result['message']}", parent=self)
            self.log_text.insert(tk.END, f"캐싱 실패: {result['message']}\n")
            return
        
        # [변경] Manager의 캐시를 직접 설정하는 대신, Manager가 내부적으로 관리하도록 둠
        # self.translation_apply_manager.translation_cache = result["translation_cache"] ...
        
        id_count = result.get("id_count", 0)
        filtered_count = result.get("filtered_count", 0)

        if filtered_count > 0:
            self.filter_status_var.set(f"필터링된 데이터: {filtered_count}개")
            self.view_filtered_button.config(state="normal")
            self.log_text.insert(tk.END, f"🔍 특수 컬럼 필터링 결과: {filtered_count}개\n")
        else:
            self.filter_status_var.set("필터링 데이터 없음")
            self.view_filtered_button.config(state="disabled")
        
        # 결과 로그
        source_type_msg = "사용자 DB" if result.get("source_type") == "DB" else "엑셀 DB 캐시"
        self.log_text.insert(tk.END, f"메모리 캐시 로딩 완료 (소스: {source_type_msg}):\n")
        self.log_text.insert(tk.END, f"- 전체 고유 STRING_ID: {id_count}개\n")
        
        if filtered_count > 0:
            self.log_text.insert(tk.END, f"- 특수 컬럼 필터링: {filtered_count}개\n")
        
        status_parts = [f"{id_count}개 항목"]
        if filtered_count > 0:
            status_parts.append(f"{filtered_count}개 필터링됨")
        
        self.status_label_apply.config(text=f"데이터 로드 완료 - {', '.join(status_parts)}")
        self.view_data_button.config(state="normal")
        
        completion_message = f"번역 데이터 로딩 완료!\n항목 수: {id_count}개"
        if filtered_count > 0:
            completion_message += f"\n특수 필터링: {filtered_count}개"
        
        messagebox.showinfo("완료", completion_message, parent=self)

    ### show_filtered_data_viewer: [변경 없음]
    def show_filtered_data_viewer(self):
        if not hasattr(self.translation_apply_manager, 'special_filtered_cache') or not self.translation_apply_manager.special_filtered_cache:
            messagebox.showinfo("정보", "표시할 필터링된 데이터가 없습니다.", parent=self)
            return

        viewer_win = tk.Toplevel(self)
        viewer_win.title("필터링된 번역 데이터 보기")
        viewer_win.geometry("1200x700")
        viewer_win.transient(self)
        viewer_win.grab_set()

        info_frame = ttk.Frame(viewer_win, padding="5")
        info_frame.pack(fill="x")
        
        filter_info = f"필터 조건: {self.special_column_input_var.get()} = '{self.special_condition_var.get()}'"
        ttk.Label(info_frame, text=filter_info, font=("Arial", 10, "bold")).pack(anchor="w")
        ttk.Label(info_frame, text=f"필터링된 항목 수: {len(self.translation_apply_manager.special_filtered_cache)}개", 
                  foreground="blue").pack(anchor="w")

        tree_frame = ttk.Frame(viewer_win, padding="5")
        tree_frame.pack(fill="both", expand=True)

        # 컬럼에 string_id 추가
        columns = ("string_id", "kr", "cn", "tw", "file_name", "sheet_name", self.special_column_input_var.get())
        tree = ttk.Treeview(tree_frame, columns=columns, show="headings")
        
        for col in columns:
            tree.heading(col, text=col.upper())
            if col == "kr": tree.column(col, width=250)
            elif col in ["cn", "tw"]: tree.column(col, width=200)
            elif col == self.special_column_input_var.get(): tree.column(col, width=100)
            else: tree.column(col, width=150)

        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        vsb.pack(side="right", fill="y")
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=tree.xview)
        hsb.pack(side="bottom", fill="x")
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        tree.pack(fill="both", expand=True)

        import json
        for string_id, data in self.translation_apply_manager.special_filtered_cache.items():
            # special_columns가 json 문자열이므로 파싱해야 함
            special_values = json.loads(data.get('special_columns', '{}'))
            values = (
                string_id,
                data.get('kr', ''),
                data.get('cn', ''),
                data.get('tw', ''),
                data.get('file_name', ''),
                data.get('sheet_name', ''),
                special_values.get(self.special_column_input_var.get(), '')
            )
            tree.insert("", "end", values=values)

        button_frame = ttk.Frame(viewer_win, padding="5")
        button_frame.pack(fill="x")
        
        ttk.Button(button_frame, text="Excel로 내보내기", 
                   command=lambda: self.export_filtered_data_standalone()).pack(side="left")
        ttk.Button(button_frame, text="닫기", 
                   command=viewer_win.destroy).pack(side="right")

    ### export_filtered_data_standalone: [변경 없음]
    def export_filtered_data_standalone(self):
        if not self.translation_apply_manager.special_filtered_cache:
            messagebox.showerror("오류", "내보낼 필터링된 데이터가 없습니다.", parent=self)
            return
        
        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel 파일", "*.xlsx")],
            title="필터링된 데이터 엑셀 저장",
            parent=self
        )
        if not save_path:
            return
        
        try:
            data_list = list(self.translation_apply_manager.special_filtered_cache.values())
            df = pd.DataFrame(data_list)
            df.to_excel(save_path, index=False)
            
            self.log_text.insert(tk.END, f"필터링된 데이터 엑셀 저장 완료: {save_path}\n")
            messagebox.showinfo("성공", f"필터링된 데이터가 성공적으로 저장되었습니다:\n{save_path}", parent=self)
            
        except Exception as e:
            messagebox.showerror("저장 오류", f"파일 저장 중 오류 발생:\n{e}", parent=self)

    def apply_translation(self):
        """[버그 수정] self.special_column_var -> self.special_column_input_var로 수정"""
        if not hasattr(self.translation_apply_manager, 'translation_cache') or not self.translation_apply_manager.translation_cache:
            messagebox.showwarning("경고", "먼저 '번역 데이터 로드'를 실행하세요.", parent=self)
            return
            
        selected_files = self.original_files_list.get_checked_items()
        if not selected_files:
            messagebox.showwarning("경고", "적용할 파일을 선택하세요.", parent=self)
            return
            
        selected_langs = [lang for lang, var in self.apply_lang_vars.items() if var.get()]
        if not selected_langs:
            messagebox.showwarning("경고", "적용할 언어를 하나 이상 선택하세요.", parent=self)
            return

        use_filtered = self.use_filtered_data_var.get()
        if use_filtered and not self.translation_apply_manager.special_filtered_cache:
            messagebox.showwarning("경고", "특수 컬럼 필터링 데이터가 없습니다. 먼저 필터링 설정을 하고 데이터를 로드하세요.", parent=self)
            return

        files_to_process = [item for item in self.original_files if item[0] in selected_files]
        open_files = self._check_files_are_open([path for name, path in files_to_process])
        if open_files:
            messagebox.showwarning("작업 중단", f"다음 파일이 열려 있어 작업을 시작할 수 없습니다:\n\n" + "\n".join(open_files), parent=self)
            return

        self.log_text.delete(1.0, tk.END)
        
        allowed_statuses = []
        if self.apply_on_new_var.get(): allowed_statuses.append('신규')
        if self.apply_on_change_var.get(): allowed_statuses.append('change')
        if self.apply_on_transferred_var.get(): allowed_statuses.append('전달')
        
        mode_text = "ID 기반" if self.apply_mode_var.get() == "id" else "KR 기반"
        lang_text = ", ".join(selected_langs)
        condition_text = ", ".join(allowed_statuses) if allowed_statuses else "모든 항목"
        data_source = "특수필터링 데이터" if use_filtered else "전체 데이터"
        
        self.log_text.insert(tk.END, "="*60 + "\n")
        self.log_text.insert(tk.END, "🚀 번역 적용 작업 시작\n")
        self.log_text.insert(tk.END, f"📋 적용 모드: {mode_text}\n")
        self.log_text.insert(tk.END, f"🎯 데이터 소스: {data_source}\n")
        self.log_text.insert(tk.END, f"🌍 적용 언어: {lang_text}\n")
        self.log_text.insert(tk.END, f"🔍 적용 조건: {condition_text}\n")
        self.log_text.insert(tk.END, f"📁 대상 파일: {len(files_to_process)}개\n")
        
        if use_filtered:
            filtered_count = len(self.translation_apply_manager.special_filtered_cache)
            self.log_text.insert(tk.END, f"🔍 특수 필터링된 캐시 사용: {filtered_count}개 항목\n")

        if use_filtered:
            filter_info = f"{self.special_column_input_var.get()} = '{self.special_condition_var.get()}'"
        
        self.log_text.insert(tk.END, "="*60 + "\n\n")
        
        self.status_label_apply.config(text="작업 중...")
        self.update()
            
        loading_popup = LoadingPopup(self, "번역 적용 중", "번역 적용 준비 중...")
            
        apply_options = {
            "mode": self.apply_mode_var.get(),
            "selected_langs": selected_langs,
            "record_date": self.record_date_var.get(),
            "kr_match_check": self.kr_match_check_var.get(),
            "kr_mismatch_delete": self.kr_mismatch_delete_var.get(),
            "kr_overwrite": self.kr_overwrite_var.get(),
            "kr_overwrite_on_kr_mode": self.kr_overwrite_on_kr_mode_var.get(),
            "allowed_statuses": allowed_statuses,
            "use_filtered_data": use_filtered
        }
            
        def apply_translations_thread():
            total_results = {
                "total_updated": 0, "total_overwritten": 0, "total_kr_mismatch_skipped": 0, 
                "total_kr_mismatch_deleted": 0, "total_conditional_skipped": 0,
            }
            processed_count = 0
            error_count = 0
            successful_files = []
            failed_files = []

            # ▼▼▼ [요청 2, 3] 수집할 리스트 추가 ▼▼▼
            modified_files = []
            total_overwritten_items = []

            start_time = time.time()

            for idx, (file_name, file_path) in enumerate(files_to_process):
                self.after(0, lambda i=idx, n=file_name: [
                    loading_popup.update_progress((i / len(files_to_process)) * 100, f"파일 처리 중 ({i+1}/{len(files_to_process)}): {n}"),
                ])

                result = self.translation_apply_manager.apply_translation_with_filter_option(
                    file_path,
                    apply_options
                )

                if result["status"] == "success":
                    processed_count += 1
                    successful_files.append(file_name)
                    for key in total_results:
                        total_results[key] += result.get(key, 0)

                    # ▼▼▼ [요청 2, 3] 결과 수집 ▼▼▼
                    if result.get("total_updated", 0) > 0 or result.get("total_overwritten", 0) > 0:
                        modified_files.append(file_name)

                    total_overwritten_items.extend(result.get("overwritten_items", []))
                else:
                    error_count += 1
                    failed_files.append((file_name, result.get("message", "알 수 없는 오류")))

            elapsed_time = time.time() - start_time

            # ▼▼▼ [요청 2, 3] 수집된 리스트를 결과 처리 함수로 전달 ▼▼▼
            self.after(0, lambda: self.process_translation_apply_result(
                total_results, processed_count, error_count, loading_popup, 
                successful_files, failed_files, elapsed_time, use_filtered,
                modified_files, total_overwritten_items) # 전달인자 추가
            )
            
        thread = threading.Thread(target=apply_translations_thread, daemon=True)
        thread.start()

    def process_translation_apply_result(self, total_results, processed_count, error_count, loading_popup, 
                                        successful_files, failed_files, elapsed_time, use_filtered,
                                        modified_files, total_overwritten_items): # 전달인자 추가
        """[수정] 변경된 파일 목록 출력 및 덮어쓴 데이터 보기 기능을 처리합니다."""
        loading_popup.close()

        minutes = int(elapsed_time // 60)
        seconds = int(elapsed_time % 60)
        time_str = f"{minutes}분 {seconds}초" if minutes > 0 else f"{seconds}초"

        self.log_text.insert(tk.END, "\n" + "="*60 + "\n")
        self.log_text.insert(tk.END, "🎉 번역 적용 작업 완료\n")
        self.log_text.insert(tk.END, "="*60 + "\n")

        data_source_info = "특수필터링 데이터" if use_filtered else "전체 데이터"
        self.log_text.insert(tk.END, f"📊 데이터 소스: {data_source_info}\n")

        self.log_text.insert(tk.END, f"⏱️  소요 시간: {time_str}\n")
        self.log_text.insert(tk.END, f"✅ 성공: {processed_count}개 파일\n")
        if error_count > 0: self.log_text.insert(tk.END, f"❌ 실패: {error_count}개 파일\n")

        # ▼▼▼ [요청 2] 변경된 파일 목록 출력 ▼▼▼
        if modified_files:
            self.log_text.insert(tk.END, f"🔄 변경된 파일 ({len(modified_files)}개):\n")
            for f_name in modified_files[:10]: # 최대 10개까지만 표시
                self.log_text.insert(tk.END, f"   - {f_name}\n")
            if len(modified_files) > 10:
                self.log_text.insert(tk.END, f"   ... 외 {len(modified_files) - 10}개\n")
        # ▲▲▲ 여기까지 추가 ▲▲▲

        total_applied = total_results["total_updated"] + total_results["total_overwritten"]
        self.log_text.insert(tk.END, f"\n📊 작업 통계:\n")
        if total_results["total_updated"] > 0: self.log_text.insert(tk.END, f"   • 신규 적용: {total_results['total_updated']:,}개\n")
        if total_results["total_overwritten"] > 0: self.log_text.insert(tk.END, f"   • 덮어쓰기: {total_results['total_overwritten']:,}개\n")
        if total_results["total_conditional_skipped"] > 0: self.log_text.insert(tk.END, f"   • 조건 불일치로 건너뜀: {total_results['total_conditional_skipped']:,}개\n")
        if total_results["total_kr_mismatch_skipped"] > 0: self.log_text.insert(tk.END, f"   • KR 불일치로 건너뜀: {total_results['total_kr_mismatch_skipped']:,}개\n")
        if total_results["total_kr_mismatch_deleted"] > 0: self.log_text.insert(tk.END, f"   • KR 불일치로 삭제: {total_results['total_kr_mismatch_deleted']:,}개\n")

        self.log_text.insert(tk.END, f"\n🎯 총 적용된 번역: {total_applied:,}개\n")

        if failed_files:
            self.log_text.insert(tk.END, f"\n❌ 실패한 파일:\n")
            for file_name, error_msg in failed_files[:5]:
                self.log_text.insert(tk.END, f"   • {file_name}: {error_msg}\n")
            if len(failed_files) > 5: self.log_text.insert(tk.END, f"   ... 외 {len(failed_files) - 5}개\n")

        self.log_text.insert(tk.END, "="*60 + "\n")
        self.log_text.see(tk.END)

        status_text = f"완료 - {total_applied:,}개 항목 적용"
        if use_filtered: status_text += " (특수필터링)"
        self.status_label_apply.config(text=status_text)

        # ▼▼▼ [요청 3] 덮어쓴 데이터 처리 및 버튼 활성화 ▼▼▼
        self.overwritten_data = total_overwritten_items
        if self.overwritten_data:
            self.view_overwritten_button.config(state="normal")
            status_text += f" ({len(self.overwritten_data)}개 덮어씀)"
            self.status_label_apply.config(text=status_text)
        else:
            self.view_overwritten_button.config(state="disabled")
        # ▲▲▲ 여기까지 추가 ▲▲▲

        completion_msg = f"번역 적용이 완료되었습니다!\n\n"
        completion_msg += f"📊 데이터 소스: {data_source_info}\n"
        completion_msg += f"⏱️ 소요 시간: {time_str}\n"
        completion_msg += f"✅ 성공: {processed_count}개 파일\n"
        if error_count > 0:
            completion_msg += f"❌ 실패: {error_count}개 파일\n"
        completion_msg += f"\n🎯 총 적용된 번역: {total_applied:,}개"
        
        if total_results["total_updated"] > 0:
            completion_msg += f"\n   • 신규 적용: {total_results['total_updated']:,}개"
        if total_results["total_overwritten"] > 0:
            completion_msg += f"\n   • 덮어쓰기: {total_results['total_overwritten']:,}개"
        
        messagebox.showinfo("완료", completion_msg, parent=self)

    ### select_translation_db_file: [변경 없음]
    def select_translation_db_file(self, *args):
        file_path = filedialog.askopenfilename(
            filetypes=[("DB 파일", "*.db"), ("모든 파일", "*.*")],
            title="번역 DB 선택", parent=self
        )
        if file_path:
            self.translation_db_var.set(file_path)
            self.excel_source_path_var.set("") 
            self.selected_sheets = []
            self.selected_sheets_display_var.set("선택된 시트 없음")
            self.cached_excel_path = None
            self.cached_sheet_names = []
            self.filter_status_var.set("필터링 데이터 없음")
            # 버튼 상태 업데이트
            self.sheet_select_button.config(state="disabled")
            self.detect_button.config(state="disabled")
            self.load_data_button.config(state="normal") # 직접 DB 선택시에는 바로 로드 가능

    ### select_original_folder: [변경 없음]
    def select_original_folder(self):
        folder = filedialog.askdirectory(title="원본 파일 폴더 선택", parent=self)
        if folder:
            self.original_folder_var.set(folder)

    ### search_original_files: [변경 없음]
    def search_original_files(self):
        folder = self.original_folder_var.get()
        if not folder or not os.path.isdir(folder):
            messagebox.showwarning("경고", "유효한 폴더를 선택하세요.", parent=self)
            return
        
        self.original_files_list.clear()
        self.original_files = []
        
        for root_dir, _, files in os.walk(folder):
            for file in files:
                if file.startswith("String") and file.endswith(".xlsx") and not file.startswith("~$"):
                    if file not in self.excluded_files:
                        file_path = os.path.join(root_dir, file)
                        self.original_files.append((file, file_path))
                        self.original_files_list.add_item(file, checked=True)
        
        if not self.original_files:
            messagebox.showinfo("알림", "String으로 시작하는 엑셀 파일을 찾지 못했습니다.", parent=self)
        else:
            messagebox.showinfo("알림", f"{len(self.original_files)}개의 엑셀 파일을 찾았습니다.", parent=self)

    ### _check_files_are_open: [변경 없음]
    def _check_files_are_open(self, file_paths_to_check):
        open_files = []
        for file_path in file_paths_to_check:
            if not os.path.exists(file_path):
                continue
            try:
                os.rename(file_path, file_path)
            except OSError:
                open_files.append(os.path.basename(file_path))
        return open_files

    ### show_loaded_data_viewer: [변경 없음]
    def show_loaded_data_viewer(self):
        if not hasattr(self.translation_apply_manager, 'translation_cache') or not self.translation_apply_manager.translation_cache:
            messagebox.showinfo("정보", "먼저 번역 데이터를 로드해주세요.", parent=self)
            return

        viewer_win = tk.Toplevel(self)
        viewer_win.title("로드된 번역 데이터 보기")
        viewer_win.geometry("1200x700")
        viewer_win.transient(self)
        viewer_win.grab_set()

        search_frame = ttk.Frame(viewer_win, padding="5")
        search_frame.pack(fill="x")
        
        ttk.Label(search_frame, text="STRING_ID:").pack(side="left", padx=(0, 2))
        id_search_var = tk.StringVar()
        id_search_entry = ttk.Entry(search_frame, textvariable=id_search_var, width=30)
        id_search_entry.pack(side="left", padx=(0, 10))

        ttk.Label(search_frame, text="KR:").pack(side="left", padx=(0, 2))
        kr_search_var = tk.StringVar()
        kr_search_entry = ttk.Entry(search_frame, textvariable=kr_search_var, width=40)
        kr_search_entry.pack(side="left", padx=(0, 10))

        tree_frame = ttk.Frame(viewer_win, padding="5")
        tree_frame.pack(fill="both", expand=True)

        columns = ("string_id", "kr", "cn", "tw", "file_name", "sheet_name")
        tree = ttk.Treeview(tree_frame, columns=columns, show="headings")
        
        tree.heading("string_id", text="STRING_ID")
        tree.heading("kr", text="KR")
        tree.heading("cn", text="CN")
        tree.heading("tw", text="TW")
        tree.heading("file_name", text="파일명")
        tree.heading("sheet_name", text="시트명")

        tree.column("string_id", width=150)
        tree.column("kr", width=250)
        tree.column("cn", width=200)
        tree.column("tw", width=200)
        tree.column("file_name", width=150)
        tree.column("sheet_name", width=150)

        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        vsb.pack(side="right", fill="y")
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=tree.xview)
        hsb.pack(side="bottom", fill="x")
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        tree.pack(fill="both", expand=True)

        status_frame = ttk.Frame(viewer_win, padding="5")
        status_frame.pack(fill="x")
        status_label = ttk.Label(status_frame, text="데이터 준비 중...")
        status_label.pack(side="left")

        all_data = []
        for string_id, data_dict in self.translation_apply_manager.translation_cache.items():
            item = data_dict.copy()
            item['string_id'] = string_id
            all_data.append(item)

        def populate_tree(data_to_show):
            tree.delete(*tree.get_children())
            
            for item in data_to_show:
                values = (
                    item.get('string_id', ''),
                    item.get('kr', ''),
                    item.get('cn', ''),
                    item.get('tw', ''),
                    item.get('file_name', ''),
                    item.get('sheet_name', '')
                )
                tree.insert("", "end", values=values)
            status_label.config(text=f"{len(data_to_show):,} / {len(all_data):,}개 항목 표시 중")

        def perform_search():
            id_query = id_search_var.get().lower().strip()
            kr_query = kr_search_var.get().lower().strip()

            if not id_query and not kr_query:
                populate_tree(all_data)
                return

            filtered_data = []
            for item in all_data:
                id_match = (id_query in item.get('string_id', '').lower()) if id_query else True
                kr_match = (kr_query in item.get('kr', '').lower()) if kr_query else True
                
                if id_match and kr_match:
                    filtered_data.append(item)
            
            populate_tree(filtered_data)

        def reset_search():
            id_search_var.set("")
            kr_search_var.set("")
            populate_tree(all_data)

        search_button = ttk.Button(search_frame, text="검색", command=perform_search)
        search_button.pack(side="left", padx=5)
        reset_button = ttk.Button(search_frame, text="초기화", command=reset_search)
        reset_button.pack(side="left", padx=5)
        
        id_search_entry.bind("<Return>", lambda event: perform_search())
        kr_search_entry.bind("<Return>", lambda event: perform_search())

        populate_tree(all_data)
        
    def force_rebuild_cache(self):
        """[신규] 사용자 확인 후 DB 캐시를 강제로 다시 구축합니다."""
        excel_path = self.excel_source_path_var.get()
        if not excel_path or not self.cached_excel_path:
            # 이 경우는 버튼이 비활성화되어 있어 거의 발생하지 않음
            messagebox.showwarning("오류", "먼저 엑셀 파일을 선택해주세요.", parent=self)
            return

        if messagebox.askyesno("캐시 재구축 확인",
                                "현재 엑셀 파일의 캐시를 강제로 다시 만드시겠습니까?\n\n(시간이 소요될 수 있습니다)",
                                parent=self):
            # 기존 캐싱 스레드 호출 함수를 재사용하되, force 플래그를 True로 전달
            self._start_db_caching_thread(excel_path, force=True)

    def show_overwritten_data_viewer(self):
        """[신규] 덮어쓰기 된 항목들을 보여주는 새 창을 엽니다."""
        if not hasattr(self, 'overwritten_data') or not self.overwritten_data:
            messagebox.showinfo("정보", "표시할 덮어쓰기 데이터가 없습니다.", parent=self)
            return

        viewer_win = tk.Toplevel(self)
        viewer_win.title(f"덮어쓴 데이터 보기 ({len(self.overwritten_data)}개)")
        viewer_win.geometry("1200x700")
        viewer_win.transient(self)
        viewer_win.grab_set()

        tree_frame = ttk.Frame(viewer_win, padding="5")
        tree_frame.pack(fill="both", expand=True)

        columns = ("file_name", "sheet_name", "string_id", "language", "kr_text", "overwritten_text")
        tree = ttk.Treeview(tree_frame, columns=columns, show="headings")
        
        # 컬럼 헤더 설정
        tree.heading("file_name", text="파일명")
        tree.heading("sheet_name", text="시트명")
        tree.heading("string_id", text="STRING_ID")
        tree.heading("language", text="언어")
        tree.heading("kr_text", text="KR 원문")
        tree.heading("overwritten_text", text="덮어쓴 내용")

        # 컬럼 너비 설정
        tree.column("file_name", width=180, anchor="w")
        tree.column("sheet_name", width=120, anchor="w")
        tree.column("string_id", width=200, anchor="w")
        tree.column("language", width=60, anchor="center")
        tree.column("kr_text", width=250, anchor="w")
        tree.column("overwritten_text", width=250, anchor="w")

        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        vsb.pack(side="right", fill="y")
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=tree.xview)
        hsb.pack(side="bottom", fill="x")
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        tree.pack(fill="both", expand=True)

        # 데이터 추가
        for item in self.overwritten_data:
            values = (
                item.get('file_name', ''),
                item.get('sheet_name', ''),
                item.get('string_id', ''),
                item.get('language', ''),
                item.get('kr_text', ''),
                item.get('overwritten_text', '')
            )
            tree.insert("", "end", values=values)

        button_frame = ttk.Frame(viewer_win, padding="5")
        button_frame.pack(fill="x")
        ttk.Button(button_frame, text="닫기", command=viewer_win.destroy).pack(side="right")