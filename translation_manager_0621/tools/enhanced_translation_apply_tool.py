import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import threading
import time
import sys
import pandas as pd
from datetime import datetime
from collections import defaultdict
from tkinter import font as tkfont

# --- 경로 문제 해결을 위한 코드 ---
# 이 부분은 환경에 맞게 조정이 필요할 수 있습니다.
try:
    from ui.common_components import ScrollableCheckList, LoadingPopup
    from tools.enhanced_translation_apply_manager import EnhancedTranslationApplyManager
except ImportError:
    # 대체 경로 설정 (프로젝트 구조에 따라 다름)
    project_root = os.path.dirname(os.path.abspath(__file__))
    if project_root not in sys.path:
        sys.path.append(project_root)
    # 상위 폴더를 추가해야 할 경우
    parent_root = os.path.dirname(project_root)
    if parent_root not in sys.path:
        sys.path.append(parent_root)
    from ui.common_components import ScrollableCheckList, LoadingPopup
    from tools.enhanced_translation_apply_manager import EnhancedTranslationApplyManager

import openpyxl

# ==============================================================================
# [신규] 번역 충돌 해결을 위한 UI 팝업 클래스
# ==============================================================================
class ConflictResolverPopup(tk.Toplevel):
    def __init__(self, parent, conflicts, callback):
        super().__init__(parent)
        self.title("번역 충돌 해결")
        self.geometry("900x600")
        self.transient(parent)
        self.grab_set()

        self.conflicts = conflicts
        self.callback = callback
        self.resolutions = {} 

        # --- [신규] 폰트 객체 정의 ---
        # 여기서 원하는 폰트, 크기, 스타일을 설정할 수 있습니다.
        # family: 글꼴 이름 (예: "맑은 고딕", "나눔고딕", "Dotum")
        # size: 글꼴 크기
        # weight: "normal"(보통), "bold"(굵게)
        self.custom_font = tkfont.Font(family="GmarketSansTTFMedium", size=15, weight="normal")
        # --------------------------------

        self.protocol("WM_DELETE_WINDOW", self._on_close)

        main_frame = ttk.Frame(self, padding="10")
        main_frame.pack(fill="both", expand=True)

        info_text = f"동일한 KR 텍스트에 대해 여러 다른 번역이 발견되었습니다. ({len(conflicts)}개 항목)\n아래 목록에서 각 KR 텍스트에 사용할 올바른 번역을 하나씩 선택해주세요."
        ttk.Label(main_frame, text=info_text, wraplength=880).pack(fill="x", pady=(0, 10))

        paned_window = ttk.PanedWindow(main_frame, orient="horizontal")
        paned_window.pack(fill="both", expand=True)

        left_frame = ttk.Frame(paned_window, padding="5")
        ttk.Label(left_frame, text="충돌 항목 (KR)", font=("", 10, "bold")).pack(anchor="w")
        
        self.kr_listbox = tk.Listbox(left_frame, exportselection=False, width=40)
        self.kr_listbox.pack(fill="both", expand=True)
        for kr_text in self.conflicts.keys():
            self.kr_listbox.insert(tk.END, kr_text)
        
        self.kr_listbox.bind("<<ListboxSelect>>", self._on_kr_select)
        paned_window.add(left_frame, weight=1)

        right_frame = ttk.Frame(paned_window, padding="5")
        self.right_frame_label = ttk.Label(right_frame, text="번역 옵션", font=("", 10, "bold"))
        self.right_frame_label.pack(anchor="w", pady=(0,5))
        
        self.options_canvas = tk.Canvas(right_frame)
        scrollbar = ttk.Scrollbar(right_frame, orient="vertical", command=self.options_canvas.yview)
        self.options_frame = ttk.Frame(self.options_canvas)
        
        self.options_frame.bind("<Configure>", lambda e: self.options_canvas.configure(scrollregion=self.options_canvas.bbox("all")))
        self.options_canvas.create_window((0,0), window=self.options_frame, anchor="nw")
        self.options_canvas.configure(yscrollcommand=scrollbar.set)
        
        self.options_canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        paned_window.add(right_frame, weight=2)
        
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill="x", pady=(10, 0))
        
        self.status_label = ttk.Label(button_frame, text="")
        self.status_label.pack(side="left")
        
        ttk.Button(button_frame, text="해결 완료", command=self._on_confirm).pack(side="right")
        ttk.Button(button_frame, text="추천값으로 전체 자동 해결", command=self._auto_resolve_all).pack(side="right", padx=5)
        ttk.Button(button_frame, text="닫기", command=self._on_close).pack(side="right", padx=5)
        
        if self.kr_listbox.size() > 0:
            self.kr_listbox.select_set(0)
            self._on_kr_select(None)
        
        self._update_status()

    def _on_kr_select(self, event):
        for widget in self.options_frame.winfo_children():
            widget.destroy()
            
        selection_indices = self.kr_listbox.curselection()
        if not selection_indices: return
            
        selected_index = selection_indices[0]
        selected_kr = self.kr_listbox.get(selected_index).lstrip("✔ ").strip()
        
        self.right_frame_label.config(text=f"번역 옵션: \"{selected_kr[:30]}...\"")
        
        options = self.conflicts[selected_kr]
        sorted_options = sorted(options, key=lambda x: x['count'], reverse=True)
        
        # --- [개선] 동률 1위일 경우 추천하지 않는 로직 ---
        is_tie = False
        if len(sorted_options) > 1 and sorted_options[0]['count'] == sorted_options[1]['count']:
            is_tie = True
        # ----------------------------------------------

        self.selected_option_var = tk.StringVar()
        
        if selected_kr in self.resolutions:
            current_cn = self.resolutions[selected_kr]['cn']
            current_tw = self.resolutions[selected_kr]['tw']
            self.selected_option_var.set(f"{current_cn}|{current_tw}")

        for i, option in enumerate(sorted_options):
            cn_text = option['cn']
            tw_text = option['tw']
            count = option['count']
            radio_value = f"{cn_text}|{tw_text}"
            
            # [개선] 동률이 아닐 경우에만 첫 번째 항목을 추천
            recommend_text = " (추천)" if i == 0 and not is_tie else ""
            
            frame = ttk.LabelFrame(self.options_frame, text=f"옵션 {i+1} - 빈도: {count}{recommend_text}", padding=5)
            frame.pack(fill="x", pady=2, padx=2)
            
            rb = ttk.Radiobutton(frame, text="", variable=self.selected_option_var, 
                                 value=radio_value,
                                 command=lambda kr=selected_kr, data=option['data'], index=selected_index: self._on_option_select(kr, data, index))
            rb.pack(side="left")
            
            text_widget = tk.Text(frame, height=3, wrap="word", relief="flat", background=self.cget('bg'))
            text_widget.config(font=self.custom_font)
            text_widget.insert(tk.END, f"CN: {cn_text}\nTW: {tw_text}")
            text_widget.config(state="disabled")
            text_widget.pack(side="left", fill="x", expand=True, padx=5)

            # 이전에 선택한 값이 없으면, 추천 항목(또는 첫 번째 항목)을 기본으로 선택
            if not self.selected_option_var.get() and i == 0:
                rb.invoke()

    def _on_option_select(self, kr_text, data, index):
        """라디오 버튼 선택 시, self.resolutions 딕셔너리를 업데이트하고 UI에 피드백"""
        self.resolutions[kr_text] = data
        
        # --- [개선] 선택 완료 시 리스트박스에 시각적 피드백 ---
        original_text = self.kr_listbox.get(index)
        if not original_text.startswith("✔"):
            self.kr_listbox.delete(index)
            self.kr_listbox.insert(index, f"✔ {original_text}")
        self.kr_listbox.itemconfig(index, {'fg': 'green'})
        # ----------------------------------------------------
        
        self._update_status()

    def _update_status(self):
        resolved_count = len(self.resolutions)
        total_count = len(self.conflicts)
        self.status_label.config(text=f"{resolved_count} / {total_count} 개 해결됨")
        
    def _on_confirm(self):
        if len(self.resolutions) != len(self.conflicts):
            messagebox.showwarning("미해결 항목", "아직 해결되지 않은 충돌 항목이 있습니다.\n모든 항목을 선택해주세요.", parent=self)
            return
        if self.callback:
            self.callback(self.resolutions)
        self.destroy()
        
    def _on_close(self):
        self.destroy()

    def _auto_resolve_all(self):
        if not messagebox.askyesno("자동 해결 확인", "모든 충돌 항목을 가장 빈도가 높은 추천값으로 자동 해결하시겠습니까?\n(빈도가 동일한 경우 첫 번째 항목으로 선택됩니다)", parent=self):
            return

        for index, kr_text_display in enumerate(self.kr_listbox.get(0, tk.END)):
            kr_text = kr_text_display.lstrip("✔ ").strip()
            if kr_text in self.resolutions: # 이미 해결된 항목은 건너뛰기
                continue
                
            options = self.conflicts[kr_text]
            recommended_option = sorted(options, key=lambda x: x['count'], reverse=True)[0]
            
            # 내부적으로 해결된 것으로 저장하고, UI에도 피드백
            self._on_option_select(kr_text, recommended_option['data'], index)

        messagebox.showinfo("완료", "모든 항목이 추천값으로 자동 선택되었습니다.\n'해결 완료' 버튼을 눌러 확정해주세요.", parent=self)


# ==============================================================================
# 메인 UI 클래스
# ==============================================================================
class EnhancedTranslationApplyTool(tk.Frame):
    def __init__(self, parent, excluded_files):
        super().__init__(parent)
        self.parent = parent
        self.translation_apply_manager = EnhancedTranslationApplyManager(self)
        
        self.translation_db_var = tk.StringVar()
        self.excel_source_path_var = tk.StringVar()
        self.original_folder_var = tk.StringVar()
        self.selected_sheets_display_var = tk.StringVar(value="선택된 시트 없음")
        self.selected_sheets = []
        self.available_languages = ["KR", "CN", "TW"]
        self.apply_lang_vars = {}
        self.record_date_var = tk.BooleanVar(value=True)
        self.kr_match_check_var = tk.BooleanVar(value=True)
        self.kr_mismatch_delete_var = tk.BooleanVar(value=False)
        self.use_filtered_data_var = tk.BooleanVar(value=False)
        self.special_column_input_var = tk.StringVar(value="#번역요청")
        self.special_condition_var = tk.StringVar()
        self.filter_status_var = tk.StringVar(value="필터링 데이터 없음")
        self.apply_on_new_var = tk.BooleanVar(value=True)
        self.apply_on_change_var = tk.BooleanVar(value=True)
        self.apply_on_transferred_var = tk.BooleanVar(value=False)
    
        self.original_files = []
        self.excluded_files = excluded_files
        self.cached_excel_path = None
        self.cached_sheet_names = []
        self.detected_special_columns = {}  
        
        self.setup_ui()

    def setup_ui(self):
        main_paned = ttk.PanedWindow(self, orient="horizontal")
        main_paned.pack(fill="both", expand=True, padx=5, pady=5)
        left_frame = ttk.Frame(main_paned, width=800); right_frame = ttk.Frame(main_paned, width=600)
        main_paned.add(left_frame, weight=4); main_paned.add(right_frame, weight=3)
        
        # 좌측 스크롤 가능 프레임 설정
        left_canvas = tk.Canvas(left_frame); left_scrollbar = ttk.Scrollbar(left_frame, orient="vertical", command=left_canvas.yview)
        left_scrollable_frame = ttk.Frame(left_canvas)
        left_scrollable_frame.bind("<Configure>", lambda e: left_canvas.configure(scrollregion=left_canvas.bbox("all")))
        left_canvas.create_window((0, 0), window=left_scrollable_frame, anchor="nw")
        left_canvas.configure(yscrollcommand=left_scrollbar.set)
        left_canvas.pack(side="left", fill="both", expand=True); left_scrollbar.pack(side="right", fill="y")
        
        # 우측 스크롤 가능 프레임 설정
        right_canvas = tk.Canvas(right_frame); right_scrollbar = ttk.Scrollbar(right_frame, orient="vertical", command=right_canvas.yview)
        right_scrollable_frame = ttk.Frame(right_canvas)
        right_scrollable_frame.bind("<Configure>", lambda e: right_canvas.configure(scrollregion=right_canvas.bbox("all")))
        right_canvas.create_window((0, 0), window=right_scrollable_frame, anchor="nw")
        right_canvas.configure(yscrollcommand=right_scrollbar.set)
        right_canvas.pack(side="left", fill="both", expand=True); right_scrollbar.pack(side="right", fill="y")


        source_selection_frame = ttk.LabelFrame(left_scrollable_frame, text="🔧 번역 데이터 소스 선택")
        source_selection_frame.pack(fill="x", padx=5, pady=5)
        db_frame = ttk.Frame(source_selection_frame)
        db_frame.pack(fill="x", padx=5, pady=3)
        ttk.Label(db_frame, text="번역 DB:").grid(row=0, column=0, padx=5, pady=2, sticky="w")
        ttk.Entry(db_frame, textvariable=self.translation_db_var).grid(row=0, column=1, padx=5, pady=2, sticky="ew")
        ttk.Button(db_frame, text="찾아보기", command=self.select_translation_db_file).grid(row=0, column=2, padx=5, pady=2)
        db_frame.columnconfigure(1, weight=1)
        ttk.Separator(source_selection_frame, orient="horizontal").pack(fill="x", padx=5, pady=5)
        excel_frame = ttk.Frame(source_selection_frame)
        excel_frame.pack(fill="x", padx=5, pady=3)
        ttk.Label(excel_frame, text="엑셀 파일:").grid(row=0, column=0, padx=5, pady=2, sticky="w")
        ttk.Entry(excel_frame, textvariable=self.excel_source_path_var, state="readonly").grid(row=0, column=1, padx=5, pady=2, sticky="ew")
        ttk.Button(excel_frame, text="찾아보기", command=self.select_excel_source_file).grid(row=0, column=2, padx=5, pady=2)
        excel_frame.columnconfigure(1, weight=1)
        sheet_frame = ttk.Frame(source_selection_frame)
        sheet_frame.pack(fill="x", padx=5, pady=3)
        ttk.Label(sheet_frame, text="시트 선택:").grid(row=0, column=0, padx=5, pady=2, sticky="w")
        ttk.Entry(sheet_frame, textvariable=self.selected_sheets_display_var, state="readonly").grid(row=0, column=1, padx=5, pady=2, sticky="ew")
        self.sheet_select_button = ttk.Button(sheet_frame, text="시트 선택", command=self.open_sheet_selection_popup, state="disabled")
        self.sheet_select_button.grid(row=0, column=2, padx=5, pady=2)
        self.rebuild_cache_button = ttk.Button(sheet_frame, text="캐시 재구축", command=self.force_rebuild_cache, state="disabled")
        self.rebuild_cache_button.grid(row=0, column=3, padx=5, pady=2) 
        sheet_frame.columnconfigure(1, weight=1)
        
        filter_frame = ttk.LabelFrame(left_scrollable_frame, text="🔍 특수 컬럼 필터링 (고급 옵션)")
        filter_frame.pack(fill="x", padx=5, pady=5)
        filter_enable_frame = ttk.Frame(filter_frame)
        filter_enable_frame.pack(fill="x", padx=5, pady=3)
        ttk.Checkbutton(filter_enable_frame, text="특수 컬럼 필터링된 데이터만 적용", variable=self.use_filtered_data_var, command=self.toggle_filter_options).pack(side="left")
        self.filter_config_frame = ttk.Frame(filter_frame)
        self.filter_config_frame.pack(fill="x", padx=15, pady=5)
        col_frame = ttk.Frame(self.filter_config_frame)
        col_frame.pack(fill="x", pady=2)
        ttk.Label(col_frame, text="컬럼명:").grid(row=0, column=0, padx=5, pady=2, sticky="w")
        ttk.Entry(col_frame, textvariable=self.special_column_input_var, width=25).grid(row=0, column=1, padx=5, pady=2, sticky="ew")
        self.detect_button = ttk.Button(col_frame, text="감지", command=self.detect_special_column, state="disabled")
        self.detect_button.grid(row=0, column=2, padx=5, pady=2)
        col_frame.columnconfigure(1, weight=1)
        condition_frame = ttk.Frame(self.filter_config_frame)
        condition_frame.pack(fill="x", pady=2)
        ttk.Label(condition_frame, text="조건값:").grid(row=0, column=0, padx=5, pady=2, sticky="w")
        ttk.Entry(condition_frame, textvariable=self.special_condition_var, width=25).grid(row=0, column=1, padx=5, pady=2, sticky="ew")
        condition_frame.columnconfigure(1, weight=1)
        status_frame = ttk.Frame(self.filter_config_frame)
        status_frame.pack(fill="x", pady=2)
        ttk.Label(status_frame, text="상태:").grid(row=0, column=0, padx=5, pady=2, sticky="w")
        ttk.Label(status_frame, textvariable=self.filter_status_var, foreground="blue").grid(row=0, column=1, padx=5, pady=2, sticky="w")
        self.toggle_filter_options()

        options_frame = ttk.LabelFrame(left_scrollable_frame, text="⚙️ 적용 옵션")
        options_frame.pack(fill="x", padx=5, pady=5)
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
        ttk.Checkbutton(self.id_based_options_frame, text="KR 일치 검사", variable=self.kr_match_check_var, command=self.toggle_kr_options).pack(anchor="w", padx=5, pady=1)
        self.id_mismatch_delete_cb = ttk.Checkbutton(self.id_based_options_frame, text="└ KR 불일치 시 다국어 제거", variable=self.kr_mismatch_delete_var)
        self.id_mismatch_delete_cb.pack(anchor="w", padx=20, pady=1)
        self.id_overwrite_cb = ttk.Checkbutton(self.id_based_options_frame, text="└ 선택 언어 덮어쓰기", variable=self.kr_overwrite_var)
        self.id_overwrite_cb.pack(anchor="w", padx=20, pady=1)
        self.kr_based_options_frame = ttk.Frame(options_frame)
        self.kr_overwrite_on_kr_mode_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(self.kr_based_options_frame, text="선택 언어 덮어쓰기", variable=self.kr_overwrite_on_kr_mode_var).pack(anchor="w", padx=5, pady=1)
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

        original_files_frame = ttk.LabelFrame(right_scrollable_frame, text="📁 번역을 적용할 원본 파일")
        original_files_frame.pack(fill="x", padx=5, pady=5)
        folder_frame = ttk.Frame(original_files_frame)
        folder_frame.pack(fill="x", padx=5, pady=3)
        ttk.Label(folder_frame, text="원본 폴더:").grid(row=0, column=0, padx=5, pady=2, sticky="w")
        ttk.Entry(folder_frame, textvariable=self.original_folder_var).grid(row=0, column=1, padx=5, pady=2, sticky="ew")
        ttk.Button(folder_frame, text="찾아보기", command=self.select_original_folder).grid(row=0, column=2, padx=5, pady=2)
        ttk.Button(folder_frame, text="파일 검색", command=self.search_original_files).grid(row=0, column=3, padx=5, pady=2)
        folder_frame.columnconfigure(1, weight=1)
        files_list_frame = ttk.LabelFrame(right_scrollable_frame, text="원본 파일 목록")
        files_list_frame.pack(fill="both", expand=True, padx=5, pady=5)
        self.original_files_list = ScrollableCheckList(files_list_frame, height=200)
        self.original_files_list.pack(fill="both", expand=True, padx=5, pady=5)
        action_frame = ttk.Frame(right_scrollable_frame)
        action_frame.pack(fill="x", padx=5, pady=5)
        
        data_button_frame = ttk.Frame(action_frame)
        data_button_frame.pack(fill="x", pady=2)
        
        self.load_data_button = ttk.Button(data_button_frame, text="번역 데이터 로드", command=self.load_translation_data, state="disabled")
        self.load_data_button.pack(side="left", padx=5)
        self.view_data_button = ttk.Button(data_button_frame, text="로드된 데이터 보기", command=self.show_loaded_data_viewer, state="disabled")
        self.view_data_button.pack(side="left", padx=5)
        # [신규] 충돌 보기 버튼 추가
        self.view_conflicts_button = ttk.Button(data_button_frame, text="번역 충돌 보기", command=self.open_conflict_resolver_popup, state="disabled")
        self.view_conflicts_button.pack(side="left", padx=5)
        self.view_overwritten_button = ttk.Button(data_button_frame, text="덮어쓴 데이터 보기", command=self.show_overwritten_data_viewer, state="disabled")
        self.view_overwritten_button.pack(side="left", padx=5)
        
        exec_button_frame = ttk.Frame(action_frame)
        exec_button_frame.pack(fill="x", pady=2)
        
        self.apply_button = ttk.Button(exec_button_frame, text="🚀 번역 적용", command=self.apply_translation, state="disabled")
        self.apply_button.pack(side="right", padx=5)
        
        log_frame = ttk.LabelFrame(right_scrollable_frame, text="📋 작업 로그")
        log_frame.pack(fill="both", expand=True, padx=5, pady=5)
        self.log_text = tk.Text(log_frame, wrap="word", height=15)
        log_scrollbar_v = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_scrollbar_v.set)
        log_scrollbar_v.pack(side="right", fill="y")
        self.log_text.pack(fill="both", expand=True)
        status_frame_bottom = ttk.Frame(right_scrollable_frame)
        status_frame_bottom.pack(fill="x", padx=5, pady=5)
        self.status_label_apply = ttk.Label(status_frame_bottom, text="대기 중... 번역 엑셀 파일을 선택하세요.")
        self.status_label_apply.pack(side="left", padx=5)
        self.progress_bar = ttk.Progressbar(status_frame_bottom, orient="horizontal", mode="determinate")
        left_canvas.pack(side="left", fill="both", expand=True)
        left_scrollbar.pack(side="right", fill="y")
        right_canvas.pack(side="left", fill="both", expand=True)
        right_scrollbar.pack(side="right", fill="y")
        self.toggle_options_by_mode()
        self.toggle_kr_options()

    def toggle_filter_options(self):
        if self.use_filtered_data_var.get():
            self.filter_config_frame.pack(fill="x", padx=15, pady=5)
        else:
            self.filter_config_frame.pack_forget()

    def detect_special_column(self):
        excel_path = self.excel_source_path_var.get()
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
                result = self.translation_apply_manager.detect_special_column_in_excel(excel_path, column_name)
                self.after(0, lambda: self.process_special_detection_result(result, loading_popup, column_name))
            except Exception as e:
                self.after(0, lambda: [
                    loading_popup.close(),
                    self.log_text.insert(tk.END, f"특수 컬럼 감지 중 오류: {str(e)}\n"),
                    messagebox.showerror("오류", f"특수 컬럼 감지 중 오류 발생: {str(e)}", parent=self)
                ])
        threading.Thread(target=detect_thread, daemon=True).start()


    def process_special_detection_result(self, result, loading_popup, column_name):
        loading_popup.close()
        
        if result["status"] == "error":
            messagebox.showerror("오류", f"특수 컬럼 감지 실패: {result['message']}", parent=self)
            return
        
        detected_info = result.get("detected_info", {})
        
        if detected_info:
            self.detected_special_columns = {column_name: detected_info}
            
            # ▼▼▼ [문제의 코드] 이 부분을 주석 처리 또는 삭제 ▼▼▼
            # if detected_info.get('unique_values'):
            #     most_common = detected_info.get('most_common', [])
            #     if most_common:
            #         # 이 라인이 사용자의 입력을 덮어쓰는 원인이었습니다.
            #         self.special_condition_var.set(most_common[0][0]) 
            # ▲▲▲ 여기까지 수정 ▲▲▲
            
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

    def toggle_options_by_mode(self, *args):
        mode = self.apply_mode_var.get()
        if mode == "id":
            self.kr_based_options_frame.pack_forget()
            self.id_based_options_frame.pack(fill="x", padx=15, pady=5)
        elif mode == "kr":
            self.id_based_options_frame.pack_forget()
            self.kr_based_options_frame.pack(fill="x", padx=15, pady=5)
        self.toggle_kr_options()

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
        file_path = filedialog.askopenfilename(filetypes=[("Excel 파일", "*.xlsx"), ("모든 파일", "*.*")], title="번역 엑셀 파일 선택", parent=self)
        if not file_path: return
        self.excel_source_path_var.set(file_path)
        self.translation_db_var.set("")
        self.selected_sheets = []
        self.selected_sheets_display_var.set("DB 캐시 생성 필요")
        self.cached_sheet_names = []
        self.filter_status_var.set("필터링 데이터 없음")
        self.initial_info_shown = False
        self.rebuild_cache_button.config(state="disabled") 
        self.sheet_select_button.config(state="disabled")
        self.detect_button.config(state="disabled")
        self.load_data_button.config(state="disabled")
        self.apply_button.config(state="disabled")
        self.view_data_button.config(state="disabled")
        self._start_db_caching_thread(file_path)

    def _start_db_caching_thread(self, excel_path, force=False):
        self.progress_bar.pack(side="left", fill="x", expand=True, padx=5)
        self.progress_bar["value"] = 0
        self.status_label_apply.config(text="DB 캐시 구축 준비 중...")
        self.sheet_select_button.config(state="disabled")
        self.rebuild_cache_button.config(state="disabled")
        self.load_data_button.config(state="disabled")
        def progress_callback(percentage, message):
            self.after(0, lambda: [
                self.progress_bar.config(value=percentage),
                self.status_label_apply.config(text=message)
            ])
        def task():
            result = self.translation_apply_manager.initiate_excel_caching(excel_path, force_rebuild=force, progress_callback=progress_callback)
            self.after(0, lambda: self._process_caching_result(result, excel_path))
        threading.Thread(target=task, daemon=True).start()

    def _process_caching_result(self, result, excel_path):
        self.progress_bar.pack_forget() 
        if result["status"] == "error":
            self.status_label_apply.config(text="DB 캐시 구축 실패.")
            messagebox.showerror("캐시 생성 오류", result["message"], parent=self)
            return
        self.cached_excel_path = excel_path
        self.cached_sheet_names = result.get("sheets", [])
        self.status_label_apply.config(text=f"DB 캐시 준비 완료. 시트를 선택하세요. (총 {len(self.cached_sheet_names)}개)")
        self.selected_sheets_display_var.set(f"{len(self.cached_sheet_names)}개 시트 발견됨")
        self.sheet_select_button.config(state="normal")
        self.rebuild_cache_button.config(state="normal")
        self.detect_button.config(state="normal")
        if not hasattr(self, 'initial_info_shown') or self.initial_info_shown is False:
                messagebox.showinfo("캐시 준비 완료", f"'{os.path.basename(excel_path)}'에 대한 DB 캐시가 준비되었습니다.\n'시트 선택' 버튼으로 작업할 시트를 골라주세요.", parent=self)
                self.initial_info_shown = True

    def load_translation_data(self):
        db_path = self.translation_db_var.get()
        excel_path = self.excel_source_path_var.get()
        if db_path:
            self.load_from_db(db_path)
        elif excel_path:
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
            loading_popup = LoadingPopup(self, "메모리 캐싱 중", "번역 데이터 로딩 중...")
            def task():
                result = self.translation_apply_manager.load_translation_cache_from_excel_with_filter(excel_path, self.selected_sheets, special_filter)
                self.after(0, lambda: self.process_cache_load_result(result, loading_popup))
            threading.Thread(target=task, daemon=True).start()
        else:
            messagebox.showwarning("경고", "번역 DB 또는 엑셀 파일을 선택하세요.", parent=self)

    def open_sheet_selection_popup(self):
        if not self.cached_sheet_names:
            messagebox.showwarning("오류", "사용 가능한 시트 목록이 없습니다.", parent=self)
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
                self.load_data_button.config(state="normal")
            else:
                self.selected_sheets_display_var.set("선택된 시트 없음")
                self.load_data_button.config(state="disabled")
            popup.destroy()
        ttk.Button(popup, text="확인", command=on_confirm).pack(pady=10)

    def load_from_db(self, db_path):
        if not os.path.isfile(db_path):
            messagebox.showwarning("경고", "유효한 번역 DB 파일을 선택하세요.", parent=self)
            return
        self.log_text.insert(tk.END, "번역 DB 캐싱 중...\n")
        loading_popup = LoadingPopup(self, "DB 캐싱 중", "번역 데이터 캐싱 중...")
        def task():
            result = self.translation_apply_manager.load_translation_cache_from_excel_with_filter(db_path, [], None) # DB 로드도 동일한 함수 사용
            self.after(0, lambda: self.process_cache_load_result(result, loading_popup))
        threading.Thread(target=task, daemon=True).start()

    def process_cache_load_result(self, result, loading_popup):
        """[수정] 충돌 보기 버튼 활성화 로직 추가"""
        loading_popup.close()
        if result["status"] == "error":
            messagebox.showerror("오류", f"캐싱 중 오류 발생: {result['message']}", parent=self)
            self.log_text.insert(tk.END, f"캐싱 실패: {result['message']}\n")
            return
        
        id_count = result.get("id_count", 0)
        conflict_count = result.get("conflict_count", 0)

        self.status_label_apply.config(text=f"데이터 로드 완료 - {id_count}개 항목")
        self.view_data_button.config(state="normal")
        
        if conflict_count > 0:
            self.apply_button.config(state="disabled")
            self.view_conflicts_button.config(state="normal") # [신규] 충돌 보기 버튼 활성화
            messagebox.showwarning("번역 충돌 발견", 
                                   f"{conflict_count}개의 KR 텍스트에서 번역 충돌이 발견되었습니다.\n'번역 충돌 보기' 버튼을 눌러 해결해주세요.",
                                   parent=self)
            self.open_conflict_resolver_popup()
        else:
            self.apply_button.config(state="normal")
            self.view_conflicts_button.config(state="disabled") # 충돌 없으면 비활성화
            messagebox.showinfo("완료", f"번역 데이터 로딩 완료!\n로드된 항목 수: {id_count}개", parent=self)


    def open_conflict_resolver_popup(self):
        conflicts = self.translation_apply_manager.get_translation_conflicts()
        if not conflicts: return
        ConflictResolverPopup(self, conflicts, self.on_conflicts_resolved)

    def on_conflicts_resolved(self, resolutions):
        """[수정] 충돌 보기 버튼 비활성화 로직 추가"""
        self.log_text.insert(tk.END, f"사용자 선택에 따라 번역 충돌 해결 중...\n")
        self.translation_apply_manager.update_resolved_translations(resolutions)
        
        self.apply_button.config(state="normal")
        self.view_conflicts_button.config(state="disabled") # [신규] 해결 완료 후 버튼 비활성화
        self.status_label_apply.config(text="번역 충돌 해결 완료. 적용 준비됨.")
        messagebox.showinfo("해결 완료", "모든 번역 충돌이 해결되었습니다.\n이제 '번역 적용'을 진행할 수 있습니다.", parent=self)
  
    def apply_translation(self):
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
        self.log_text.insert(tk.END, "="*60 + "\n🚀 번역 적용 작업 시작\n")
        self.log_text.insert(tk.END, f"- 적용 모드: {mode_text}, 적용 언어: {lang_text}\n")
        self.log_text.insert(tk.END, f"- 대상 파일: {len(files_to_process)}개\n" + "="*60 + "\n")
        
        loading_popup = LoadingPopup(self, "번역 적용 중", "번역 적용 준비 중...")
        apply_options = {
            "mode": self.apply_mode_var.get(), "selected_langs": selected_langs,
            "record_date": self.record_date_var.get(), "kr_match_check": self.kr_match_check_var.get(),
            "kr_mismatch_delete": self.kr_mismatch_delete_var.get(), "kr_overwrite": self.kr_overwrite_var.get(),
            "kr_overwrite_on_kr_mode": self.kr_overwrite_on_kr_mode_var.get(),
            "allowed_statuses": allowed_statuses, "use_filtered_data": use_filtered
        }
            
        def apply_translations_thread():
            total_results = defaultdict(int)
            processed_count = 0
            error_count = 0
            modified_files = []
            total_overwritten_items = []
            start_time = time.time()
            for idx, (file_name, file_path) in enumerate(files_to_process):
                self.after(0, lambda i=idx, n=file_name: loading_popup.update_progress((i / len(files_to_process)) * 100, f"파일 처리 중 ({i+1}/{len(files_to_process)}): {n}"))
                result = self.translation_apply_manager.apply_translation_with_filter_option(file_path, apply_options)
                if result["status"] == "success":
                    processed_count += 1
                    for key, value in result.items():
                        if key.startswith("total_"): total_results[key] += value
                    if result.get("total_updated", 0) > 0 or result.get("total_overwritten", 0) > 0:
                        modified_files.append(file_name)
                    total_overwritten_items.extend(result.get("overwritten_items", []))
                else: error_count += 1
            elapsed_time = time.time() - start_time
            self.after(0, lambda: self.process_translation_apply_result(
                total_results, processed_count, error_count, loading_popup, 
                elapsed_time, use_filtered, modified_files, total_overwritten_items)
            )
        threading.Thread(target=apply_translations_thread, daemon=True).start()

    def process_translation_apply_result(self, total_results, processed_count, error_count, loading_popup, 
                                        elapsed_time, use_filtered,
                                        modified_files, total_overwritten_items):
        loading_popup.close()
        minutes, seconds = divmod(int(elapsed_time), 60)
        time_str = f"{minutes}분 {seconds}초" if minutes > 0 else f"{seconds}초"

        self.log_text.insert(tk.END, "\n" + "="*60 + "\n🎉 번역 적용 작업 완료\n" + "="*60 + "\n")
        self.log_text.insert(tk.END, f"⏱️ 소요 시간: {time_str}, 성공: {processed_count}개, 실패: {error_count}개\n")
        if modified_files:
            self.log_text.insert(tk.END, f"🔄 변경된 파일 ({len(modified_files)}개): {', '.join(modified_files[:5])}{'...' if len(modified_files)>5 else ''}\n")
        
        self.log_text.insert(tk.END, f"📊 작업 통계:\n")
        if total_results['total_updated'] > 0: self.log_text.insert(tk.END, f"  • 신규 적용: {total_results['total_updated']:,}개\n")
        if total_results['total_overwritten'] > 0: self.log_text.insert(tk.END, f"  • 덮어쓰기: {total_results['total_overwritten']:,}개\n")
        if total_results['total_conditional_skipped'] > 0: self.log_text.insert(tk.END, f"  • 조건 불일치 스킵: {total_results['total_conditional_skipped']:,}개\n")
        if total_results['total_kr_mismatch_skipped'] > 0: self.log_text.insert(tk.END, f"  • KR 불일치 스킵: {total_results['total_kr_mismatch_skipped']:,}개\n")
        
        total_applied = total_results["total_updated"] + total_results["total_overwritten"]
        self.log_text.insert(tk.END, f"🎯 총 적용된 번역: {total_applied:,}개\n" + "="*60 + "\n")
        
        self.overwritten_data = total_overwritten_items
        if self.overwritten_data:
            self.view_overwritten_button.config(state="normal")
        else:
            self.view_overwritten_button.config(state="disabled")
        messagebox.showinfo("완료", f"작업 완료!\n\n총 적용: {total_applied:,}개\n소요 시간: {time_str}", parent=self)

    def select_translation_db_file(self, *args):
        file_path = filedialog.askopenfilename(filetypes=[("DB 파일", "*.db"), ("모든 파일", "*.*")], title="번역 DB 선택", parent=self)
        if file_path:
            self.translation_db_var.set(file_path)
            self.excel_source_path_var.set("") 
            self.load_data_button.config(state="normal")

    def select_original_folder(self):
        folder = filedialog.askdirectory(title="원본 파일 폴더 선택", parent=self)
        if folder: self.original_folder_var.set(folder)

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

    def _check_files_are_open(self, file_paths_to_check):
        open_files = []
        for file_path in file_paths_to_check:
            if not os.path.exists(file_path): continue
            try: os.rename(file_path, file_path)
            except OSError: open_files.append(os.path.basename(file_path))
        return open_files

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
        for col in columns: tree.heading(col, text=col.upper())
        tree.column("string_id", width=200); tree.column("kr", width=300); tree.column("cn", width=250); tree.column("tw", width=250)
        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        vsb.pack(side="right", fill="y")
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=tree.xview)
        hsb.pack(side="bottom", fill="x")
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        tree.pack(fill="both", expand=True)

        status_label = ttk.Label(viewer_win, text="", padding=5)
        status_label.pack(side="bottom", fill="x")

        all_data = list(self.translation_apply_manager.translation_cache.values())
        def populate_tree(data_to_show):
            tree.delete(*tree.get_children())
            for item in data_to_show:
                values = (item.get('string_id',''), item.get('kr',''), item.get('cn',''), item.get('tw',''), item.get('file_name',''), item.get('sheet_name',''))
                tree.insert("", "end", values=values)
            status_label.config(text=f"{len(data_to_show):,} / {len(all_data):,}개 항목 표시 중")
        def perform_search():
            id_query = id_search_var.get().lower().strip()
            kr_query = kr_search_var.get().lower().strip()
            if not id_query and not kr_query: populate_tree(all_data); return
            filtered_data = [item for item in all_data if (id_query in item.get('string_id', '').lower()) and (kr_query in item.get('kr', '').lower())]
            populate_tree(filtered_data)
        
        search_button = ttk.Button(search_frame, text="검색", command=perform_search)
        search_button.pack(side="left", padx=5)
        id_search_entry.bind("<Return>", lambda event: perform_search())
        kr_search_entry.bind("<Return>", lambda event: perform_search())
        populate_tree(all_data)

    def force_rebuild_cache(self):
        excel_path = self.excel_source_path_var.get()
        if not excel_path: return
        if messagebox.askyesno("캐시 재구축 확인", "현재 엑셀 파일의 캐시를 강제로 다시 만드시겠습니까?", parent=self):
            self._start_db_caching_thread(excel_path, force=True)

    def show_overwritten_data_viewer(self):
        if not hasattr(self, 'overwritten_data') or not self.overwritten_data:
            messagebox.showinfo("정보", "표시할 덮어쓰기 데이터가 없습니다.", parent=self)
            return
        viewer_win = tk.Toplevel(self)
        viewer_win.title(f"덮어쓴 데이터 보기 ({len(self.overwritten_data)}개)")
        viewer_win.geometry("1300x700")
        viewer_win.transient(self)
        viewer_win.grab_set()
        tree_frame = ttk.Frame(viewer_win, padding="5")
        tree_frame.pack(fill="both", expand=True)
        columns = ("file_name", "sheet_name", "string_id", "language", "kr_text", "original_text", "overwritten_text")
        tree = ttk.Treeview(tree_frame, columns=columns, show="headings")
        tree.heading("file_name", text="파일명"); tree.heading("sheet_name", text="시트명"); tree.heading("string_id", text="STRING_ID"); tree.heading("language", text="언어"); tree.heading("kr_text", text="KR 원문"); tree.heading("original_text", text="원본 내용"); tree.heading("overwritten_text", text="덮어쓴 내용")
        tree.column("file_name", width=150, anchor="w"); tree.column("sheet_name", width=100, anchor="w"); tree.column("string_id", width=180, anchor="w"); tree.column("language", width=50, anchor="center"); tree.column("kr_text", width=250, anchor="w"); tree.column("original_text", width=250, anchor="w"); tree.column("overwritten_text", width=250, anchor="w")
        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        vsb.pack(side="right", fill="y")
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=tree.xview)
        hsb.pack(side="bottom", fill="x")
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        tree.pack(fill="both", expand=True)
        for item in self.overwritten_data:
            values = (item.get('file_name', ''), item.get('sheet_name', ''), item.get('string_id', ''), item.get('language', ''), item.get('kr_text', ''), item.get('original_text', ''), item.get('overwritten_text', ''))
            tree.insert("", "end", values=values)
        button_frame = ttk.Frame(viewer_win, padding="5")
        button_frame.pack(fill="x")
        ttk.Button(button_frame, text="닫기", command=viewer_win.destroy).pack(side="right")