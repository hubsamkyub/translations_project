# tools/translate/translation_apply_tool.py (수정 후)

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import threading
import time
import sys # sys 모듈 추가

# --- 경로 문제 해결을 위한 코드 ---
project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if project_root not in sys.path:
    sys.path.append(project_root)
# ---------------------------------

from ui.common_components import ScrollableCheckList, LoadingPopup
from tools.translation_apply_manager import TranslationApplyManager
import openpyxl

class TranslationApplyTool(tk.Frame):
    def __init__(self, parent, excluded_files):
        super().__init__(parent)
        self.parent = parent
        self.translation_apply_manager = TranslationApplyManager(self)
        
        # --- UI 변수 선언 ---
        # DB 소스
        self.translation_db_var = tk.StringVar()
        # 엑셀 소스
        self.excel_source_path_var = tk.StringVar()
        self.excel_sheet_var = tk.StringVar()
        # 공통
        self.original_folder_var = tk.StringVar()
        self.record_date_var = tk.BooleanVar(value=True)
        self.available_languages = ["KR", "EN", "CN", "TW", "TH"]
        self.apply_lang_vars = {}

        # --- 내부 데이터 ---
        self.original_files = []
        self.excluded_files = excluded_files
        
        self.setup_ui()

    def setup_ui(self):
        """번역 적용 탭 UI 구성 (좌/우 분할 레이아웃)"""

        # --- 상단 소스 선택 프레임 (좌/우 분할) ---
        source_selection_frame = ttk.Frame(self)
        source_selection_frame.pack(fill="x", padx=5, pady=5)
        source_selection_frame.columnconfigure(0, weight=1)
        source_selection_frame.columnconfigure(1, weight=1)

        # --- 좌측 프레임: 번역 DB 선택 ---
        db_frame = ttk.LabelFrame(source_selection_frame, text="옵션 1: 번역 DB 선택")
        db_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 5))
        
        ttk.Label(db_frame, text="번역 DB:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        db_entry = ttk.Entry(db_frame, textvariable=self.translation_db_var)
        db_entry.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        ttk.Button(db_frame, text="찾아보기", command=self.select_translation_db_file).grid(row=0, column=2, padx=5, pady=5)
        db_frame.columnconfigure(1, weight=1)

        # --- 우측 프레임: 번역 엑셀 파일 선택 ---
        excel_frame = ttk.LabelFrame(source_selection_frame, text="옵션 2: 번역 엑셀 파일 선택")
        excel_frame.grid(row=0, column=1, sticky="nsew", padx=(5, 0))

        ttk.Label(excel_frame, text="엑셀 파일:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        excel_entry = ttk.Entry(excel_frame, textvariable=self.excel_source_path_var)
        excel_entry.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        ttk.Button(excel_frame, text="찾아보기", command=self.select_excel_source_file).grid(row=0, column=2, padx=5, pady=5)
        
        ttk.Label(excel_frame, text="시트 선택:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.sheet_combobox = ttk.Combobox(excel_frame, textvariable=self.excel_sheet_var, state="readonly")
        self.sheet_combobox.grid(row=1, column=1, columnspan=2, padx=5, pady=5, sticky="ew")
        excel_frame.columnconfigure(1, weight=1)

        # --- 원본 파일 및 옵션 (공통 영역) ---
        original_files_frame = ttk.LabelFrame(self, text="번역을 적용할 원본 파일")
        original_files_frame.pack(fill="x", padx=5, pady=5)
        
        ttk.Label(original_files_frame, text="원본 폴더:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        ttk.Entry(original_files_frame, textvariable=self.original_folder_var).grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        ttk.Button(original_files_frame, text="찾아보기", command=self.select_original_folder).grid(row=0, column=2, padx=5, pady=5)
        ttk.Button(original_files_frame, text="파일 검색", command=self.search_original_files).grid(row=0, column=3, padx=5, pady=5)
        original_files_frame.columnconfigure(1, weight=1)
        
        files_list_frame = ttk.LabelFrame(self, text="원본 파일 목록")
        files_list_frame.pack(fill="both", expand=True, padx=5, pady=5)
        self.original_files_list = ScrollableCheckList(files_list_frame)
        self.original_files_list.pack(fill="both", expand=True, padx=5, pady=5)

        options_frame = ttk.LabelFrame(self, text="적용 옵션")
        options_frame.pack(fill="x", padx=5, pady=5)
        for i, lang in enumerate(self.available_languages):
            var = tk.BooleanVar(value=True if lang in ["CN", "TW"] else False)
            self.apply_lang_vars[lang] = var
            ttk.Checkbutton(options_frame, text=lang, variable=var).grid(row=i // 5, column=i % 5, padx=10, pady=5, sticky="w")
        ttk.Checkbutton(options_frame, text="번역 적용 표시", variable=self.record_date_var).grid(row=1, column=0, columnspan=5, padx=5, pady=5, sticky="w")
        
        action_frame = ttk.Frame(self)
        action_frame.pack(fill="x", padx=5, pady=5)
        ttk.Button(action_frame, text="번역 적용", command=self.apply_translation).pack(side="right", padx=5, pady=5)
        ttk.Button(action_frame, text="번역 데이터 로드", command=self.load_translation_data).pack(side="right", padx=5, pady=5)
        
        log_frame = ttk.LabelFrame(self, text="작업 로그")
        log_frame.pack(fill="both", expand=True, padx=5, pady=5)
        self.log_text = tk.Text(log_frame, wrap="word")
        scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        self.log_text.pack(fill="both", expand=True)
        
        status_frame = ttk.Frame(self)
        status_frame.pack(fill="x", padx=5, pady=5)
        self.status_label_apply = ttk.Label(status_frame, text="대기 중...")
        self.status_label_apply.pack(side="left", padx=5)

    def select_excel_source_file(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel 파일", "*.xlsx"), ("모든 파일", "*.*")],
            title="번역 엑셀 파일 선택", parent=self
        )
        if file_path:
            self.excel_source_path_var.set(file_path)
            self.translation_db_var.set("") # 다른 옵션 초기화
            self.excel_sheet_var.set("")
            self.sheet_combobox.set('')
            self.sheet_combobox['values'] = []
            
            # 백그라운드 스레드에서 시트 목록 로드
            threading.Thread(target=self._populate_sheets, args=(file_path,), daemon=True).start()

    def _populate_sheets(self, file_path):
        """엑셀 파일에서 시트 목록을 읽어 콤보박스를 채웁니다."""
        try:
            self.after(0, lambda: self.sheet_combobox.set("시트 목록 읽는 중..."))
            workbook = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
            sheet_names = workbook.sheetnames
            
            def update_combobox():
                self.sheet_combobox['values'] = sheet_names
                if sheet_names:
                    self.sheet_combobox.set(sheet_names[0]) # 첫 번째 시트를 기본값으로
                self.sheet_combobox.config(state="readonly")

            self.after(0, update_combobox)
        except Exception as e:
            self.after(0, lambda: messagebox.showerror("오류", f"엑셀 파일 시트를 읽는 중 오류 발생: {e}", parent=self))
            self.after(0, lambda: self.sheet_combobox.set("시트 읽기 실패"))

    def load_translation_data(self):
        """DB 또는 엑셀 파일로부터 번역 데이터를 로드하는 분기 처리"""
        db_path = self.translation_db_var.get()
        excel_path = self.excel_source_path_var.get()

        if db_path:
            self.load_from_db(db_path)
        elif excel_path:
            sheet_name = self.excel_sheet_var.get()
            if not sheet_name or sheet_name.startswith("시트"):
                messagebox.showwarning("경고", "데이터를 읽어올 시트를 선택하세요.", parent=self)
                return
            self.load_from_excel(excel_path, sheet_name)
        else:
            messagebox.showwarning("경고", "번역 DB 또는 엑셀 파일을 선택하세요.", parent=self)

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

    def load_from_excel(self, file_path, sheet_name):
        self.log_text.insert(tk.END, f"'{os.path.basename(file_path)}' 파일의 '{sheet_name}' 시트 캐싱 중...\n")
        loading_popup = LoadingPopup(self, "엑셀 캐싱 중", "번역 데이터 캐싱 중...")
        
        def task():
            result = self.translation_apply_manager.load_translation_cache_from_excel(file_path, sheet_name)
            self.after(0, lambda: self.process_cache_load_result(result, loading_popup))
        
        threading.Thread(target=task, daemon=True).start()
   
   
    def select_translation_db_file(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("DB 파일", "*.db"), ("모든 파일", "*.*")],
            title="번역 DB 선택", parent=self
        )
        if file_path:
            self.translation_db_var.set(file_path)
            self.excel_source_path_var.set("") # 다른 옵션 초기화
            self.excel_sheet_var.set("")
            self.sheet_combobox.set('')
            self.sheet_combobox['values'] = []

    def select_original_folder(self):
        folder = filedialog.askdirectory(title="원본 파일 폴더 선택", parent=self)
        if folder:
            self.original_folder_var.set(folder)

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

    def load_translation_cache(self):
        db_path = self.translation_db_var.get()
        if not db_path or not os.path.isfile(db_path):
            messagebox.showwarning("경고", "유효한 번역 DB 파일을 선택하세요.", parent=self)
            return
        
        self.log_text.insert(tk.END, "번역 DB 캐싱 중...\n")
        self.update()
        
        loading_popup = LoadingPopup(self, "번역 DB 캐싱 중", "번역 데이터 캐싱 중...")
        
        def load_cache():
            try:
                result = self.translation_apply_manager.load_translation_cache(db_path)
                self.after(0, lambda: self.process_cache_load_result(result, loading_popup))
            except Exception as e:
                error_msg = str(e)
                self.after(0, lambda: [
                    loading_popup.close(),
                    self.log_text.insert(tk.END, f"캐싱 중 오류 발생: {error_msg}\n"),
                    self.status_label_apply.config(text="오류 발생"),
                    messagebox.showerror("오류", f"DB 캐싱 중 오류 발생: {error_msg}", parent=self)
                ])
                
        thread = threading.Thread(target=load_cache, daemon=True)
        thread.start()
        
    def process_cache_load_result(self, result, loading_popup):
        loading_popup.close()
        
        if result["status"] == "error":
            messagebox.showerror("오류", f"캐싱 중 오류 발생: {result['message']}", parent=self)
            self.log_text.insert(tk.END, f"캐싱 실패: {result['message']}\n")
            return
            
        self.translation_apply_manager.translation_cache = result["translation_cache"]
        self.translation_apply_manager.translation_file_cache = result["translation_file_cache"]
        self.translation_apply_manager.translation_sheet_cache = result["translation_sheet_cache"]
        self.translation_apply_manager.duplicate_ids = result["duplicate_ids"]
        
        file_count = result["file_count"]
        sheet_count = result["sheet_count"]
        id_count = result["id_count"]
        
        duplicate_count = sum(1 for ids in result["duplicate_ids"].values() if len(ids) > 1)
        if duplicate_count > 0:
            self.log_text.insert(tk.END, f"\n주의: {duplicate_count}개의 STRING_ID가 여러 파일에 중복 존재합니다.\n")
            dup_examples = [(id, files) for id, files in result["duplicate_ids"].items() if len(files) > 1][:5]
            for id, files in dup_examples:
                self.log_text.insert(tk.END, f"  - {id}: {', '.join(files)}\n")
            if len(dup_examples) < duplicate_count:
                self.log_text.insert(tk.END, f"  ... 외 {duplicate_count - len(dup_examples)}개\n")
        
        self.log_text.insert(tk.END, f"번역 DB 캐싱 완료:\n")
        self.log_text.insert(tk.END, f"- 파일별 캐시: {file_count}개 파일, {sum(len(ids) for ids in result['translation_file_cache'].values())}개 항목\n")
        self.log_text.insert(tk.END, f"- 시트별 캐시: {sheet_count}개 시트, {sum(len(ids) for ids in result['translation_sheet_cache'].values())}개 항목\n")
        self.log_text.insert(tk.END, f"- 전체 고유 STRING_ID: {id_count}개\n")
        
        self.status_label_apply.config(text=f"번역 DB 캐싱 완료 - {id_count}개 항목")
        
        messagebox.showinfo(
            "완료", 
            f"번역 DB 캐싱 완료!\n파일 수: {file_count}개\n시트 수: {sheet_count}개\n항목 수: {id_count}개", 
            parent=self
        )

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
            
        self.log_text.delete(1.0, tk.END)
        self.log_text.insert(tk.END, "번역 적용 작업 시작...\n")
        self.status_label_apply.config(text="작업 중...")
        self.update()
            
        self.progress_bar["maximum"] = len(selected_files)
        self.progress_bar["value"] = 0
            
        loading_popup = LoadingPopup(self, "번역 적용 중", "번역 적용 준비 중...")
            
        def apply_translations():
            total_updated = 0
            processed_count = 0
            error_count = 0
            problem_files = {
                "external_links": [], "permission_denied": [], "file_corrupted": [],
                "file_not_found": [], "unknown_error": []
            }
            
            for idx, file_name in enumerate(selected_files):
                file_path = next((path for name, path in self.original_files if name == file_name), None)
                if not file_path:
                    continue
                    
                self.after(0, lambda i=idx, n=file_name: [
                    loading_popup.update_progress((i / len(selected_files)) * 100, f"파일 처리 중 ({i+1}/{len(selected_files)}): {n}"),
                    self.log_text.insert(tk.END, f"\n파일 {n} 처리 중...\n"),
                    self.log_text.see(tk.END),
                    self.progress_bar.configure(value=i+1)
                ])
                    
                try:
                    result = self.translation_apply_manager.apply_translation(
                        file_path, selected_langs, self.record_date_var.get()
                    )
                        
                    if result["status"] == "success":
                        update_count = result["total_updated"]
                        total_updated += update_count
                        processed_count += 1
                        self.after(0, lambda c=update_count: [
                            self.log_text.insert(tk.END, f"  {c}개 항목 업데이트 완료\n"),
                            self.log_text.see(tk.END)
                        ])
                    elif result["status"] == "info":
                        processed_count += 1
                        self.after(0, lambda m=result["message"]: [
                            self.log_text.insert(tk.END, f"  {m}\n"),
                            self.log_text.see(tk.END)
                        ])
                    else:
                        error_count += 1
                        error_type = result.get("error_type", "unknown_error")
                        if error_type in problem_files:
                            problem_files[error_type].append({"file_name": file_name, "message": result["message"]})
                        self.after(0, lambda m=result["message"]: [
                            self.log_text.insert(tk.END, f"  오류 발생: {m}\n"),
                            self.log_text.see(tk.END)
                        ])
                        
                except Exception as e:
                    error_count += 1
                    error_msg = str(e)
                    problem_files["unknown_error"].append({"file_name": file_name, "message": error_msg})
                    self.after(0, lambda: [
                        self.log_text.insert(tk.END, f"  오류 발생: {error_msg}\n"),
                        self.log_text.see(tk.END)
                    ])
                    
            self.after(0, lambda: self.process_translation_apply_result(
                total_updated, processed_count, error_count, loading_popup, problem_files))

        thread = threading.Thread(target=apply_translations, daemon=True)
        thread.start()
            
    def process_translation_apply_result(self, results, loading_popup):
        loading_popup.close()
        total_updated = results['total_updated']
        self.log_text.insert(tk.END, f"\n번역 적용 작업 완료!\n총 {total_updated}개 항목이 업데이트되었습니다.\n")
        self.status_label_apply.config(text=f"번역 적용 완료 - {total_updated}개 항목")
        
        messagebox.showinfo("완료", f"번역 적용 작업이 완료되었습니다.\n총 {total_updated}개 항목이 업데이트되었습니다.", parent=self)

        problem_summary = []
        total_problem_files = 0
        
        for error_type, files in problem_files.items():
            if files:
                file_names = [f["file_name"] for f in files]
                problem_summary.append(f"🔗 {error_type} ({len(files)}개):\n   " + "\n   ".join(file_names))
                total_problem_files += len(files)

        completion_msg = f"번역 적용 작업이 완료되었습니다.\n총 {total_updated}개 항목이 업데이트되었습니다."
        
        if total_problem_files > 0:
            problem_detail = "\n\n⚠️ 처리하지 못한 파일들:\n\n" + "\n\n".join(problem_summary)
            completion_msg += problem_detail
            self.log_text.insert(tk.END, f"\n처리하지 못한 파일 ({total_problem_files}개):\n")
            for summary in problem_summary:
                self.log_text.insert(tk.END, f"{summary}\n")
        
        messagebox.showinfo("완료", completion_msg, parent=self)