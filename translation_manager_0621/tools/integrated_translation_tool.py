# integrated_translation_tool.py

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import threading
import time
import sys
import pandas as pd
from datetime import datetime

# --- 경로 문제 해결을 위한 코드 ---
project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if project_root not in sys.path:
    sys.path.append(project_root)
# ---------------------------------

from ui.common_components import ScrollableCheckList, LoadingPopup
from integrated_translation_manager import IntegratedTranslationManager

class IntegratedTranslationTool(tk.Frame):
    def __init__(self, parent, excluded_files):
        super().__init__(parent)
        self.parent = parent
        self.manager = IntegratedTranslationManager(self)
        self.excluded_files = excluded_files
        
        # UI 변수들
        self.excel_folder_var = tk.StringVar()
        self.individual_file_var = tk.StringVar()
        self.master_db_var = tk.StringVar()
        self.output_excel_var = tk.StringVar()
        self.output_db_var = tk.StringVar()
        
        # 언어 선택 변수
        self.available_languages = ["KR", "EN", "CN", "TW", "TH"]
        self.lang_vars = {}
        
        # 비교 옵션 변수
        self.include_new_var = tk.BooleanVar(value=True)
        self.include_deleted_var = tk.BooleanVar(value=True)
        self.include_modified_var = tk.BooleanVar(value=True)
        
        # 출력 옵션 변수
        self.export_new_var = tk.BooleanVar(value=True)
        self.export_deleted_var = tk.BooleanVar(value=True)
        self.export_modified_var = tk.BooleanVar(value=True)
        self.export_duplicates_var = tk.BooleanVar(value=True)
        self.save_db_var = tk.BooleanVar(value=False)
        
        # 내부 데이터
        self.excel_files = []
        self.current_results = None
        
        self.setup_ui()

    def setup_ui(self):
        """통합 번역 도구 UI 구성"""
        
        # 스크롤 가능한 메인 프레임
        main_canvas = tk.Canvas(self)
        scrollbar = ttk.Scrollbar(self, orient="vertical", command=main_canvas.yview)
        scrollable_frame = ttk.Frame(main_canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: main_canvas.configure(scrollregion=main_canvas.bbox("all"))
        )
        
        main_canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        main_canvas.configure(yscrollcommand=scrollbar.set)
        
        # --- 1. 파일 선택 영역 ---
        file_frame = ttk.LabelFrame(scrollable_frame, text="📁 번역 파일 선택")
        file_frame.pack(fill="x", padx=5, pady=5)
        
        # 폴더 선택
        folder_frame = ttk.Frame(file_frame)
        folder_frame.pack(fill="x", padx=5, pady=3)
        ttk.Label(folder_frame, text="엑셀 폴더:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        ttk.Entry(folder_frame, textvariable=self.excel_folder_var, width=50).grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        ttk.Button(folder_frame, text="찾아보기", command=self.select_excel_folder).grid(row=0, column=2, padx=5, pady=5)
        ttk.Button(folder_frame, text="폴더 검색", command=self.search_excel_files).grid(row=0, column=3, padx=5, pady=5)
        folder_frame.columnconfigure(1, weight=1)
        
        # 개별 파일 추가
        individual_frame = ttk.Frame(file_frame)
        individual_frame.pack(fill="x", padx=5, pady=3)
        ttk.Label(individual_frame, text="개별 파일:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        ttk.Entry(individual_frame, textvariable=self.individual_file_var, state="readonly").grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        ttk.Button(individual_frame, text="파일 추가", command=self.add_excel_files).grid(row=0, column=2, padx=5, pady=5)
        individual_frame.columnconfigure(1, weight=1)
        
        # 파일 목록
        files_list_frame = ttk.LabelFrame(file_frame, text="선택된 파일 목록")
        files_list_frame.pack(fill="both", expand=True, padx=5, pady=5)
        self.excel_files_list = ScrollableCheckList(files_list_frame, width=700, height=120)
        self.excel_files_list.pack(fill="both", expand=True, padx=5, pady=5)
        
        # --- 2. 비교 설정 영역 ---
        compare_frame = ttk.LabelFrame(scrollable_frame, text="⚙️ 비교 설정")
        compare_frame.pack(fill="x", padx=5, pady=5)
        
        # 마스터 DB 선택 (선택적)
        master_db_frame = ttk.Frame(compare_frame)
        master_db_frame.pack(fill="x", padx=5, pady=3)
        ttk.Label(master_db_frame, text="마스터 DB (선택적):").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        ttk.Entry(master_db_frame, textvariable=self.master_db_var, width=50).grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        ttk.Button(master_db_frame, text="찾아보기", command=self.select_master_db).grid(row=0, column=2, padx=5, pady=5)
        master_db_frame.columnconfigure(1, weight=1)
        
        # 언어 선택
        lang_frame = ttk.LabelFrame(compare_frame, text="추출할 언어")
        lang_frame.pack(fill="x", padx=5, pady=5)
        lang_checkboxes_frame = ttk.Frame(lang_frame)
        lang_checkboxes_frame.pack(fill="x", padx=5, pady=5)
        
        for i, lang in enumerate(self.available_languages):
            var = tk.BooleanVar(value=True)
            self.lang_vars[lang] = var
            ttk.Checkbutton(lang_checkboxes_frame, text=lang, variable=var).grid(row=0, column=i, padx=10, pady=5, sticky="w")
        
        # 비교 옵션
        comparison_options_frame = ttk.LabelFrame(compare_frame, text="비교 항목")
        comparison_options_frame.pack(fill="x", padx=5, pady=5)
        options_frame = ttk.Frame(comparison_options_frame)
        options_frame.pack(fill="x", padx=5, pady=5)
        
        ttk.Checkbutton(options_frame, text="신규 항목", variable=self.include_new_var).pack(side="left", padx=10)
        ttk.Checkbutton(options_frame, text="삭제된 항목", variable=self.include_deleted_var).pack(side="left", padx=10)
        ttk.Checkbutton(options_frame, text="변경된 항목", variable=self.include_modified_var).pack(side="left", padx=10)
        
        # --- 3. 출력 설정 영역 ---
        output_frame = ttk.LabelFrame(scrollable_frame, text="💾 출력 설정")
        output_frame.pack(fill="x", padx=5, pady=5)
        
        # 엑셀 출력 설정
        excel_output_frame = ttk.Frame(output_frame)
        excel_output_frame.pack(fill="x", padx=5, pady=3)
        ttk.Label(excel_output_frame, text="결과 엑셀 파일:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        ttk.Entry(excel_output_frame, textvariable=self.output_excel_var, width=50).grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        ttk.Button(excel_output_frame, text="경로 선택", command=self.select_output_excel).grid(row=0, column=2, padx=5, pady=5)
        excel_output_frame.columnconfigure(1, weight=1)
        
        # DB 출력 설정 (선택적)
        db_output_frame = ttk.Frame(output_frame)
        db_output_frame.pack(fill="x", padx=5, pady=3)
        ttk.Checkbutton(db_output_frame, text="DB로도 저장", variable=self.save_db_var).grid(row=0, column=0, padx=5, pady=5, sticky="w")
        ttk.Entry(db_output_frame, textvariable=self.output_db_var, width=50).grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        ttk.Button(db_output_frame, text="경로 선택", command=self.select_output_db).grid(row=0, column=2, padx=5, pady=5)
        db_output_frame.columnconfigure(1, weight=1)
        
        # 출력 항목 선택
        export_options_frame = ttk.LabelFrame(output_frame, text="출력할 시트")
        export_options_frame.pack(fill="x", padx=5, pady=5)
        export_checkboxes_frame = ttk.Frame(export_options_frame)
        export_checkboxes_frame.pack(fill="x", padx=5, pady=5)
        
        ttk.Checkbutton(export_checkboxes_frame, text="신규 항목", variable=self.export_new_var).pack(side="left", padx=10)
        ttk.Checkbutton(export_checkboxes_frame, text="삭제된 항목", variable=self.export_deleted_var).pack(side="left", padx=10)
        ttk.Checkbutton(export_checkboxes_frame, text="변경된 항목", variable=self.export_modified_var).pack(side="left", padx=10)
        ttk.Checkbutton(export_checkboxes_frame, text="중복 항목", variable=self.export_duplicates_var).pack(side="left", padx=10)
        
        # --- 4. 실행 및 결과 영역 ---
        action_frame = ttk.Frame(scrollable_frame)
        action_frame.pack(fill="x", padx=5, pady=10)
        
        # 프리셋 버튼들
        preset_frame = ttk.LabelFrame(action_frame, text="🔧 빠른 설정")
        preset_frame.pack(fill="x", padx=5, pady=5)
        preset_buttons_frame = ttk.Frame(preset_frame)
        preset_buttons_frame.pack(fill="x", padx=5, pady=5)
        
        ttk.Button(preset_buttons_frame, text="전체 비교", command=self.preset_full_comparison).pack(side="left", padx=5)
        ttk.Button(preset_buttons_frame, text="신규만", command=self.preset_new_only).pack(side="left", padx=5)
        ttk.Button(preset_buttons_frame, text="변경만", command=self.preset_modified_only).pack(side="left", padx=5)
        ttk.Button(preset_buttons_frame, text="설정 초기화", command=self.reset_settings).pack(side="left", padx=5)
        
        # 실행 버튼들
        execute_frame = ttk.Frame(action_frame)
        execute_frame.pack(fill="x", padx=5, pady=5)
        
        self.execute_button = ttk.Button(execute_frame, text="🚀 통합 실행", command=self.run_integrated_process)
        self.execute_button.pack(side="right", padx=5)
        
        self.preview_button = ttk.Button(execute_frame, text="👁️ 중복 미리보기", command=self.preview_duplicates, state="disabled")
        self.preview_button.pack(side="right", padx=5)
        
        # --- 5. 로그 및 상태 ---
        log_frame = ttk.LabelFrame(scrollable_frame, text="📋 작업 로그")
        log_frame.pack(fill="both", expand=True, padx=5, pady=5)
        
        self.log_text = tk.Text(log_frame, wrap="word", height=12)
        log_scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_scrollbar.set)
        log_scrollbar.pack(side="right", fill="y")
        self.log_text.pack(fill="both", expand=True)
        
        # 상태 표시
        status_frame = ttk.Frame(scrollable_frame)
        status_frame.pack(fill="x", padx=5, pady=5)
        self.status_label = ttk.Label(status_frame, text="대기 중...")
        self.status_label.pack(side="left", padx=5)
        
        # 스크롤바 패킹
        main_canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

    def select_excel_folder(self):
        """엑셀 폴더 선택"""
        folder = filedialog.askdirectory(title="번역 엑셀 폴더 선택", parent=self)
        if folder:
            self.excel_folder_var.set(folder)

    def search_excel_files(self):
        """폴더에서 엑셀 파일 검색"""
        folder = self.excel_folder_var.get()
        if not folder or not os.path.isdir(folder):
            messagebox.showwarning("경고", "유효한 폴더를 선택하세요.", parent=self)
            return
        
        self.excel_files_list.clear()
        self.excel_files = []
        
        for root, _, files in os.walk(folder):
            for file in files:
                if file.endswith(".xlsx") and not file.startswith("~$"):
                    if file not in self.excluded_files:
                        file_name_without_ext = os.path.splitext(file)[0].lower()
                        if file_name_without_ext.startswith("string"):
                            file_path = os.path.join(root, file)
                            self.excel_files.append((file, file_path))
                            self.excel_files_list.add_item(file, checked=True)
        
        if not self.excel_files:
            messagebox.showinfo("알림", "엑셀 파일을 찾지 못했습니다.", parent=self)
        else:
            self.log_text.insert(tk.END, f"{len(self.excel_files)}개의 엑셀 파일을 찾았습니다.\n")
            messagebox.showinfo("알림", f"{len(self.excel_files)}개의 엑셀 파일을 찾았습니다.", parent=self)

    def add_excel_files(self):
        """개별 엑셀 파일 추가"""
        file_paths = filedialog.askopenfilenames(
            title="추가할 번역 엑셀 파일 선택",
            filetypes=[("Excel 파일", "*.xlsx"), ("모든 파일", "*.*")],
            parent=self
        )
        if not file_paths:
            return
        
        added_count = 0
        for file_path in file_paths:
            file_name = os.path.basename(file_path)
            if not any(f[1] == file_path for f in self.excel_files):
                self.excel_files.append((file_name, file_path))
                self.excel_files_list.add_item(file_name, checked=True)
                added_count += 1
        
        if added_count > 0:
            self.log_text.insert(tk.END, f"{added_count}개의 파일이 목록에 추가되었습니다.\n")
            self.individual_file_var.set(f"{added_count}개 파일 추가됨")

    def select_master_db(self):
        """마스터 DB 파일 선택"""
        file_path = filedialog.askopenfilename(
            filetypes=[("DB 파일", "*.db"), ("모든 파일", "*.*")],
            title="마스터 DB 파일 선택",
            parent=self
        )
        if file_path:
            self.master_db_var.set(file_path)

    def select_output_excel(self):
        """출력 엑셀 파일 경로 선택"""
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel 파일", "*.xlsx")],
            title="결과 엑셀 파일 저장",
            parent=self
        )
        if file_path:
            self.output_excel_var.set(file_path)

    def select_output_db(self):
        """출력 DB 파일 경로 선택"""
        file_path = filedialog.asksaveasfilename(
            defaultextension=".db",
            filetypes=[("DB 파일", "*.db")],
            title="결과 DB 파일 저장",
            parent=self
        )
        if file_path:
            self.output_db_var.set(file_path)

    def preset_full_comparison(self):
        """전체 비교 프리셋"""
        self.include_new_var.set(True)
        self.include_deleted_var.set(True)
        self.include_modified_var.set(True)
        self.export_new_var.set(True)
        self.export_deleted_var.set(True)
        self.export_modified_var.set(True)
        self.export_duplicates_var.set(True)
        self.log_text.insert(tk.END, "프리셋 적용: 전체 비교 모드\n")

    def preset_new_only(self):
        """신규만 프리셋"""
        self.include_new_var.set(True)
        self.include_deleted_var.set(False)
        self.include_modified_var.set(False)
        self.export_new_var.set(True)
        self.export_deleted_var.set(False)
        self.export_modified_var.set(False)
        self.export_duplicates_var.set(True)
        self.log_text.insert(tk.END, "프리셋 적용: 신규 항목만\n")

    def preset_modified_only(self):
        """변경된 항목만 프리셋"""
        self.include_new_var.set(False)
        self.include_deleted_var.set(False)
        self.include_modified_var.set(True)
        self.export_new_var.set(False)
        self.export_deleted_var.set(False)
        self.export_modified_var.set(True)
        self.export_duplicates_var.set(False)
        self.log_text.insert(tk.END, "프리셋 적용: 변경된 항목만\n")

    def reset_settings(self):
        """설정 초기화"""
        # 언어 설정 초기화
        for var in self.lang_vars.values():
            var.set(True)
        
        # 비교 옵션 초기화
        self.include_new_var.set(True)
        self.include_deleted_var.set(True)
        self.include_modified_var.set(True)
        
        # 출력 옵션 초기화
        self.export_new_var.set(True)
        self.export_deleted_var.set(True)
        self.export_modified_var.set(True)
        self.export_duplicates_var.set(True)
        self.save_db_var.set(False)
        
        self.log_text.insert(tk.END, "설정이 초기화되었습니다.\n")

    def run_integrated_process(self):
        """통합 프로세스 실행"""
        # 입력 검증
        selected_files = self.excel_files_list.get_checked_items()
        if not selected_files:
            messagebox.showwarning("경고", "번역 파일을 선택하세요.", parent=self)
            return
        
        selected_langs = [lang for lang, var in self.lang_vars.items() if var.get()]
        if not selected_langs:
            messagebox.showwarning("경고", "하나 이상의 언어를 선택하세요.", parent=self)
            return
        
        output_excel = self.output_excel_var.get()
        if not output_excel:
            messagebox.showwarning("경고", "결과 엑셀 파일 경로를 지정하세요.", parent=self)
            return
        
        # DB 저장 옵션이 선택되었는데 경로가 없는 경우
        if self.save_db_var.get() and not self.output_db_var.get():
            messagebox.showwarning("경고", "DB 저장이 선택되었지만 DB 파일 경로가 지정되지 않았습니다.", parent=self)
            return
        
        # 설정 수집
        excel_files = [(name, path) for name, path in self.excel_files if name in selected_files]
        master_db_path = self.master_db_var.get() if self.master_db_var.get() else None
        
        comparison_options = {
            "include_new": self.include_new_var.get(),
            "include_deleted": self.include_deleted_var.get(),
            "include_modified": self.include_modified_var.get()
        }
        
        export_options = {
            "export_new": self.export_new_var.get(),
            "export_deleted": self.export_deleted_var.get(),
            "export_modified": self.export_modified_var.get(),
            "export_duplicates": self.export_duplicates_var.get()
        }
        
        # UI 비활성화
        self.execute_button.config(state="disabled")
        self.log_text.delete(1.0, tk.END)
        self.log_text.insert(tk.END, "통합 프로세스 시작...\n")
        self.status_label.config(text="처리 중...")
        
        # 로딩 팝업
        loading_popup = LoadingPopup(self, "통합 처리 중", "번역 파일 통합 처리를 시작합니다...")
        start_time = time.time()
        
        def progress_callback(message, current, total):
            self.after(0, lambda: [
                loading_popup.update_progress((current / total) * 100, f"{current}/{total} - {message}"),
                self.log_text.insert(tk.END, f"{message}\n"),
                self.log_text.see(tk.END)
            ])
        
        def process_thread():
            try:
                # 통합 프로세스 실행
                results = self.manager.integrated_process(
                    excel_files, selected_langs, comparison_options, master_db_path, progress_callback
                )
                
                self.after(0, lambda: self.on_process_complete(
                    results, export_options, output_excel, loading_popup, start_time
                ))
                
            except Exception as e:
                self.after(0, lambda: [
                    loading_popup.close(),
                    self.log_text.insert(tk.END, f"\n오류 발생: {str(e)}\n"),
                    self.status_label.config(text="오류 발생"),
                    self.execute_button.config(state="normal"),
                    messagebox.showerror("오류", f"처리 중 오류 발생: {str(e)}", parent=self)
                ])
        
        threading.Thread(target=process_thread, daemon=True).start()

    def on_process_complete(self, results, export_options, output_excel, loading_popup, start_time):
        """프로세스 완료 처리"""
        loading_popup.close()
        
        if results["status"] == "error":
            self.log_text.insert(tk.END, f"\n오류 발생: {results['message']}\n")
            self.status_label.config(text="오류 발생")
            self.execute_button.config(state="normal")
            messagebox.showerror("오류", f"처리 중 오류 발생: {results['message']}", parent=self)
            return
        
        self.current_results = results
        elapsed_time = time.time() - start_time
        time_str = f"{int(elapsed_time // 60)}분 {int(elapsed_time % 60)}초"
        
        # 결과 로그
        summary = results["summary"]
        self.log_text.insert(tk.END, f"\n=== 처리 완료 (소요 시간: {time_str}) ===\n")
        self.log_text.insert(tk.END, f"📊 결과 요약:\n")
        self.log_text.insert(tk.END, f"• 마스터 데이터: {results['master_count']}개\n")
        self.log_text.insert(tk.END, f"• 타겟 데이터: {results['target_count']}개\n")
        self.log_text.insert(tk.END, f"• 신규 항목: {summary['new_items']}개\n")
        self.log_text.insert(tk.END, f"• 삭제된 항목: {summary['deleted_items']}개\n")
        self.log_text.insert(tk.END, f"• 변경된 항목: {summary['modified_items']}개\n")
        self.log_text.insert(tk.END, f"• 중복 STRING_ID: {summary['duplicate_ids']}개\n")
        
        # 엑셀 파일 내보내기
        self.log_text.insert(tk.END, f"\n📄 엑셀 파일 생성 중...\n")
        export_result = self.manager.export_results_to_excel(output_excel, export_options)
        
        if export_result["status"] == "success":
            self.log_text.insert(tk.END, f"✅ 엑셀 파일 저장 완료: {output_excel}\n")
        else:
            self.log_text.insert(tk.END, f"❌ 엑셀 파일 저장 실패: {export_result['message']}\n")
        
        # DB 저장 (선택적)
        if self.save_db_var.get() and self.output_db_var.get():
            self.log_text.insert(tk.END, f"\n💾 DB 파일 생성 중...\n")
            db_result = self.manager.save_to_db(self.output_db_var.get(), "target")
            
            if db_result["status"] == "success":
                self.log_text.insert(tk.END, f"✅ DB 파일 저장 완료: {self.output_db_var.get()}\n")
            else:
                self.log_text.insert(tk.END, f"❌ DB 파일 저장 실패: {db_result['message']}\n")
        
        # UI 상태 업데이트
        self.status_label.config(text=f"완료 - {summary['new_items']}신규, {summary['modified_items']}변경, {summary['duplicate_ids']}중복")
        self.execute_button.config(state="normal")
        
        # 중복 미리보기 버튼 활성화
        if summary['duplicate_ids'] > 0:
            self.preview_button.config(state="normal")
        
        # 완료 메시지
        completion_message = (
            f"🎉 통합 처리가 완료되었습니다!\n\n"
            f"📊 처리 결과:\n"
            f"• 신규: {summary['new_items']}개\n"
            f"• 변경: {summary['modified_items']}개\n"
            f"• 삭제: {summary['deleted_items']}개\n"
            f"• 중복: {summary['duplicate_ids']}개\n"
            f"⏱️ 소요 시간: {time_str}\n\n"
            f"📁 결과 파일: {output_excel}"
        )
        
        messagebox.showinfo("완료", completion_message, parent=self)

    def preview_duplicates(self):
        """중복 데이터 미리보기"""
        if not self.current_results or not self.manager.duplicate_data:
            messagebox.showinfo("정보", "표시할 중복 데이터가 없습니다.", parent=self)
            return
        
        # 중복 데이터 미리보기 창 생성
        popup = tk.Toplevel(self)
        popup.title("중복 STRING_ID 미리보기")
        popup.geometry("1200x700")
        popup.transient(self)
        popup.grab_set()
        
        # 트리뷰 생성
        tree_frame = ttk.Frame(popup, padding=10)
        tree_frame.pack(fill="both", expand=True)
        
        columns = ("string_id", "kr", "en", "file_name", "sheet_name", "status")
        tree = ttk.Treeview(tree_frame, columns=columns, show="headings")
        
        for col in columns:
            tree.heading(col, text=col.upper())
            tree.column(col, width=150 if col != "kr" else 200)
        
        # 스크롤바
        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        vsb.pack(side="right", fill="y")
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=tree.xview)
        hsb.pack(side="bottom", fill="x")
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        tree.pack(fill="both", expand=True)
        
        # 데이터 추가
        tree.tag_configure('group', background='#E8E8E8')
        
        for string_id, items in self.manager.duplicate_data.items():
            parent_id = tree.insert("", "end", text=string_id, 
                                  values=(f"📋 {string_id} ({len(items)}개)",), 
                                  open=True, tags=('group',))
            for item in items:
                values = (
                    item.get('string_id', ''),
                    item.get('kr', ''),
                    item.get('en', ''),
                    item.get('file_name', ''),
                    item.get('sheet_name', ''),
                    item.get('status', '')
                )
                tree.insert(parent_id, "end", values=values)
        
        # 버튼 프레임
        button_frame = ttk.Frame(popup, padding=10)
        button_frame.pack(fill="x")
        
        ttk.Button(button_frame, text="Excel로 내보내기", 
                  command=lambda: self.export_duplicates_standalone()).pack(side="left")
        ttk.Button(button_frame, text="닫기", 
                  command=popup.destroy).pack(side="right")

    def export_duplicates_standalone(self):
        """중복 데이터만 별도 엑셀로 내보내기"""
        if not self.manager.duplicate_data:
            messagebox.showerror("오류", "내보낼 중복 데이터가 없습니다.", parent=self)
            return
        
        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel 파일", "*.xlsx")],
            title="중복 데이터 엑셀 저장",
            parent=self
        )
        if not save_path:
            return
        
        try:
            flat_list = []
            for string_id, items in self.manager.duplicate_data.items():
                for item in items:
                    flat_list.append(item)
            
            df = pd.DataFrame(flat_list)
            df.to_excel(save_path, index=False)
            
            self.log_text.insert(tk.END, f"중복 데이터 엑셀 저장 완료: {save_path}\n")
            messagebox.showinfo("성공", f"중복 데이터가 성공적으로 저장되었습니다:\n{save_path}", parent=self)
            
        except Exception as e:
            messagebox.showerror("저장 오류", f"파일 저장 중 오류 발생:\n{e}", parent=self)