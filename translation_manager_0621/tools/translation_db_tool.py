# tools/translate/translation_db_tool.py (최종)

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import threading
import time
import sys
import pandas as pd

# --- 경로 문제 해결을 위한 코드 ---
project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if project_root not in sys.path:
    sys.path.append(project_root)
# ---------------------------------

from ui.common_components import ScrollableCheckList, LoadingPopup
from tools.translation_db_manager import TranslationDBManager

class TranslationDBTool(tk.Frame):
    def __init__(self, parent, excluded_files):
        super().__init__(parent)
        self.parent = parent
        self.db_manager = TranslationDBManager(self)
        self.update_option = tk.StringVar(value="default")
        self.debug_string_id_var = tk.StringVar()
        
        # UI 변수
        self.trans_excel_folder_var = tk.StringVar()
        self.individual_file_var = tk.StringVar()
        self.output_db_var = tk.StringVar()
        self.update_db_var = tk.StringVar()
        self.batch_size_var = tk.IntVar(value=500)
        self.read_only_var = tk.BooleanVar(value=True)
        self.available_languages = ["KR", "EN", "CN", "TW", "TH"]
        self.lang_vars = {}
        
        # 내부 데이터
        self.trans_excel_files = []
        self.duplicate_data = None
        self.excluded_files = excluded_files
        
        self.setup_ui()

    def setup_ui(self):
        """번역 DB 구축 탭 UI 구성"""
        # --- 1. 번역 파일 선택 프레임 ---
        excel_frame = ttk.LabelFrame(self, text="번역 파일 선택")
        excel_frame.pack(fill="x", padx=5, pady=5)
        
        folder_frame = ttk.Frame(excel_frame)
        folder_frame.pack(fill="x", padx=5, pady=5)
        ttk.Label(folder_frame, text="엑셀 폴더:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        ttk.Entry(folder_frame, textvariable=self.trans_excel_folder_var, width=50).grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        ttk.Button(folder_frame, text="찾아보기", command=self.select_trans_excel_folder).grid(row=0, column=2, padx=5, pady=5)
        ttk.Button(folder_frame, text="폴더 내 파일 검색", command=self.search_translation_excel_files).grid(row=0, column=3, padx=5, pady=5)
        folder_frame.columnconfigure(1, weight=1)

        file_select_frame = ttk.Frame(excel_frame)
        file_select_frame.pack(fill="x", padx=5, pady=5)
        ttk.Label(file_select_frame, text="개별 파일 추가:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        ttk.Entry(file_select_frame, textvariable=self.individual_file_var, state="readonly").grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        ttk.Button(file_select_frame, text="파일 추가", command=self.add_excel_files).grid(row=0, column=2, padx=5, pady=5)
        file_select_frame.columnconfigure(1, weight=1)

        files_frame = ttk.LabelFrame(self, text="번역 엑셀 파일 목록")
        files_frame.pack(fill="both", expand=True, padx=5, pady=5)
        self.trans_excel_files_list = ScrollableCheckList(files_frame, width=700, height=150)
        self.trans_excel_files_list.pack(fill="both", expand=True, padx=5, pady=5)
        
        # --- 2. DB 출력 설정 ---
        output_frame = ttk.LabelFrame(self, text="DB 출력 설정")
        output_frame.pack(fill="x", padx=5, pady=5)
        
        db_build_frame = ttk.Frame(output_frame)
        db_build_frame.pack(fill="x", padx=5, pady=2)
        ttk.Label(db_build_frame, text="새 DB 파일:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        ttk.Entry(db_build_frame, textvariable=self.output_db_var, width=50).grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        ttk.Button(db_build_frame, text="찾아보기", command=self.save_db_file).grid(row=0, column=2, padx=5, pady=5)
        db_build_frame.columnconfigure(1, weight=1)
        
        db_update_frame = ttk.Frame(output_frame)
        db_update_frame.pack(fill="x", padx=5, pady=2)
        ttk.Label(db_update_frame, text="기존 DB 파일:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        ttk.Entry(db_update_frame, textvariable=self.update_db_var, width=50).grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        ttk.Button(db_update_frame, text="찾아보기", command=self.select_update_db_file).grid(row=0, column=2, padx=5, pady=5)
        db_update_frame.columnconfigure(1, weight=1)
        
        # --- 3. 추출 및 업데이트 옵션 ---
        languages_frame = ttk.LabelFrame(self, text="추출할 언어")
        languages_frame.pack(fill="x", padx=5, pady=5)
        for i, lang in enumerate(self.available_languages):
            var = tk.BooleanVar(value=True)
            self.lang_vars[lang] = var
            ttk.Checkbutton(languages_frame, text=lang, variable=var).grid(row=i // 5, column=i % 5, padx=10, pady=5, sticky="w")

        update_options_frame = ttk.LabelFrame(self, text="DB 업데이트 옵션")
        update_options_frame.pack(fill="x", padx=5, pady=5)
        ttk.Radiobutton(update_options_frame, text="기본 업데이트 (STRING_ID 기준, KR 제외)", variable=self.update_option, value="default").pack(anchor="w", padx=5)
        ttk.Radiobutton(update_options_frame, text="KR 추가 비교 (STRING_ID + KR 기준)", variable=self.update_option, value="kr_additional_compare").pack(anchor="w", padx=5)
        ttk.Radiobutton(update_options_frame, text="KR 비교 (KR 기준)", variable=self.update_option, value="kr_compare").pack(anchor="w", padx=5)
        ttk.Radiobutton(update_options_frame, text="KR 덮어쓰기 (STRING_ID 기준, KR 포함)", variable=self.update_option, value="kr_overwrite").pack(anchor="w", padx=5)
        
        # --- 4. 실행 버튼 ---
        action_frame = ttk.Frame(self)
        action_frame.pack(fill="x", padx=5, pady=5)
        self.show_duplicates_button = ttk.Button(action_frame, text="중복 데이터 추출", command=self.show_duplicates_window, state="disabled")
        self.show_duplicates_button.pack(side="right", padx=5, pady=5)
        ttk.Button(action_frame, text="번역 DB 업데이트", command=self.update_translation_db).pack(side="right", padx=5, pady=5)
        ttk.Button(action_frame, text="번역 DB 구축", command=self.build_translation_db).pack(side="right", padx=5, pady=5)
        
        # --- 5. 로그 및 상태 ---
        log_frame = ttk.LabelFrame(self, text="작업 로그")
        log_frame.pack(fill="both", expand=True, padx=5, pady=5)
        self.db_log_text = tk.Text(log_frame, wrap="word", height=10)
        scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=self.db_log_text.yview)
        self.db_log_text.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        self.db_log_text.pack(fill="both", expand=True)
        
        status_frame = ttk.Frame(self)
        status_frame.pack(fill="x", padx=5, pady=5, side="bottom")
        self.status_label_db = ttk.Label(status_frame, text="대기 중...")
        self.status_label_db.pack(side="left", padx=5)
    
    def select_trans_excel_folder(self):
        folder = filedialog.askdirectory(title="번역 엑셀 폴더 선택", parent=self)
        if folder: self.trans_excel_folder_var.set(folder)

    def add_excel_files(self):
        file_paths = filedialog.askopenfilenames(
            title="추가할 번역 엑셀 파일 선택",
            filetypes=[("Excel 파일", "*.xlsx"), ("모든 파일", "*.*")],
            parent=self
        )
        if not file_paths: return
        added_count = 0
        for file_path in file_paths:
            file_name = os.path.basename(file_path)
            if not any(f[1] == file_path for f in self.trans_excel_files):
                self.trans_excel_files.append((file_name, file_path))
                self.trans_excel_files_list.add_item(file_name, checked=True)
                added_count += 1
        if added_count > 0: self.db_log_text.insert(tk.END, f"{added_count}개의 파일이 목록에 추가되었습니다.\n")
        self.individual_file_var.set(f"{added_count}개 파일 추가됨")

    def save_db_file(self):
        file_path = filedialog.asksaveasfilename(
            defaultextension=".db", filetypes=[("DB 파일", "*.db"), ("모든 파일", "*.*")],
            title="새 번역 DB 파일 저장", parent=self
        )
        if file_path: self.output_db_var.set(file_path)
    
    def select_update_db_file(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("DB 파일", "*.db"), ("모든 파일", "*.*")],
            title="기존 번역 DB 파일 선택", parent=self
        )
        if file_path: self.update_db_var.set(file_path)

    def search_translation_excel_files(self):
        folder = self.trans_excel_folder_var.get()
        if not folder or not os.path.isdir(folder):
            messagebox.showwarning("경고", "유효한 폴더를 선택하세요.", parent=self)
            return
        
        self.trans_excel_files_list.clear()
        self.trans_excel_files = []
        
        for root, _, files in os.walk(folder):
            for file in files:
                if file.endswith(".xlsx") and not file.startswith("~$"):
                    if file not in self.excluded_files:
                        file_name_without_ext = os.path.splitext(file)[0].lower()
                        if file_name_without_ext.startswith("string"):
                            file_path = os.path.join(root, file)
                            self.trans_excel_files.append((file, file_path))
                            self.trans_excel_files_list.add_item(file, checked=True)
        
        if not self.trans_excel_files: messagebox.showinfo("알림", "엑셀 파일을 찾지 못했습니다.", parent=self)
        else: messagebox.showinfo("알림", f"{len(self.trans_excel_files)}개의 엑셀 파일을 찾았습니다.", parent=self)

    def build_translation_db(self):
        selected_file_names = self.trans_excel_files_list.get_checked_items()
        if not selected_file_names:
            messagebox.showwarning("경고", "번역 파일을 선택하세요.", parent=self)
            return
        
        db_path = self.output_db_var.get()
        if not db_path:
            messagebox.showwarning("경고", "DB 파일 경로를 지정하세요.", parent=self)
            return
        
        selected_langs = [lang for lang, var in self.lang_vars.items() if var.get()]
        if not selected_langs:
            messagebox.showwarning("경고", "하나 이상의 언어를 선택하세요.", parent=self)
            return
        
        if os.path.exists(db_path) and not messagebox.askyesno("확인", f"'{db_path}' 파일이 이미 존재합니다. 덮어쓰시겠습니까?", parent=self):
            return
        
        self.db_log_text.delete(1.0, tk.END)
        self.db_log_text.insert(tk.END, "번역 DB 구축 시작...\n")
        self.status_label_db.config(text="번역 DB 구축 중...")
        self.show_duplicates_button.config(state="disabled")
        self.duplicate_data = None
        self.update()
        
        excel_files = [(name, path) for name, path in self.trans_excel_files if name in selected_file_names]
        
        loading_popup = LoadingPopup(self, "번역 DB 구축 중", "번역 DB 구축 준비 중...")
        start_time = time.time()
        
        def progress_callback(message, current, total):
            self.after(0, lambda: [
                loading_popup.update_progress((current / total) * 100, f"{current}/{total} - {message}"),
                self.db_log_text.insert(tk.END, f"{message}\n"),
                self.db_log_text.see(tk.END)
            ])
        
        def build_db_thread():
            try:
                result = self.db_manager.build_translation_db(
                    excel_files, db_path, selected_langs, self.batch_size_var.get(), self.read_only_var.get(), progress_callback
                )
                self.after(0, lambda: self.process_db_build_result(result, loading_popup, start_time))
            except Exception as e:
                self.after(0, lambda: [
                    loading_popup.close(),
                    self.db_log_text.insert(tk.END, f"\n오류 발생: {str(e)}\n"),
                    self.status_label_db.config(text="오류 발생"),
                    messagebox.showerror("오류", f"DB 구축 중 오류 발생: {str(e)}", parent=self)
                ])
                
        threading.Thread(target=build_db_thread, daemon=True).start()

    def process_db_build_result(self, result, loading_popup, start_time):
        loading_popup.close()
        
        if result["status"] == "error":
            self.db_log_text.insert(tk.END, f"\n오류 발생: {result['message']}\n")
            self.status_label_db.config(text="오류 발생")
            messagebox.showerror("오류", f"DB 구축 중 오류 발생: {result['message']}", parent=self)
            return
            
        elapsed_time = time.time() - start_time
        time_str = f"{int(elapsed_time // 60)}분 {int(elapsed_time % 60)}초"
        
        self.db_log_text.insert(tk.END, f"\n번역 DB 구축 완료! (소요 시간: {time_str})\n")
        self.db_log_text.insert(tk.END, f"파일 처리: {result['processed_count']}/{len(self.trans_excel_files_list.get_checked_items())} (오류: {result['error_count']})\n")
        self.db_log_text.insert(tk.END, f"총 {result['total_rows']}개 고유 항목이 DB에 추가되었습니다.\n")
        
        self.status_label_db.config(text=f"번역 DB 구축 완료 - {result['total_rows']}개 항목")
        
        final_message = f"번역 DB 구축이 완료되었습니다.\n총 {result['total_rows']}개 항목이 추가되었습니다.\n소요 시간: {time_str}"
        
        self.duplicate_data = result.get("duplicates")
        if self.duplicate_data:
            self.db_log_text.insert(tk.END, f"\n주의: {len(self.duplicate_data)}개의 중복 STRING_ID가 발견되었습니다.\n")
            self.show_duplicates_button.config(state="normal")
            final_message += f"\n\n중복된 STRING_ID {len(self.duplicate_data)}개를 발견했습니다.\n'중복 데이터 추출' 버튼으로 확인 및 저장이 가능합니다."
        else:
            self.show_duplicates_button.config(state="disabled")

        messagebox.showinfo("완료", final_message, parent=self)

    def update_translation_db(self):
        selected_file_names = self.trans_excel_files_list.get_checked_items()
        if not selected_file_names:
            messagebox.showwarning("경고", "번역 파일을 선택하세요.", parent=self)
            return
        
        db_path = self.update_db_var.get()
        if not db_path or not os.path.exists(db_path):
            messagebox.showwarning("경고", "유효한 기존 DB 파일을 지정하세요.", parent=self)
            return
        
        selected_langs = [lang for lang, var in self.lang_vars.items() if var.get()]
        if not selected_langs:
            messagebox.showwarning("경고", "하나 이상의 언어를 선택하세요.", parent=self)
            return
        
        self.db_log_text.delete(1.0, tk.END)
        self.db_log_text.insert(tk.END, "번역 DB 업데이트 시작...\n")
        self.status_label_db.config(text="번역 DB 업데이트 중...")
        self.show_duplicates_button.config(state="disabled")
        self.duplicate_data = None
        self.update()
        
        excel_files = [(name, path) for name, path in self.trans_excel_files if name in selected_file_names]

        loading_popup = LoadingPopup(self, "번역 DB 업데이트 중", "번역 DB 업데이트 준비 중...")
        start_time = time.time()
        
        def progress_callback(message, current, total):
            self.after(0, lambda: [
                loading_popup.update_progress((current / total) * 100, f"{current}/{total} - {message}"),
                self.db_log_text.insert(tk.END, f"{message}\n"),
                self.db_log_text.see(tk.END)
            ])
        
        def update_db_thread():
            try:
                selected_option = self.update_option.get()
                debug_id = self.debug_string_id_var.get()
                result = self.db_manager.update_translation_db(
                    excel_files=excel_files, db_path=db_path, language_list=selected_langs, 
                    batch_size=self.batch_size_var.get(), use_read_only=self.read_only_var.get(), 
                    progress_callback=progress_callback, update_option=selected_option,
                    debug_string_id=debug_id if debug_id else None
                )
                self.after(0, lambda: self.process_db_update_result(result, loading_popup, start_time))
            except Exception as e:
                self.after(0, lambda: [
                    loading_popup.close(),
                    self.db_log_text.insert(tk.END, f"\n오류 발생: {str(e)}\n"),
                    self.status_label_db.config(text="오류 발생"),
                    messagebox.showerror("오류", f"DB 업데이트 중 오류 발생: {str(e)}", parent=self)
                ])
                
        threading.Thread(target=update_db_thread, daemon=True).start()

    def process_db_update_result(self, result, loading_popup, start_time):
        loading_popup.close()
        
        if result["status"] == "error":
            self.db_log_text.insert(tk.END, f"\n오류 발생: {result['message']}\n")
            self.status_label_db.config(text="오류 발생")
            messagebox.showerror("오류", f"DB 업데이트 중 오류 발생: {result['message']}", parent=self)
            return
            
        elapsed_time = time.time() - start_time
        time_str = f"{int(elapsed_time // 60)}분 {int(elapsed_time % 60)}초"
        
        self.db_log_text.insert(tk.END, f"\n번역 DB 업데이트 완료! (소요 시간: {time_str})\n")
        self.db_log_text.insert(tk.END, f"파일 처리: {result['processed_count']}/{len(self.trans_excel_files_list.get_checked_items())} (오류: {result['error_count']})\n")
        self.db_log_text.insert(tk.END, f"신규 추가: {result.get('new_rows', 0)}개\n")
        self.db_log_text.insert(tk.END, f"기존 업데이트: {result.get('updated_rows', 0)}개\n")
        self.db_log_text.insert(tk.END, f"삭제 표시: {result.get('deleted_rows', 0)}개\n")
        
        self.status_label_db.config(text=f"번역 DB 업데이트 완료")
        
        update_summary = (f"번역 DB 업데이트가 완료되었습니다.\n\n"
                          f"📊 처리 통계:\n"
                          f"• 신규 추가: {result.get('new_rows', 0)}개\n"
                          f"• 기존 업데이트: {result.get('updated_rows', 0)}개\n"
                          f"⏱️ 소요 시간: {time_str}")
        
        self.duplicate_data = result.get("duplicates")
        if self.duplicate_data:
            self.db_log_text.insert(tk.END, f"\n주의: {len(self.duplicate_data)}개의 중복 STRING_ID가 발견되었습니다.\n")
            self.show_duplicates_button.config(state="normal")
            update_summary += f"\n\n중복된 STRING_ID {len(self.duplicate_data)}개를 발견했습니다.\n'중복 데이터 추출' 버튼으로 확인 및 저장이 가능합니다."
        else:
            self.show_duplicates_button.config(state="disabled")

        messagebox.showinfo("완료", update_summary, parent=self)

    def show_duplicates_window(self):
        if not self.duplicate_data:
            messagebox.showinfo("정보", "추출할 중복 데이터가 없습니다.", parent=self)
            return

        popup = tk.Toplevel(self)
        popup.title("중복 데이터 미리보기 및 추출")
        popup.geometry("1000x600")
        popup.transient(self)
        popup.grab_set()

        tree_frame = ttk.Frame(popup, padding=10)
        tree_frame.pack(fill="both", expand=True)

        columns = ("string_id", "kr", "file_name", "sheet_name", "status")
        tree = ttk.Treeview(tree_frame, columns=columns, show="headings")
        
        for col in columns:
            tree.heading(col, text=col.upper())
            tree.column(col, width=150)
        
        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        vsb.pack(side="right", fill="y")
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=tree.xview)
        hsb.pack(side="bottom", fill="x")
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        tree.pack(fill="both", expand=True)

        tree.tag_configure('group', background='#E8E8E8')

        for string_id, items in self.duplicate_data.items():
            parent_id = tree.insert("", "end", text=string_id, values=(string_id,), open=True, tags=('group',))
            for item in items:
                values = (
                    item.get('string_id', ''), item.get('kr', ''),
                    item.get('file_name', ''), item.get('sheet_name', ''),
                    item.get('status', '')
                )
                tree.insert(parent_id, "end", values=values)
        
        button_frame = ttk.Frame(popup, padding=10)
        button_frame.pack(fill="x")
        
        ttk.Button(button_frame, text="Excel로 내보내기", command=self.export_duplicates_to_excel).pack(side="left")
        ttk.Button(button_frame, text="닫기", command=popup.destroy).pack(side="right")

    def export_duplicates_to_excel(self):
        if not self.duplicate_data:
            messagebox.showerror("오류", "내보낼 데이터가 없습니다.", parent=self)
            return

        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel 파일", "*.xlsx")],
            title="중복 데이터 저장",
            parent=self
        )
        if not save_path:
            return

        flat_list = []
        for string_id, items in self.duplicate_data.items():
            for item in items:
                flat_list.append(item)
        
        df = pd.DataFrame(flat_list)
        try:
            df.to_excel(save_path, index=False)
            messagebox.showinfo("성공", f"중복 데이터가 성공적으로 저장되었습니다:\n{save_path}", parent=self)
        except Exception as e:
            messagebox.showerror("저장 오류", f"파일 저장 중 오류 발생:\n{e}", parent=self)