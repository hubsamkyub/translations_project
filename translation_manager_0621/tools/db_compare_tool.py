# tools/db_compare_tool.py (수정 후)

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import pandas as pd
import os
import sys

# --- 경로 문제 해결을 위한 코드 ---
project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if project_root not in sys.path:
    sys.path.append(project_root)
# ---------------------------------

from tools.db_compare_manager import DBCompareManager
from ui.common_components import LoadingPopup

class DBCompareTool(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        self.db_compare_manager = DBCompareManager()

        self.master_db_var = tk.StringVar()
        self.target_db_var = tk.StringVar()
        self.output_excel_var = tk.StringVar()

        self.available_languages = ["KR", "EN", "CN", "TW", "TH"]
        self.lang_vars = {}
        self.language_list = []

        # [신규] 추출 옵션 변수
        self.export_new_var = tk.BooleanVar(value=True)
        self.export_deleted_var = tk.BooleanVar(value=True)
        self.export_modified_var = tk.BooleanVar(value=True)
        
        self.setup_ui()

    def setup_ui(self):
        db_frame = ttk.LabelFrame(self, text="DB 파일 선택")
        db_frame.pack(fill="x", padx=5, pady=5)

        ttk.Label(db_frame, text="Master DB:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        ttk.Entry(db_frame, textvariable=self.master_db_var, width=70).grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        ttk.Button(db_frame, text="찾아보기", command=lambda: self.select_db_file(self.master_db_var)).grid(row=0, column=2, padx=5, pady=5)

        ttk.Label(db_frame, text="Target DB:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        ttk.Entry(db_frame, textvariable=self.target_db_var, width=70).grid(row=1, column=1, padx=5, pady=5, sticky="ew")
        ttk.Button(db_frame, text="찾아보기", command=lambda: self.select_db_file(self.target_db_var)).grid(row=1, column=2, padx=5, pady=5)
        
        db_frame.columnconfigure(1, weight=1)

        lang_frame = ttk.LabelFrame(self, text="비교할 언어 선택")
        lang_frame.pack(fill="x", padx=5, pady=5)
        for i, lang in enumerate(self.available_languages):
            var = tk.BooleanVar(value=True if lang in ["KR", "EN"] else False)
            self.lang_vars[lang] = var
            ttk.Checkbutton(lang_frame, text=lang, variable=var).grid(row=0, column=i, padx=10, pady=5)
            
        # [신규] 추출 옵션 프레임
        export_options_frame = ttk.LabelFrame(self, text="추출 옵션 (유형 선택)")
        export_options_frame.pack(fill="x", padx=5, pady=5)
        ttk.Checkbutton(export_options_frame, text="신규 항목", variable=self.export_new_var).pack(side="left", padx=10, pady=5)
        ttk.Checkbutton(export_options_frame, text="삭제된 항목", variable=self.export_deleted_var).pack(side="left", padx=10, pady=5)
        ttk.Checkbutton(export_options_frame, text="변경된 항목", variable=self.export_modified_var).pack(side="left", padx=10, pady=5)


        output_frame = ttk.LabelFrame(self, text="결과 저장")
        output_frame.pack(fill="x", padx=5, pady=5)

        ttk.Label(output_frame, text="출력 파일:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        ttk.Entry(output_frame, textvariable=self.output_excel_var, width=70).grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        ttk.Button(output_frame, text="경로 선택", command=self.select_output_file).grid(row=0, column=2, padx=5, pady=5)
        output_frame.columnconfigure(1, weight=1)

        action_frame = ttk.Frame(self)
        action_frame.pack(fill="x", padx=5, pady=10)
        ttk.Button(action_frame, text="추출 실행", command=self.run_extraction).pack(side="right", padx=5)
        
        log_frame = ttk.LabelFrame(self, text="진행 로그")
        log_frame.pack(fill="both", expand=True, padx=5, pady=5)
        self.log_text = tk.Text(log_frame, wrap="word", height=15)
        self.log_text.pack(fill="both", expand=True)

    def select_db_file(self, var):
        file_path = filedialog.askopenfilename(filetypes=[("DB 파일", "*.db"), ("모든 파일", "*.*")])
        if file_path:
            var.set(file_path)

    def select_output_file(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel 파일", "*.xlsx")])
        if file_path:
            self.output_excel_var.set(file_path)

    def run_extraction(self):
        master_db = self.master_db_var.get()
        target_db = self.target_db_var.get()
        output_file = self.output_excel_var.get()
        self.language_list = [lang for lang, var in self.lang_vars.items() if var.get()]

        if not all([master_db, target_db, output_file]):
            messagebox.showwarning("입력 오류", "모든 파일 경로를 지정해야 합니다.", parent=self)
            return
        if not self.language_list:
            messagebox.showwarning("입력 오류", "하나 이상의 언어를 선택해야 합니다.", parent=self)
            return
        if not any([self.export_new_var.get(), self.export_deleted_var.get(), self.export_modified_var.get()]):
            messagebox.showwarning("입력 오류", "하나 이상의 추출 유형을 선택해야 합니다.", parent=self)
            return

        self.log_text.delete(1.0, tk.END)
        self.log_text.insert(tk.END, "DB 비교 및 추출을 시작합니다...\n")
        
        loading_popup = LoadingPopup(self, "DB 비교 중", "DB 비교를 준비하고 있습니다...")

        def progress_callback(message, current, total):
            self.after(0, loading_popup.update_progress, (current / total) * 100, message)
            self.log_text.insert(tk.END, f"{message}\n")
            self.log_text.see(tk.END)
        
        def extraction_thread():
            try:
                results = self.db_compare_manager.compare_and_extract(master_db, target_db, self.language_list, progress_callback)
                self.after(0, self.on_extraction_complete, results, output_file, loading_popup)
            except Exception as e:
                self.after(0, lambda: [
                    loading_popup.close(),
                    messagebox.showerror("추출 오류", f"오류 발생: {e}", parent=self),
                    self.log_text.insert(tk.END, f"오류 발생: {e}\n")
                ])

        thread = threading.Thread(target=extraction_thread, daemon=True)
        thread.start()

    def on_extraction_complete(self, results, output_file, loading_popup):
        loading_popup.close()
        self.log_text.insert(tk.END, "데이터 비교 완료. Excel 파일 생성 중...\n")

        all_diffs = []
        lang_lower_list = [lang.lower() for lang in self.language_list]

        # [수정됨] 체크된 옵션에 따라 데이터 필터링 및 리스트 구성
        if self.export_new_var.get():
            for item in results.get('new_in_target', []):
                row = {'유형': '신규', 'STRING_ID': item['string_id']}
                for lang in lang_lower_list:
                    row[lang] = item.get(lang)
                all_diffs.append(row)

        if self.export_deleted_var.get():
            for item in results.get('new_in_master', []):
                row = {'유형': '삭제', 'STRING_ID': item['string_id']}
                for lang in lang_lower_list:
                    row[lang] = item.get(lang)
                all_diffs.append(row)

        if self.export_modified_var.get():
            for item in results.get('modified', []):
                row = {'유형': '변경', 'STRING_ID': item['string_id'], 'KR': item.get('kr', '')}
                for lang in lang_lower_list:
                    if lang == 'kr': continue # KR은 이미 추가됨
                    master_col = f'{lang}_master'
                    target_col = f'{lang}_target'
                    if master_col in item:
                        row[f'{lang}_master'] = item.get(master_col, '')
                        row[f'{lang}_target'] = item.get(target_col, '')
                all_diffs.append(row)
        
        if not all_diffs:
            messagebox.showinfo("알림", "선택된 유형에 해당하는 변경 사항이 없습니다.", parent=self)
            self.log_text.insert(tk.END, "추출할 데이터가 없습니다.\n")
            return
            
        final_df = pd.DataFrame(all_diffs)
        
        # 컬럼 순서 정리
        cols = ['유형', 'STRING_ID', 'kr']
        for lang in lang_lower_list:
            if lang == 'kr': continue
            cols.append(lang) # 신규/삭제용
            cols.append(f'{lang}_master') # 변경용
            cols.append(f'{lang}_target') # 변경용

        # 데이터프레임에 존재하는 컬럼만으로 순서 재정의
        final_cols = [col for col in cols if col in final_df.columns]
        final_df = final_df[final_cols]

        try:
            final_df.to_excel(output_file, index=False)
            self.log_text.insert(tk.END, f"추출 완료! 파일이 '{output_file}'에 저장되었습니다.\n")
            self.log_text.insert(tk.END, f"신규: {len(results.get('new_in_target', []))}건, 삭제: {len(results.get('new_in_master', []))}건, 변경: {len(results.get('modified', []))}건\n")
            messagebox.showinfo("완료", f"추출이 완료되었습니다.\n파일이 저장된 경로: {output_file}", parent=self)
        except Exception as e:
            messagebox.showerror("파일 저장 오류", f"Excel 파일 저장 중 오류 발생: {e}", parent=self)
            self.log_text.insert(tk.END, f"파일 저장 오류: {e}\n")