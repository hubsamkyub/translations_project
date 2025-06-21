# tools/db_compare_tool.py

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import threading
import pandas as pd

from ui.common_components import ScrollableCheckList, LoadingPopup
from tools.db_compare_manager import DBCompareManager

class DBCompareTool(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent  # 부모 위젯(탭) 저장

        # DB 비교 로직을 처리할 매니저 인스턴스 생성
        self.db_compare_manager = DBCompareManager(self)
        self.compare_results = []
        self.db_pairs = []

        # UI 구성
        self.setup_ui()

    def setup_ui(self):
        """DB 비교 탭의 UI를 구성합니다."""
        # 상단 프레임 (좌우 분할)
        top_frame = ttk.Frame(self)
        top_frame.pack(fill="x", padx=5, pady=5)
        
        # 좌측 프레임 (개별 DB 비교 + 폴더 DB 비교)
        left_frame = ttk.Frame(top_frame)
        left_frame.pack(side="left", fill="both", expand=True, padx=(0, 5))
        
        # ===== 개별 DB 비교 섹션 =====
        single_db_frame = ttk.LabelFrame(left_frame, text="개별 DB 비교")
        single_db_frame.pack(fill="x", padx=5, pady=5)
        
        ttk.Label(single_db_frame, text="원본 DB:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.original_db_var = tk.StringVar()
        ttk.Entry(single_db_frame, textvariable=self.original_db_var, width=40).grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        ttk.Button(single_db_frame, text="찾아보기", 
                command=lambda: self.select_db_file("original")).grid(row=0, column=2, padx=5, pady=5)
        
        ttk.Label(single_db_frame, text="비교 DB:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.compare_db_var = tk.StringVar()
        ttk.Entry(single_db_frame, textvariable=self.compare_db_var, width=40).grid(row=1, column=1, padx=5, pady=5, sticky="ew")
        ttk.Button(single_db_frame, text="찾아보기", 
                command=lambda: self.select_db_file("compare")).grid(row=1, column=2, padx=5, pady=5)
        
        single_db_frame.columnconfigure(1, weight=1)
        
        # ===== 폴더 DB 비교 섹션 =====
        folder_db_frame = ttk.LabelFrame(left_frame, text="폴더 DB 비교")
        folder_db_frame.pack(fill="x", padx=5, pady=5)
        
        ttk.Label(folder_db_frame, text="원본 폴더:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.original_folder_db_var = tk.StringVar()
        ttk.Entry(folder_db_frame, textvariable=self.original_folder_db_var, width=40).grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        ttk.Button(folder_db_frame, text="찾아보기", 
                command=lambda: self.select_db_folder("original")).grid(row=0, column=2, padx=5, pady=5)
        
        ttk.Label(folder_db_frame, text="비교 폴더:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.compare_folder_db_var = tk.StringVar()
        ttk.Entry(folder_db_frame, textvariable=self.compare_folder_db_var, width=40).grid(row=1, column=1, padx=5, pady=5, sticky="ew")
        ttk.Button(folder_db_frame, text="찾아보기", 
                command=lambda: self.select_db_folder("compare")).grid(row=1, column=2, padx=5, pady=5)
        
        ttk.Button(folder_db_frame, text="DB 목록 보기", 
                command=self.show_db_list).grid(row=2, column=1, padx=5, pady=5, sticky="e")
        
        folder_db_frame.columnconfigure(1, weight=1)
        
        # 우측 프레임 (DB 목록 + 비교 옵션)
        right_frame = ttk.Frame(top_frame)
        right_frame.pack(side="right", fill="both", expand=True, padx=(5, 0))
        
        self.db_list_frame = ttk.LabelFrame(right_frame, text="비교할 DB 목록")
        self.db_list_frame.pack(fill="both", expand=True, padx=5, pady=5)
        
        self.db_checklist = ScrollableCheckList(self.db_list_frame, width=350, height=150)
        self.db_checklist.pack(fill="both", expand=True, padx=5, pady=5)
        
        options_frame = ttk.LabelFrame(right_frame, text="비교 옵션")
        options_frame.pack(fill="x", padx=5, pady=5)

        string_options_frame = ttk.LabelFrame(options_frame, text="STRING DB 옵션")
        string_options_frame.pack(fill="x", padx=5, pady=2)

        self.changed_kr_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(string_options_frame, text="KR 값이 변경된 항목 추출", 
                    variable=self.changed_kr_var).grid(row=0, column=0, padx=5, pady=2, sticky="w")

        self.new_items_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(string_options_frame, text="비교본에만 있는 새 항목 추출", 
                    variable=self.new_items_var).grid(row=1, column=0, padx=5, pady=2, sticky="w")

        self.deleted_items_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(string_options_frame, text="원본에서 삭제된 항목 추출", 
                    variable=self.deleted_items_var).grid(row=2, column=0, padx=5, pady=2, sticky="w")

        lang_options_frame = ttk.LabelFrame(options_frame, text="언어 옵션 (TRANSLATION DB용)")
        lang_options_frame.pack(fill="x", padx=5, pady=2)
        
        ttk.Label(lang_options_frame, text="비교 언어:").grid(row=0, column=0, padx=5, pady=2, sticky="w")
        lang_frame = ttk.Frame(lang_options_frame)
        lang_frame.grid(row=0, column=1, columnspan=2, padx=5, pady=2, sticky="w")
        
        self.compare_lang_vars = {}
        available_languages = ["KR", "EN", "CN", "TW", "TH"]
        for i, lang in enumerate(available_languages):
            var = tk.BooleanVar(value=True)
            self.compare_lang_vars[lang] = var
            ttk.Checkbutton(lang_frame, text=lang, variable=var).grid(
                row=0, column=i, padx=5, sticky="w")
        
        action_frame = ttk.Frame(self)
        action_frame.pack(fill="x", padx=5, pady=5)
        
        ttk.Button(action_frame, text="개별 DB 비교", 
                command=self.compare_individual_databases).pack(side="right", padx=5, pady=5)
        ttk.Button(action_frame, text="폴더 DB 비교", 
                command=self.compare_folder_databases).pack(side="right", padx=5, pady=5)
        ttk.Button(action_frame, text="결과 내보내기", 
                command=self.export_compare_results).pack(side="right", padx=5, pady=5)
        
        bottom_frame = ttk.Frame(self)
        bottom_frame.pack(fill="both", expand=True, padx=5, pady=5)
        
        result_frame = ttk.LabelFrame(bottom_frame, text="비교 결과")
        result_frame.pack(fill="both", expand=True, padx=5, pady=5)
        
        tree_frame = ttk.Frame(result_frame)
        tree_frame.pack(fill="both", expand=True)
        
        columns = ("db_name", "file_name", "sheet_name", "string_id", "type", "kr", "original_kr")
        self.compare_result_tree = ttk.Treeview(tree_frame, columns=columns, show="headings")
        
        self.compare_result_tree.heading("db_name", text="DB명")
        self.compare_result_tree.heading("file_name", text="파일명")
        self.compare_result_tree.heading("sheet_name", text="시트명")
        self.compare_result_tree.heading("string_id", text="STRING_ID")
        self.compare_result_tree.heading("type", text="변경 유형")
        self.compare_result_tree.heading("kr", text="KR")
        self.compare_result_tree.heading("original_kr", text="원본 KR")        
        
        self.compare_result_tree.column("db_name", width=120)
        self.compare_result_tree.column("file_name", width=120)
        self.compare_result_tree.column("sheet_name", width=120)
        self.compare_result_tree.column("string_id", width=100)
        self.compare_result_tree.column("type", width=150)
        self.compare_result_tree.column("kr", width=200)
        self.compare_result_tree.column("original_kr", width=200)
        
        scrollbar_y = ttk.Scrollbar(tree_frame, orient="vertical", command=self.compare_result_tree.yview)
        self.compare_result_tree.configure(yscrollcommand=scrollbar_y.set)
        
        scrollbar_x = ttk.Scrollbar(result_frame, orient="horizontal", command=self.compare_result_tree.xview)
        self.compare_result_tree.configure(xscrollcommand=scrollbar_x.set)
        
        scrollbar_y.pack(side="right", fill="y")
        self.compare_result_tree.pack(side="left", fill="both", expand=True)
        scrollbar_x.pack(side="bottom", fill="x")
        
        status_frame = ttk.Frame(self)
        status_frame.pack(fill="x", padx=5, pady=5)
        
        self.status_label_compare = ttk.Label(status_frame, text="대기 중...")
        self.status_label_compare.pack(side="left", fill="x", expand=True, padx=5)
        
        self.progress_label = ttk.Label(status_frame, text="진행 상황:")
        self.progress_label.pack(side="left", padx=5)
        
        self.progress_bar_compare = ttk.Progressbar(status_frame, length=300, mode="determinate")
        self.progress_bar_compare.pack(side="right", padx=5)
        
    def select_db_file(self, db_type):
        file_path = filedialog.askopenfilename(
            filetypes=[("DB 파일", "*.db"), ("모든 파일", "*.*")],
            title=f"{db_type.capitalize()} DB 파일 선택",
            parent=self
        )
        if file_path:
            if db_type == "original":
                self.original_db_var.set(file_path)
            else:
                self.compare_db_var.set(file_path)
            self.after(100, self.focus_force)
            self.after(100, self.lift)

    def select_db_folder(self, folder_type):
        folder = filedialog.askdirectory(title=f"{folder_type.capitalize()} DB 폴더 선택", parent=self)
        if folder:
            if folder_type == "original":
                self.original_folder_db_var.set(folder)
            else:
                self.compare_folder_db_var.set(folder)
            self.after(100, self.focus_force)
            self.after(100, self.lift)

    def show_db_list(self):
        original_folder = self.original_folder_db_var.get()
        compare_folder = self.compare_folder_db_var.get()
        
        if not original_folder or not os.path.isdir(original_folder):
            messagebox.showwarning("경고", "유효한 원본 DB 폴더를 선택하세요.", parent=self)
            return
        
        if not compare_folder or not os.path.isdir(compare_folder):
            messagebox.showwarning("경고", "유효한 비교 DB 폴더를 선택하세요.", parent=self)
            return
        
        original_dbs = {f for f in os.listdir(original_folder) if f.endswith('.db')}
        compare_dbs = {f for f in os.listdir(compare_folder) if f.endswith('.db')}
        
        common_dbs = original_dbs.intersection(compare_dbs)
        
        if not common_dbs:
            messagebox.showinfo("알림", "두 폴더에 공통된 DB 파일이 없습니다.", parent=self)
            return
        
        self.db_checklist.clear()
        self.db_pairs = []
        
        for db_file in sorted(common_dbs):
            self.db_checklist.add_item(db_file, checked=True)
            self.db_pairs.append({
                'file_name': db_file,
                'original_path': os.path.join(original_folder, db_file),
                'compare_path': os.path.join(compare_folder, db_file)
            })
        
        messagebox.showinfo("알림", f"{len(common_dbs)}개의 공통 DB 파일을 찾았습니다.", parent=self)

    def compare_individual_databases(self):
        original_db_path = self.original_db_var.get()
        compare_db_path = self.compare_db_var.get()

        if not original_db_path or not os.path.isfile(original_db_path):
            messagebox.showwarning("경고", "유효한 원본 DB 파일을 선택하세요.", parent=self)
            return

        if not compare_db_path or not os.path.isfile(compare_db_path):
            messagebox.showwarning("경고", "유효한 비교 DB 파일을 선택하세요.", parent=self)
            return

        self.compare_result_tree.delete(*self.compare_result_tree.get_children())
        self.compare_results = []

        loading_popup = LoadingPopup(self, "DB 비교 중", "DB 타입 확인 및 비교 중...")
        
        def run_comparison():
            try:
                result = self.db_compare_manager.auto_compare_databases(
                    original_db_path,
                    compare_db_path,
                    self.get_compare_options()
                )
                self.after(0, lambda: self.process_unified_compare_results(result, loading_popup))
            except Exception as e:
                self.after(0, lambda: [
                    loading_popup.close(),
                    messagebox.showerror("오류", f"DB 비교 중 오류 발생: {str(e)}", parent=self)
                ])
                
        thread = threading.Thread(target=run_comparison, daemon=True)
        thread.start()

    def compare_folder_databases(self):
        if not self.db_pairs:
            messagebox.showwarning("경고", "비교할 DB 파일 목록이 없습니다. 'DB 목록 보기'를 먼저 실행하세요.", parent=self)
            return
        
        selected_db_names = self.db_checklist.get_checked_items()
        selected_db_pairs = [pair for pair in self.db_pairs if pair['file_name'] in selected_db_names]
        
        if not selected_db_pairs:
            messagebox.showwarning("경고", "비교할 DB 파일을 선택하세요.", parent=self)
            return
        
        self.compare_result_tree.delete(*self.compare_result_tree.get_children())
        self.compare_results = []
        
        loading_popup = LoadingPopup(self, "폴더 DB 비교 중", "DB 비교 작업 준비 중...")
        
        def run_comparison():
            try:
                def progress_callback(status_text, current, total):
                    self.after(0, lambda: loading_popup.update_progress(
                        (current / total) * 100, status_text))
                
                result = self.db_compare_manager.auto_compare_folder_databases(
                    selected_db_pairs,
                    self.get_compare_options(),
                    progress_callback
                )
                self.after(0, lambda: self.process_unified_compare_results(result, loading_popup))
            except Exception as e:
                self.after(0, lambda: [
                    loading_popup.close(),
                    messagebox.showerror("오류", f"폴더 DB 비교 중 오류 발생: {str(e)}", parent=self)
                ])

        thread = threading.Thread(target=run_comparison, daemon=True)
        thread.start()

    def get_compare_options(self):
        return {
            "changed_kr": self.changed_kr_var.get(),
            "new_items": self.new_items_var.get(),
            "deleted_items": self.deleted_items_var.get(),
            "languages": [lang.lower() for lang, var in self.compare_lang_vars.items() if var.get()]
        }

    def process_unified_compare_results(self, result, loading_popup):
        loading_popup.close()
        
        if result["status"] != "success":
            messagebox.showerror("오류", result["message"], parent=self)
            return
            
        self.compare_results = result["compare_results"]
        
        self.compare_result_tree.delete(*self.compare_result_tree.get_children())
        for idx, item in enumerate(self.compare_results):
            db_name = item.get("db_name", item.get("file_name", ""))
            file_name = item.get("file_name", "")
            sheet_name = item.get("sheet_name", "")
            string_id = item.get("string_id", "")
            item_type = item.get("type", "")
            kr = item.get("kr", "")
            original_kr = item.get("original_kr", "")
            
            self.compare_result_tree.insert(
                "", "end", iid=idx,
                values=(db_name, file_name, sheet_name, string_id, item_type, kr, original_kr)
            )
        
        total_changes = result.get("total_changes", len(self.compare_results))
        db_type = result.get("db_type", "DB")
        
        self.status_label_compare.config(
            text=f"{db_type} 비교 완료: {total_changes}개 차이점 발견"
        )
        
        if total_changes > 0:
            summary_msg = f"{db_type} 비교가 완료되었습니다.\n\n🔍 총 {total_changes}개의 차이점을 발견했습니다."
            if "new_items" in result:
                summary_msg += f"\n\n📊 세부 결과:\n• 신규 항목: {result.get('new_items', 0)}개\n• 삭제된 항목: {result.get('deleted_items', 0)}개\n• 변경된 항목: {result.get('changed_items', 0)}개"
            messagebox.showinfo("완료", summary_msg, parent=self)
        else:
            messagebox.showinfo("완료", f"두 {db_type}가 동일합니다.", parent=self)

    def export_compare_results(self):
        if not self.compare_results:
            messagebox.showwarning("경고", "내보낼 비교 결과가 없습니다.", parent=self)
            return
        
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel 파일", "*.xlsx")],
            title="비교 결과 저장",
            parent=self
        )
        
        if not file_path:
            return
            
        loading_popup = LoadingPopup(self, "결과 내보내기", "엑셀 파일로 저장 중...")
        
        def export_data():
            try:
                df = pd.DataFrame(self.compare_results)
                df.to_excel(file_path, index=False)
                self.after(0, lambda: [
                    loading_popup.close(),
                    messagebox.showinfo("완료", f"비교 결과가 {file_path}에 저장되었습니다.", parent=self)
                ])
            except Exception as e:
                self.after(0, lambda: [
                    loading_popup.close(),
                    messagebox.showerror("오류", f"데이터 저장 실패: {str(e)}", parent=self)
                ])
                
        thread = threading.Thread(target=export_data, daemon=True)
        thread.start()