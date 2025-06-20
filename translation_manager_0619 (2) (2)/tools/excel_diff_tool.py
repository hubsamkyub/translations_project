# tools/excel_diff_tool.py

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import threading
import pandas as pd

# 내부 유틸리티 및 컴포넌트 임포트
from ui.common_components import ImprovedFileList, LoadingPopup
from utils.config_utils import load_config, save_config
from tools.excel_diff_manager import ExcelDiffManager

class ExcelDiffTool(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.root = parent
        
        # 로직 매니저 인스턴스 생성
        self.diff_manager = ExcelDiffManager(self)
        
        # 설정 파일 로드
        self.config = load_config()
        
        # UI 변수 초기화
        self.source_path = tk.StringVar(value=self.config.get("source_path", ""))
        self.target_path = tk.StringVar(value=self.config.get("target_path", ""))
        self.status_var = tk.StringVar(value="준비 완료")
        
        # UI 빌드
        self.build_ui()

        # 저장된 경로가 있으면 파일 목록 자동 업데이트
        if self.source_path.get() and os.path.exists(self.source_path.get()):
            self.root.after(100, lambda: self.update_file_list(self.source_path.get()))

    def build_ui(self):
        """UI 구성 - 기존 구조 유지"""
        self.build_path_ui()
        self.build_file_and_option_ui()
        self.build_result_ui()

    def build_path_ui(self):
        """경로 선택 UI 구성"""
        path_frame = ttk.LabelFrame(self, text="파일/폴더 선택")
        path_frame.pack(fill="x", padx=10, pady=5)
        
        source_frame = ttk.Frame(path_frame)
        source_frame.pack(fill="x", padx=5, pady=5)
        
        ttk.Label(source_frame, text="원본 경로:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        source_entry = ttk.Entry(source_frame, textvariable=self.source_path, width=50)
        source_entry.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        source_entry.bind('<Return>', lambda event: self.update_file_list(self.source_path.get()))
        
        ttk.Button(source_frame, text="폴더", command=lambda: self.select_folder("source")).grid(row=0, column=2, padx=2)
        ttk.Button(source_frame, text="파일", command=lambda: self.select_file("source")).grid(row=0, column=3, padx=2)
        ttk.Button(source_frame, text="새로고침", command=lambda: self.update_file_list(self.source_path.get())).grid(row=0, column=4, padx=2)
        
        source_frame.columnconfigure(1, weight=1)

        target_frame = ttk.Frame(path_frame)
        target_frame.pack(fill="x", padx=5, pady=5)
        
        ttk.Label(target_frame, text="비교본 경로:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        ttk.Entry(target_frame, textvariable=self.target_path, width=50).grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        ttk.Button(target_frame, text="폴더", command=lambda: self.select_folder("target")).grid(row=0, column=2, padx=2)
        ttk.Button(target_frame, text="파일", command=lambda: self.select_file("target")).grid(row=0, column=3, padx=2)
        target_frame.columnconfigure(1, weight=1)

    def build_file_and_option_ui(self):
        """파일 목록과 옵션 UI 구성"""
        main_frame = ttk.Frame(self)
        main_frame.pack(fill="both", expand=True, padx=10, pady=5)
        
        left_frame = ttk.LabelFrame(main_frame, text="비교 대상 파일")
        left_frame.pack(side="left", fill="both", expand=True, padx=(0, 5))
        
        self.source_file_list = ImprovedFileList(left_frame, width=40, height=15)
        self.source_file_list.pack(fill="both", expand=True, padx=5, pady=5)
        
        right_frame = ttk.Frame(main_frame)
        right_frame.pack(side="right", fill="y", padx=(5, 0))
        
        button_frame = ttk.LabelFrame(right_frame, text="실행")
        button_frame.pack(fill="x", pady=5)
        
        ttk.Button(button_frame, text="비교 시작", command=self.start_comparison, width=15).pack(pady=5, padx=10)
        ttk.Button(button_frame, text="결과 내보내기", command=self.export_results, width=15).pack(pady=5, padx=10)

    def build_result_ui(self):
        """결과 표시 UI 구성"""
        result_frame = ttk.LabelFrame(self, text="비교 결과")
        result_frame.pack(fill="both", expand=True, padx=10, pady=5)
        
        self.summary_tree = ttk.Treeview(result_frame, columns=("파일", "신규", "수정", "삭제"), show="headings")
        vsb = ttk.Scrollbar(result_frame, orient="vertical", command=self.summary_tree.yview)
        hsb = ttk.Scrollbar(result_frame, orient="horizontal", command=self.summary_tree.xview)
        self.summary_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        self.summary_tree.heading("파일", text="파일")
        self.summary_tree.heading("신규", text="신규")
        self.summary_tree.heading("수정", text="수정")
        self.summary_tree.heading("삭제", text="삭제")
        
        self.summary_tree.column("파일", width=300)
        
        vsb.pack(side="right", fill="y")
        hsb.pack(side="bottom", fill="x")
        self.summary_tree.pack(fill="both", expand=True)
        self.summary_tree.bind("<Double-1>", self.show_detail_popup)
        
        status_frame = ttk.Frame(self)
        status_frame.pack(fill="x", padx=10, pady=5, side="bottom")
        ttk.Label(status_frame, textvariable=self.status_var).pack(side="left")

    def select_folder(self, target):
        folder = filedialog.askdirectory()
        if folder:
            if target == "source":
                self.source_path.set(folder)
                self.target_path.set(folder)
                self.update_file_list(folder)
            else:
                self.target_path.set(folder)
            self.save_paths()

    def select_file(self, target):
        file = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
        if file:
            if target == "source":
                self.source_path.set(file)
                self.target_path.set(file)
                self.update_file_list(file)
            else:
                self.target_path.set(file)
            self.save_paths()

    def update_file_list(self, path):
        self.source_file_list.clear()
        if os.path.isfile(path):
            self.source_file_list.add_item(os.path.basename(path), checked=True)
        elif os.path.isdir(path):
            files = self.diff_manager.get_excel_files(path)
            for f in sorted(files):
                self.source_file_list.add_item(os.path.relpath(f, path).replace('\\', '/'), checked=True)
        self.status_var.set(f"파일 목록 업데이트 완료: {self.source_file_list.listbox.size()}개 파일")

    def save_paths(self):
        self.config["source_path"] = self.source_path.get()
        self.config["target_path"] = self.target_path.get()
        save_config(self.config)
        self.status_var.set("경로가 저장되었습니다.")

    def start_comparison(self):
        source_path = self.source_path.get()
        target_path = self.target_path.get()
        if not source_path or not target_path:
            messagebox.showwarning("경로 오류", "원본과 비교본 경로를 모두 지정해주세요.")
            return

        selected_files = self.source_file_list.get_checked_items()
        loading_popup = LoadingPopup(self.root, title="비교 진행 중", message="준비 중...")
        
        thread = threading.Thread(
            target=self.diff_manager.start_comparison_logic,
            args=(source_path, target_path, loading_popup, selected_files)
        )
        thread.daemon = True
        thread.start()
        
    def update_results_ui(self, loading_popup):
        if loading_popup: loading_popup.close()
        self.summary_tree.delete(*self.summary_tree.get_children())
        
        file_summary = {}
        for file, sheet, row, col, src, tgt, status in self.diff_manager.diff_results:
            if file not in file_summary:
                file_summary[file] = {"추가": 0, "수정": 0, "삭제": 0}
            
            if "추가" in status: file_summary[file]["추가"] += 1
            elif "삭제" in status: file_summary[file]["삭제"] += 1
            elif "변경" in status: file_summary[file]["수정"] += 1
            
        for file, counts in file_summary.items():
            self.summary_tree.insert("", "end", values=(file, counts["추가"], counts["수정"], counts["삭제"]))
            
        total_changes = len(self.diff_manager.diff_results)
        self.status_var.set(f"비교 완료. 총 {total_changes}개의 변경 사항 발견.")

    def show_error(self, message, loading_popup):
        if loading_popup: loading_popup.close()
        self.status_var.set(f"오류 발생: {message}")
        messagebox.showerror("오류", message)

    def export_results(self):
        if not self.diff_manager.diff_results:
            messagebox.showinfo("알림", "내보낼 비교 결과가 없습니다.")
            return
        
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel 파일", "*.xlsx")],
            title="비교 결과 저장"
        )
        if not file_path: return
        
        df = pd.DataFrame(self.diff_manager.diff_results, columns=["파일", "시트", "행/PK", "열", "원본 값", "비교본 값", "상태"])
        df.to_excel(file_path, index=False)
        messagebox.showinfo("완료", f"비교 결과가 '{file_path}'에 저장되었습니다.")

    def show_detail_popup(self, event):
        item = self.summary_tree.focus()
        if not item: return
        file_name = self.summary_tree.item(item, "values")[0]
        
        file_results = [r for r in self.diff_manager.diff_results if r[0] == file_name]
        
        popup = tk.Toplevel(self.root)
        popup.title(f"상세 결과 - {file_name}")
        popup.geometry("1000x600")

        tree = ttk.Treeview(popup, columns=("시트", "행/PK", "열", "원본 값", "비교본 값", "상태"), show="headings")
        for col in tree["columns"]:
            tree.heading(col, text=col)
        tree.pack(fill="both", expand=True)

        for result in file_results:
            tree.insert("", "end", values=result[1:])