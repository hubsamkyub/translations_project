import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import threading
from .excel_diff_manager import ExcelDiffManager # 새로 만든 매니저를 import

class ExcelDiffTool(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.manager = ExcelDiffManager()
        self.setup_ui()

    def setup_ui(self):
        # 파일 선택 프레임
        file_frame = ttk.LabelFrame(self, text="파일 선택")
        file_frame.pack(fill="x", padx=10, pady=5)

        self.file_path1_var = tk.StringVar()
        self.file_path2_var = tk.StringVar()

        self._create_file_input(file_frame, "파일 1:", self.file_path1_var, 0)
        self._create_file_input(file_frame, "파일 2:", self.file_path2_var, 1)

        # 설정 프레임
        options_frame = ttk.LabelFrame(self, text="비교 설정")
        options_frame.pack(fill="x", padx=10, pady=5)

        self.key_column_var = tk.StringVar(value="STRING_ID")
        self.header1_var = tk.StringVar(value="1")
        self.header2_var = tk.StringVar(value="1")
        self.sheet_name1_var = tk.StringVar(value="Sheet1")
        self.sheet_name2_var = tk.StringVar(value="Sheet1")

        self._create_option_input(options_frame, "Key 컬럼:", self.key_column_var, 0)
        self._create_option_input(options_frame, "파일 1 헤더 행:", self.header1_var, 1)
        self._create_option_input(options_frame, "파일 2 헤더 행:", self.header2_var, 2)
        self._create_option_input(options_frame, "파일 1 시트 이름:", self.sheet_name1_var, 3)
        self._create_option_input(options_frame, "파일 2 시트 이름:", self.sheet_name2_var, 4)

        # 실행 프레임
        action_frame = ttk.Frame(self)
        action_frame.pack(fill="x", padx=10, pady=10)

        self.compare_button = ttk.Button(action_frame, text="비교 시작", command=self.start_comparison_thread)
        self.compare_button.pack(side="right", padx=5)

        self.status_label = ttk.Label(action_frame, text="대기 중...")
        self.status_label.pack(side="left", padx=5)

        self.progress_bar = ttk.Progressbar(action_frame, orient="horizontal", mode="indeterminate")
        self.progress_bar.pack(side="left", fill="x", expand=True)

        # 로그 프레임
        log_frame = ttk.LabelFrame(self, text="결과 로그")
        log_frame.pack(fill="both", expand=True, padx=10, pady=5)

        self.log_text = tk.Text(log_frame, wrap="word", height=10)
        self.log_text.pack(fill="both", expand=True, side="left")

        scrollbar = ttk.Scrollbar(log_frame, command=self.log_text.yview)
        scrollbar.pack(side="right", fill="y")
        self.log_text.config(yscrollcommand=scrollbar.set)

    def _create_file_input(self, parent, label_text, var, row):
        """파일 경로 입력을 위한 UI 요소를 생성합니다."""
        ttk.Label(parent, text=label_text).grid(row=row, column=0, padx=5, pady=2, sticky="w")
        entry = ttk.Entry(parent, textvariable=var, width=60)
        entry.grid(row=row, column=1, padx=5, pady=2, sticky="ew")
        button = ttk.Button(parent, text="찾아보기", command=lambda: self.browse_file(var))
        button.grid(row=row, column=2, padx=5, pady=2)
        parent.columnconfigure(1, weight=1)

    def _create_option_input(self, parent, label_text, var, row):
        """옵션 입력을 위한 UI 요소를 생성합니다."""
        ttk.Label(parent, text=label_text).grid(row=row, column=0, padx=5, pady=2, sticky="w")
        ttk.Entry(parent, textvariable=var, width=20).grid(row=row, column=1, padx=5, pady=2, sticky="w")

    def browse_file(self, var):
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if file_path:
            var.set(file_path)

    def log(self, message):
        self.log_text.insert(tk.END, f"[{pd.Timestamp.now().strftime('%H:%M:%S')}] {message}\n")
        self.log_text.see(tk.END)

    def start_comparison_thread(self):
        """UI 멈춤 현상을 방지하기 위해 별도 스레드에서 비교를 시작합니다."""
        # 입력값 유효성 검사
        if not all([self.file_path1_var.get(), self.file_path2_var.get(), self.key_column_var.get()]):
            messagebox.showerror("입력 오류", "파일 1, 파일 2, Key 컬럼은 모두 필수입니다.")
            return

        # UI 비활성화 및 초기화
        self.compare_button.config(state="disabled")
        self.status_label.config(text="비교 준비 중...")
        self.progress_bar.start(10)

        # 스레드 생성 및 시작
        thread = threading.Thread(target=self._run_comparison_worker, daemon=True)
        thread.start()

    def _run_comparison_worker(self):
        """실제 데이터 비교를 수행하는 워커 함수 (백그라운드 스레드에서 실행)"""
        self.log("데이터 비교를 시작합니다...")
        self.status_label.config(text="엑셀 파일 읽는 중...")

        result = self.manager.run_comparison(
            file_path1=self.file_path1_var.get(),
            file_path2=self.file_path2_var.get(),
            key_column=self.key_column_var.get(),
            header1=self.header1_var.get(),
            header2=self.header2_var.get(),
            sheet_name1=self.sheet_name1_var.get(),
            sheet_name2=self.sheet_name2_var.get()
        )

        # UI 업데이트는 메인 스레드에서 안전하게 처리
        self.after(0, self._process_comparison_result, result)

    def _process_comparison_result(self, result):
        """워커 스레드의 결과를 받아 UI를 업데이트합니다."""
        self.progress_bar.stop()
        self.compare_button.config(state="normal")
        self.status_label.config(text="대기 중...")

        self.log(result["message"])

        if result["status"] == "success":
            save_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")],
                title="비교 결과 보고서 저장"
            )
            if save_path:
                self.status_label.config(text="보고서 저장 중...")
                save_result = self.manager.save_report_to_excel(result["report_df"], save_path)
                self.log(save_result["message"])
                if save_result["status"] == "success":
                     messagebox.showinfo("완료", save_result["message"])
                else:
                     messagebox.showerror("오류", save_result["message"])
                self.status_label.config(text="대기 중...")
        else:
            messagebox.showerror("비교 오류", result["message"])