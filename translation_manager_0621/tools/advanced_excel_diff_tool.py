#advanced_excel_diff_tool

# tools/advanced_excel_diff_tool.py
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading

from .advanced_excel_diff_manager import AdvancedExcelDiffManager
from ui.common_components import ScrollableCheckList

class AdvancedExcelDiffTool(ttk.Frame):
    """
    고급 엑셀 비교 기능의 UI와 사용자 상호작용을 담당합니다.
    실제 데이터 처리는 AdvancedExcelDiffManager에 위임합니다.
    """
    def __init__(self, parent):
        super().__init__(parent)
        self.manager = AdvancedExcelDiffManager()
        self.last_report_df = None
        
        # --- [신규] 폴더에서 생성된 데이터프레임을 저장할 변수 ---
        self.source_df_from_folder = None
        self.target_df_from_folder = None
        
        # UI 컨트롤 변수 선언
        self.file_path_old = tk.StringVar()
        self.file_path_new = tk.StringVar()
        self.header_old = tk.StringVar(value="1")
        self.header_new = tk.StringVar(value="1")

        # --- [개선 2] Key 컬럼 선택을 위한 BooleanVar 추가 ---
        self.key_vars = {
            "STRING_ID": tk.BooleanVar(value=True),
            "KR": tk.BooleanVar(value=False),
            "CN": tk.BooleanVar(value=False),
            "TW": tk.BooleanVar(value=False)
        }

        self.filter_column = tk.StringVar()
        self.filter_text = tk.StringVar()

        # --- [신규] 폴더 로드 UI를 위한 변수 ---
        self.folder_path_var = tk.StringVar()
        
        self._setup_ui()

    def _setup_ui(self):
        """전체 UI 레이아웃을 구성합니다."""
        # --- 1. 폴더에서 데이터 로드 패널 (신규) ---
        folder_load_panel = ttk.LabelFrame(self, text="[옵션] 폴더 단위로 데이터 불러오기")
        folder_load_panel.pack(fill="x", expand=True, padx=10, pady=5)
        self._create_folder_panel(folder_load_panel)

        # --- 2. 개별 파일 선택 패널 ---            
        file_panel = ttk.Frame(self)
        file_panel.pack(fill="x", expand=True, padx=10, pady=5)

        self.old_file_panel, self.sheet_checklist_old = self._create_file_panel(file_panel, "1. 원본 파일", self.file_path_old, self._load_sheets_old)
        self.old_file_panel.pack(side="left", fill="both", expand=True, padx=5)

        self.new_file_panel, self.sheet_checklist_new = self._create_file_panel(file_panel, "2. 대상 파일", self.file_path_new, self._load_sheets_new)
        self.new_file_panel.pack(side="right", fill="both", expand=True, padx=5)

        options_panel = ttk.LabelFrame(self, text="3. 비교 옵션 설정")
        options_panel.pack(fill="x", expand=True, padx=10, pady=5)

        # --- [개선 2] Key 컬럼 선택 UI를 체크박스로 변경 ---
        key_frame = ttk.Frame(options_panel)
        key_frame.pack(fill="x", padx=5, pady=2)
        ttk.Label(key_frame, text="비교 Key 컬럼:").pack(side="left", padx=5)
        for key_name, var in self.key_vars.items():
            ttk.Checkbutton(key_frame, text=key_name, variable=var).pack(side="left", padx=5)

        filter_frame = ttk.LabelFrame(options_panel, text="행 필터링 (선택 사항)")
        filter_frame.pack(fill="x", padx=5, pady=5)

        # 필터링 설정
        filter_frame = ttk.LabelFrame(options_panel, text="행 필터링 (선택 사항)")
        filter_frame.pack(fill="x", padx=5, pady=5)
        ttk.Label(filter_frame, text="필터 적용할 컬럼명:").pack(side="left", padx=5)
        ttk.Entry(filter_frame, textvariable=self.filter_column, width=20).pack(side="left", padx=5)
        ttk.Label(filter_frame, text="포함할 텍스트:").pack(side="left", padx=5)
        ttk.Entry(filter_frame, textvariable=self.filter_text, width=20).pack(side="left", padx=5)

        # --- 실행 및 로그 패널 ---
        action_panel = ttk.Frame(self)
        action_panel.pack(fill="x", padx=10, pady=10)

        self.compare_button = ttk.Button(action_panel, text="4. 비교 시작", command=self._start_comparison_thread, style="Accent.TButton")
        self.compare_button.pack(side="left", padx=5)

        self.export_button = ttk.Button(action_panel, text="5. 결과 저장", state="disabled", command=self._save_report)
        self.export_button.pack(side="left", padx=5)

        self.status_label = ttk.Label(action_panel, text="대기 중...")
        self.status_label.pack(side="right", padx=5)

        # 로그
        log_frame = ttk.LabelFrame(self, text="결과 로그")
        log_frame.pack(fill="both", expand=True, padx=10, pady=5)
        self.log_text = tk.Text(log_frame, wrap="word", height=12)
        self.log_text.pack(fill="both", expand=True)

    def _create_folder_panel(self, parent):
        """폴더에서 데이터를 로드하는 UI 패널을 생성합니다."""
        top_frame = ttk.Frame(parent)
        top_frame.pack(fill="x", padx=5, pady=5)

        ttk.Label(top_frame, text="폴더 경로:").pack(side="left")
        ttk.Entry(top_frame, textvariable=self.folder_path_var).pack(side="left", fill="x", expand=True, padx=5)
        ttk.Button(top_frame, text="폴더 선택", command=self._select_folder).pack(side="left", padx=5)
        ttk.Button(top_frame, text="파일 검색", command=self._search_files_in_folder).pack(side="left")

        list_frame = ttk.Frame(parent)
        list_frame.pack(fill="x", padx=5, pady=5)

        self.folder_file_checklist = ScrollableCheckList(list_frame, height=100)
        self.folder_file_checklist.pack(side="left", fill="x", expand=True)

        button_frame = ttk.Frame(list_frame)
        button_frame.pack(side="left", padx=10)
        ttk.Button(button_frame, text="선택 파일을 '원본'으로 설정", command=lambda: self._build_df_from_folder(is_source=True)).pack(pady=5, fill="x")
        ttk.Button(button_frame, text="선택 파일을 '대상'으로 설정", command=lambda: self._build_df_from_folder(is_source=False)).pack(pady=5, fill="x")


    def _create_file_panel(self, parent, title, path_var, load_cmd):
        """파일과 시트 선택을 위한 UI 패널 한 개를 생성합니다."""
        frame = ttk.LabelFrame(parent, text=title)
        top_frame = ttk.Frame(frame)
        top_frame.pack(fill="x", padx=5, pady=5)
        ttk.Entry(top_frame, textvariable=path_var, width=40).pack(side="left", fill="x", expand=True)
        ttk.Button(top_frame, text="파일 선택", command=lambda: self._browse_file(path_var)).pack(side="left", padx=5)

        sheet_frame = ttk.Frame(frame)
        sheet_frame.pack(fill="both", expand=True, padx=5, pady=5)
        ttk.Button(sheet_frame, text="시트 불러오기", command=load_cmd).pack(fill="x")

        # --- [개선 1] Listbox 대신 ScrollableCheckList 사용 및 크기 확장 ---
        checklist = ScrollableCheckList(sheet_frame, height=150) # 높이 150으로 확장
        checklist.pack(fill="both", expand=True, pady=(5,0))

        return frame, checklist

    def _select_folder(self):
        folder_path = filedialog.askdirectory()
        if folder_path:
            self.folder_path_var.set(folder_path)

    def _search_files_in_folder(self):
        folder_path = self.folder_path_var.get()
        if not folder_path:
            messagebox.showwarning("경고", "먼저 폴더를 선택해주세요.")
            return
        try:
            files = self.manager.find_string_excels_in_folder(folder_path)
            self.folder_file_checklist.clear()
            for f in files:
                self.folder_file_checklist.add_item(f, checked=True)
            self._log(f"{len(files)}개의 'string...' 엑셀 파일을 찾았습니다.")
        except Exception as e:
            messagebox.showerror("오류", str(e))

    def _build_df_from_folder(self, is_source):
        """폴더 내 선택된 파일들을 읽어 데이터프레임으로 만듭니다."""
        folder_path = self.folder_path_var.get()
        selected_files = self.folder_file_checklist.get_checked_items()
        header_row = self.header_old.get() if is_source else self.header_new.get()

        if not folder_path or not selected_files:
            messagebox.showwarning("경고", "폴더와 하나 이상의 파일을 선택해야 합니다.")
            return

        target_label = "'원본'" if is_source else "'대상'"
        self.status_label.config(text=f"{target_label} 데이터 구성 중...")
        self._log(f"{target_label}으로 사용할 데이터를 폴더에서 구성합니다...")

        def worker():
            try:
                df = self.manager.create_dataframe_from_files(folder_path, selected_files, header_row)
                self.after(0, self._on_df_build_complete, df, is_source)
            except Exception as e:
                self.after(0, self._on_df_build_error, e)

        threading.Thread(target=worker, daemon=True).start()

    def _on_df_build_complete(self, df, is_source):
        target_label = "'원본'" if is_source else "'대상'"
        if is_source:
            self.source_df_from_folder = df
            self.file_path_old.set(f"[폴더에서 로드됨: {len(df)} 행]")
        else:
            self.target_df_from_folder = df
            self.file_path_new.set(f"[폴더에서 로드됨: {len(df)} 행]")

        self.status_label.config(text="대기 중...")
        self._log(f"{target_label} 데이터 구성 완료. {len(df)}개의 행이 로드되었습니다.")
        messagebox.showinfo("완료", f"{target_label} 데이터 구성이 완료되었습니다.")

    def _on_df_build_error(self, error):
        self.status_label.config(text="오류 발생.")
        self._log(f"데이터 구성 중 오류: {error}")
        messagebox.showerror("오류", f"폴더에서 데이터를 구성하는 중 오류가 발생했습니다:\n{error}")

    def _gather_parameters(self):
        """UI 파라미터 수집 로직 수정. 폴더 데이터 우선 사용."""
        selected_keys = [key for key, var in self.key_vars.items() if var.get()]
        if not selected_keys: raise ValueError("하나 이상의 Key 컬럼을 선택해야 합니다.")

        params = {
            "key_columns": selected_keys,
            "filter_options": { 'column': self.filter_column.get(), 'text': self.filter_text.get() }
        }

        # 원본 데이터 설정
        if self.source_df_from_folder is not None:
            params["df_old"] = self.source_df_from_folder
        else:
            params["file_path_old"] = self.file_path_old.get()
            params["sheets_old"] = self.sheet_checklist_old.get_checked_items()
            params["header_old"] = self.header_old.get()
            if not params["file_path_old"] or not params["sheets_old"]:
                raise ValueError("원본 파일과 시트를 선택하거나, 폴더에서 원본 데이터를 구성해야 합니다.")

        # 대상 데이터 설정
        if self.target_df_from_folder is not None:
            params["df_new"] = self.target_df_from_folder
        else:
            params["file_path_new"] = self.file_path_new.get()
            params["sheets_new"] = self.sheet_checklist_new.get_checked_items()
            params["header_new"] = self.header_new.get()
            if not params["file_path_new"] or not params["sheets_new"]:
                raise ValueError("대상 파일과 시트를 선택하거나, 폴더에서 대상 데이터를 구성해야 합니다.")

        return params
    
    def _browse_file(self, path_var):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
        if file_path:
            path_var.set(file_path)

    def _load_sheets_old(self):
        self._load_sheets_into_checklist(self.file_path_old.get(), self.sheet_checklist_old)

    def _load_sheets_new(self):
        self._load_sheets_into_checklist(self.file_path_new.get(), self.sheet_checklist_new)

    def _load_sheets_into_checklist(self, file_path, checklist):
        """[개선 1] 시트 이름을 체크리스트에 추가합니다."""
        if not file_path:
            messagebox.showwarning("경고", "먼저 엑셀 파일을 선택해주세요.")
            return
        try:
            sheet_names = self.manager.get_sheet_names(file_path)
            checklist.clear()
            for name in sheet_names:
                checklist.add_item(name, checked=True) # 기본적으로 모두 체크된 상태로 추가
        except Exception as e:
            messagebox.showerror("오류", str(e))
            
    def _log(self, message):
        self.log_text.insert(tk.END, f"[{pd.Timestamp.now().strftime('%H:%M:%S')}] {message}\n")
        self.log_text.see(tk.END)

    def _start_comparison_thread(self):
        """사용자 입력값을 검증하고 백그라운드 스레드에서 비교를 시작합니다."""
        try:
            # 입력값 검증
            params = self._gather_parameters()
        except ValueError as e:
            messagebox.showerror("입력 오류", str(e))
            return

        # UI 상태 변경
        self.compare_button.config(state="disabled")
        self.export_button.config(state="disabled")
        self.status_label.config(text="비교 작업 실행 중...")
        self.log_text.delete(1.0, tk.END)
        self._log("비교를 시작합니다...")

        # 백그라운드 스레드 실행
        thread = threading.Thread(target=self._run_comparison_worker, args=(params,), daemon=True)
        thread.start()

    def _gather_parameters(self):
        """UI에서 사용자가 입력한 모든 파라미터를 수집하고 검증합니다."""
        # --- [개선 2] 체크박스에서 Key 컬럼 목록 생성 ---
        selected_keys = [key for key, var in self.key_vars.items() if var.get()]

        params = {
            "file_path_old": self.file_path_old.get(),
            "file_path_new": self.file_path_new.get(),
            "header_old": self.header_old.get(),
            "header_new": self.header_new.get(),
            "key_columns": selected_keys,
            # --- [개선 1] 체크리스트에서 선택된 시트 목록 가져오기 ---
            "sheets_old": self.sheet_checklist_old.get_checked_items(),
            "sheets_new": self.sheet_checklist_new.get_checked_items(),
            "filter_options": {
                'column': self.filter_column.get(),
                'text': self.filter_text.get()
            }
        }
        if not all([params["file_path_old"], params["file_path_new"], params["key_columns"]]):
            raise ValueError("원본/대상 파일과 하나 이상의 Key 컬럼은 반드시 선택해야 합니다.")
        if not all([params["sheets_old"], params["sheets_new"]]):
            raise ValueError("원본과 대상 파일 각각 하나 이상의 시트를 선택해야 합니다.")
        return params

    def _run_comparison_worker(self, params):
        """백그라운드 스레드에서 실행될 실제 작업 함수입니다."""
        result = self.manager.run_comparison(**params)
        self.after(0, self._process_comparison_result, result)

    def _process_comparison_result(self, result):
        """비교 결과를 받아 UI를 업데이트합니다."""
        self._log(result["message"])
        self.compare_button.config(state="normal")

        if result["status"] == "success":
            self.last_report_df = result["report_df"]
            self.export_button.config(state="normal")
            self.status_label.config(text="비교 완료. 결과 저장이 가능합니다.")
            messagebox.showinfo("완료", "비교가 완료되었습니다. 로그를 확인하고 결과를 저장하세요.")
        else:
            self.last_report_df = None
            self.status_label.config(text="오류 발생.")
            messagebox.showerror("비교 오류", result["message"])

    def _save_report(self):
        """결과 보고서를 엑셀 파일로 저장합니다."""
        if self.last_report_df is None:
            messagebox.showwarning("경고", "저장할 비교 결과가 없습니다.")
            return

        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            title="비교 결과 보고서 저장"
        )
        if save_path:
            try:
                df = self.last_report_df
                # 스타일링을 위해 ExcelWriter 사용
                with pd.ExcelWriter(save_path, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False, sheet_name='diff_report')
                    workbook = writer.book
                    worksheet = writer.sheets['diff_report']

                    # 상태에 따라 행 배경색 변경
                    for idx, status in enumerate(df['상태'], 2): # 2부터 시작 (헤더 제외)
                        fill = None
                        if status == '추가':
                            fill = tk.PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')
                        elif status == '삭제':
                            fill = tk.PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')
                        elif status == '변경':
                            fill = tk.PatternFill(start_color='FFEB9C', end_color='FFEB9C', fill_type='solid')

                        if fill:
                            for cell in worksheet[f'A{idx}':f'{chr(ord("A") + len(df.columns) - 1)}{idx}'][0]:
                                cell.fill = fill
                self._log(f"보고서가 성공적으로 저장되었습니다: {save_path}")
                messagebox.showinfo("성공", f"보고서가 성공적으로 저장되었습니다.")
            except Exception as e:
                self._log(f"보고서 저장 중 오류 발생: {e}")
                messagebox.showerror("저장 오류", f"보고서를 저장하는 중 오류가 발생했습니다:\n{e}")