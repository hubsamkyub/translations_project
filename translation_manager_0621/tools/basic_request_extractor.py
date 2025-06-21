# tools/basic_request_extractor.py (Refactored)

import tkinter as tk
from tkinter import ttk, filedialog
import os
import threading
from ui.common_components import show_message

class BasicRequestExtractor(tk.Frame):
    def __init__(self, parent, parent_app):
        super().__init__(parent)
        self.parent_app = parent_app

        #체크박스 기본 설정
        self.extract_new_var = tk.BooleanVar(value=True)
        self.extract_change_var = tk.BooleanVar(value=False)
        self.mark_as_transferred_var = tk.BooleanVar(value=False)
        self.save_to_db_var = tk.BooleanVar(value=False)
        self.full_results = []

        self.setup_ui()
        self._toggle_db_path_entry() # 초기 UI 상태 설정

    def setup_ui(self):
        """'기본 추출' 탭의 UI 구성"""
        # --- 추출 조건 ---
        condition_frame = ttk.LabelFrame(self, text="2. 추출 조건")
        condition_frame.pack(fill="x", padx=5, pady=5)
        
        ttk.Checkbutton(condition_frame, text="#번역요청 컬럼 값이 '신규'인 행", variable=self.extract_new_var).pack(anchor="w", padx=5)
        ttk.Checkbutton(condition_frame, text="#번역요청 컬럼 값이 'change'인 행", variable=self.extract_change_var).pack(anchor="w", padx=5)
        ttk.Separator(condition_frame, orient="horizontal").pack(fill="x", pady=5)
        ttk.Checkbutton(condition_frame, text="추출 후 #번역요청 컬럼을 '전달'로 변경", variable=self.mark_as_transferred_var).pack(anchor="w", padx=5)

        # --- 출력 설정 ---
        output_frame = ttk.LabelFrame(self, text="3. 출력 설정")
        output_frame.pack(fill="x", padx=5, pady=5)

        # [수정] DB 저장 옵션 체크박스 추가
        db_option_frame = ttk.Frame(output_frame)
        db_option_frame.pack(fill="x", padx=5, pady=5)
        ttk.Checkbutton(db_option_frame, text="추출 결과를 DB 파일로 저장", variable=self.save_to_db_var, command=self._toggle_db_path_entry).pack(anchor="w")

        self.db_path_frame = ttk.Frame(output_frame)
        self.db_path_frame.pack(fill="x", padx=5, pady=5)
        ttk.Label(self.db_path_frame, text="DB 파일:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.db_entry = ttk.Entry(self.db_path_frame, textvariable=self.parent_app.output_db_var, width=50)
        self.db_entry.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        self.db_button = ttk.Button(self.db_path_frame, text="저장 위치", command=self.parent_app.select_output_db)
        self.db_button.grid(row=0, column=2, padx=5, pady=5)
        self.db_path_frame.columnconfigure(1, weight=1)

        # --- 실행 버튼 ---
        action_frame = ttk.Frame(self)
        action_frame.pack(fill="x", padx=5, pady=5)
        ttk.Button(action_frame, text="추출 결과 Excel로 내보내기", command=self.export_results).pack(side="right", padx=(0, 5))
        ttk.Button(action_frame, text="기본 추출 실행", command=self.start_extraction).pack(side="right", padx=5)

        # --- 결과 표시 테이블 ---
        result_frame = ttk.LabelFrame(self, text="추출 결과 (DB 저장 비활성 시)")
        result_frame.pack(fill="both", expand=True, padx=5, pady=5)
        
        columns = ("file_name", "sheet_name", "string_id", "kr", "request_type")
        self.result_tree = ttk.Treeview(result_frame, columns=columns, show="headings", height=8)
        self.result_tree.heading("file_name", text="파일명")
        self.result_tree.heading("sheet_name", text="시트명")
        self.result_tree.heading("string_id", text="STRING_ID")
        self.result_tree.heading("kr", text="KR")
        self.result_tree.heading("request_type", text="요청 타입")
        
        for col in columns: self.result_tree.column(col, width=150)
            
        vsb = ttk.Scrollbar(result_frame, orient="vertical", command=self.result_tree.yview)
        hsb = ttk.Scrollbar(result_frame, orient="horizontal", command=self.result_tree.xview)
        self.result_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        vsb.pack(side="right", fill="y")
        hsb.pack(side="bottom", fill="x")
        self.result_tree.pack(fill="both", expand=True)

    def _toggle_db_path_entry(self):
        """DB 경로 입력 위젯 활성화/비활성화"""
        state = "normal" if self.save_to_db_var.get() else "disabled"
        self.db_entry.config(state=state)
        self.db_button.config(state=state)

    def start_extraction(self):
        """기본 추출 작업을 시작"""
        if self.parent_app._is_task_running(): return
        selected_files = self.parent_app._get_selected_files()
        if not selected_files: return

        save_to_db = self.save_to_db_var.get()
        db_path = self.parent_app.output_db_var.get()

        if save_to_db and not db_path:
            show_message(self, "warning", "경고", "저장할 DB 파일 경로를 지정하세요.")
            return
        
        if save_to_db and os.path.exists(db_path):
            if not show_message(self, "yesno", "확인", f"'{os.path.basename(db_path)}' 파일이 이미 존재합니다. 덮어쓰시겠습니까?"):
                return

        conditions = [c.lower() for c, v in {"신규": self.extract_new_var, "change": self.extract_change_var}.items() if v.get()]
        if not conditions:
            show_message(self, "warning", "경고", "추출할 조건을 하나 이상 선택하세요.")
            return

        self.parent_app.log_text.delete(1.0, tk.END)
        self.parent_app.log_message("기본 추출 작업을 시작합니다...")
        self.result_tree.delete(*self.result_tree.get_children()) # 결과 테이블 초기화

        self.parent_app.extraction_thread = threading.Thread(
            target=self.parent_app.extraction_manager.run_basic_extraction,
            args=(selected_files, db_path, conditions, self.mark_as_transferred_var.get(), save_to_db, self._display_results_in_tree)
        )
        self.parent_app.extraction_thread.daemon = True
        self.parent_app.extraction_thread.start()

    # tools/basic_request_extractor.py
    def _display_results_in_tree(self, results):
        """추출 결과를 UI의 Treeview에 표시하고, 전체 결과를 self.full_results에 저장"""
        # ▼▼▼ [추가] 전체 결과를 클래스 변수에 저장 ▼▼▼
        self.full_results = results

        self.result_tree.delete(*self.result_tree.get_children())
        for row in self.full_results: # results -> self.full_results
            # (file_name, sheet_name, string_id, kr, cn, tw, request_type, additional_info)
            # 필요한 데이터만 선택하여 표시
            display_row = (row[0], row[1], row[2], row[3], row[6])
            self.result_tree.insert("", "end", values=display_row)

    # tools/basic_request_extractor.py
    def export_results(self):
        """추출 결과를 엑셀로 내보내기"""
        db_path = self.parent_app.output_db_var.get()
        if self.save_to_db_var.get():
            if not db_path or not os.path.exists(db_path):
                show_message(self, "warning", "경고", "먼저 추출 작업을 실행하여 DB 파일을 생성해야 합니다.")
                return
        else:
            # ▼▼▼ [수정] self.result_tree.get_children() 대신 self.full_results 사용 ▼▼▼
            if not self.full_results:
                show_message(self, "warning", "경고", "먼저 추출 작업을 실행하여 결과를 확인해야 합니다.")
                return

        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel 파일", "*.xlsx")], title="번역 요청 내보내기")
        if not save_path: return

        if self.save_to_db_var.get():
            self.parent_app.extraction_manager.export_to_excel(db_path, save_path)
        else:
            # Treeview가 아닌, 전체 데이터가 담긴 self.full_results를 사용
            if self.full_results:
                import pandas as pd
                # ▼▼▼ [수정] 전체 컬럼 이름 목록으로 DataFrame 생성 ▼▼▼
                columns = ["file_name", "sheet_name", "string_id", "kr", "cn", "tw", "request_type", "additional_info"]
                df = pd.DataFrame(self.full_results, columns=columns)
                df.to_excel(save_path, index=False)
                show_message(self, "info", "완료", f"데이터를 엑셀로 내보냈습니다.\n파일: {save_path}")
            else:
                show_message(self, "info", "알림", "내보낼 데이터가 없습니다.")