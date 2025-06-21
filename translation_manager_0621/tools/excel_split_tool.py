# tools/excel_split_tool.py

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import threading
import pythoncom
import win32com.client

class ExcelSplitTool(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        
        # 변수 초기화
        self.excel_input_file_path = tk.StringVar()
        self.excel_output_dir_path = tk.StringVar()
        self.excel_output_dir_path.set(os.getcwd())  # 기본값: 현재 디렉토리
        
        # UI 구성
        self.setup_ui()

    def setup_ui(self):
        """엑셀 시트 분리 탭 UI 구성"""
        # 메인 프레임
        main_frame = ttk.Frame(self, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 입력 파일 선택
        ttk.Label(main_frame, text="원본 엑셀 파일:").grid(row=0, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.excel_input_file_path, width=50).grid(row=0, column=1, sticky=tk.W+tk.E, pady=5, padx=5)
        ttk.Button(main_frame, text="찾아보기", command=self.browse_excel_input_file).grid(row=0, column=2, sticky=tk.W, pady=5)
        
        # 출력 디렉토리 선택
        ttk.Label(main_frame, text="출력 폴더:").grid(row=1, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.excel_output_dir_path, width=50).grid(row=1, column=1, sticky=tk.W+tk.E, pady=5, padx=5)
        ttk.Button(main_frame, text="찾아보기", command=self.browse_excel_output_dir).grid(row=1, column=2, sticky=tk.W, pady=5)
        
        # 진행 상황 표시
        self.excel_progress = ttk.Progressbar(main_frame, orient=tk.HORIZONTAL, length=100, mode='determinate')
        self.excel_progress.grid(row=2, column=0, columnspan=3, sticky=tk.W+tk.E, pady=10)
        
        # 로그 표시 영역
        log_frame = ttk.LabelFrame(main_frame, text="처리 로그")
        log_frame.grid(row=3, column=0, columnspan=3, sticky=tk.W+tk.E+tk.N+tk.S, pady=5)
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        
        self.excel_log_text = tk.Text(log_frame, wrap=tk.WORD, width=70, height=10)
        self.excel_log_text.grid(row=0, column=0, sticky=tk.W+tk.E+tk.N+tk.S, padx=5, pady=5)
        
        scrollbar = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.excel_log_text.yview)
        scrollbar.grid(row=0, column=1, sticky=tk.N+tk.S)
        self.excel_log_text['yscrollcommand'] = scrollbar.set
        
        # 버튼 영역
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=4, column=0, columnspan=3, sticky=tk.W+tk.E, pady=10)
        button_frame.columnconfigure(0, weight=1)
        
        ttk.Button(button_frame, text="시트 분리 실행", command=self.start_excel_split_processing).pack(side="right", padx=5)
        
        # 리사이징 설정
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(3, weight=1)

    def browse_excel_input_file(self):
        file_path = filedialog.askopenfilename(
            title="원본 엑셀 파일 선택",
            filetypes=[("Excel 파일", "*.xlsx;*.xls"), ("모든 파일", "*.*")],
            parent=self
        )
        if file_path:
            self.excel_input_file_path.set(file_path)
            self.excel_log("원본 파일이 선택되었습니다: " + file_path)
            self.after(100, self.focus_force)
            self.after(100, self.lift)
    
    def browse_excel_output_dir(self):
        dir_path = filedialog.askdirectory(title="출력 폴더 선택", parent=self)
        if dir_path:
            self.excel_output_dir_path.set(dir_path)
            self.excel_log("출력 폴더가 선택되었습니다: " + dir_path)
            self.after(100, self.focus_force)
            self.after(100, self.lift)
    
    def excel_log(self, message):
        self.excel_log_text.insert(tk.END, message + "\n")
        self.excel_log_text.see(tk.END)
    
    def start_excel_split_processing(self):
        input_file = self.excel_input_file_path.get().strip()
        output_dir = self.excel_output_dir_path.get().strip()
        
        if not input_file:
            messagebox.showerror("오류", "원본 엑셀 파일을 선택해주세요.", parent=self)
            return
        
        if not output_dir:
            messagebox.showerror("오류", "출력 폴더를 선택해주세요.", parent=self)
            return
        
        self.excel_progress['value'] = 0
        threading.Thread(target=self.split_excel_by_sheets, args=(input_file, output_dir), daemon=True).start()
    
    def split_excel_by_sheets(self, input_file, output_dir):
        try:
            self.excel_log("처리 준비 중...")
            pythoncom.CoInitialize()
            
            excel = win32com.client.Dispatch("Excel.Application")
            
            try:
                # 엑셀을 완전히 숨기고 모든 화면 업데이트 중지
                excel.Visible = False  # 엑셀 화면 숨김
                excel.DisplayAlerts = False  # 경고 메시지 숨김
                excel.ScreenUpdating = False  # 화면 업데이트 중지
                excel.EnableEvents = False  # 이벤트 비활성화
                excel.Interactive = False  # 사용자 상호작용 비활성화
                
                # Calculation 속성은 설정할 수 없는 경우가 있으므로 안전하게 처리
                try:
                    excel.Calculation = -4135  # xlCalculationManual
                except:
                    self.excel_log("계산 모드 설정을 건너뜁니다.")
                    
            except Exception as setup_error:
                self.excel_log(f"엑셀 설정 중 일부 오류 (계속 진행): {str(setup_error)}")
            
            try:
                self.excel_log("엑셀 파일 열기 중...")
                workbook = excel.Workbooks.Open(os.path.abspath(input_file))
                
                total_sheets = workbook.Sheets.Count
                self.excel_progress['maximum'] = total_sheets
                self.excel_log(f"총 {total_sheets}개의 시트를 발견했습니다.")
                
                for idx in range(1, total_sheets + 1):
                    sheet = workbook.Sheets(idx)
                    sheet_name = sheet.Name
                    
                    try:
                        self.excel_log(f"시트 처리 중 ({idx}/{total_sheets}): {sheet_name}")
                        
                        safe_sheet_name = sheet_name.replace('/', '_').replace('\\', '_').replace('*', '_') \
                                         .replace('?', '_').replace(':', '_').replace('[', '_').replace(']', '_')
                        
                        new_file_path = os.path.join(output_dir, f"{safe_sheet_name}.xlsx")
                        
                        # 새 워크북 생성 (화면에 표시되지 않도록)
                        new_workbook = excel.Workbooks.Add()
                        
                        # 시트 복사
                        sheet.Copy(Before=new_workbook.Sheets(1))
                        
                        # 기본 시트 삭제 (새로 추가된 시트만 남김)
                        if new_workbook.Sheets.Count > 1:
                            for i in range(new_workbook.Sheets.Count, 1, -1):
                                if new_workbook.Sheets(i).Name != sheet_name:
                                    new_workbook.Sheets(i).Delete()
                        
                        # 파일 저장
                        new_workbook.SaveAs(os.path.abspath(new_file_path))
                        new_workbook.Close(SaveChanges=False)
                        
                        self.excel_log(f"생성된 파일: {safe_sheet_name}.xlsx")
                        
                        self.after(0, lambda i=idx: self.excel_progress.configure(value=i))
                        
                    except Exception as sheet_error:
                        self.excel_log(f"시트 '{sheet_name}' 처리 중 오류 발생: {str(sheet_error)}")
                
                self.excel_log("모든 시트가 처리되었습니다.")
                
            finally:
                # 엑셀 설정 복원 (안전하게 처리)
                try:
                    excel.ScreenUpdating = True
                    excel.EnableEvents = True
                    excel.Interactive = True
                    try:
                        excel.Calculation = -4105  # xlCalculationAutomatic
                    except:
                        pass  # 계산 모드 복원 실패 시 무시
                except:
                    pass  # 설정 복원 실패 시 무시
                
                workbook.Close(SaveChanges=False)
                excel.Quit()
                
            self.after(0, lambda: messagebox.showinfo("완료", "모든 시트가 처리되었습니다.", parent=self))
            
        except Exception as e:
            error_msg = str(e)
            self.excel_log(f"오류 발생: {error_msg}")
            self.after(0, lambda error=error_msg: messagebox.showerror("오류", f"처리 중 오류가 발생했습니다: {error}", parent=self))
        finally:
            pythoncom.CoUninitialize()