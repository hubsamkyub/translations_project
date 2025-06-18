import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import win32com.client
import pythoncom

class ExcelSplitterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("엑셀 시트 분리기")
        self.root.geometry("600x400")
        self.root.resizable(True, True)
        
        # 변수 초기화
        self.input_file_path = tk.StringVar()
        self.output_dir_path = tk.StringVar()
        self.output_dir_path.set(os.getcwd())  # 기본값: 현재 디렉토리
        
        # UI 구성
        self.create_widgets()
        
    def create_widgets(self):
        # 프레임 생성
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 입력 파일 선택
        ttk.Label(main_frame, text="원본 엑셀 파일:").grid(row=0, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.input_file_path, width=50).grid(row=0, column=1, sticky=tk.W+tk.E, pady=5, padx=5)
        ttk.Button(main_frame, text="찾아보기", command=self.browse_input_file).grid(row=0, column=2, sticky=tk.W, pady=5)
        
        # 출력 디렉토리 선택
        ttk.Label(main_frame, text="출력 폴더:").grid(row=1, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.output_dir_path, width=50).grid(row=1, column=1, sticky=tk.W+tk.E, pady=5, padx=5)
        ttk.Button(main_frame, text="찾아보기", command=self.browse_output_dir).grid(row=1, column=2, sticky=tk.W, pady=5)
        
        # 진행 상황 표시
        self.progress = ttk.Progressbar(main_frame, orient=tk.HORIZONTAL, length=100, mode='determinate')
        self.progress.grid(row=2, column=0, columnspan=3, sticky=tk.W+tk.E, pady=10)
        
        # 로그 표시 영역
        log_frame = ttk.LabelFrame(main_frame, text="처리 로그")
        log_frame.grid(row=3, column=0, columnspan=3, sticky=tk.W+tk.E+tk.N+tk.S, pady=5)
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        
        self.log_text = tk.Text(log_frame, wrap=tk.WORD, width=70, height=10)
        self.log_text.grid(row=0, column=0, sticky=tk.W+tk.E+tk.N+tk.S, padx=5, pady=5)
        
        scrollbar = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        scrollbar.grid(row=0, column=1, sticky=tk.N+tk.S)
        self.log_text['yscrollcommand'] = scrollbar.set
        
        # 버튼 영역
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=4, column=0, columnspan=3, sticky=tk.W+tk.E, pady=10)
        button_frame.columnconfigure(0, weight=1)
        button_frame.columnconfigure(1, weight=1)
        
        ttk.Button(button_frame, text="시트 분리 실행", command=self.start_processing).grid(row=0, column=0, padx=5)
        ttk.Button(button_frame, text="종료", command=self.root.destroy).grid(row=0, column=1, padx=5)
        
        # 리사이징 설정
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(3, weight=1)
        
    def browse_input_file(self):
        file_path = filedialog.askopenfilename(
            title="원본 엑셀 파일 선택",
            filetypes=[("Excel 파일", "*.xlsx;*.xls"), ("모든 파일", "*.*")]
        )
        if file_path:
            self.input_file_path.set(file_path)
            self.log("원본 파일이 선택되었습니다: " + file_path)
    
    def browse_output_dir(self):
        dir_path = filedialog.askdirectory(title="출력 폴더 선택")
        if dir_path:
            self.output_dir_path.set(dir_path)
            self.log("출력 폴더가 선택되었습니다: " + dir_path)
    
    def log(self, message):
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
    
    def start_processing(self):
        input_file = self.input_file_path.get().strip()
        output_dir = self.output_dir_path.get().strip()
        
        if not input_file:
            messagebox.showerror("오류", "원본 엑셀 파일을 선택해주세요.")
            return
        
        if not output_dir:
            messagebox.showerror("오류", "출력 폴더를 선택해주세요.")
            return
        
        # 백그라운드 스레드에서 처리
        self.progress['value'] = 0
        threading.Thread(target=self.split_excel_by_sheets, args=(input_file, output_dir), daemon=True).start()
    
    def split_excel_by_sheets(self, input_file, output_dir):
        try:
            self.log("처리 준비 중...")
            # COM 스레드 초기화
            pythoncom.CoInitialize()
            
            # Excel 애플리케이션 시작
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False  # 백그라운드에서 실행
            excel.DisplayAlerts = False  # 알림 끄기
            
            try:
                self.log("엑셀 파일 열기 중...")
                # 원본 엑셀 파일 열기
                workbook = excel.Workbooks.Open(os.path.abspath(input_file))
                
                # 시트 개수 확인
                total_sheets = workbook.Sheets.Count
                self.progress['maximum'] = total_sheets
                self.log(f"총 {total_sheets}개의 시트를 발견했습니다.")
                
                # 각 시트에 대해 반복
                for idx in range(1, total_sheets + 1):
                    sheet = workbook.Sheets(idx)
                    sheet_name = sheet.Name
                    
                    try:
                        self.log(f"시트 처리 중 ({idx}/{total_sheets}): {sheet_name}")
                        
                        # 파일명에 사용할 수 없는 문자 처리
                        safe_sheet_name = sheet_name.replace('/', '_').replace('\\', '_').replace('*', '_') \
                                         .replace('?', '_').replace(':', '_').replace('[', '_').replace(']', '_')
                        
                        # 새 파일 경로
                        new_file_path = os.path.join(output_dir, f"{safe_sheet_name}.xlsx")
                        
                        # 시트 복사 (Excel의 "이동/복사" 기능 사용)
                        sheet.Copy(Before=None)  # 새 통합 문서로 복사
                        
                        # 새로 생성된 통합 문서는 활성 통합 문서가 됨
                        new_workbook = excel.ActiveWorkbook
                        
                        # 저장 및 닫기
                        new_workbook.SaveAs(os.path.abspath(new_file_path))
                        new_workbook.Close(SaveChanges=False)
                        
                        self.log(f"생성된 파일: {safe_sheet_name}.xlsx")
                        
                        # 진행 상황 업데이트
                        self.progress['value'] = idx
                        
                    except Exception as sheet_error:
                        self.log(f"시트 '{sheet_name}' 처리 중 오류 발생: {str(sheet_error)}")
                
                self.log("모든 시트가 처리되었습니다.")
                
            finally:
                # 원본 통합 문서 닫기
                workbook.Close(SaveChanges=False)
                # Excel 종료
                excel.Quit()
                
            self.root.after(0, lambda: messagebox.showinfo("완료", "모든 시트가 처리되었습니다."))
            
        except Exception as e:
            error_msg = str(e)
            self.log(f"오류 발생: {error_msg}")
            self.root.after(0, lambda error=error_msg: messagebox.showerror("오류", f"처리 중 오류가 발생했습니다: {error}"))
        finally:
            # COM 스레드 해제
            pythoncom.CoUninitialize()

# 메인 애플리케이션 실행
if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelSplitterApp(root)
    root.mainloop()