import sqlite3
import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import os

# 데이터베이스와 테이블 데이터를 엑셀로 내보내는 핵심 로직
def export_to_excel(db_path, table_name, excel_path, status_label, progress_bar, export_button):
    """
    지정된 DB의 테이블을 엑셀 파일로 내보냅니다.
    이 함수는 별도의 스레드에서 실행됩니다.
    """
    try:
        # 상태 업데이트 및 버튼 비활성화
        status_label.config(text="변환 중...", foreground="blue")
        export_button.config(state="disabled")
        progress_bar.start(10)

        # 데이터베이스 연결 및 데이터 읽기
        conn = sqlite3.connect(db_path)
        query = f'SELECT * FROM "{table_name}"'
        df = pd.read_sql_query(query, conn)
        conn.close()

        # 엑셀 파일로 저장
        df.to_excel(excel_path, index=False)
        
        # 완료 메시지 표시
        status_label.config(text=f"엑셀 파일 저장 완료!", foreground="green")
        messagebox.showinfo("완료", f"테이블 '{table_name}'이(가)\n'{excel_path}'\n파일로 성공적으로 저장되었습니다.")

    except sqlite3.Error as e:
        error_message = f"데이터베이스 오류: {e}\n테이블 이름이 정확한지 확인하세요."
        status_label.config(text="오류 발생!", foreground="red")
        messagebox.showerror("오류", error_message)
    except FileNotFoundError:
        error_message = "오류: 지정된 데이터베이스 파일을 찾을 수 없습니다."
        status_label.config(text="오류 발생!", foreground="red")
        messagebox.showerror("오류", error_message)
    except Exception as e:
        error_message = f"알 수 없는 오류가 발생했습니다: {e}"
        status_label.config(text="오류 발생!", foreground="red")
        messagebox.showerror("오류", error_message)
    finally:
        # 작업 완료 후 버튼 재활성화 및 프로그레스 바 중지
        progress_bar.stop()
        progress_bar['value'] = 0
        export_button.config(state="normal")


class DbToExcelConverter(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("DB to Excel 변환기")
        self.geometry("600x280")
        
        # 메인 프레임
        main_frame = ttk.Frame(self, padding="10 10 10 10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # --- 위젯 생성 ---
        # 1. DB 파일 선택
        ttk.Label(main_frame, text="SQLite DB 파일:").grid(row=0, column=0, sticky="w", pady=2)
        self.db_path_entry = ttk.Entry(main_frame, width=60)
        self.db_path_entry.grid(row=1, column=0, columnspan=2, sticky="ew", pady=2)
        self.db_select_button = ttk.Button(main_frame, text="파일 선택...", command=self.select_db_file)
        self.db_select_button.grid(row=1, column=2, sticky="w", padx=(5, 0))

        # 2. 테이블 선택 (Combobox로 변경)
        ttk.Label(main_frame, text="테이블 선택:").grid(row=2, column=0, sticky="w", pady=(10, 2))
        self.table_name_combobox = ttk.Combobox(main_frame, state="disabled", width=58)
        self.table_name_combobox.grid(row=3, column=0, columnspan=2, sticky="ew", pady=2)

        # 3. 엑셀 파일 저장 경로
        ttk.Label(main_frame, text="저장할 엑셀 파일:").grid(row=4, column=0, sticky="w", pady=(10, 2))
        self.excel_path_entry = ttk.Entry(main_frame, width=60)
        self.excel_path_entry.grid(row=5, column=0, columnspan=2, sticky="ew", pady=2)
        self.excel_save_button = ttk.Button(main_frame, text="저장 위치...", command=self.select_save_path)
        self.excel_save_button.grid(row=5, column=2, sticky="w", padx=(5, 0))

        # 4. 변환 시작 버튼
        self.export_button = ttk.Button(main_frame, text="엑셀로 내보내기", command=self.start_export_thread)
        self.export_button.grid(row=6, column=0, columnspan=3, pady=(20, 5), sticky="ew")
        
        # 5. 프로그레스 바와 상태 라벨
        self.progress_bar = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress_bar.grid(row=7, column=0, columnspan=3, sticky="ew", pady=5)
        self.status_label = ttk.Label(main_frame, text="DB 파일을 선택하세요.", anchor="center")
        self.status_label.grid(row=8, column=0, columnspan=3, sticky="ew")

        # 그리드 컬럼 가중치 설정
        main_frame.columnconfigure(0, weight=1)

    def select_db_file(self):
        # 이전에 선택된 테이블 목록 초기화
        self.table_name_combobox.set('')
        self.table_name_combobox.config(values=[], state="disabled")
        
        file_path = filedialog.askopenfilename(
            title="SQLite 데이터베이스 파일 선택",
            filetypes=(("DB files", "*.db"), ("SQLite3 files", "*.sqlite3"), ("All files", "*.*"))
        )
        if file_path:
            self.db_path_entry.delete(0, tk.END)
            self.db_path_entry.insert(0, file_path)
            
            dir_name = os.path.dirname(file_path)
            base_name = "output.xlsx"
            default_excel_path = os.path.join(dir_name, base_name)
            self.excel_path_entry.delete(0, tk.END)
            self.excel_path_entry.insert(0, default_excel_path)
            
            # DB 파일 선택 후 테이블 목록 업데이트
            self.update_table_list(file_path)

    def update_table_list(self, db_path):
        """선택된 DB 파일에서 테이블 목록을 읽어와 Combobox를 업데이트합니다."""
        try:
            conn = sqlite3.connect(db_path)
            cursor = conn.cursor()
            # 모든 테이블 이름을 조회하는 쿼리
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
            tables = [row[0] for row in cursor.fetchall()]
            conn.close()

            if tables:
                self.table_name_combobox.config(values=tables, state="readonly")
                self.table_name_combobox.current(0) # 첫 번째 테이블을 기본값으로 선택
                self.status_label.config(text=f"{len(tables)}개의 테이블을 찾았습니다. 테이블을 선택하세요.", foreground="black")
            else:
                self.status_label.config(text="선택한 DB 파일에 테이블이 없습니다.", foreground="orange")
                messagebox.showwarning("정보", "선택하신 데이터베이스 파일에 테이블이 존재하지 않습니다.")

        except sqlite3.Error as e:
            self.table_name_combobox.set('')
            self.table_name_combobox.config(values=[], state="disabled")
            self.status_label.config(text="DB 파일을 읽는 중 오류가 발생했습니다.", foreground="red")
            messagebox.showerror("DB 오류", f"데이터베이스 파일을 읽을 수 없습니다.\n파일이 손상되었거나 SQLite 파일이 아닐 수 있습니다.\n\n오류: {e}")

    def select_save_path(self):
        file_path = filedialog.asksaveasfilename(
            title="엑셀 파일로 저장",
            defaultextension=".xlsx",
            filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*"))
        )
        if file_path:
            self.excel_path_entry.delete(0, tk.END)
            self.excel_path_entry.insert(0, file_path)
            
    def start_export_thread(self):
        db_path = self.db_path_entry.get()
        table_name = self.table_name_combobox.get() # Entry 대신 Combobox에서 값 가져오기
        excel_path = self.excel_path_entry.get()

        if not all([db_path, table_name, excel_path]):
            messagebox.showwarning("입력 오류", "모든 필드를 선택하거나 입력해주세요.")
            return

        # 스레드를 생성하여 export_to_excel 함수 실행
        export_thread = threading.Thread(
            target=export_to_excel,
            args=(db_path, table_name, excel_path, self.status_label, self.progress_bar, self.export_button)
        )
        export_thread.daemon = True
        export_thread.start()


if __name__ == "__main__":
    app = DbToExcelConverter()
    app.mainloop()