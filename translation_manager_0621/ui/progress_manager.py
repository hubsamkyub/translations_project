import tkinter as tk
from tkinter import ttk
import time

class ProgressManager:
    """진행 상태 창과 관련 기능을 관리하는 클래스"""
    
    def __init__(self, parent):
        """
        ProgressManager 초기화
        
        Args:
            parent: 부모 윈도우 (tkinter 창)
        """
        self.parent = parent
        self.progress_window = None
        self.progress_var = None
        self.progress_log = None
        self.total_items = 1  # 최소 1로 설정하여 나눗셈 오류 방지
        self.processed_items = 0
    
    def start_progress_window(self, total_count=0, title="처리 중"):
        """
        진행 상황 창 시작
        
        Args:
            total_count: 전체 작업 아이템 수
            title: 창 제목 (기본값: "처리 중")
        """
        # 이미 창이 있으면 닫기
        if self.progress_window:
            self.progress_window.destroy()
            
        self.progress_window = tk.Toplevel(self.parent)
        self.progress_window.title(f"🔄 {title}")
        self.progress_window.geometry("500x400")
        self.progress_window.transient(self.parent)
        
        # 창 닫기 버튼 비활성화 (작업 완료 전에 창을 닫지 못하게)
        self.progress_window.protocol("WM_DELETE_WINDOW", lambda: None)
        
        self.progress_var = tk.DoubleVar(value=0)

        # 진행 창 타이틀
        tk.Label(self.progress_window, text=title, font=("Arial", 12, "bold")).pack(pady=(10, 5))

        # 진행바
        progressbar = ttk.Progressbar(self.progress_window, variable=self.progress_var, maximum=100)
        progressbar.pack(fill="x", padx=10, pady=10)

        # 로그 창 프레임
        log_frame = tk.Frame(self.progress_window)
        log_frame.pack(fill="both", expand=True, padx=10, pady=5)
        
        # 로그 창
        self.progress_log = tk.Text(log_frame, height=15, state="normal")
        log_scrollbar = tk.Scrollbar(log_frame, orient="vertical", command=self.progress_log.yview)
        self.progress_log.configure(yscrollcommand=log_scrollbar.set)
        
        self.progress_log.pack(side="left", fill="both", expand=True)
        log_scrollbar.pack(side="right", fill="y")

        # 초기 메시지
        self.update_progress("작업 준비 중...", success=True)

        # 최소 1로 설정하여 나눗셈 오류 방지
        self.total_items = max(1, total_count) 
        self.processed_items = 0

        # 중앙 정렬
        self.progress_window.update_idletasks()
        width = self.progress_window.winfo_width()
        height = self.progress_window.winfo_height()
        x = (self.progress_window.winfo_screenwidth() // 2) - (width // 2)
        y = (self.progress_window.winfo_screenheight() // 2) - (height // 2)
        self.progress_window.geometry('{}x{}+{}+{}'.format(width, height, x, y))

        # 초기에는 창을 포커스 시킴
        self.progress_window.focus_set()
        
        # 강제 업데이트
        self.progress_window.update()
        
        return self.progress_window
    
    def update_progress(self, message, success=True):
        """
        진행 상황 업데이트
        
        Args:
            message: 표시할 메시지
            success: 성공(True) 또는 실패(False) 여부
        """
        if self.progress_window and self.progress_log:
            if "시작" in message or "완료" in message:
                self.processed_items += 1
                # 0으로 나누는 것을 방지
                if self.total_items > 0:
                    percentage = (self.processed_items / self.total_items) * 100
                    self.progress_var.set(percentage)

            self.progress_log.config(state="normal")
            icon = "✔️" if success else "❌"
            timestamp = time.strftime("%H:%M:%S", time.localtime())
            self.progress_log.insert("end", f"[{timestamp}] {icon} {message}\n")
            self.progress_log.see("end")
            self.progress_log.config(state="disabled")

            # 화면 강제 갱신
            try:
                self.progress_window.update()
            except:
                pass  # 창이 닫힌 경우 예외 발생 방지
    
    def finish_progress(self, final_message="모든 작업 완료!", external_link_files=None):
        """
        진행 상황 완료 처리
        
        Args:
            final_message: 최종 메시지
            external_link_files: 외부 링크가 있는 파일 목록 (선택 사항)
        """
        if not self.progress_window or not self.progress_log:
            return
            
        # 창 닫기 버튼 다시 활성화
        self.progress_window.protocol("WM_DELETE_WINDOW", self.progress_window.destroy)
        
        # topmost 속성 해제 (필요한 경우)
        self.progress_window.attributes('-topmost', False)
        
        self.progress_log.config(state="normal")
        
        # 외부 링크 파일이 있는지 확인
        if external_link_files and len(external_link_files) > 0:
            external_link_count = len(external_link_files)
            self.progress_log.insert("end", f"\n⚠️ 주의: {external_link_count}개 파일은 외부 링크로 인해 처리하지 못했습니다.\n")
            
            for fp in external_link_files:
                self.progress_log.insert("end", f"   - {os.path.basename(fp) if isinstance(fp, str) else fp}\n")
        
        self.progress_log.insert("end", f"\n✅ {final_message}\n")
        self.progress_log.see("end")
        self.progress_log.config(state="disabled")

        # 완료 메시지를 더욱 눈에 띄게 표시
        completion_frame = tk.Frame(self.progress_window, bg="#e6ffe6")
        completion_frame.pack(fill="x", padx=10, pady=5)
        
        # 외부 링크 파일이 있는 경우 메시지 조정
        completion_msg = final_message
        if external_link_files and len(external_link_files) > 0:
            external_link_count = len(external_link_files)
            completion_msg += f" ({external_link_count}개 파일은 외부 링크로 인해 처리되지 않음)"
                
        tk.Label(completion_frame, 
                text=completion_msg,
                font=("Arial", 10, "bold"),
                bg="#e6ffe6").pack(pady=5)

        close_button = tk.Button(self.progress_window, 
                                text="닫기", 
                                command=self.progress_window.destroy,
                                font=("Arial", 10, "bold"))
        close_button.pack(pady=10)
        
        # 닫기 버튼에 포커스
        close_button.focus_set()

        self.progress_window.update_idletasks()
    
    def is_active(self):
        """진행 창이 활성 상태인지 확인"""
        return self.progress_window is not None and self.progress_window.winfo_exists()