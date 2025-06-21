import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import json
import time
import gc

# 설정 유틸리티 함수는 config_utils.py에서 가져와서 사용
from utils.config_utils import load_config, save_config

class ScrollableCheckList(tk.Frame):
    def __init__(self, parent, width=300, height=150, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)
        self.config = load_config()
        
        self.original_folder_var = tk.StringVar()
        default_path = self.config.get("data_path", "")
        self.original_folder_var.set(default_path)
        
        self.canvas = tk.Canvas(self, width=width, height=height, borderwidth=0)
        self.scrollbar = tk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.inner_frame = tk.Frame(self.canvas)
        self.canvas_window = self.canvas.create_window((0, 0), window=self.inner_frame, anchor="nw")

        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")

        self.inner_frame.bind("<Configure>", self._on_frame_configure)
        self.canvas.bind("<Configure>", self._on_canvas_configure)
        
        # 전체 선택/해제 변수 및 컨트롤
        self.check_all_var = tk.BooleanVar(value=True)
        self.check_all_frame = tk.Frame(self)
        self.check_all_frame.pack(fill="x", before=self.canvas)
        
        self.check_all_cb = tk.Checkbutton(
            self.check_all_frame, 
            text="전체 선택/해제", 
            variable=self.check_all_var,
            command=self._toggle_all
        )
        self.check_all_cb.pack(side="left", padx=5, pady=2)
        
        self.vars_dict = {}
        
        # 마우스 휠 이벤트 바인딩
        self._bind_mouse_wheel(self.canvas)
        self._bind_mouse_wheel(self.inner_frame)
        
        # 위젯에 포커스가 있을 때만 마우스 휠 작동하도록 포커스 이벤트 추가
        self.canvas.bind("<Enter>", lambda e: self.canvas.focus_set())
        self.canvas.bind("<Leave>", lambda e: self.root.focus_set() if hasattr(self, 'root') else None)

    def _bind_mouse_wheel(self, widget):
        """여러 플랫폼에 대한 마우스 휠 이벤트 바인딩"""
        widget.bind("<MouseWheel>", self._on_mouse_wheel)       # Windows
        widget.bind("<Button-4>", self._on_mouse_wheel)         # Linux 위로 스크롤
        widget.bind("<Button-5>", self._on_mouse_wheel)         # Linux 아래로 스크롤
        widget.bind("<Button-2>", self._on_mouse_wheel)         # macOS/Linux 중간 버튼

    def _on_mouse_wheel(self, event):
        """마우스 휠 이벤트 처리 함수"""
        # Windows의 경우
        if event.num == 5 or event.delta < 0:  # 아래로 스크롤
            self.canvas.yview_scroll(1, "units")
        elif event.num == 4 or event.delta > 0:  # 위로 스크롤
            self.canvas.yview_scroll(-1, "units")
        return "break"  # 이벤트 전파 중단

    def _on_frame_configure(self, event):
        # 내부 프레임 크기가 변경되면 스크롤 영역 업데이트
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
    
    def _on_canvas_configure(self, event):
        # 캔버스 크기가 변경되면 내부 프레임의 너비 조정
        width = event.width
        self.canvas.itemconfig(self.canvas_window, width=width)
    
    def _toggle_all(self):
        # 전체 선택/해제 토글 처리
        checked = self.check_all_var.get()
        self.set_all_checked(checked)

    def add_item(self, text, checked=True):
        var = tk.BooleanVar(value=checked)
        cb = tk.Checkbutton(self.inner_frame, text=text, variable=var, anchor="w")
        cb.pack(fill="x", anchor="w")
        
        # 새로 추가된 체크박스에도 마우스 휠 이벤트 바인딩
        self._bind_mouse_wheel(cb)
        
        self.vars_dict[text] = var
        
        # 체크박스 상태 변경 시 전체 선택 상태 업데이트
        # 일시적으로 trace 비활성화 플래그 추가하여 초기 로딩 시 충돌 방지
        if not hasattr(self, '_loading'):
            # 체크박스 상태 변경 시 전체 선택 상태 업데이트
            var.trace_add("write", lambda *args: self._update_check_all_state())

    def clear(self):
        self._loading = True  # 로딩 플래그 설정
        for w in self.inner_frame.winfo_children():
            w.destroy()
        self.vars_dict.clear()
        self.check_all_var.set(True)  # 기본값으로 초기화
        self._loading = False  # 로딩 완료

    def get_checked_items(self):
        return [t for (t, v) in self.vars_dict.items() if v.get()]

    def set_all_checked(self, checked=True):
        for var in self.vars_dict.values():
            var.set(checked)
    
    def _update_check_all_state(self):
        # 로딩 중이거나 목록이 비어있으면 처리하지 않음
        if hasattr(self, '_loading') and self._loading:
            return
        if not self.vars_dict:
            return
        
        # 콜백 무한루프 방지
        if hasattr(self, '_updating_state') and self._updating_state:
            return
        
        self._updating_state = True
        try:
            all_checked = all(var.get() for var in self.vars_dict.values())
            self.check_all_var.set(all_checked)
        finally:
            self._updating_state = False

# 공통 유틸리티 함수
def show_message(parent, message_type, title, message):
    """통합 메시지 표시 함수"""
    if message_type == "info":
        messagebox.showinfo(title, message, parent=parent)
    elif message_type == "warning":
        messagebox.showwarning(title, message, parent=parent)
    elif message_type == "error":
        messagebox.showerror(title, message, parent=parent)
    elif message_type == "yesno":
        return messagebox.askyesno(title, message, parent=parent)
    
def select_folder(parent, title_text, initial_dir=None):
    """폴더 선택 다이얼로그 함수"""
    folder = filedialog.askdirectory(title=title_text, parent=parent, initialdir=initial_dir)
    if folder:
        # 포커스를 다시 부모 창으로
        parent.after(100, parent.focus_force)
        parent.after(100, parent.lift)
        return folder
    return None

def select_file(parent, title_text, filetypes, initialdir=None):
    """파일 선택 다이얼로그 함수"""
    file_path = filedialog.askopenfilename(
        filetypes=filetypes,
        title=title_text,
        parent=parent,
        initialdir=initialdir
    )
    if file_path:
        # 포커스를 다시 부모 창으로
        parent.after(100, parent.focus_force)
        parent.after(100, parent.lift)
        return file_path
    return None

def save_file(parent, title_text, filetypes, defaultextension, initialdir=None):
    """파일 저장 다이얼로그 함수"""
    file_path = filedialog.asksaveasfilename(
        defaultextension=defaultextension,
        filetypes=filetypes,
        title=title_text,
        parent=parent,
        initialdir=initialdir
    )
    if file_path:
        # 포커스를 다시 부모 창으로
        parent.after(100, parent.focus_force)
        parent.after(100, parent.lift)
        return file_path
    return None

class LoadingPopup:
    """진행 상황 표시 팝업"""
    def __init__(self, parent, title="로딩 중...", message="작업 준비 중..."):
        self.popup = tk.Toplevel(parent)
        self.popup.title(title)
        self.popup.geometry("400x150")
        self.popup.transient(parent)
        self.popup.grab_set()
        
        self.message_label = ttk.Label(self.popup, text=message)
        self.message_label.pack(pady=20)
        
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(self.popup, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(fill="x", padx=20, pady=10)
        
        self.status_var = tk.StringVar(value="준비 중...")
        self.status_label = ttk.Label(self.popup, textvariable=self.status_var)
        self.status_label.pack(pady=10)
        
        # 즉시 표시하도록 업데이트
        self.popup.update()
    
    def update_progress(self, percentage, status_text=None):
        """진행률 및 상태 텍스트 업데이트"""
        self.progress_var.set(percentage)
        if status_text:
            self.status_var.set(status_text)
        self.popup.update_idletasks()
    
    def update_message(self, message):
        """메시지 라벨 업데이트"""
        self.message_label.config(text=message)
        self.popup.update_idletasks()
    
    def close(self):
        """팝업 닫기"""
        self.popup.destroy()