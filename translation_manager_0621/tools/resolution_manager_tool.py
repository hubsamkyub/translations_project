# tools/resolution_manager_tool.py (신규 파일)

import tkinter as tk
from tkinter import ttk, messagebox, simpledialog

class ResolutionManagerTool(tk.Frame):
    def __init__(self, parent, manager):
        super().__init__(parent)
        self.manager = manager

        self.setup_ui()
        self.load_data()

    def setup_ui(self):
        # --- 상단 프레임: 제어 버튼 및 검색 ---
        top_frame = ttk.Frame(self)
        top_frame.pack(fill="x", padx=10, pady=10)

        ttk.Button(top_frame, text="새로고침", command=self.load_data).pack(side="left")
        ttk.Label(top_frame, text="   KR 검색: ").pack(side="left")
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda n, i, m: self.filter_tree())
        ttk.Entry(top_frame, textvariable=self.search_var, width=40).pack(side="left")

        # --- Treeview 프레임: 데이터 목록 ---
        tree_frame = ttk.Frame(self)
        tree_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        columns = ("kr", "cn", "tw", "date")
        self.tree = ttk.Treeview(tree_frame, columns=columns, show="headings")
        
        self.tree.heading("kr", text="KR 텍스트")
        self.tree.heading("cn", text="CN 번역")
        self.tree.heading("tw", text="TW 번역")
        self.tree.heading("date", text="수정일")

        self.tree.column("kr", width=400)
        self.tree.column("cn", width=300)
        self.tree.column("tw", width=300)
        self.tree.column("date", width=150, anchor="center")

        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        vsb.pack(side="right", fill="y")
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview)
        hsb.pack(side="bottom", fill="x")
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.pack(fill="both", expand=True)

        # --- 하단 프레임: 편집 및 삭제 버튼 ---
        bottom_frame = ttk.Frame(self)
        bottom_frame.pack(fill="x", padx=10, pady=(0, 10))

        ttk.Button(bottom_frame, text="선택 항목 수정", command=self.edit_selected).pack(side="left")
        ttk.Button(bottom_frame, text="선택 항목 삭제", command=self.delete_selected).pack(side="left", padx=10)
        self.status_label = ttk.Label(bottom_frame, text="")
        self.status_label.pack(side="right")

    def load_data(self):
        """DB에서 데이터를 불러와 Treeview에 표시합니다."""
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        self.all_data = self.manager.get_all_resolutions()
        for row in self.all_data:
            self.tree.insert("", "end", values=(row['kr_text'], row['cn_text'], row['tw_text'], row['resolved_at']))
        self.status_label.config(text=f"총 {len(self.all_data)}개 항목")

    def filter_tree(self):
        """검색어에 따라 Treeview 내용을 필터링합니다."""
        search_term = self.search_var.get().lower()
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        for row in self.all_data:
            if search_term in row['kr_text'].lower():
                self.tree.insert("", "end", values=(row['kr_text'], row['cn_text'], row['tw_text'], row['resolved_at']))

    def edit_selected(self):
        """선택된 항목을 수정하는 팝업을 띄웁니다."""
        selected_item = self.tree.focus()
        if not selected_item:
            messagebox.showwarning("선택 필요", "수정할 항목을 목록에서 선택하세요.", parent=self)
            return
            
        item_values = self.tree.item(selected_item, "values")
        kr_text, old_cn, old_tw = item_values[0], item_values[1], item_values[2]
        
        new_cn = simpledialog.askstring("CN 번역 수정", f"KR: {kr_text[:30]}...", initialvalue=old_cn, parent=self)
        if new_cn is None: return # 사용자가 취소

        new_tw = simpledialog.askstring("TW 번역 수정", f"KR: {kr_text[:30]}...", initialvalue=old_tw, parent=self)
        if new_tw is None: return # 사용자가 취소

        if self.manager.update_resolution(kr_text, new_cn, new_tw):
            messagebox.showinfo("성공", "수정이 완료되었습니다.", parent=self)
            self.load_data() # 목록 새로고침
        else:
            messagebox.showerror("오류", "수정 중 오류가 발생했습니다.", parent=self)

    def delete_selected(self):
        """선택된 항목을 삭제합니다."""
        selected_item = self.tree.focus()
        if not selected_item:
            messagebox.showwarning("선택 필요", "삭제할 항목을 목록에서 선택하세요.", parent=self)
            return
            
        kr_text = self.tree.item(selected_item, "values")[0]
        
        if messagebox.askyesno("삭제 확인", f"정말로 아래 항목을 삭제하시겠습니까?\n\nKR: {kr_text[:50]}...", parent=self):
            if self.manager.delete_resolution(kr_text):
                messagebox.showinfo("성공", "삭제가 완료되었습니다.", parent=self)
                self.load_data()
            else:
                messagebox.showerror("오류", "삭제 중 오류가 발생했습니다.", parent=self)