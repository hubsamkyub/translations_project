# tools/translation_verification_tool.py

import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd

class TranslationVerificationTool(tk.Toplevel):
    def __init__(self, parent, duplicate_data, db_path, manager, completion_callback):
        super().__init__(parent)
        self.parent = parent
        self.duplicate_data = duplicate_data
        self.db_path = db_path
        self.manager = manager
        self.completion_callback = completion_callback

        self.resolved_data = {} # 사용자가 수정한 최종 데이터를 저장

        self.title("번역 확인 및 수정")
        self.geometry("1200x700")
        self.transient(parent)
        self.grab_set()

        self.setup_ui()
        self.load_data_to_tree()

    def setup_ui(self):
        main_frame = ttk.Frame(self, padding=10)
        main_frame.pack(fill="both", expand=True)

        # 상단 설명
        info_label = ttk.Label(main_frame, text="동일한 STRING_ID에 여러 번역이 존재합니다. 아래 목록을 확인하고 올바른 번역으로 수정한 후, [확정 및 DB 반영] 버튼을 눌러주세요.", wraplength=1100)
        info_label.pack(fill="x", pady=5)

        # 트리뷰 프레임
        tree_frame = ttk.LabelFrame(main_frame, text="중복 항목 목록")
        tree_frame.pack(fill="both", expand=True, pady=5)
        
        columns = ("string_id", "kr", "en", "cn", "tw", "th")
        self.tree = ttk.Treeview(tree_frame, columns=columns, show="headings")
        for col in columns:
            self.tree.heading(col, text=col.upper())
            self.tree.column(col, width=180)
        
        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        vsb.pack(side="right", fill="y")
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(fill="both", expand=True)
        self.tree.bind("<<TreeviewSelect>>", self.on_item_select)

        # 편집 프레임
        edit_frame = ttk.LabelFrame(main_frame, text="선택 항목 편집")
        edit_frame.pack(fill="x", pady=5)

        self.edit_vars = {lang: tk.StringVar() for lang in columns}
        
        for i, lang in enumerate(columns):
            ttk.Label(edit_frame, text=f"{lang.upper()}:").grid(row=i, column=0, padx=5, pady=2, sticky="w")
            entry = ttk.Entry(edit_frame, textvariable=self.edit_vars[lang], width=100)
            if lang == "string_id":
                entry.config(state="readonly") # STRING_ID는 편집 불가
            entry.grid(row=i, column=1, padx=5, pady=2, sticky="ew")
        edit_frame.columnconfigure(1, weight=1)

        # 하단 버튼 프레임
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill="x", pady=10)
        ttk.Button(button_frame, text="엑셀로 내보내기", command=self.export_to_excel).pack(side="left")
        ttk.Button(button_frame, text="닫기", command=self.destroy).pack(side="right")
        self.apply_button = ttk.Button(button_frame, text="확정 및 DB 반영", command=self.apply_resolved_data)
        self.apply_button.pack(side="right", padx=5)
        ttk.Button(button_frame, text="선택 항목으로 최종본 저장", command=self.save_resolution).pack(side="right", padx=5)

    def load_data_to_tree(self):
        """중복 데이터를 트리뷰에 로드"""
        parent_id = ""
        for string_id, items in self.duplicate_data.items():
            # 부모 노드 (STRING_ID)
            parent_id = self.tree.insert("", "end", text=string_id, values=(string_id,), open=True, tags=('group',))
            # 자식 노드 (각 번역 내용)
            for item in items:
                values = (
                    string_id,
                    item.get('kr', ''), item.get('en', ''), item.get('cn', ''),
                    item.get('tw', ''), item.get('th', '')
                )
                self.tree.insert(parent_id, "end", values=values)
        
        self.tree.tag_configure('group', background='#E8E8E8')

    def on_item_select(self, event):
        """트리뷰 항목 선택 시 편집 창에 데이터 표시"""
        selected_item = self.tree.focus()
        if not selected_item: return
        
        # 부모 노드를 선택하면 편집창 비우기
        if self.tree.parent(selected_item) == "":
            for lang_var in self.edit_vars.values():
                lang_var.set("")
            return

        values = self.tree.item(selected_item, "values")
        for i, col in enumerate(self.tree['columns']):
            self.edit_vars[col].set(values[i])

    def save_resolution(self):
        """사용자가 수정한 내용을 최종본으로 저장"""
        string_id = self.edit_vars["string_id"].get()
        if not string_id:
            messagebox.showwarning("선택 오류", "수정할 항목을 목록에서 선택하세요.", parent=self)
            return
            
        # 수정된 내용을 resolved_data에 저장
        self.resolved_data[string_id] = {lang: var.get() for lang, var in self.edit_vars.items()}
        
        # UI에 확정된 내용 표시 (예: 배경색 변경)
        selected_item = self.tree.focus()
        parent_item = self.tree.parent(selected_item)
        if parent_item:
             # 모든 형제 항목들의 태그를 초기화하고 선택된 항목에만 태그 적용
            for child in self.tree.get_children(parent_item):
                self.tree.item(child, tags=())
            self.tree.item(selected_item, tags=('resolved',))
            self.tree.tag_configure('resolved', background='lightgreen')
        
        messagebox.showinfo("저장 완료", f"'{string_id}'에 대한 수정사항이 임시 저장되었습니다.\n모든 항목을 수정한 후 '확정 및 DB 반영'을 눌러주세요.", parent=self)

    def apply_resolved_data(self):
        """최종 확정된 데이터를 DB에 업데이트"""
        if not self.resolved_data:
            messagebox.showwarning("확인 필요", "하나 이상의 항목을 수정한 후 [선택 항목으로 최종본 저장]을 먼저 눌러주세요.", parent=self)
            return

        if messagebox.askyesno("최종 확인", f"{len(self.resolved_data)}개의 수정된 항목을 DB에 최종 반영하시겠습니까?", parent=self):
            # Manager를 통해 DB 업데이트 로직 실행
            self.apply_button.config(state="disabled")
            self.manager.run_step2_apply_resolved_data(self.db_path, self.resolved_data, self.on_apply_complete)
    
    def on_apply_complete(self, result):
        if result.get("status") == "success":
            messagebox.showinfo("반영 완료", f"수정된 {result.get('updated_rows', 0)}개 항목이 DB에 성공적으로 반영되었습니다.", parent=self)
            self.completion_callback(result) # 부모 창에 완료 신호 전달
            self.destroy()
        else:
            messagebox.showerror("반영 실패", f"DB 반영 중 오류 발생:\n{result.get('message')}", parent=self)
            self.apply_button.config(state="normal")

    def export_to_excel(self):
        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel 파일", "*.xlsx")])
        if not save_path: return
        
        export_list = []
        for string_id, items in self.duplicate_data.items():
             for item in items:
                export_list.append(item)

        df = pd.DataFrame(export_list)
        df.to_excel(save_path, index=False)
        messagebox.showinfo("내보내기 완료", f"중복 항목 목록을 엑셀 파일로 저장했습니다.\n{save_path}", parent=self)