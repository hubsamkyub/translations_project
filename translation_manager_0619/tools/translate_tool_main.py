import tkinter as tk
import os
import logging
import json
import pandas as pd
import time
import gc
import threading
import pythoncom  # 추가
import win32com.client  # 추가
from tkinter import filedialog, messagebox, ttk

# 리팩토링한 모듈 불러오기
from ui.common_components import (
    ScrollableCheckList, select_folder, select_file, show_message, 
    LoadingPopup, save_file
)
from tools.translate.translation_db_manager import TranslationDBManager
from tools.db_compare_manager import DBCompareManager
from tools.translate.translation_apply_manager import TranslationApplyManager
from utils.config_utils import load_config, save_config
from tools.translate.string_sync_manager import StringSyncManager
from tools.translate.word_replacement_manager import WordReplacementManager


# TranslationAutomationTool 클래스를 다시 정의
class TranslationAutomationTool(tk.Frame):
    def __init__(self, root):
        # 기본 로그 설정
        logging.basicConfig(
            filename='translation_tool.log',
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            filemode='w'
        )
        
        self.root = root
        self.root.title("번역 자동화 툴")
        self.root.geometry("1400x800")
        
        # 모듈 초기화
        self.db_manager = TranslationDBManager(root)
        self.db_compare_manager = DBCompareManager(root)
        self.translation_apply_manager = TranslationApplyManager(self)
        
        # 상단에 확장 기능 버튼 추가
        extension_frame = ttk.Frame(root)
        extension_frame.pack(fill="x", padx=10, pady=5)
        
        # 노트북으로 탭 생성
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill="both", expand=True, padx=10, pady=10)
        
        # 탭 프레임 생성
        self.text_extract_frame = ttk.Frame(self.notebook)
        self.db_compare_frame = ttk.Frame(self.notebook)
        self.translation_db_frame = ttk.Frame(self.notebook)  # 번역 DB 구축 탭
        self.translation_apply_frame = ttk.Frame(self.notebook)        
        self.translation_request_frame = ttk.Frame(self.notebook)  # 번역 요청 추출 탭
        self.string_sync_frame = ttk.Frame(self.notebook)  # STRING 동기화 탭 추가
        self.excel_split_frame = ttk.Frame(self.notebook)  # 엑셀 시트 분리 탭 추가
        self.word_replacement_frame = ttk.Frame(self.notebook)  # 단어 치환 탭 추가
    
        # 노트북에 탭 추가
        self.notebook.add(self.db_compare_frame, text="DB 비교 추출")
        self.notebook.add(self.translation_db_frame, text="번역 DB 구축")
        self.notebook.add(self.translation_apply_frame, text="번역 적용")
        self.notebook.add(self.translation_request_frame, text="번역 요청 추출")
        self.notebook.add(self.string_sync_frame, text="STRING 동기화")  # 새 탭 추가
        self.notebook.add(self.excel_split_frame, text="엑셀 시트 분리")  # 엑셀 시트 분리 탭 추가
        self.notebook.add(self.word_replacement_frame, text="단어 치환")  # 새 탭 추가
        
        # 제외 파일 목록
        self.excluded_files = self.load_excluded_files()
        
        # 번역 캐시 및 설정 변수 추가
        self.translation_cache = {}
        self.available_languages = ["KR", "EN", "CN", "TW", "TH"]  # 지원 언어 목록
        
        # 언어 매핑 추가 (대체 언어 키)
        self.language_mapping = {
            "ZH": "CN",  # ZH는 CN과 동일하게 처리
        }
        
        # 각 탭 구성
        self.setup_db_compare_tab()
        self.setup_translation_db_tab()
        self.setup_translation_apply_tab()
        self.setup_string_sync_tab()  # STRING 동기화 탭 설정 추가
        self.setup_excel_split_tab()
        self.setup_word_replacement_tab()  # 단어 치환 탭 설정 추가
        
        # 번역 요청 추출 탭은 translation_request_extractor.py에서 가져오기
        try:
            from tools.translate.translation_request_extractor import TranslationRequestExtractor
            self.translation_request_extractor = TranslationRequestExtractor(self.translation_request_frame)
            self.translation_request_extractor.root = self.root  # 부모 윈도우 설정
            self.translation_request_extractor.pack(fill="both", expand=True)
        except ImportError:
            ttk.Label(self.translation_request_frame, text="번역 요청 추출 모듈을 불러올 수 없습니다.").pack(pady=20)

    def load_excluded_files(self):
        """제외 파일 목록 로드"""
        try:
            with open("제외 파일 목록.txt", "r", encoding="utf-8") as f:
                return [line.strip() for line in f.readlines() if line.strip()]
        except Exception:
            return []

    #DB 비교 추출
    def setup_db_compare_tab(self):
        """DB 비교 탭 설정 (통합 버전)"""
        # 상단 프레임 (좌우 분할)
        top_frame = ttk.Frame(self.db_compare_frame)
        top_frame.pack(fill="x", padx=5, pady=5)
        
        # 좌측 프레임 (개별 DB 비교 + 폴더 DB 비교)
        left_frame = ttk.Frame(top_frame)
        left_frame.pack(side="left", fill="both", expand=True, padx=(0, 5))
        
        # ===== 개별 DB 비교 섹션 =====
        single_db_frame = ttk.LabelFrame(left_frame, text="개별 DB 비교")
        single_db_frame.pack(fill="x", padx=5, pady=5)
        
        ttk.Label(single_db_frame, text="원본 DB:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.original_db_var = tk.StringVar()
        ttk.Entry(single_db_frame, textvariable=self.original_db_var, width=40).grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        ttk.Button(single_db_frame, text="찾아보기", 
                command=lambda: self.select_db_file("original")).grid(row=0, column=2, padx=5, pady=5)
        
        ttk.Label(single_db_frame, text="비교 DB:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.compare_db_var = tk.StringVar()
        ttk.Entry(single_db_frame, textvariable=self.compare_db_var, width=40).grid(row=1, column=1, padx=5, pady=5, sticky="ew")
        ttk.Button(single_db_frame, text="찾아보기", 
                command=lambda: self.select_db_file("compare")).grid(row=1, column=2, padx=5, pady=5)
        
        single_db_frame.columnconfigure(1, weight=1)
        
        # ===== 폴더 DB 비교 섹션 =====
        folder_db_frame = ttk.LabelFrame(left_frame, text="폴더 DB 비교")
        folder_db_frame.pack(fill="x", padx=5, pady=5)
        
        ttk.Label(folder_db_frame, text="원본 폴더:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.original_folder_db_var = tk.StringVar()
        ttk.Entry(folder_db_frame, textvariable=self.original_folder_db_var, width=40).grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        ttk.Button(folder_db_frame, text="찾아보기", 
                command=lambda: self.select_db_folder("original")).grid(row=0, column=2, padx=5, pady=5)
        
        ttk.Label(folder_db_frame, text="비교 폴더:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.compare_folder_db_var = tk.StringVar()
        ttk.Entry(folder_db_frame, textvariable=self.compare_folder_db_var, width=40).grid(row=1, column=1, padx=5, pady=5, sticky="ew")
        ttk.Button(folder_db_frame, text="찾아보기", 
                command=lambda: self.select_db_folder("compare")).grid(row=1, column=2, padx=5, pady=5)
        
        ttk.Button(folder_db_frame, text="DB 목록 보기", 
                command=self.show_db_list).grid(row=2, column=1, padx=5, pady=5, sticky="e")
        
        folder_db_frame.columnconfigure(1, weight=1)
        
        # 우측 프레임 (DB 목록 + 비교 옵션)
        right_frame = ttk.Frame(top_frame)
        right_frame.pack(side="right", fill="both", expand=True, padx=(5, 0))
        
        # DB 목록 표시 영역
        self.db_list_frame = ttk.LabelFrame(right_frame, text="비교할 DB 목록")
        self.db_list_frame.pack(fill="both", expand=True, padx=5, pady=5)
        
        # 체크박스 목록 컨테이너
        self.db_checklist = ScrollableCheckList(self.db_list_frame, width=350, height=150)
        self.db_checklist.pack(fill="both", expand=True, padx=5, pady=5)
        
        # 비교 옵션 설정 (통합)
        options_frame = ttk.LabelFrame(right_frame, text="비교 옵션")
        options_frame.pack(fill="x", padx=5, pady=5)

        # STRING DB 비교 옵션
        string_options_frame = ttk.LabelFrame(options_frame, text="STRING DB 옵션")
        string_options_frame.pack(fill="x", padx=5, pady=2)

        self.changed_kr_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(string_options_frame, text="KR 값이 변경된 항목 추출", 
                    variable=self.changed_kr_var).grid(row=0, column=0, padx=5, pady=2, sticky="w")

        self.new_items_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(string_options_frame, text="비교본에만 있는 새 항목 추출", 
                    variable=self.new_items_var).grid(row=1, column=0, padx=5, pady=2, sticky="w")

        self.deleted_items_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(string_options_frame, text="원본에서 삭제된 항목 추출", 
                    variable=self.deleted_items_var).grid(row=2, column=0, padx=5, pady=2, sticky="w")

        # TRANSLATION DB 비교 언어 옵션 (통합)
        lang_options_frame = ttk.LabelFrame(options_frame, text="언어 옵션 (TRANSLATION DB용)")
        lang_options_frame.pack(fill="x", padx=5, pady=2)
        
        ttk.Label(lang_options_frame, text="비교 언어:").grid(row=0, column=0, padx=5, pady=2, sticky="w")
        lang_frame = ttk.Frame(lang_options_frame)
        lang_frame.grid(row=0, column=1, columnspan=2, padx=5, pady=2, sticky="w")
        
        self.compare_lang_vars = {}
        for i, lang in enumerate(self.available_languages):
            var = tk.BooleanVar(value=True)
            self.compare_lang_vars[lang] = var
            ttk.Checkbutton(lang_frame, text=lang, variable=var).grid(
                row=0, column=i, padx=5, sticky="w")
        
        # 비교 실행 버튼 프레임
        action_frame = ttk.Frame(self.db_compare_frame)
        action_frame.pack(fill="x", padx=5, pady=5)
        
        ttk.Button(action_frame, text="개별 DB 비교", 
                command=self.compare_individual_databases).pack(side="right", padx=5, pady=5)
        ttk.Button(action_frame, text="폴더 DB 비교", 
                command=self.compare_folder_databases).pack(side="right", padx=5, pady=5)
        ttk.Button(action_frame, text="결과 내보내기", 
                command=self.export_compare_results).pack(side="right", padx=5, pady=5)
        
        # 하단 프레임 (결과 표시)
        bottom_frame = ttk.Frame(self.db_compare_frame)
        bottom_frame.pack(fill="both", expand=True, padx=5, pady=5)
        
        # 결과 표시 영역
        result_frame = ttk.LabelFrame(bottom_frame, text="비교 결과")
        result_frame.pack(fill="both", expand=True, padx=5, pady=5)
        
        # 트리뷰 결과 표시를 위한 프레임
        tree_frame = ttk.Frame(result_frame)
        tree_frame.pack(fill="both", expand=True)
        
        # 트리뷰로 결과 표시 - 통합된 컬럼 구성
        columns = ("db_name", "file_name", "sheet_name", "string_id", "type", "kr", "original_kr")
        self.compare_result_tree = ttk.Treeview(tree_frame, columns=columns, show="headings")
        
        # 컬럼 설정
        self.compare_result_tree.heading("db_name", text="DB명")
        self.compare_result_tree.heading("file_name", text="파일명")
        self.compare_result_tree.heading("sheet_name", text="시트명")
        self.compare_result_tree.heading("string_id", text="STRING_ID")
        self.compare_result_tree.heading("type", text="변경 유형")
        self.compare_result_tree.heading("kr", text="KR")
        self.compare_result_tree.heading("original_kr", text="원본 KR")        
        
        self.compare_result_tree.column("db_name", width=120)
        self.compare_result_tree.column("file_name", width=120)
        self.compare_result_tree.column("sheet_name", width=120)
        self.compare_result_tree.column("string_id", width=100)
        self.compare_result_tree.column("type", width=150)
        self.compare_result_tree.column("kr", width=200)
        self.compare_result_tree.column("original_kr", width=200)
        
        # 스크롤바 연결
        scrollbar_y = ttk.Scrollbar(tree_frame, orient="vertical", command=self.compare_result_tree.yview)
        self.compare_result_tree.configure(yscrollcommand=scrollbar_y.set)
        
        scrollbar_x = ttk.Scrollbar(result_frame, orient="horizontal", command=self.compare_result_tree.xview)
        self.compare_result_tree.configure(xscrollcommand=scrollbar_x.set)
        
        # 배치
        scrollbar_y.pack(side="right", fill="y")
        self.compare_result_tree.pack(side="left", fill="both", expand=True)
        scrollbar_x.pack(side="bottom", fill="x")
        
        # 상태 및 진행 표시 프레임
        status_frame = ttk.Frame(self.db_compare_frame)
        status_frame.pack(fill="x", padx=5, pady=5)
        
        self.status_label_compare = ttk.Label(status_frame, text="대기 중...")
        self.status_label_compare.pack(side="left", fill="x", expand=True, padx=5)
        
        self.progress_label = ttk.Label(status_frame, text="진행 상황:")
        self.progress_label.pack(side="left", padx=5)
        
        self.progress_bar_compare = ttk.Progressbar(status_frame, length=300, mode="determinate")
        self.progress_bar_compare.pack(side="right", padx=5)
        
        # 내부 데이터 저장용
        self.compare_results = []
        self.db_pairs = []  # 폴더 비교에 사용할 DB 파일 쌍


    def setup_translation_db_tab(self):
        """번역 DB 구축 탭 구성"""
        # 엑셀 파일 선택 프레임
        excel_frame = ttk.LabelFrame(self.translation_db_frame, text="번역 파일 선택")
        excel_frame.pack(fill="x", padx=5, pady=5)
        
        # 폴더 선택 행
        folder_frame = ttk.Frame(excel_frame)
        folder_frame.pack(fill="x", padx=5, pady=5)
        
        ttk.Label(folder_frame, text="엑셀 폴더:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.trans_excel_folder_var = tk.StringVar()
        ttk.Entry(folder_frame, textvariable=self.trans_excel_folder_var, width=50).grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        ttk.Button(folder_frame, text="찾아보기", 
                command=lambda: self.select_folder(self.trans_excel_folder_var, "번역 엑셀 폴더 선택")).grid(row=0, column=2, padx=5, pady=5)
        ttk.Button(folder_frame, text="파일 검색", 
                command=self.search_translation_excel_files).grid(row=0, column=3, padx=5, pady=5)
        
        folder_frame.columnconfigure(1, weight=1)
        
        # 파일 목록 프레임
        files_frame = ttk.LabelFrame(self.translation_db_frame, text="번역 엑셀 파일 목록")
        files_frame.pack(fill="both", expand=True, padx=5, pady=5)
        
        self.trans_excel_files_list = ScrollableCheckList(files_frame, width=700, height=150)
        self.trans_excel_files_list.pack(fill="both", expand=True, padx=5, pady=5)
        
        # DB 출력 설정 프레임
        output_frame = ttk.LabelFrame(self.translation_db_frame, text="DB 출력 설정")
        output_frame.pack(fill="x", padx=5, pady=5)
        
        # DB 파일 선택 행 (구축용)
        db_build_frame = ttk.Frame(output_frame)
        db_build_frame.pack(fill="x", padx=5, pady=2)
        
        ttk.Label(db_build_frame, text="새 DB 파일:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.output_db_var = tk.StringVar()
        ttk.Entry(db_build_frame, textvariable=self.output_db_var, width=50).grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        ttk.Button(db_build_frame, text="찾아보기", 
                command=lambda: self.save_db_file(self.output_db_var, "새 번역 DB 파일 저장")).grid(row=0, column=2, padx=5, pady=5)
        
        db_build_frame.columnconfigure(1, weight=1)
        
        # DB 파일 선택 행 (업데이트용)
        db_update_frame = ttk.Frame(output_frame)
        db_update_frame.pack(fill="x", padx=5, pady=2)
        
        ttk.Label(db_update_frame, text="기존 DB 파일:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.update_db_var = tk.StringVar()
        ttk.Entry(db_update_frame, textvariable=self.update_db_var, width=50).grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        ttk.Button(db_update_frame, text="찾아보기", 
                command=lambda: self.select_file(self.update_db_var, "기존 번역 DB 파일 선택", [("DB 파일", "*.db")])).grid(row=0, column=2, padx=5, pady=5)
        
        db_update_frame.columnconfigure(1, weight=1)
        
        # 언어 선택 프레임
        languages_frame = ttk.LabelFrame(self.translation_db_frame, text="추출할 언어")
        languages_frame.pack(fill="x", padx=5, pady=5)
        
        self.lang_vars = {}
        for i, lang in enumerate(self.available_languages):
            var = tk.BooleanVar(value=True)
            self.lang_vars[lang] = var
            ttk.Checkbutton(languages_frame, text=lang, variable=var).grid(
                row=i // 3, column=i % 3, padx=20, pady=5, sticky="w")
        
        # 성능 옵션 프레임
        perf_frame = ttk.LabelFrame(self.translation_db_frame, text="성능 설정")
        perf_frame.pack(fill="x", padx=5, pady=5)
        
        ttk.Label(perf_frame, text="배치 크기:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.batch_size_var = tk.IntVar(value=500)
        ttk.Spinbox(perf_frame, from_=100, to=2000, increment=100, 
                   textvariable=self.batch_size_var, width=5).grid(row=0, column=1, padx=5, pady=5, sticky="w")
        
        self.read_only_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(perf_frame, text="읽기 전용 모드 사용 (빠름)", 
                       variable=self.read_only_var).grid(row=0, column=2, padx=20, pady=5, sticky="w")
        
        # 실행 버튼 프레임
        action_frame = ttk.Frame(self.translation_db_frame)
        action_frame.pack(fill="x", padx=5, pady=5)
        
        # 기존 버튼과 새 업데이트 버튼 추가
        ttk.Button(action_frame, text="번역 DB 구축", 
                command=self.build_translation_db).pack(side="right", padx=5, pady=5)
        ttk.Button(action_frame, text="번역 DB 업데이트", 
                command=self.update_translation_db).pack(side="right", padx=5, pady=5)
        
        # 로그 표시 영역
        log_frame = ttk.LabelFrame(self.translation_db_frame, text="작업 로그")
        log_frame.pack(fill="both", expand=True, padx=5, pady=5)
        
        self.db_log_text = tk.Text(log_frame, wrap="word", height=10)
        scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=self.db_log_text.yview)
        self.db_log_text.configure(yscrollcommand=scrollbar.set)
        
        scrollbar.pack(side="right", fill="y")
        self.db_log_text.pack(fill="both", expand=True)
        
        # 상태 표시줄
        status_frame = ttk.Frame(self.translation_db_frame)
        status_frame.pack(fill="x", padx=5, pady=5)
        
        self.status_label_db = ttk.Label(status_frame, text="대기 중...")
        self.status_label_db.pack(side="left", padx=5)
        
        self.progress_bar_db = ttk.Progressbar(status_frame, length=400, mode="determinate")
        self.progress_bar_db.pack(side="right", fill="x", expand=True, padx=5)

    def update_translation_db(self):
        """번역 DB 업데이트 실행"""
        # 입력 유효성 검증
        selected_files = self.trans_excel_files_list.get_checked_items()
        if not selected_files:
            messagebox.showwarning("경고", "번역 파일을 선택하세요.", parent=self.root)
            return
        
        db_path = self.update_db_var.get()
        if not db_path:
            messagebox.showwarning("경고", "기존 DB 파일 경로를 지정하세요.", parent=self.root)
            return
        
        if not os.path.exists(db_path):
            messagebox.showwarning("경고", "기존 DB 파일이 존재하지 않습니다.", parent=self.root)
            return
        
        # 선택된 언어 확인
        selected_langs = [lang for lang, var in self.lang_vars.items() if var.get()]
        if not selected_langs:
            messagebox.showwarning("경고", "하나 이상의 언어를 선택하세요.", parent=self.root)
            return
        
        # 로그 초기화
        self.db_log_text.delete(1.0, tk.END)
        self.db_log_text.insert(tk.END, "번역 DB 업데이트 시작...\n")
        self.status_label_db.config(text="번역 DB 업데이트 중...")
        self.root.update()
        
        # 파일 경로 리스트 만들기
        excel_files = [(file, path) for file, path in self.trans_excel_files if file in selected_files]
        
        # 성능 설정 가져오기
        batch_size = self.batch_size_var.get()
        use_read_only = self.read_only_var.get()
        
        # 진행 창 생성
        loading_popup = LoadingPopup(self.root, "번역 DB 업데이트 중", "번역 DB 업데이트 준비 중...")
        
        # 시작 시간 기록
        start_time = time.time()
        
        # 진행 콜백 함수
        def progress_callback(message, current, total):
            self.root.after(0, lambda: [
                loading_popup.update_progress((current / total) * 100, f"{current}/{total} - {message}"),
                self.db_log_text.insert(tk.END, f"{message}\n"),
                self.db_log_text.see(tk.END)
            ])
        
        # 작업 스레드 함수
        def update_db():
            try:
                # DB 업데이트 실행
                result = self.db_manager.update_translation_db(
                    excel_files, 
                    db_path, 
                    selected_langs, 
                    batch_size, 
                    use_read_only,
                    progress_callback
                )
                
                # 결과 처리 (메인 스레드에서)
                self.root.after(0, lambda: self.process_db_update_result(
                    result, loading_popup, start_time))
                
            except Exception as e:
                self.root.after(0, lambda: [
                    loading_popup.close(),
                    self.db_log_text.insert(tk.END, f"\n오류 발생: {str(e)}\n"),
                    self.status_label_db.config(text="오류 발생"),
                    messagebox.showerror("오류", f"DB 업데이트 중 오류 발생: {str(e)}", parent=self.root)
                ])
                
        # 백그라운드 스레드 실행
        thread = threading.Thread(target=update_db)
        thread.daemon = True
        thread.start()

    def process_db_update_result(self, result, loading_popup, start_time):
        """DB 업데이트 결과 처리"""
        loading_popup.close()
        
        if result["status"] == "error":
            self.db_log_text.insert(tk.END, f"\n오류 발생: {result['message']}\n")
            self.status_label_db.config(text="오류 발생")
            messagebox.showerror("오류", f"DB 업데이트 중 오류 발생: {result['message']}", parent=self.root)
            return
            
        # 작업 시간 계산
        elapsed_time = time.time() - start_time
        time_str = f"{int(elapsed_time // 60)}분 {int(elapsed_time % 60)}초"
        
        # 작업 완료 메시지
        self.db_log_text.insert(tk.END, f"\n번역 DB 업데이트 완료! (소요 시간: {time_str})\n")
        self.db_log_text.insert(tk.END, f"파일 처리: {result['processed_count']}/{len(self.trans_excel_files_list.get_checked_items())} (오류: {result['error_count']})\n")
        self.db_log_text.insert(tk.END, f"신규 추가: {result.get('new_rows', 0)}개\n")
        self.db_log_text.insert(tk.END, f"기존 업데이트: {result.get('updated_rows', 0)}개\n")
        self.db_log_text.insert(tk.END, f"삭제 표시: {result.get('deleted_rows', 0)}개\n")
        self.db_log_text.insert(tk.END, f"총 처리된 항목: {result['total_rows']}개\n")
        
        self.status_label_db.config(text=f"번역 DB 업데이트 완료 - {result['total_rows']}개 항목")
        
        # 상세 통계 표시
        update_summary = (
            f"번역 DB 업데이트가 완료되었습니다.\n\n"
            f"📊 처리 통계:\n"
            f"• 신규 추가: {result.get('new_rows', 0)}개\n"
            f"• 기존 업데이트: {result.get('updated_rows', 0)}개\n"
            f"• 삭제 표시: {result.get('deleted_rows', 0)}개\n"
            f"• 총 처리: {result['total_rows']}개\n\n"
            f"⏱️ 소요 시간: {time_str}"
        )
        
        messagebox.showinfo("완료", update_summary, parent=self.root)

    # DB 비교 탭 함수
    def select_folder(self, var, title):
        """폴더 선택 다이얼로그 공통 함수"""
        folder = filedialog.askdirectory(title=title, parent=self.root)
        if folder:
            var.set(folder)
            # 포커스를 다시 자동화 툴 창으로
            self.root.after(100, self.root.focus_force)
            self.root.after(100, self.root.lift)

    def select_file(self, var, title, filetypes):
        """파일 선택 다이얼로그 공통 함수"""
        file_path = filedialog.askopenfilename(
            filetypes=filetypes,
            title=title,
            parent=self.root
        )
        if file_path:
            var.set(file_path)
            # 포커스를 다시 자동화 툴 창으로
            self.root.after(100, self.root.focus_force)
            self.root.after(100, self.root.lift)

    def save_db_file(self, var, title):
        """DB 파일 저장 다이얼로그"""
        file_path = filedialog.asksaveasfilename(
            defaultextension=".db",
            filetypes=[("DB 파일", "*.db"), ("모든 파일", "*.*")],
            title=title,
            parent=self.root
        )
        if file_path:
            var.set(file_path)
            # 포커스를 다시 자동화 툴 창으로
            self.root.after(100, self.root.focus_force)
            self.root.after(100, self.root.lift)

    def select_db_file(self, db_type):
        """DB 파일 선택"""
        file_path = filedialog.askopenfilename(
            filetypes=[("DB 파일", "*.db"), ("모든 파일", "*.*")],
            title=f"{db_type.capitalize()} DB 파일 선택",
            parent=self.root
        )
        if file_path:
            if db_type == "original":
                self.original_db_var.set(file_path)
            else:
                self.compare_db_var.set(file_path)
            
            # 포커스를 다시 자동화 툴 창으로
            self.root.after(100, self.root.focus_force)
            self.root.after(100, self.root.lift)

    def select_db_folder(self, folder_type):
        """DB 폴더 선택"""
        folder = filedialog.askdirectory(title=f"{folder_type.capitalize()} DB 폴더 선택", parent=self.root)
        if folder:
            if folder_type == "original":
                self.original_folder_db_var.set(folder)
            else:
                self.compare_folder_db_var.set(folder)
            
            # 포커스를 다시 자동화 툴 창으로
            self.root.after(100, self.root.focus_force)
            self.root.after(100, self.root.lift)
            

    # translate_tool_main.py에 추가해야 할 메서드들
    def compare_individual_databases(self):
        """개별 DB 비교 (자동 타입 판단)"""
        original_db_path = self.original_db_var.get()
        compare_db_path = self.compare_db_var.get()

        if not original_db_path or not os.path.isfile(original_db_path):
            messagebox.showwarning("경고", "유효한 원본 DB 파일을 선택하세요.", parent=self.root)
            return

        if not compare_db_path or not os.path.isfile(compare_db_path):
            messagebox.showwarning("경고", "유효한 비교 DB 파일을 선택하세요.", parent=self.root)
            return

        # 결과 초기화
        self.compare_result_tree.delete(*self.compare_result_tree.get_children())
        self.compare_results = []

        # 진행 창 생성
        loading_popup = LoadingPopup(self.root, "DB 비교 중", "DB 타입 확인 및 비교 중...")
        
        # 작업 스레드 함수
        def run_comparison():
            try:
                # DB 타입 자동 판단 및 비교 실행
                result = self.db_compare_manager.auto_compare_databases(
                    original_db_path,
                    compare_db_path,
                    self.get_compare_options()
                )
                
                # 결과 처리 (메인 스레드에서)
                self.root.after(0, lambda: self.process_unified_compare_results(result, loading_popup))
                
            except Exception as e:
                self.root.after(0, lambda: [
                    loading_popup.close(),
                    messagebox.showerror("오류", f"DB 비교 중 오류 발생: {str(e)}", parent=self.root)
                ])
                
        # 백그라운드 스레드 실행
        thread = threading.Thread(target=run_comparison)
        thread.daemon = True
        thread.start()

    def compare_folder_databases(self):
        """폴더 DB 비교 (자동 타입 판단)"""
        if not self.db_pairs:
            messagebox.showwarning("경고", "비교할 DB 파일 목록이 없습니다. 'DB 목록 보기'를 먼저 실행하세요.", parent=self.root)
            return
        
        # 선택된 DB 파일 쌍 필터링
        selected_db_names = self.db_checklist.get_checked_items()
        selected_db_pairs = [pair for pair in self.db_pairs if pair['file_name'] in selected_db_names]
        
        if not selected_db_pairs:
            messagebox.showwarning("경고", "비교할 DB 파일을 선택하세요.", parent=self.root)
            return
        
        # 결과 초기화
        self.compare_result_tree.delete(*self.compare_result_tree.get_children())
        self.compare_results = []
        
        # 진행 창 생성
        loading_popup = LoadingPopup(self.root, "폴더 DB 비교 중", "DB 비교 작업 준비 중...")
        
        # 작업용 스레드 함수
        def run_comparison():
            try:
                # 폴더 DB 비교 실행 (자동 타입 판단)
                def progress_callback(status_text, current, total):
                    self.root.after(0, lambda: loading_popup.update_progress(
                        (current / total) * 100, status_text))
                
                result = self.db_compare_manager.auto_compare_folder_databases(
                    selected_db_pairs,
                    self.get_compare_options(),
                    progress_callback
                )
                
                # 결과 처리 (메인 스레드에서)
                self.root.after(0, lambda: self.process_unified_compare_results(result, loading_popup))
                
            except Exception as e:
                self.root.after(0, lambda: [
                    loading_popup.close(),
                    messagebox.showerror("오류", f"폴더 DB 비교 중 오류 발생: {str(e)}", parent=self.root)
                ])

        # 백그라운드 스레드 실행
        thread = threading.Thread(target=run_comparison)
        thread.daemon = True
        thread.start()

    def get_compare_options(self):
        """비교 옵션 수집"""
        return {
            # STRING DB 옵션
            "changed_kr": self.changed_kr_var.get(),
            "new_items": self.new_items_var.get(),
            "deleted_items": self.deleted_items_var.get(),
            # TRANSLATION DB 언어 옵션
            "languages": [lang.lower() for lang, var in self.compare_lang_vars.items() if var.get()]
        }

    def process_unified_compare_results(self, result, loading_popup):
        """통합된 비교 결과 처리"""
        loading_popup.close()
        
        if result["status"] != "success":
            messagebox.showerror("오류", result["message"], parent=self.root)
            return
            
        # 결과 저장
        self.compare_results = result["compare_results"]
        
        # 결과 표시 업데이트 (통합된 형태)
        self.compare_result_tree.delete(*self.compare_result_tree.get_children())
        for idx, item in enumerate(self.compare_results):
            # 통합된 결과 표시
            db_name = item.get("db_name", item.get("file_name", ""))
            file_name = item.get("file_name", "")
            sheet_name = item.get("sheet_name", "")
            string_id = item.get("string_id", "")
            item_type = item.get("type", "")
            kr = item.get("kr", "")
            original_kr = item.get("original_kr", "")
            
            self.compare_result_tree.insert(
                "", "end", iid=idx,
                values=(db_name, file_name, sheet_name, string_id, item_type, kr, original_kr)
            )
        
        # 상태 표시줄 업데이트
        total_changes = result.get("total_changes", len(self.compare_results))
        db_type = result.get("db_type", "DB")
        
        self.status_label_compare.config(
            text=f"{db_type} 비교 완료: {total_changes}개 차이점 발견"
        )
        
        # 결과 메시지
        if total_changes > 0:
            summary_msg = (
                f"{db_type} 비교가 완료되었습니다.\n\n"
                f"🔍 총 {total_changes}개의 차이점을 발견했습니다."
            )
            
            # 세부 통계가 있으면 추가
            if "new_items" in result:
                summary_msg += (
                    f"\n\n📊 세부 결과:\n"
                    f"• 신규 항목: {result.get('new_items', 0)}개\n"
                    f"• 삭제된 항목: {result.get('deleted_items', 0)}개\n"
                    f"• 변경된 항목: {result.get('changed_items', 0)}개"
                )
            
            messagebox.showinfo("완료", summary_msg, parent=self.root)
        else:
            messagebox.showinfo("완료", f"두 {db_type}가 동일합니다.", parent=self.root)

    # show_db_list 메서드도 수정이 필요합니다:
    def show_db_list(self):
        """폴더 내 DB 파일 목록 표시 (모든 DB 파일)"""
        original_folder = self.original_folder_db_var.get()
        compare_folder = self.compare_folder_db_var.get()
        
        if not original_folder or not os.path.isdir(original_folder):
            messagebox.showwarning("경고", "유효한 원본 DB 폴더를 선택하세요.", parent=self.root)
            return
        
        if not compare_folder or not os.path.isdir(compare_folder):
            messagebox.showwarning("경고", "유효한 비교 DB 폴더를 선택하세요.", parent=self.root)
            return
        
        # DB 파일 목록 가져오기 (모든 .db 파일)
        original_dbs = {f for f in os.listdir(original_folder) if f.endswith('.db')}
        compare_dbs = {f for f in os.listdir(compare_folder) if f.endswith('.db')}
        
        # 공통 DB 파일만 찾기
        common_dbs = original_dbs.intersection(compare_dbs)
        
        if not common_dbs:
            messagebox.showinfo("알림", "두 폴더에 공통된 DB 파일이 없습니다.", parent=self.root)
            return
        
        # DB 목록 업데이트
        self.db_checklist.clear()
        self.db_pairs = []
        
        for db_file in sorted(common_dbs):
            self.db_checklist.add_item(db_file, checked=True)
            self.db_pairs.append({
                'file_name': db_file,
                'original_path': os.path.join(original_folder, db_file),
                'compare_path': os.path.join(compare_folder, db_file)
            })
        
        messagebox.showinfo("알림", f"{len(common_dbs)}개의 공통 DB 파일을 찾았습니다.", parent=self.root)


    def export_compare_results(self):
        """비교 결과를 엑셀 파일로 내보내기"""
        if not self.compare_results:
            messagebox.showwarning("경고", "내보낼 비교 결과가 없습니다.")
            return
        
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel 파일", "*.xlsx")],
            title="비교 결과 저장",
            parent=self.root
        )
        
        if not file_path:
            return
            
        # 진행 창 생성
        loading_popup = LoadingPopup(self.root, "결과 내보내기", "엑셀 파일로 저장 중...")
        
        # 작업 함수
        def export_data():
            try:
                # 데이터프레임으로 변환하여 저장
                df = pd.DataFrame(self.compare_results)
                df.to_excel(file_path, index=False)
                
                # 완료 처리 (메인 스레드에서)
                self.root.after(0, lambda: [
                    loading_popup.close(),
                    messagebox.showinfo("완료", f"비교 결과가 {file_path}에 저장되었습니다.", parent=self.root)
                ])
            except Exception as e:
                self.root.after(0, lambda: [
                    loading_popup.close(),
                    messagebox.showerror("오류", f"데이터 저장 실패: {str(e)}", parent=self.root)
                ])
                
        # 백그라운드 스레드 실행
        thread = threading.Thread(target=export_data)
        thread.daemon = True
        thread.start()

    # 번역 DB 구축 탭 함수들
    def search_translation_excel_files(self):
        """번역 엑셀 파일 검색"""
        folder = self.trans_excel_folder_var.get()
        if not folder or not os.path.isdir(folder):
            messagebox.showwarning("경고", "유효한 폴더를 선택하세요.", parent=self.root)
            return
        
        self.trans_excel_files_list.clear()
        self.trans_excel_files = []
        
        # 폴더와 하위 폴더 검색
        for root, _, files in os.walk(folder):
            for file in files:
                if file.endswith(".xlsx"):
                    if file not in self.excluded_files:
                        # 파일명에서 확장자 제거 후 string으로 시작하는지 확인 (대소문자 구분 없음)
                        file_name_without_ext = os.path.splitext(file)[0].lower()
                        if file_name_without_ext.startswith("string"):
                            file_path = os.path.join(root, file)
                            self.trans_excel_files.append((file, file_path))
                            self.trans_excel_files_list.add_item(file, checked=True)
        
        if not self.trans_excel_files:
            messagebox.showinfo("알림", "엑셀 파일을 찾지 못했습니다.", parent=self.root)
        else:
            messagebox.showinfo("알림", f"{len(self.trans_excel_files)}개의 엑셀 파일을 찾았습니다.", parent=self.root)
        
        # 🔧 여기에 추가! 기존 messagebox 다음에 넣으세요
        # ScrollableCheckList 상태 확인
        print("\n[체크박스 상태 확인]")
        try:
            # ScrollableCheckList의 실제 메서드명을 확인해야 할 수도 있습니다
            all_items = []
            checked_items = self.trans_excel_files_list.get_checked_items()
            
            # 전체 항목 가져오기 (ScrollableCheckList 구현에 따라 다를 수 있음)
            # get_all_items() 메서드가 없다면 아래처럼 확인:
            for file, path in self.trans_excel_files:
                all_items.append(file)
            
            print(f"전체 항목: {len(all_items)}개")
            print(f"체크된 항목: {len(checked_items)}개")
            
            for item in all_items:
                is_checked = item in checked_items
                status = "✓" if is_checked else "✗"
                print(f"  {status} {item}")
                
                if item in ["String@_New.xlsx", "String_EventDialouge_New.xlsx"]:
                    print(f"    ⚠️ 문제 파일 체크 상태: {is_checked}")
                    
        except Exception as e:
            print(f"[체크박스 상태 확인 오류] {e}")

            
    def build_translation_db(self):
        """번역 DB 구축 실행"""
        # 입력 유효성 검증
        selected_files = self.trans_excel_files_list.get_checked_items()
        if not selected_files:
            messagebox.showwarning("경고", "번역 파일을 선택하세요.", parent=self.root)
            return
        
        db_path = self.output_db_var.get()
        if not db_path:
            messagebox.showwarning("경고", "DB 파일 경로를 지정하세요.", parent=self.root)
            return
        
        # 선택된 언어 확인
        selected_langs = [lang for lang, var in self.lang_vars.items() if var.get()]
        if not selected_langs:
            messagebox.showwarning("경고", "하나 이상의 언어를 선택하세요.", parent=self.root)
            return
        
        # 이미 파일 존재 여부 체크
        if os.path.exists(db_path):
            if not messagebox.askyesno("확인", f"'{db_path}' 파일이 이미 존재합니다. 덮어쓰시겠습니까?", parent=self.root):
                return
        
        # 로그 초기화
        self.db_log_text.delete(1.0, tk.END)
        self.db_log_text.insert(tk.END, "번역 DB 구축 시작...\n")
        self.status_label_db.config(text="번역 DB 구축 중...")
        self.root.update()
        
        # 파일 경로 리스트 만들기
        excel_files = [(file, path) for file, path in self.trans_excel_files if file in selected_files]
        
        # 성능 설정 가져오기
        batch_size = self.batch_size_var.get()
        use_read_only = self.read_only_var.get()
        
        # 진행 창 생성
        loading_popup = LoadingPopup(self.root, "번역 DB 구축 중", "번역 DB 구축 준비 중...")
        
        # 시작 시간 기록
        start_time = time.time()
        
        # 진행 콜백 함수
        def progress_callback(message, current, total):
            self.root.after(0, lambda: [
                loading_popup.update_progress((current / total) * 100, f"{current}/{total} - {message}"),
                self.db_log_text.insert(tk.END, f"{message}\n"),
                self.db_log_text.see(tk.END)
            ])
        
        # 작업 스레드 함수
        def build_db():
            try:
                # DB 구축 실행
                result = self.db_manager.build_translation_db(
                    excel_files, 
                    db_path, 
                    selected_langs, 
                    batch_size, 
                    use_read_only,
                    progress_callback
                )
                
                # 결과 처리 (메인 스레드에서)
                self.root.after(0, lambda: self.process_db_build_result(
                    result, loading_popup, start_time))
                
            except Exception as e:
                self.root.after(0, lambda: [
                    loading_popup.close(),
                    self.db_log_text.insert(tk.END, f"\n오류 발생: {str(e)}\n"),
                    self.status_label_db.config(text="오류 발생"),
                    messagebox.showerror("오류", f"DB 구축 중 오류 발생: {str(e)}", parent=self.root)
                ])
                
        # 백그라운드 스레드 실행
        thread = threading.Thread(target=build_db)
        thread.daemon = True
        thread.start()



    def process_db_build_result(self, result, loading_popup, start_time):
        """DB 구축 결과 처리"""
        loading_popup.close()
        
        if result["status"] == "error":
            self.db_log_text.insert(tk.END, f"\n오류 발생: {result['message']}\n")
            self.status_label_db.config(text="오류 발생")
            messagebox.showerror("오류", f"DB 구축 중 오류 발생: {result['message']}", parent=self.root)
            return
            
        # 작업 시간 계산
        elapsed_time = time.time() - start_time
        time_str = f"{int(elapsed_time // 60)}분 {int(elapsed_time % 60)}초"
        
        # 작업 완료 메시지
        self.db_log_text.insert(tk.END, f"\n번역 DB 구축 완료! (소요 시간: {time_str})\n")
        self.db_log_text.insert(tk.END, f"파일 처리: {result['processed_count']}/{len(self.trans_excel_files_list.get_checked_items())} (오류: {result['error_count']})\n")
        self.db_log_text.insert(tk.END, f"총 {result['total_rows']}개 항목이 DB에 추가되었습니다.\n")
        
        self.status_label_db.config(text=f"번역 DB 구축 완료 - {result['total_rows']}개 항목")
        
        messagebox.showinfo(
            "완료", 
            f"번역 DB 구축이 완료되었습니다.\n총 {result['total_rows']}개 항목이 추가되었습니다.\n소요 시간: {time_str}", 
            parent=self.root
        )

    # 번역 적용 탭 함수들
    def search_original_files(self):
        """원본 엑셀 파일 검색"""
        folder = self.original_folder_var.get()
        if not folder or not os.path.isdir(folder):
            messagebox.showwarning("경고", "유효한 폴더를 선택하세요.", parent=self.root)
            return
        
        self.original_files_list.clear()
        self.original_files = []
        
        # 폴더와 하위 폴더 검색
        for root, _, files in os.walk(folder):
            for file in files:
                if file.startswith("String") and file.endswith(".xlsx"):
                    if file not in self.excluded_files:
                        file_path = os.path.join(root, file)
                        self.original_files.append((file, file_path))
                        self.original_files_list.add_item(file, checked=True)
        
        if not self.original_files:
            messagebox.showinfo("알림", "String으로 시작하는 엑셀 파일을 찾지 못했습니다.", parent=self.root)
        else:
            messagebox.showinfo("알림", f"{len(self.original_files)}개의 엑셀 파일을 찾았습니다.", parent=self.root)

    def load_translation_cache(self):
        """번역 DB를 메모리에 캐싱하여 사용"""
        db_path = self.translation_db_var.get()
        if not db_path or not os.path.isfile(db_path):
            messagebox.showwarning("경고", "유효한 번역 DB 파일을 선택하세요.", parent=self.root)
            return
        
        self.log_text.insert(tk.END, "번역 DB 캐싱 중...\n")
        self.root.update()
        
        # 진행 창 생성
        loading_popup = LoadingPopup(self.root, "번역 DB 캐싱 중", "번역 데이터 캐싱 중...")
        
        # 작업 스레드 함수
        def load_cache():
            try:
                # 번역 캐시 로드
                result = self.translation_apply_manager.load_translation_cache(db_path)
                
                # 결과 처리 (메인 스레드에서)
                self.root.after(0, lambda: self.process_cache_load_result(result, loading_popup))
                
            except Exception as e:
                error_msg = str(e)
                self.root.after(0, lambda: [
                    loading_popup.close(),
                    self.log_text.insert(tk.END, f"캐싱 중 오류 발생: {error_msg}\n"),
                    self.status_label_apply.config(text="오류 발생"),
                    messagebox.showerror("오류", f"DB 캐싱 중 오류 발생: {error_msg}", parent=self.root)
                ])
                
        # 백그라운드 스레드 실행
        thread = threading.Thread(target=load_cache)
        thread.daemon = True
        thread.start()
        
    def process_cache_load_result(self, result, loading_popup):
        """캐시 로드 결과 처리"""
        loading_popup.close()
        
        if "status" in result and result["status"] == "error":
            self.log_text.insert(tk.END, f"캐싱 중 오류 발생: {result['message']}\n")
            self.status_label_apply.config(text="오류 발생")
            messagebox.showerror("오류", f"DB 캐싱 중 오류 발생: {result['message']}", parent=self.root)
            return
            
        # 캐시 데이터 저장
        self.translation_apply_manager.translation_cache = result["translation_cache"]
        self.translation_apply_manager.translation_file_cache = result["translation_file_cache"]
        self.translation_apply_manager.translation_sheet_cache = result["translation_sheet_cache"]
        self.translation_apply_manager.duplicate_ids = result["duplicate_ids"]
        
        # 통계 정보
        file_count = result["file_count"]
        sheet_count = result["sheet_count"]
        id_count = result["id_count"]
        
        # 중복 STRING_ID 로깅
        duplicate_count = sum(1 for ids in result["duplicate_ids"].values() if len(ids) > 1)
        if duplicate_count > 0:
            self.log_text.insert(tk.END, f"\n주의: {duplicate_count}개의 STRING_ID가 여러 파일에 중복 존재합니다.\n")
            
            # 일부 중복 예시 기록 (최대 5개)
            dup_examples = [(id, files) for id, files in result["duplicate_ids"].items() if len(files) > 1][:5]
            for id, files in dup_examples:
                self.log_text.insert(tk.END, f"  - {id}: {', '.join(files)}\n")
            
            if len(dup_examples) < duplicate_count:
                self.log_text.insert(tk.END, f"  ... 외 {duplicate_count - len(dup_examples)}개\n")
        
        self.log_text.insert(tk.END, f"번역 DB 캐싱 완료:\n")
        self.log_text.insert(tk.END, f"- 파일별 캐시: {file_count}개 파일, {sum(len(ids) for ids in result['translation_file_cache'].values())}개 항목\n")
        self.log_text.insert(tk.END, f"- 시트별 캐시: {sheet_count}개 시트, {sum(len(ids) for ids in result['translation_sheet_cache'].values())}개 항목\n")
        self.log_text.insert(tk.END, f"- 전체 고유 STRING_ID: {id_count}개\n")
        
        self.status_label_apply.config(text=f"번역 DB 캐싱 완료 - {id_count}개 항목")
        
        messagebox.showinfo(
            "완료", 
            f"번역 DB 캐싱 완료!\n파일 수: {file_count}개\n시트 수: {sheet_count}개\n항목 수: {id_count}개", 
            parent=self.root
        )

    def apply_translation(self):
        """번역 적용 작업 실행"""
        # 입력 유효성 검사
        if not hasattr(self.translation_apply_manager, 'translation_cache') or not self.translation_apply_manager.translation_cache:
            messagebox.showwarning("경고", "먼저 번역 DB를 캐시에 로드하세요.", parent=self.root)
            return
            
        selected_files = self.original_files_list.get_checked_items()
        if not selected_files:
            messagebox.showwarning("경고", "적용할 파일을 선택하세요.", parent=self.root)
            return
            
        selected_langs = [lang for lang, var in self.apply_lang_vars.items() if var.get()]
        if not selected_langs:
            messagebox.showwarning("경고", "적용할 언어를 하나 이상 선택하세요.", parent=self.root)
            return
            
        # 진행 관련 초기화
        self.log_text.delete(1.0, tk.END)
        self.log_text.insert(tk.END, "번역 적용 작업 시작...\n")
        self.status_label_apply.config(text="작업 중...")
        self.root.update()
            
        self.progress_bar["maximum"] = len(selected_files)
        self.progress_bar["value"] = 0
            
        # 진행 창 생성
        loading_popup = LoadingPopup(self.root, "번역 적용 중", "번역 적용 준비 중...")
            
        # 통계 변수
        total_updated = 0
        processed_count = 0
        error_count = 0
            
        # 작업 스레드 함수
        def apply_translations():
            nonlocal total_updated, processed_count, error_count
            
            # 문제 파일 목록 수집용
            problem_files = {
                "external_links": [],
                "permission_denied": [],
                "file_corrupted": [],
                "file_not_found": [],
                "unknown_error": []
            }
            
            try:
                # 각 파일 처리
                for idx, file_name in enumerate(selected_files):
                    file_path = next((path for name, path in self.original_files if name == file_name), None)
                    if not file_path:
                        continue
                        
                    # 진행 상태 업데이트
                    self.root.after(0, lambda i=idx, n=file_name: [
                        loading_popup.update_progress(
                            (i / len(selected_files)) * 100,
                            f"파일 처리 중 ({i+1}/{len(selected_files)}): {n}"
                        ),
                        self.log_text.insert(tk.END, f"\n파일 {n} 처리 중...\n"),
                        self.log_text.see(tk.END),
                        self.progress_bar.configure(value=i+1)
                    ])
                        
                    try:
                        # 번역 적용
                        result = self.translation_apply_manager.apply_translation(
                            file_path,
                            selected_langs,
                            self.record_date_var.get()
                        )
                            
                        if result["status"] == "success":
                            update_count = result["total_updated"]
                            total_updated += update_count
                            processed_count += 1
                                
                            # 로그 업데이트
                            self.root.after(0, lambda c=update_count: [
                                self.log_text.insert(tk.END, f"  {c}개 항목 업데이트 완료\n"),
                                self.log_text.see(tk.END)
                            ])
                        elif result["status"] == "info":
                            processed_count += 1
                            self.root.after(0, lambda m=result["message"]: [
                                self.log_text.insert(tk.END, f"  {m}\n"),
                                self.log_text.see(tk.END)
                            ])
                        else:  # 오류
                            error_count += 1
                            error_type = result.get("error_type", "unknown_error")
                            
                            # 문제 파일 분류해서 저장
                            if error_type in problem_files:
                                problem_files[error_type].append({
                                    "file_name": file_name,
                                    "message": result["message"]
                                })
                            
                            self.root.after(0, lambda m=result["message"]: [
                                self.log_text.insert(tk.END, f"  오류 발생: {m}\n"),
                                self.log_text.see(tk.END)
                            ])
                            
                    except Exception as e:
                        error_count += 1
                        error_msg = str(e)
                        problem_files["unknown_error"].append({
                            "file_name": file_name,
                            "message": error_msg
                        })
                        self.root.after(0, lambda: [
                            self.log_text.insert(tk.END, f"  오류 발생: {error_msg}\n"),
                            self.log_text.see(tk.END)
                        ])
                        
                # 작업 완료 처리 (메인 스레드에서) - 문제 파일 목록도 전달
                self.root.after(0, lambda: self.process_translation_apply_result(
                    total_updated, processed_count, error_count, loading_popup, problem_files))
                    
            except Exception as e:
                error_msg = str(e)
                self.root.after(0, lambda: [
                    loading_popup.close(),
                    self.log_text.insert(tk.END, f"\n작업 중 치명적 오류 발생: {error_msg}\n"),
                    self.status_label_apply.config(text="오류 발생"),
                    messagebox.showerror("오류", f"번역 적용 중 오류 발생: {error_msg}", parent=self.root)
                ])


        # 백그라운드 스레드 실행
        thread = threading.Thread(target=apply_translations)
        thread.daemon = True
        thread.start()

    
            
    def process_translation_apply_result(self, total_updated, processed_count, error_count, loading_popup, problem_files):
        """번역 적용 결과 처리"""
        loading_popup.close()
            
        # 작업 완료 메시지
        self.log_text.insert(tk.END, f"\n번역 적용 작업 완료!\n")
        self.log_text.insert(tk.END, f"파일 처리: {processed_count}/{len(self.original_files_list.get_checked_items())} (오류: {error_count})\n")
        self.log_text.insert(tk.END, f"총 {total_updated}개 항목이 업데이트되었습니다.\n")
            
        self.status_label_apply.config(text=f"번역 적용 완료 - {total_updated}개 항목")
        
        # 문제 파일 목록 생성
        problem_summary = []
        total_problem_files = 0
        
        if problem_files["external_links"]:
            files = [f["file_name"] for f in problem_files["external_links"]]
            problem_summary.append(f"🔗 외부 링크 오류 ({len(files)}개):\n   " + "\n   ".join(files))
            total_problem_files += len(files)
        
        if problem_files["permission_denied"]:
            files = [f["file_name"] for f in problem_files["permission_denied"]]
            problem_summary.append(f"🔒 접근 권한 오류 ({len(files)}개):\n   " + "\n   ".join(files))
            total_problem_files += len(files)
        
        if problem_files["file_corrupted"]:
            files = [f["file_name"] for f in problem_files["file_corrupted"]]
            problem_summary.append(f"💥 파일 손상 ({len(files)}개):\n   " + "\n   ".join(files))
            total_problem_files += len(files)
        
        if problem_files["file_not_found"]:
            files = [f["file_name"] for f in problem_files["file_not_found"]]
            problem_summary.append(f"❌ 파일 없음 ({len(files)}개):\n   " + "\n   ".join(files))
            total_problem_files += len(files)
        
        if problem_files["unknown_error"]:
            files = [f["file_name"] for f in problem_files["unknown_error"]]
            problem_summary.append(f"⚠️ 기타 오류 ({len(files)}개):\n   " + "\n   ".join(files))
            total_problem_files += len(files)
        
        # 기본 완료 메시지
        completion_msg = f"번역 적용 작업이 완료되었습니다.\n총 {total_updated}개 항목이 업데이트되었습니다."
        
        # 문제 파일이 있으면 추가 정보 표시
        if total_problem_files > 0:
            problem_detail = "\n\n⚠️ 처리하지 못한 파일들:\n\n" + "\n\n".join(problem_summary)
            completion_msg += problem_detail
            
            # 로그에도 문제 파일 목록 추가
            self.log_text.insert(tk.END, f"\n처리하지 못한 파일 ({total_problem_files}개):\n")
            for summary in problem_summary:
                self.log_text.insert(tk.END, f"{summary}\n")
        
        messagebox.showinfo("완료", completion_msg, parent=self.root)

        
    def setup_excel_split_tab(self):
        """엑셀 시트 분리 탭 설정"""
        # 변수 초기화
        self.excel_input_file_path = tk.StringVar()
        self.excel_output_dir_path = tk.StringVar()
        self.excel_output_dir_path.set(os.getcwd())  # 기본값: 현재 디렉토리
        
        # 메인 프레임
        main_frame = ttk.Frame(self.excel_split_frame, padding="10")
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
        button_frame.columnconfigure(1, weight=1)
        
        ttk.Button(button_frame, text="시트 분리 실행", command=self.start_excel_split_processing).grid(row=0, column=0, padx=5)
        
        # 리사이징 설정
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(3, weight=1)
    
    def browse_excel_input_file(self):
        """원본 엑셀 파일 선택"""
        file_path = filedialog.askopenfilename(
            title="원본 엑셀 파일 선택",
            filetypes=[("Excel 파일", "*.xlsx;*.xls"), ("모든 파일", "*.*")],
            parent=self.root
        )
        if file_path:
            self.excel_input_file_path.set(file_path)
            self.excel_log("원본 파일이 선택되었습니다: " + file_path)
            # 포커스를 다시 자동화 툴 창으로
            self.root.after(100, self.root.focus_force)
            self.root.after(100, self.root.lift)
    
    def browse_excel_output_dir(self):
        """출력 폴더 선택"""
        dir_path = filedialog.askdirectory(title="출력 폴더 선택", parent=self.root)
        if dir_path:
            self.excel_output_dir_path.set(dir_path)
            self.excel_log("출력 폴더가 선택되었습니다: " + dir_path)
            # 포커스를 다시 자동화 툴 창으로
            self.root.after(100, self.root.focus_force)
            self.root.after(100, self.root.lift)
    
    def excel_log(self, message):
        """엑셀 분리 로그 메시지 추가"""
        self.excel_log_text.insert(tk.END, message + "\n")
        self.excel_log_text.see(tk.END)
    
    def start_excel_split_processing(self):
        """엑셀 시트 분리 작업 시작"""
        input_file = self.excel_input_file_path.get().strip()
        output_dir = self.excel_output_dir_path.get().strip()
        
        if not input_file:
            messagebox.showerror("오류", "원본 엑셀 파일을 선택해주세요.", parent=self.root)
            return
        
        if not output_dir:
            messagebox.showerror("오류", "출력 폴더를 선택해주세요.", parent=self.root)
            return
        
        # 백그라운드 스레드에서 처리
        self.excel_progress['value'] = 0
        threading.Thread(target=self.split_excel_by_sheets, args=(input_file, output_dir), daemon=True).start()
    
    def split_excel_by_sheets(self, input_file, output_dir):
        """엑셀 시트별로 분리하는 메서드"""
        try:
            self.excel_log("처리 준비 중...")
            # COM 스레드 초기화
            pythoncom.CoInitialize()
            
            # Excel 애플리케이션 시작
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False  # 백그라운드에서 실행
            excel.DisplayAlerts = False  # 알림 끄기
            
            try:
                self.excel_log("엑셀 파일 열기 중...")
                # 원본 엑셀 파일 열기
                workbook = excel.Workbooks.Open(os.path.abspath(input_file))
                
                # 시트 개수 확인
                total_sheets = workbook.Sheets.Count
                self.excel_progress['maximum'] = total_sheets
                self.excel_log(f"총 {total_sheets}개의 시트를 발견했습니다.")
                
                # 각 시트에 대해 반복
                for idx in range(1, total_sheets + 1):
                    sheet = workbook.Sheets(idx)
                    sheet_name = sheet.Name
                    
                    try:
                        self.excel_log(f"시트 처리 중 ({idx}/{total_sheets}): {sheet_name}")
                        
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
                        
                        self.excel_log(f"생성된 파일: {safe_sheet_name}.xlsx")
                        
                        # 진행 상황 업데이트 (메인 스레드에서)
                        self.root.after(0, lambda i=idx: self.excel_progress.configure(value=i))
                        
                    except Exception as sheet_error:
                        self.excel_log(f"시트 '{sheet_name}' 처리 중 오류 발생: {str(sheet_error)}")
                
                self.excel_log("모든 시트가 처리되었습니다.")
                
            finally:
                # 원본 통합 문서 닫기
                workbook.Close(SaveChanges=False)
                # Excel 종료
                excel.Quit()
                
            self.root.after(0, lambda: messagebox.showinfo("완료", "모든 시트가 처리되었습니다.", parent=self.root))
            
        except Exception as e:
            error_msg = str(e)
            self.excel_log(f"오류 발생: {error_msg}")
            self.root.after(0, lambda error=error_msg: messagebox.showerror("오류", f"처리 중 오류가 발생했습니다: {error}", parent=self.root))
        finally:
            # COM 스레드 해제
            pythoncom.CoUninitialize()

       
    def setup_translation_apply_tab(self):
        # 번역 파일 선택 부분
        trans_frame = ttk.LabelFrame(self.translation_apply_frame, text="번역 DB 선택")
        trans_frame.pack(fill="x", padx=5, pady=5)
        
        ttk.Label(trans_frame, text="번역 DB:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.translation_db_var = tk.StringVar()
        ttk.Entry(trans_frame, textvariable=self.translation_db_var, width=50).grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        ttk.Button(trans_frame, text="찾아보기", 
                command=lambda: self.select_file(self.translation_db_var, "번역 DB 선택", [("DB 파일", "*.db")])).grid(row=0, column=2, padx=5, pady=5)
        
        ttk.Label(trans_frame, text="원본 폴더:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.original_folder_var = tk.StringVar()
        ttk.Entry(trans_frame, textvariable=self.original_folder_var, width=50).grid(row=1, column=1, padx=5, pady=5, sticky="ew")
        ttk.Button(trans_frame, text="찾아보기", 
                command=lambda: self.select_folder(self.original_folder_var, "원본 파일 폴더 선택")).grid(row=1, column=2, padx=5, pady=5)
        ttk.Button(trans_frame, text="파일 검색", 
                command=self.search_original_files).grid(row=1, column=3, padx=5, pady=5)
        
        trans_frame.columnconfigure(1, weight=1)
        
        # 파일 목록 표시
        files_frame = ttk.LabelFrame(self.translation_apply_frame, text="원본 파일 목록")
        files_frame.pack(fill="both", expand=True, padx=5, pady=5)
        
        self.original_files_list = ScrollableCheckList(files_frame, width=700, height=150)
        self.original_files_list.pack(fill="both", expand=True, padx=5, pady=5)

        
        # 옵션 설정
        options_frame = ttk.LabelFrame(self.translation_apply_frame, text="적용 옵션")
        options_frame.pack(fill="x", padx=5, pady=5)
        
        # 언어 선택 - 2행 3열로 배치
        self.apply_lang_vars = {}
        for i, lang in enumerate(self.available_languages):
            var = tk.BooleanVar(value=True if lang in ["CN", "TW"] else False)
            self.apply_lang_vars[lang] = var
            ttk.Checkbutton(options_frame, text=lang, variable=var).grid(
                row=i // 3, column=i % 3, padx=20, pady=5, sticky="w")
        
        # 언어 매핑 정보 추가
        ttk.Label(options_frame, text="언어 매핑: ZH → CN (자동 처리)", 
                font=("", 9, "italic")).grid(
            row=2, column=1, columnspan=2, padx=5, pady=1, sticky="w")

        # 번역 적용일 기록 옵션
        self.record_date_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(options_frame, text="번역 적용 표시 (#번역적용 컬럼)", 
                    variable=self.record_date_var).grid(
            row=2, column=0, columnspan=3, padx=5, pady=5, sticky="w")
        
        # 작업 실행 버튼
        action_frame = ttk.Frame(self.translation_apply_frame)
        action_frame.pack(fill="x", padx=5, pady=5)
        
        ttk.Button(action_frame, text="번역 적용", 
                command=self.apply_translation).pack(side="right", padx=5, pady=5)
        ttk.Button(action_frame, text="번역 DB 캐시 로드", 
                command=self.load_translation_cache).pack(side="right", padx=5, pady=5)
        
        # 로그 표시 영역
        log_frame = ttk.LabelFrame(self.translation_apply_frame, text="작업 로그")
        log_frame.pack(fill="both", expand=True, padx=5, pady=5)
        
        self.log_text = tk.Text(log_frame, wrap="word")
        scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        
        scrollbar.pack(side="right", fill="y")
        self.log_text.pack(fill="both", expand=True)
        
        # 상태와 진행 표시
        status_frame = ttk.Frame(self.translation_apply_frame)
        status_frame.pack(fill="x", padx=5, pady=5)
        
        self.status_label_apply = ttk.Label(status_frame, text="대기 중...")
        self.status_label_apply.pack(side="left", padx=5)
        
        self.progress_bar = ttk.Progressbar(status_frame, length=400, mode="determinate")
        self.progress_bar.pack(side="right", fill="x", expand=True, padx=5)

    
    def setup_string_sync_tab(self):
        """STRING 동기화 탭 설정"""
        try:
            self.string_sync_manager = StringSyncManager(self.string_sync_frame, self.root)
            self.string_sync_manager.pack(fill="both", expand=True)
        except ImportError:
            ttk.Label(self.string_sync_frame, text="STRING 동기화 모듈을 불러올 수 없습니다.").pack(pady=20)

    def setup_word_replacement_tab(self):
        """단어 치환 탭 설정"""
        try:
            self.word_replacement_manager = WordReplacementManager(self.word_replacement_frame, self.root)
            self.word_replacement_manager.pack(fill="both", expand=True)
        except ImportError:
            ttk.Label(self.word_replacement_frame, text="단어 치환 모듈을 불러올 수 없습니다.").pack(pady=20)


# 파일 마지막에 추가
if __name__ == "__main__":
    root = tk.Tk()
    app = TranslationAutomationTool(root)
    root.mainloop()