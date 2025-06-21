# tools/excel_diff_manager.py

import os
import pandas as pd
import logging
import json

# utils.cache_utils 및 기타 필요한 유틸리티 임포트
from utils.cache_utils import load_cached_data, hash_paths, get_file_mtime
from utils.common_utils import get_columns_from_db_safe

class ExcelDiffManager:
    def __init__(self, parent_ui):
        self.parent_ui = parent_ui
        self.source_cache = {}
        self.target_cache = {}
        self.diff_results = []
        # 프로젝트 루트 디렉토리를 클래스 생성 시점에 결정
        self.root_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

    def log(self, message):
        """UI에 로그 메시지를 전달합니다."""
        if self.parent_ui and hasattr(self.parent_ui, 'status_var'):
            self.parent_ui.root.after(0, lambda: self.parent_ui.status_var.set(message))
        logging.info(message)
        print(message)

    def get_pk_from_excel_cache(self, db_name, sheet_name=None):
        """캐시와 설정 파일에서 PK 정보를 추출합니다."""
        try:
            db_name_lower = db_name.lower().split('@')[0].replace(".xlsx", "")

            # 1. 메모리에 로드된 캐시 확인
            for cache in [self.source_cache, self.target_cache]:
                for file_info in cache.values():
                    for sheet_name_in_cache, sheet_info in file_info.get("sheets", {}).items():
                        sheet_db_name = sheet_name_in_cache.lower().split('@')[0].replace(".xlsx", "")
                        if (sheet_name and sheet_name.lower() == sheet_name_in_cache.lower()) or (sheet_db_name == db_name_lower):
                            if sheet_info.get("pk"):
                                return sheet_info["pk"]

            # 2. Excel_Indent.json 파일 확인
            indent_path = os.path.join(self.root_dir, "Excel_Indent.json")
            if os.path.exists(indent_path):
                with open(indent_path, 'r', encoding='utf-8') as f:
                    indent_data = json.load(f)
                for key, value in indent_data.items():
                    if key.lower() == db_name_lower and "OrderBy" in value:
                        return [col.strip() for col in value["OrderBy"].split(",")]
            
            return []
        except Exception as e:
            self.log(f"PK 정보 추출 오류: {e}")
            return []

    def get_sheet_header_row_from_cache(self, file_path, sheet_name, cache):
        """캐시에서 시트의 헤더 행 번호를 가져옵니다."""
        # 이 함수는 원본 파일에 유지하거나, Manager로 옮겨서 사용할 수 있습니다.
        # 여기서는 Manager로 옮기는 예시를 보입니다.
        try:
            # 상대 경로로 변환하여 캐시 키로 사용
            rel_path = os.path.relpath(file_path, os.path.dirname(file_path)).replace("\\", "/") # 간단한 상대경로
            if rel_path in cache and sheet_name in cache[rel_path].get("sheets", {}):
                return cache[rel_path]["sheets"][sheet_name].get("header_row", 2)
            return 2 # 기본값
        except Exception as e:
            self.log(f"헤더 행 검색 오류: {e}")
            return 2

    def compare_excel_file(self, file_name, source_file, target_file):
        """엑셀 파일 직접 비교 로직"""
        file_diff = []
        try:
            with pd.ExcelFile(source_file) as source_xls, pd.ExcelFile(target_file) as target_xls:
                source_sheets = [s for s in source_xls.sheet_names if not s.startswith('#')]
                target_sheets = [s for s in target_xls.sheet_names if not s.startswith('#')]
                common_sheets = set(source_sheets) & set(target_sheets)

                for sheet in common_sheets:
                    source_header_row = self.get_sheet_header_row_from_cache(source_file, sheet, self.source_cache)
                    target_header_row = self.get_sheet_header_row_from_cache(target_file, sheet, self.target_cache)
                    
                    source_df = pd.read_excel(source_xls, sheet_name=sheet, header=source_header_row, engine='openpyxl')
                    target_df = pd.read_excel(target_xls, sheet_name=sheet, header=target_header_row, engine='openpyxl')

                    # #으로 시작하는 컬럼 및 행 제외
                    source_df = source_df[[c for c in source_df.columns if not str(c).startswith('#')]]
                    target_df = target_df[[c for c in target_df.columns if not str(c).startswith('#')]]
                    if not source_df.empty:
                        source_df = source_df[~source_df.iloc[:, 0].astype(str).str.startswith('#', na=False)]
                    if not target_df.empty:
                        target_df = target_df[~target_df.iloc[:, 0].astype(str).str.startswith('#', na=False)]
                    
                    # PK 컬럼 확인 및 비교 수행
                    db_name = sheet.split('@')[0].replace(".xlsx", "")
                    pk_columns = self.get_pk_from_excel_cache(db_name, sheet_name=sheet)
                    pk_columns = [col for col in pk_columns if col in source_df.columns and col in target_df.columns] if pk_columns else []

                    if pk_columns:
                        sheet_diff = self.compare_dataframes_pk(source_df, target_df, pk_columns, file_name, sheet)
                    else:
                        sheet_diff = self.compare_dataframes_position(source_df, target_df, file_name, sheet)
                    
                    file_diff.extend(sheet_diff)
            return file_diff
        except Exception as e:
            self.log(f"파일 비교 오류 ({file_name}): {e}")
            return []

    def compare_dataframes_pk(self, df_source, df_target, pk_columns, file_name, sheet_name):
        """PK를 기준으로 DataFrame 비교"""
        diff_results = []
        try:
            for col in pk_columns:
                df_source[col] = df_source[col].fillna('').astype(str)
                df_target[col] = df_target[col].fillna('').astype(str)

            if len(pk_columns) > 1:
                df_source.set_index(pk_columns, inplace=True)
                df_target.set_index(pk_columns, inplace=True)
            else:
                df_source.set_index(pk_columns[0], inplace=True)
                df_target.set_index(pk_columns[0], inplace=True)

            added_indices = df_target.index.difference(df_source.index)
            removed_indices = df_source.index.difference(df_target.index)
            common_indices = df_source.index.intersection(df_target.index)

            for idx in added_indices:
                row_data = df_target.loc[idx]
                for col, val in row_data.items():
                    if pd.notna(val):
                        diff_results.append((file_name, sheet_name, str(idx), col, None, val, "추가됨"))

            for idx in removed_indices:
                row_data = df_source.loc[idx]
                for col, val in row_data.items():
                    if pd.notna(val):
                        diff_results.append((file_name, sheet_name, str(idx), col, val, None, "삭제됨"))
            
            for idx in common_indices:
                source_row = df_source.loc[idx]
                target_row = df_target.loc[idx]
                if not source_row.equals(target_row):
                    for col in df_source.columns:
                        source_val = source_row[col]
                        target_val = target_row.get(col) # get으로 안전하게 접근
                        if str(source_val) != str(target_val):
                            diff_results.append((file_name, sheet_name, str(idx), col, source_val, target_val, "변경됨"))
        except Exception as e:
            self.log(f"PK 비교 오류 ({sheet_name}): {e}")
        return diff_results

    def compare_dataframes_position(self, df_source, df_target, file_name, sheet_name):
        """위치 기반으로 DataFrame 비교"""
        diff_results = []
        try:
            max_rows = max(len(df_source), len(df_target))
            for i in range(max_rows):
                row_id = f"행 {i+1}"
                is_added = i >= len(df_source)
                is_removed = i >= len(df_target)

                if is_added:
                    for col, val in df_target.iloc[i].items():
                        diff_results.append((file_name, sheet_name, row_id, col, None, val, "추가됨"))
                elif is_removed:
                    for col, val in df_source.iloc[i].items():
                        diff_results.append((file_name, sheet_name, row_id, col, val, None, "삭제됨"))
                else:
                    source_row = df_source.iloc[i]
                    target_row = df_target.iloc[i]
                    if not source_row.equals(target_row):
                        for col in df_source.columns:
                            source_val = source_row[col]
                            target_val = target_row.get(col)
                            if str(source_val) != str(target_val):
                                diff_results.append((file_name, sheet_name, row_id, col, source_val, target_val, "변경됨"))
        except Exception as e:
            self.log(f"위치 기반 비교 오류 ({sheet_name}): {e}")
        return diff_results

    def start_comparison_logic(self, source_path, target_path, loading_popup, selected_files=None):
        """파일 비교 작업을 수행하는 핵심 로직"""
        try:
            source_files = self.get_excel_files(source_path, selected_files)
            target_files = self.get_excel_files(target_path, selected_files)
            
            source_file_dict = {os.path.basename(f): f for f in source_files}
            target_file_dict = {os.path.basename(f): f for f in target_files}
            common_file_names = set(source_file_dict.keys()) & set(target_file_dict.keys())
            common_files = [(name, source_file_dict[name], target_file_dict[name]) for name in common_file_names]

            self.diff_results = []
            total_files = len(common_files)

            for idx, (file_name, source_file, target_file) in enumerate(common_files):
                self.parent_ui.root.after(0, lambda i=idx, n=file_name: loading_popup.update_message(f"파일 비교 중 ({i+1}/{total_files}): {n}"))
                file_diff = self.compare_excel_file(file_name, source_file, target_file)
                self.diff_results.extend(file_diff)
            
            self.parent_ui.root.after(0, lambda: self.parent_ui.update_results_ui(loading_popup))

        except Exception as e:
            self.parent_ui.root.after(0, lambda: self.parent_ui.show_error(str(e), loading_popup))

    def get_excel_files(self, path, selected_files=None):
        """경로에서 엑셀 파일 목록 가져오기"""
        result = []
        if os.path.isfile(path):
            if path.endswith(('.xlsx', '.xls')) and not os.path.basename(path).startswith('~$'):
                return [path]
            return []
        
        for root, _, files in os.walk(path):
            for file in files:
                if file.endswith(('.xlsx', '.xls')) and not file.startswith('~$'):
                    rel_path = os.path.relpath(os.path.join(root, file), path).replace('\\', '/')
                    if selected_files is None or rel_path in selected_files:
                        result.append(os.path.join(root, file))
        return result