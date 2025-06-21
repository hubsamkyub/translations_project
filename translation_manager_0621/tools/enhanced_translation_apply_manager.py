import os
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import pandas as pd
import sqlite3
import tkinter as tk
import signal
import win32com.client as pythoncom
import xlwings as xw
from collections import defaultdict
import hashlib
import json

class EnhancedTranslationApplyManager:
    def __init__(self, parent_window=None):
        self.parent_ui = parent_window
        
        # 기본 캐시
        self.translation_cache = {}
        self.kr_reverse_cache = {}

        # [수정] KR 기반 번역 충돌 데이터 수집용 딕셔너리
        # 구조: { "KR텍스트": [ {"cn": "번역1", "tw": "번역1", "count": 2, "data": 대표데이터}, ... ] }
        self.kr_translation_conflicts = defaultdict(list)
        
        # --- [신규] 영구적인 충돌 해결 DB 관리 ---
        self.db_dir = os.path.join(os.getcwd(), "user_data")
        os.makedirs(self.db_dir, exist_ok=True)
        self.resolution_db_path = os.path.join(self.db_dir, "user_resolutions.db")
        self.user_resolutions = {} # 사용자가 해결한 내용을 담을 딕셔너리
        self._init_resolution_db() # DB 파일 및 테이블 초기화
        self.load_user_resolutions() # 기존 해결 내용 로드
        # -----------------------------------------
        
        # 지원 언어 제한
        self.supported_languages = ["KR", "CN", "TW"]
        
        # DB 캐시 폴더
        self.cache_dir = os.path.join(os.getcwd(), "temp_db")
        os.makedirs(self.cache_dir, exist_ok=True)
            
        # 설정 파일에서 특수 컬럼명 불러오기
        self.SPECIAL_COLUMN_NAMES = self._load_config()

    def _load_config(self):
        """config.json 파일에서 설정을 불러옵니다."""
        config_path = 'config.json'
        default_columns = ["번역요청", "수정요청", "Help", "원문", "번역요청2", "번역추가", "번역적용"]
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                config = json.load(f)
                return config.get("SPECIAL_COLUMN_NAMES", default_columns)
        except FileNotFoundError:
            self.log_message(f"⚠️ 설정 파일({config_path})을 찾을 수 없어 기본값을 사용합니다.")
            return default_columns
        except Exception as e:
            self.log_message(f"❌ 설정 파일 로드 중 오류 발생: {e}. 기본값을 사용합니다.")
            return default_columns

    def _init_resolution_db(self):
        """[신규] 충돌 해결 내용을 저장할 DB를 초기화합니다."""
        try:
            conn = sqlite3.connect(self.resolution_db_path)
            cursor = conn.cursor()
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS resolved_translations (
                    kr_text TEXT PRIMARY KEY,
                    cn_text TEXT,
                    tw_text TEXT,
                    resolved_at TEXT
                )
            ''')
            conn.commit()
            conn.close()
        except Exception as e:
            self.log_message(f"❌ 사용자 해결 DB 초기화 오류: {e}")

    def load_user_resolutions(self):
        try:
            conn = sqlite3.connect(self.resolution_db_path)
            conn.row_factory = sqlite3.Row
            cursor = conn.cursor()
            cursor.execute("SELECT kr_text, cn_text, tw_text FROM resolved_translations")
            rows = cursor.fetchall()
            self.user_resolutions.clear() # 기존 내용을 비우고 새로 로드
            for row in rows:
                self.user_resolutions[row['kr_text']] = {'cn': row['cn_text'], 'tw': row['tw_text']}
            conn.close()
            if self.user_resolutions:
                self.log_message(f"✅ 기존에 해결한 번역 충돌 {len(self.user_resolutions)}건을 불러왔습니다.")
        except Exception as e:
            self.log_message(f"❌ 사용자 해결 내용 로드 오류: {e}")

    def _save_resolution_to_db(self, kr_text, selected_data):
        try:
            conn = sqlite3.connect(self.resolution_db_path)
            cursor = conn.cursor()
            resolved_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            cursor.execute('''
                INSERT OR REPLACE INTO resolved_translations (kr_text, cn_text, tw_text, resolved_at)
                VALUES (?, ?, ?, ?)
            ''', (kr_text, selected_data['cn'], selected_data['tw'], resolved_at))
            conn.commit()
            conn.close()
        except Exception as e:
            self.log_message(f"❌ 해결 내용 DB 저장 오류: {e}")

    # --- [신규] 해결 DB 관리용 함수들 ---
    def get_all_resolutions(self):
        """DB에서 모든 해결 내역을 가져옵니다."""
        try:
            conn = sqlite3.connect(self.resolution_db_path)
            conn.row_factory = sqlite3.Row
            cursor = conn.cursor()
            cursor.execute("SELECT kr_text, cn_text, tw_text, resolved_at FROM resolved_translations ORDER BY resolved_at DESC")
            rows = cursor.fetchall()
            conn.close()
            return [dict(row) for row in rows]
        except Exception as e:
            self.log_message(f"❌ 해결 내역 전체 조회 오류: {e}")
            return []

    def update_resolution(self, kr_text, new_cn, new_tw):
        """DB의 특정 해결 내역을 수정합니다."""
        try:
            conn = sqlite3.connect(self.resolution_db_path)
            cursor = conn.cursor()
            resolved_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            cursor.execute('''
                UPDATE resolved_translations 
                SET cn_text = ?, tw_text = ?, resolved_at = ?
                WHERE kr_text = ?
            ''', (new_cn, new_tw, resolved_at, kr_text))
            conn.commit()
            conn.close()
            self.load_user_resolutions() # 메모리에도 변경사항 반영
            return True
        except Exception as e:
            self.log_message(f"❌ 해결 내역 수정 오류: {e}")
            return False

    def delete_resolution(self, kr_text):
        """DB에서 특정 해결 내역을 삭제합니다."""
        try:
            conn = sqlite3.connect(self.resolution_db_path)
            cursor = conn.cursor()
            cursor.execute("DELETE FROM resolved_translations WHERE kr_text = ?", (kr_text,))
            conn.commit()
            conn.close()
            self.load_user_resolutions() # 메모리에도 변경사항 반영
            return True
        except Exception as e:
            self.log_message(f"❌ 해결 내역 삭제 오류: {e}")
            return False
                    
    def safe_strip(self, value):
        if value is None:
            return ""
        return str(value).strip()

    def safe_lower(self, value):
        if value is None:
            return ""
        return str(value).lower().strip()

    def log_message(self, message):
        if self.parent_ui and hasattr(self.parent_ui, 'log_text'):
            self.parent_ui.log_text.insert(tk.END, f"{message}\n")
            self.parent_ui.log_text.see(tk.END)
            self.parent_ui.update_idletasks()
        else:
            print(message)

    def _get_db_path(self, excel_path):
        file_hash = hashlib.md5(excel_path.encode()).hexdigest()
        return os.path.join(self.cache_dir, f"cache_{file_hash}.db")

    def _is_db_cache_valid(self, excel_path, db_path):
        if not os.path.exists(db_path):
            return False
        try:
            excel_mod_time = os.path.getmtime(excel_path)
            conn = sqlite3.connect(db_path)
            cursor = conn.cursor()
            cursor.execute("SELECT value FROM metadata WHERE key = 'source_mod_time'")
            stored_mod_time = cursor.fetchone()
            conn.close()
            if stored_mod_time and float(stored_mod_time[0]) == excel_mod_time:
                self.log_message("✅ 유효한 DB 캐시를 발견했습니다.")
                return True
            else:
                self.log_message("⚠️ 원본 파일이 변경되었습니다. DB 캐시를 재구성합니다.")
                return False
        except Exception as e:
            self.log_message(f"DB 캐시 유효성 검사 오류: {e}. 캐시를 재구성합니다.")
            return False

    def _build_db_from_excel(self, excel_path, db_path, progress_callback=None):
        try:
            self.log_message(f"⚙️ DB 캐시 구축 시작: {os.path.basename(excel_path)}")
            conn = sqlite3.connect(db_path)
            cursor = conn.cursor()
            cursor.execute('DROP TABLE IF EXISTS metadata')
            cursor.execute('DROP TABLE IF EXISTS translation_data')
            cursor.execute('CREATE TABLE metadata (key TEXT PRIMARY KEY, value TEXT)')
            cursor.execute('''
                CREATE TABLE translation_data (
                    string_id TEXT PRIMARY KEY, kr TEXT, cn TEXT, tw TEXT,
                    file_name TEXT, sheet_name TEXT, special_columns TEXT 
                )
            ''')
            wb = load_workbook(excel_path, read_only=True, data_only=True)
            all_sheets = wb.sheetnames
            for idx, sheet_name in enumerate(all_sheets):
                if progress_callback:
                    progress_callback((idx / len(all_sheets)) * 100, f"시트 처리 중 ({idx+1}/{len(all_sheets)}): {sheet_name}")
                if not sheet_name.lower().startswith("string") or sheet_name.startswith("#"):
                    continue
                ws = wb[sheet_name]
                header_map = {}
                header_row_idx = -1
                for i, row in enumerate(ws.iter_rows(min_row=1, max_row=10, values_only=True)):
                    cleaned_row = [self.safe_lower(str(cell)) for cell in row if cell is not None]
                    if 'string_id' in cleaned_row:
                        header_row_idx = i + 1
                        for col_idx, header_val in enumerate(row, 1):
                            if header_val:
                                header_map[self.safe_strip(str(header_val))] = col_idx
                        break
                if 'string_id' not in [self.safe_lower(k) for k in header_map.keys()]:
                    continue
                string_id_col_key = [k for k in header_map if self.safe_lower(k) == 'string_id'][0]
                string_id_col = header_map[string_id_col_key]
                rows_to_insert = []
                for row_data in ws.iter_rows(min_row=header_row_idx + 1, values_only=True):
                    if not row_data or len(row_data) < string_id_col or not row_data[string_id_col-1]:
                        continue
                    string_id = self.safe_strip(str(row_data[string_id_col-1]))
                    if not string_id:
                        continue
                    data = {}
                    special_data = {}
                    for header, col in header_map.items():
                        value = self.safe_strip(row_data[col-1]) if len(row_data) >= col else ""
                        normalized_header = header.lstrip('#').replace(' ', '')
                        if normalized_header in self.SPECIAL_COLUMN_NAMES:
                            special_data[normalized_header] = value
                        else:
                            data[self.safe_lower(header)] = value
                    special_columns_json = json.dumps(special_data, ensure_ascii=False)
                    kr_val = data.get("kr", "")
                    cn_val = data.get("cn", "")
                    tw_val = data.get("tw", "")
                    file_val = data.get("filename", data.get("file_name", os.path.basename(excel_path)))
                    sheet_val = data.get("sheetname", data.get("sheet_name", sheet_name))
                    rows_to_insert.append((string_id, kr_val, cn_val, tw_val, file_val, sheet_val, special_columns_json))
                cursor.executemany('INSERT OR REPLACE INTO translation_data (string_id, kr, cn, tw, file_name, sheet_name, special_columns) VALUES (?, ?, ?, ?, ?, ?, ?)', rows_to_insert)
            excel_mod_time = os.path.getmtime(excel_path)
            cursor.execute("INSERT INTO metadata (key, value) VALUES (?, ?)", ('source_mod_time', str(excel_mod_time)))
            conn.commit()
            conn.close()
            self.log_message("✅ DB 캐시 구축 완료.")
            return {"status": "success"}
        except Exception as e:
            self.log_message(f"❌ DB 캐시 구축 중 오류 발생: {e}")
            import traceback
            traceback.print_exc()
            return {"status": "error", "message": str(e)}

    def initiate_excel_caching(self, excel_path, force_rebuild=False, progress_callback=None):
        db_path = self._get_db_path(excel_path)
        if force_rebuild:
            self.log_message("ℹ️ 사용자의 요청으로 DB 캐시를 강제 재구성합니다.")
        elif self._is_db_cache_valid(excel_path, db_path):
            try:
                conn = sqlite3.connect(db_path)
                cursor = conn.cursor()
                cursor.execute("SELECT DISTINCT sheet_name FROM translation_data ORDER BY sheet_name")
                sheets = [row[0] for row in cursor.fetchall()]
                conn.close()
                return {"status": "success", "source_type": "DB_Cache", "sheets": sheets}
            except Exception as e:
                self.log_message(f"캐시된 시트 목록 로드 오류: {e}")
        build_result = self._build_db_from_excel(excel_path, db_path, progress_callback)
        if build_result["status"] == "success":
            try:
                conn = sqlite3.connect(db_path)
                cursor = conn.cursor()
                cursor.execute("SELECT DISTINCT sheet_name FROM translation_data ORDER BY sheet_name")
                sheets = [row[0] for row in cursor.fetchall()]
                conn.close()
                return {"status": "success", "source_type": "DB_Cache", "sheets": sheets}
            except Exception as e:
                self.log_message(f"재구축 후 시트 목록 로드 오류: {e}")
                return {"status": "error", "message": "DB 재구축 후 시트 목록을 가져오는 데 실패했습니다."}
        else:
            return build_result
        
    def load_translation_cache_from_excel_with_filter(self, excel_path, sheet_names, special_column_filter=None):
        self.log_message("⚙️ DB 캐시로부터 메모리 캐시 로딩 시작...")
        db_path = self._get_db_path(excel_path)
        if not os.path.exists(db_path):
            return {"status": "error", "message": "DB 캐시 파일이 없습니다."}

        try:
            self.translation_cache = {}
            self.kr_reverse_cache = {}
            self.kr_translation_conflicts = defaultdict(list)
            
            conn = sqlite3.connect(db_path)
            conn.row_factory = sqlite3.Row
            cursor = conn.cursor()
            query = "SELECT * FROM translation_data"
            params = []
            if sheet_names:
                query += f" WHERE sheet_name IN ({','.join('?' for _ in sheet_names)})"
                params.extend(sheet_names)
            cursor.execute(query, params)
            rows = cursor.fetchall()
            conn.close()

            temp_translation_cache = {}
            normalized_filter_column = ""
            if special_column_filter:
                raw_filter_column = special_column_filter.get('column_name', '')
                normalized_filter_column = raw_filter_column.lstrip('#').replace(' ', '')
                filter_value = special_column_filter.get('condition_value', '')

            for row in rows:
                data = dict(row)
                if special_column_filter:
                    special_columns = json.loads(data.get('special_columns', '{}'))
                    if normalized_filter_column in special_columns:
                        special_cell_value = special_columns[normalized_filter_column]
                        if self.safe_lower(filter_value) in self.safe_lower(special_cell_value):
                            temp_translation_cache[data["string_id"]] = data
                    else:
                        continue
                else:
                    temp_translation_cache[data["string_id"]] = data
            self.translation_cache = temp_translation_cache
            
            kr_candidates = defaultdict(lambda: defaultdict(list))
            for string_id, data in self.translation_cache.items():
                kr_text = self.safe_strip(data.get("kr", ""))
                if not kr_text: continue
                
                # [수정] 사용자가 이미 해결한 내용이 있다면, 그것으로 데이터를 통일
                if kr_text in self.user_resolutions:
                    cn_text = self.user_resolutions[kr_text]['cn']
                    tw_text = self.user_resolutions[kr_text]['tw']
                else:
                    cn_text = self.safe_strip(data.get("cn", ""))
                    tw_text = self.safe_strip(data.get("tw", ""))

                translation_pair = (cn_text, tw_text)
                kr_candidates[kr_text][translation_pair].append(data)

            for kr_text, pairs in kr_candidates.items():
                if len(pairs) == 1:
                    self.kr_reverse_cache[kr_text] = list(pairs.values())[0][0]
                else:
                     for (cn, tw), data_list in pairs.items():
                        self.kr_translation_conflicts[kr_text].append({
                            "cn": cn, "tw": tw, "count": len(data_list), "data": data_list[0]
                        })

            self.log_message(f"🔧 메모리 캐시 구성 완료:")
            self.log_message(f"  - 최종 로드된 STRING_ID: {len(self.translation_cache)}개")
            if special_column_filter:
                self.log_message(f"  - 필터 조건 적용: '{normalized_filter_column}' = '{filter_value}'")
            if self.kr_translation_conflicts:
                self.log_message(f"  - ⚠️ KR 기반 번역 충돌 감지: {len(self.kr_translation_conflicts)}개 텍스트")

            return {
                "status": "success", "source_type": "Excel_DB_Cache",
                "id_count": len(self.translation_cache),
                "filtered_count": len(self.translation_cache) if special_column_filter else 0,
                "conflict_count": len(self.kr_translation_conflicts)
            }
        except Exception as e:
            self.log_message(f"❌ DB 캐시 로딩 오류: {str(e)}")
            import traceback
            traceback.print_exc()
            return {"status": "error", "message": str(e)}

    def get_translation_conflicts(self):
        return self.kr_translation_conflicts
        
    def update_resolved_translations(self, resolutions):
        """[수정] 해결된 내용을 메모리와 영구 DB에 모두 저장"""
        resolved_count = 0
        for kr_text, selected_data in resolutions.items():
            # 1. 메모리 캐시 업데이트 (현재 세션용)
            if kr_text not in self.kr_reverse_cache:
                self.kr_reverse_cache[kr_text] = selected_data
                resolved_count += 1
            
            # 2. 영구 DB에 저장 (다음 세션용)
            self._save_resolution_to_db(kr_text, selected_data)

        self.log_message(f"✅ 번역 충돌 해결: {resolved_count}개 항목이 업데이트 및 영구 저장되었습니다.")
        for kr_text in resolutions:
            if kr_text in self.kr_translation_conflicts:
                del self.kr_translation_conflicts[kr_text]

    def apply_translation_with_filter_option(self, file_path, options):
        mode = options.get("mode", "id")
        selected_langs = options.get("selected_langs", [])
        record_date = options.get("record_date", True)
        use_filtered_data = options.get("use_filtered_data", False)

        kr_match_check = options.get("kr_match_check", True)
        kr_mismatch_delete = options.get("kr_mismatch_delete", False)
        kr_overwrite = options.get("kr_overwrite", False)
        kr_overwrite_on_kr_mode = options.get("kr_overwrite_on_kr_mode", False)
        allowed_statuses = options.get("allowed_statuses", [])
        allowed_statuses_lower = [status.lower() for status in allowed_statuses] if allowed_statuses else []

        active_cache = self.translation_cache
        cache_type = "특수필터링" if use_filtered_data else "전체"

        if not active_cache:
            return {"status": "error", "message": f"{cache_type} 번역 캐시가 로드되지 않았습니다."}
        if mode == 'kr' and not self.kr_reverse_cache:
            return {"status": "error", "message": "KR 기반 적용을 위한 역방향 캐시가 없습니다."}

        file_name = os.path.basename(file_path)

        option_summary = [f"{mode.upper()} 기반", f"{cache_type} 캐시"]
        if mode == 'id' and kr_match_check:
            option_summary.append("KR일치검사")
            if kr_mismatch_delete: option_summary.append("불일치시삭제")
            if kr_overwrite: option_summary.append("덮어쓰기")
        elif mode == 'kr' and kr_overwrite_on_kr_mode:
            option_summary.append("덮어쓰기")
        if allowed_statuses: option_summary.append(f"조건:{','.join(allowed_statuses)}")
        self.log_message(f"📁 {file_name} 처리시작 [{' | '.join(option_summary)}]")

        workbook = None
        try:
            workbook = load_workbook(file_path)
            string_sheets = [sheet for sheet in workbook.sheetnames if sheet.lower().startswith("string") and not sheet.startswith("#")]
            if not string_sheets:
                self.log_message(f"   ⚠️ String 시트 없음")
                return {"status": "info", "message": "파일에 String 시트가 없습니다"}

            file_modified = False
            results = defaultdict(int)
            overwritten_items = []

            fill_green = PatternFill(start_color="DAF2D0", end_color="DAF2D0", fill_type="solid")
            fill_orange = PatternFill(start_color="FFDDC1", end_color="FFDDC1", fill_type="solid")
            fill_blue = PatternFill(start_color="D0E7FF", end_color="D0E7FF", fill_type="solid")

            for sheet_name in string_sheets:
                worksheet = workbook[sheet_name]
                string_id_col, header_row = self.find_string_id_position(worksheet)
                if not string_id_col or not header_row:
                    self.log_message(f"   ⚠️ {sheet_name}: STRING_ID 컬럼 없음")
                    continue

                supported_langs = [lang for lang in selected_langs if lang in self.supported_languages]
                lang_cols = self.find_language_columns(worksheet, header_row, supported_langs + ['KR'])
                request_col_idx = self.find_target_columns(worksheet, header_row, ["#번역요청"]).get("#번역요청")
                sheet_stats = defaultdict(int)
                lang_apply_count = {lang: 0 for lang in supported_langs if lang != 'KR'}
                
                for row_idx in range(header_row + 1, worksheet.max_row + 1):
                    if allowed_statuses_lower and request_col_idx:
                        request_val = self.safe_lower(str(worksheet.cell(row=row_idx, column=request_col_idx).value or ''))
                        if request_val not in allowed_statuses_lower:
                            sheet_stats["conditional_skipped"] += 1
                            continue

                    trans_data = None
                    key_value = ''
                    if mode == 'id':
                        key_value = self.safe_strip(str(worksheet.cell(row=row_idx, column=string_id_col).value or ''))
                        if key_value: trans_data = active_cache.get(key_value)
                    else: 
                        if 'KR' in lang_cols:
                            key_value = self.safe_strip(str(worksheet.cell(row=row_idx, column=lang_cols['KR']).value or ''))
                            if key_value: trans_data = self.kr_reverse_cache.get(key_value)

                    if not key_value or not trans_data: continue

                    row_modified_this_iteration = False
                    if mode == 'id' and kr_match_check:
                        current_kr_val = self.safe_strip(str(worksheet.cell(row=row_idx, column=lang_cols['KR']).value or ''))
                        cache_kr_val = self.safe_strip(str(trans_data.get('kr', '')))
                        if current_kr_val != cache_kr_val:
                            if kr_mismatch_delete:
                                deleted_count = 0
                                for lang, col_idx in lang_cols.items():
                                    if lang != 'KR' and worksheet.cell(row=row_idx, column=col_idx).value:
                                        worksheet.cell(row=row_idx, column=col_idx).value = ""
                                        deleted_count += 1
                                        row_modified_this_iteration = True
                                if deleted_count > 0: sheet_stats["kr_mismatch_deleted"] += 1
                            else:
                                sheet_stats["kr_mismatch_skipped"] += 1
                            continue 

                    for lang in supported_langs:
                        if lang == 'KR': continue
                        lang_lower = lang.lower()
                        col_idx = lang_cols.get(lang)
                        if not col_idx: continue
                        cell = worksheet.cell(row=row_idx, column=col_idx)
                        current_val = self.safe_strip(str(cell.value or ''))
                        cached_val = self.safe_strip(str(trans_data.get(lang_lower, '')))
                        if cached_val and current_val != cached_val:
                            should_overwrite = (mode == 'id' and kr_match_check and kr_overwrite) or \
                                            (mode == 'kr' and kr_overwrite_on_kr_mode)
                            if should_overwrite:
                                original_text = cell.value
                                cell.value = cached_val
                                cell.fill = fill_orange
                                sheet_stats["overwritten"] += 1
                                lang_apply_count[lang] += 1
                                row_modified_this_iteration = True
                                overwritten_items.append({
                                    "file_name": file_name, "sheet_name": sheet_name,
                                    "string_id": trans_data.get('string_id', key_value),
                                    "language": lang, "kr_text": trans_data.get('kr', ''),
                                    "original_text": self.safe_strip(original_text),
                                    "overwritten_text": cached_val
                                })
                            elif not current_val:
                                cell.value = cached_val
                                cell.fill = fill_blue if use_filtered_data else fill_green
                                sheet_stats["updated"] += 1
                                lang_apply_count[lang] += 1
                                row_modified_this_iteration = True
                    if row_modified_this_iteration:
                        file_modified = True
                        if record_date and request_col_idx:
                            worksheet.cell(row=row_idx, column=request_col_idx).value = "특수필터적용" if use_filtered_data else "적용"

                if sheet_stats["updated"] > 0 or sheet_stats["overwritten"] > 0:
                    lang_details = [f"{lang}:{count}" for lang, count in lang_apply_count.items() if count > 0]
                    log_parts = []
                    if sheet_stats["updated"] > 0: log_parts.append(f"신규:{sheet_stats['updated']}")
                    if sheet_stats["overwritten"] > 0: log_parts.append(f"덮어씀:{sheet_stats['overwritten']}")
                    if lang_details: log_parts.append(f"[{', '.join(lang_details)}]")
                    self.log_message(f"   ✅ {sheet_name}: {' | '.join(log_parts)}")
                else:
                    skip_reasons = []
                    if sheet_stats["conditional_skipped"] > 0: skip_reasons.append(f"조건불일치:{sheet_stats['conditional_skipped']}")
                    if sheet_stats["kr_mismatch_skipped"] > 0: skip_reasons.append(f"KR불일치:{sheet_stats['kr_mismatch_skipped']}")
                    if sheet_stats["kr_mismatch_deleted"] > 0: skip_reasons.append(f"KR불일치삭제:{sheet_stats['kr_mismatch_deleted']}")
                    if skip_reasons: self.log_message(f"   ⚠️ {sheet_name}: 적용없음 ({' | '.join(skip_reasons)})")
                    else: self.log_message(f"   ⚠️ {sheet_name}: 적용없음 (번역데이터 없음)")

                for key, val in sheet_stats.items():
                    results[f"total_{key}"] += val

            if file_modified:
                self.log_message(f"   💾 변경사항 저장 중...")
                workbook.save(file_path)
                summary_parts = []
                if results["total_updated"] > 0: summary_parts.append(f"신규 {results['total_updated']}개")
                if results["total_overwritten"] > 0: summary_parts.append(f"덮어씀 {results['total_overwritten']}개")
                total_applied = results["total_updated"] + results["total_overwritten"]
                cache_info = f"({cache_type}캐시)" if use_filtered_data else ""
                self.log_message(f"   ✅ {file_name} 완료: {' | '.join(summary_parts)} (총 {total_applied}개 적용) {cache_info}")
            else:
                skip_summary = []
                if results["total_conditional_skipped"] > 0: skip_summary.append(f"조건 {results['total_conditional_skipped']}개")
                if results["total_kr_mismatch_skipped"] > 0: skip_summary.append(f"KR불일치 {results['total_kr_mismatch_skipped']}개")
                if skip_summary: self.log_message(f"   ⚠️ {file_name} 완료: 변경없음 ({' | '.join(skip_summary)} 건너뜀)")
                else: self.log_message(f"   ⚠️ {file_name} 완료: 변경없음 (번역 데이터 없음)")

            final_results = {key: val for key, val in results.items()}
            return {"status": "success", **final_results, "overwritten_items": overwritten_items}
        except Exception as e:
            self.log_message(f"   ❌ {file_name} 오류: {str(e)}")
            import traceback
            traceback.print_exc()
            return {"status": "error", "message": str(e), "error_type": "processing_error"}
        finally:
            if workbook:
                workbook.close()
 
    def find_string_id_position(self, worksheet):
        for row in range(2, 6):
            for col in range(1, min(10, worksheet.max_column + 1)):
                cell_value = worksheet.cell(row=row, column=col).value
                if cell_value and isinstance(cell_value, str):
                    if "string_id" in self.safe_lower(cell_value):
                        return col, row
        for row in worksheet.iter_rows(min_row=1, max_row=1, max_col=5):
            for cell in row:
                if cell.value and isinstance(cell.value, str):
                    if "string_id" in self.safe_lower(cell.value):
                        return cell.column, cell.row
        return None, None

    def find_language_columns(self, worksheet, header_row, langs):
        if not header_row: return {}
        lang_cols = {}
        langs_upper = [lang.upper() for lang in langs]
        for row in worksheet.iter_rows(min_row=header_row, max_row=header_row):
            for cell in row:
                if not cell.value: continue
                header_text = self.safe_strip(cell.value).upper()
                if header_text in langs_upper:
                    lang_cols[header_text] = cell.column
        return lang_cols

    def find_target_columns(self, worksheet, header_row, target_columns=None):
        if not header_row: return {}
        found_columns = {}
        all_targets = ["#번역요청"]
        if target_columns: all_targets.extend(target_columns)
        all_targets = list(set(all_targets))
        for cell in worksheet[header_row]:
            if cell.value and isinstance(cell.value, str):
                cell_value_clean = self.safe_strip(cell.value).lower()
                for target in all_targets:
                    if cell_value_clean == target.lower():
                        found_columns[target] = cell.column
                        break
        return found_columns

    def detect_special_column_in_excel(self, excel_path, target_column_name):
        db_path = self._get_db_path(excel_path)
        if not os.path.exists(db_path):
            return {"status": "error", "message": "DB 캐시가 생성되지 않았습니다."}
        self.log_message(f"⚙️ [특수컬럼감지] DB 캐시에서 '{target_column_name}' 분석 시작...")
        normalized_target_column = target_column_name.lstrip('#').replace(' ', '')
        self.log_message(f"   (정규화된 검색어: '{normalized_target_column}')")
        try:
            conn = sqlite3.connect(db_path)
            cursor = conn.cursor()
            cursor.execute("SELECT special_columns, sheet_name FROM translation_data")
            rows = cursor.fetchall()
            conn.close()
            values_count = defaultdict(int)
            found_in_sheets = set()
            non_empty_rows = 0
            for special_json, sheet_name in rows:
                if not special_json: continue
                special_data = json.loads(special_json)
                if normalized_target_column in special_data:
                    value = self.safe_strip(special_data[normalized_target_column])
                    if value:
                        values_count[value] += 1
                        found_in_sheets.add(sheet_name)
                        non_empty_rows += 1
            if not values_count:
                self.log_message(f"⚠️ 특수 컬럼 '{target_column_name}'에 대한 데이터를 찾을 수 없습니다.")
                return {"status": "success", "detected_info": {}}
            most_common = sorted(values_count.items(), key=lambda x: x[1], reverse=True)[:5]
            analysis_result = {
                'non_empty_rows': non_empty_rows, 'unique_values': dict(values_count),
                'most_common': most_common, 'found_in_sheets': list(found_in_sheets)
            }
            self.log_message(f"✅ 특수 컬럼 '{target_column_name}' 분석 완료:")
            self.log_message(f"  - 발견된 시트: {len(found_in_sheets)}개")
            self.log_message(f"  - 데이터 항목: {non_empty_rows}개 / 고유값: {len(values_count)}개")
            self.log_message(f"  - 최빈값: {most_common}")
            return {"status": "success", "detected_info": analysis_result}
        except Exception as e:
            self.log_message(f"❌ 특수 컬럼 분석 오류: {e}")
            import traceback
            traceback.print_exc()
            return {"status": "error", "message": str(e)}