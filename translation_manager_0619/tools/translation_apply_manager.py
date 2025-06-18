import os
import time
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import pandas as pd
import sqlite3

class TranslationApplyManager:
    def __init__(self, parent_window=None):
        self.parent_ui = parent_ui
        self.translation_cache = {}
        self.translation_file_cache = {}
        self.translation_sheet_cache = {}
        self.duplicate_ids = {}
        self.kr_reverse_cache = {}
        
    def log_message(self, message):
        """UI의 로그 텍스트 영역에 메시지를 기록합니다."""
        if self.parent_ui and hasattr(self.parent_ui, 'log_text'):
            self.parent_ui.log_text.insert(tk.END, f"{message}\n")
            self.parent_ui.log_text.see(tk.END)
            self.parent_ui.update_idletasks()
        else:
            print(message)

    def load_translation_cache_from_excel(self, file_path, sheet_name):
        """엑셀 파일의 특정 시트에서 번역 데이터를 읽어 캐시를 생성합니다."""
        try:
            self.log_message(f"⚙️ 엑셀 파일 로딩 시작: {os.path.basename(file_path)} - 시트: {sheet_name}")
            header_row_index = self._find_header_row(file_path, sheet_name)
            if header_row_index is None:
                message = "지정한 시트의 1~6행에서 'STRING_ID' 컬럼을 찾을 수 없습니다."
                self.log_message(f"❌ {message}")
                return {"status": "error", "message": message}
            
            df = pd.read_excel(file_path, sheet_name=sheet_name, header=header_row_index, dtype=str)
            df.fillna('', inplace=True)

            df.columns = [str(col).lower() for col in df.columns]
            
            if 'string_id' not in df.columns:
                message = "엑셀 시트에 'string_id' 컬럼이 없습니다."
                self.log_message(f"❌ {message}")
                return {"status": "error", "message": message}

            self.build_cache_from_dataframe(df)

            # UI가 필요로 하는 모든 캐시 정보를 포함하여 반환
            return {
                "status": "success",
                "source_type": "Excel",
                "translation_cache": self.translation_cache,
                "translation_file_cache": self.translation_file_cache,
                "translation_sheet_cache": self.translation_sheet_cache,
                "duplicate_ids": self.duplicate_ids,
                "kr_reverse_cache": self.kr_reverse_cache,
                "id_count": len(self.translation_cache)
            }
        except Exception as e:
            self.log_message(f"❌ 엑셀 캐시 로딩 오류: {str(e)}")
            return {"status": "error", "message": str(e)}
        
    def _find_header_row(self, file_path, sheet_name):
        """엑셀 시트의 1~6행에서 'string_id'를 포함하는 헤더 행을 찾습니다."""
        for i in range(6):
            try:
                df_peek = pd.read_excel(file_path, sheet_name=sheet_name, header=i, nrows=0)
                if 'string_id' in [str(col).lower() for col in df_peek.columns]:
                    self.log_message(f"✅ 헤더 행 발견: {i + 1}번째 행")
                    return i
            except Exception:
                continue
        return None
    
    def build_cache_from_dataframe(self, df):
        """Pandas DataFrame으로부터 정교한 다중 캐시를 구축합니다."""
        self.translation_cache = {}
        self.translation_file_cache = {}
        self.translation_sheet_cache = {}
        self.duplicate_ids = {}
        self.kr_reverse_cache = {}
        
        self.log_message(f"🔧 데이터프레임으로부터 캐시 구축 시작: {len(df)}개 행")

        for _, row in df.iterrows():
            string_id = str(row.get('string_id', '')).strip()
            if not string_id:
                continue

            file_name = str(row.get('filename', row.get('file_name', ''))).strip()
            sheet_name = str(row.get('sheetname', row.get('sheet_name', ''))).strip()

            norm_file_name = file_name.lower()
            norm_sheet_name = sheet_name.lower()

            data = {
                "kr": str(row.get("kr", "")),
                "en": str(row.get("en", "")),
                "cn": str(row.get("cn", "")),
                "tw": str(row.get("tw", "")),
                "th": str(row.get("th", "")),
                "file_name": file_name,
                "sheet_name": sheet_name
            }
            
            if norm_file_name:
                self.translation_file_cache.setdefault(norm_file_name, {})[string_id] = data
            
            if norm_sheet_name:
                self.translation_sheet_cache.setdefault(norm_sheet_name, {})[string_id] = data
            
            self.translation_cache[string_id] = data

            if string_id not in self.duplicate_ids:
                self.duplicate_ids[string_id] = []
            self.duplicate_ids[string_id].append(file_name)

            kr_text = data["kr"].strip()
            if kr_text and kr_text not in self.kr_reverse_cache:
                kr_cache_data = data.copy()
                kr_cache_data['string_id'] = string_id
                self.kr_reverse_cache[kr_text] = kr_cache_data
        
        self.log_message(f"🔧 캐시 구성 완료 (ID: {len(self.translation_cache)}, 파일: {len(self.translation_file_cache)}, 시트: {len(self.translation_sheet_cache)}, KR역방향: {len(self.kr_reverse_cache)})")
        
             
    def find_string_id_position(self, worksheet):
        """STRING_ID 위치 찾기"""
        for row in range(2, 6):  # 2행부터 5행까지 검색
            for col in range(1, min(10, worksheet.max_column + 1)):  # 최대 10개 컬럼까지만 검색
                cell_value = worksheet.cell(row=row, column=col).value
                if isinstance(cell_value, str) and "STRING_ID" in cell_value.upper():
                    return col, row
                    
        # 1행도 검색
        for row in worksheet.iter_rows(min_row=1, max_row=1, max_col=5):
            for cell in row:
                if isinstance(cell.value, str) and "STRING_ID" in cell.value.upper():
                    return cell.column, cell.row
                    
        return None, None

    def find_language_columns(self, worksheet, header_row, langs):
        """언어 컬럼 위치 찾기"""
        if not header_row:
            return {}
            
        lang_cols = {}
        
        # 지정한 헤더 행에서만 검색
        for row in worksheet.iter_rows(min_row=header_row, max_row=header_row):
            for cell in row:
                if not cell.value:
                    continue
                    
                header_text = str(cell.value).strip()
                
                # 직접 매칭
                if header_text in langs:
                    lang_cols[header_text] = cell.column
                    
        return lang_cols

    def find_target_columns(self, worksheet, header_row, target_columns=None):
        """지정된 대상 컬럼들 찾기 (번역 적용 표시용)"""
        if not header_row or not target_columns:
            return {}
            
        found_columns = {}
        
        # 기본 대상 컬럼들 (기존 "#번역요청" 관련)
        default_targets = ["#번역요청", "#번역 요청"]
        
        # target_columns가 리스트가 아니면 리스트로 변환
        if isinstance(target_columns, str):
            target_columns = [target_columns]
        elif target_columns is None:
            target_columns = []
        
        # 모든 대상 컬럼 목록 생성
        all_targets = default_targets + target_columns
        
        for cell in worksheet[header_row]:
            if cell.value and isinstance(cell.value, str):
                cell_value = cell.value.strip()
                
                # 공백 제거 후 비교 (기존 로직)
                cell_value_no_space = cell_value.replace(" ", "")
                
                for target in all_targets:
                    target_no_space = target.replace(" ", "")
                    if cell_value_no_space == target_no_space or cell_value == target:
                        found_columns[target] = cell.column
                        break
                        
        return found_columns

    def find_translation_request_column(self, worksheet, header_row):
        """#번역요청 컬럼 찾기 (공백 무시) - 기존 호환성 유지"""
        if not header_row:
            return None
            
        for cell in worksheet[header_row]:
            if cell.value and isinstance(cell.value, str):
                # 공백 제거 후 비교
                if cell.value.replace(" ", "") in ["#번역요청", "#번역 요청"]:
                    return cell.column
                    
        return None


    def apply_translation(self, file_path, selected_langs, record_date=True, target_columns=None, smart_translation=True):
        """파일에 번역 적용 (우선순위: 파일명 > 시트명 > STRING_ID)
        
        Args:
            file_path: 처리할 엑셀 파일 경로
            selected_langs: 적용할 언어 목록
            record_date: 번역 적용일 기록 여부
            target_columns: 번역 적용 표시할 추가 컬럼들 (예: ["Change", "신규"])
            smart_translation: 스마트 번역 적용 여부 (KR 일치 시 다른 번역 활용)
        """
        if not self.translation_cache:
            return {
                "status": "error",
                "message": "번역 캐시가 로드되지 않았습니다. 먼저 캐시를 로드하세요.",
                "error_type": "cache_not_loaded"
            }
        
        # 현재 날짜 포맷
        today = time.strftime("%y.%m.%d")
        file_name = os.path.basename(file_path)
        
        # 🔧 파일 처리 시작 로그
        self.log_message(f"📁 파일 처리 시작: {file_name}")
        
        # target_columns 로그 출력
        if target_columns:
            self.log_message(f"  🎯 추가 대상 컬럼: {target_columns}")
        
        # 스마트 번역 기능 로그 출력
        if smart_translation:
            self.log_message(f"  🧠 스마트 번역 기능: 활성화")
        else:
            self.log_message(f"  🧠 스마트 번역 기능: 비활성화")
        
        try:
            # 현재 작업 중인 파일명 추출 (대소문자 무시)
            current_file_name = os.path.basename(file_path).lower()
            self.log_message(f"  🔍 정규화된 파일명: {current_file_name}")
            
            # 워크북 로드 - 구체적인 에러 처리 추가
            self.log_message(f"  📖 엑셀 파일 열기 시도...")
            try:
                workbook = load_workbook(file_path, data_only=True)  # 외부 링크 값만 가져오기
                self.log_message(f"  ✅ 엑셀 파일 열기 성공")
            except FileNotFoundError:
                self.log_message(f"  ❌ 파일을 찾을 수 없음")
                return {
                    "status": "error",
                    "message": "파일을 찾을 수 없습니다",
                    "error_type": "file_not_found"
                }
            except PermissionError:
                self.log_message(f"  ❌ 파일 접근 권한 없음")
                return {
                    "status": "error", 
                    "message": "파일 접근 권한이 없습니다. 파일이 다른 프로그램에서 열려있는지 확인하세요",
                    "error_type": "permission_denied"
                }
            except Exception as load_error:
                error_msg = str(load_error).lower()
                self.log_message(f"  ❌ 파일 열기 오류: {load_error}")
                if "external" in error_msg or "링크" in error_msg or "link" in error_msg:
                    return {
                        "status": "error",
                        "message": "외부 링크 오류가 발생했습니다. 파일의 외부 참조를 제거하거나 값으로 변환하세요",
                        "error_type": "external_links"
                    }
                elif "corrupt" in error_msg or "damaged" in error_msg or "zip" in error_msg or "invalid" in error_msg:
                    return {
                        "status": "error", 
                        "message": "파일이 손상되었거나 올바른 엑셀 형식이 아닙니다",
                        "error_type": "file_corrupted"
                    }
                else:
                    return {
                        "status": "error",
                        "message": f"파일 열기 실패: {load_error}",
                        "error_type": "unknown_error"
                    }
            
            # ===== 능동적 외부 링크 검사 추가 =====
            self.log_message(f"  🔗 외부 링크 검사 중...")
            external_links_found = self.check_external_links(workbook)
            if external_links_found:
                self.log_message(f"  ❌ 외부 링크 발견: {len(external_links_found)}개")
                for i, link in enumerate(external_links_found[:3]):
                    self.log_message(f"    - {link}")
                if len(external_links_found) > 3:
                    self.log_message(f"    ... 외 {len(external_links_found) - 3}개")
                workbook.close()
                return {
                    "status": "error",
                    "message": f"외부 링크가 발견되었습니다: {', '.join(external_links_found[:3])}{'...' if len(external_links_found) > 3 else ''}",
                    "error_type": "external_links"
                }
            else:
                self.log_message(f"  ✅ 외부 링크 없음")
            
            # String 시트 찾기 (대소문자 구분 없이)
            self.log_message(f"  📋 String 시트 검색 중...")
            string_sheets = [sheet for sheet in workbook.sheetnames 
                        if sheet.lower().startswith("string") and not sheet.startswith("#")]
            
            if not string_sheets:
                self.log_message(f"  ❌ String 시트를 찾을 수 없음")
                workbook.close()
                return {
                    "status": "error",
                    "message": "파일에 String 시트가 없습니다",
                    "error_type": "no_string_sheets"
                }
            
            self.log_message(f"  ✅ String 시트 발견: {len(string_sheets)}개 ({', '.join(string_sheets)})")
            
            total_updated = 0
            fill_green = PatternFill(start_color="DAF2D0", end_color="DAF2D0", fill_type="solid")
            
            # 검색 결과 통계
            file_match_count = 0
            sheet_match_count = 0
            id_only_match_count = 0
            
            # 원문 변경 케이스 추적
            kr_changed_items = []  # 원문이 변경된 항목들
            kr_changed_count = 0   # 원문 변경 카운트
            
            # 스마트 번역 적용 추적
            smart_applied_items = []  # 스마트 번역이 적용된 항목들
            smart_applied_count = 0   # 스마트 번역 적용 카운트
            
            # 각 시트 처리
            for sheet_name in string_sheets:
                self.log_message(f"  📝 시트 처리 시작: {sheet_name}")
                worksheet = workbook[sheet_name]
                
                # 시트명 정규화 (대소문자 무시)
                norm_sheet_name = sheet_name.lower()
                self.log_message(f"    🔍 정규화된 시트명: {norm_sheet_name}")
                
                # STRING_ID 위치 찾기
                self.log_message(f"    📍 STRING_ID 컬럼 검색 중...")
                string_id_col, header_row = self.find_string_id_position(worksheet)
                if not string_id_col or not header_row:
                    self.log_message(f"    ❌ STRING_ID 컬럼을 찾을 수 없음")
                    continue
                self.log_message(f"    ✅ STRING_ID 컬럼 발견: {string_id_col}열, {header_row}행")
                
                # 언어 컬럼 위치 찾기
                self.log_message(f"    🌐 언어 컬럼 검색 중: {selected_langs}")
                lang_cols = self.find_language_columns(worksheet, header_row, selected_langs)
                if not lang_cols:
                    self.log_message(f"    ❌ 언어 컬럼을 찾을 수 없음")
                    continue
                self.log_message(f"    ✅ 언어 컬럼 발견: {dict(lang_cols)}")
                
                # 🔧 대상 컬럼들 위치 찾기 (수정된 부분)
                apply_cols = {}
                if record_date:
                    self.log_message(f"    🔍 대상 컬럼 검색 중...")
                    apply_cols = self.find_target_columns(worksheet, header_row, target_columns)
                    
                    if apply_cols:
                        self.log_message(f"    ✅ 발견된 대상 컬럼: {apply_cols}")
                    else:
                        self.log_message(f"    ❌ 대상 컬럼이 없습니다. 번역 적용 표시를 건너뜁니다.")
                        # 🔧 컬럼이 없어도 번역은 계속 진행
                
                # 시트별 통계
                sheet_updated = 0
                sheet_file_match = 0
                sheet_sheet_match = 0
                sheet_id_match = 0
                
                # 중복 STRING_ID 추적
                duplicate_ids_in_sheet = set()
                
                # 데이터 행 개수 확인
                data_rows = worksheet.max_row - header_row
                self.log_message(f"    📊 처리할 데이터 행 수: {data_rows}개")
                
                # 🔧 캐시 디버깅 정보 (시트 시작할 때 한 번만)
                self.log_message(f"    🔍 캐시 디버깅 정보:")
                self.log_message(f"      - current_file_name: '{current_file_name}'")
                self.log_message(f"      - norm_sheet_name: '{norm_sheet_name}'")
                self.log_message(f"      - 파일명 캐시에 있는 키들: {list(self.translation_file_cache.keys())[:5]}...")
                self.log_message(f"      - 시트명 캐시에 있는 키들: {list(self.translation_sheet_cache.keys())[:5]}...")
                
                if current_file_name in self.translation_file_cache:
                    file_cache_ids = list(self.translation_file_cache[current_file_name].keys())
                    self.log_message(f"      - 파일명 '{current_file_name}' 캐시의 STRING_ID 수: {len(file_cache_ids)}")
                    self.log_message(f"      - 파일명 캐시의 첫 5개 ID: {file_cache_ids[:5]}")
                else:
                    self.log_message(f"      - 파일명 '{current_file_name}' 캐시에 없음")
                
                if norm_sheet_name in self.translation_sheet_cache:
                    sheet_cache_ids = list(self.translation_sheet_cache[norm_sheet_name].keys())
                    self.log_message(f"      - 시트명 '{norm_sheet_name}' 캐시의 STRING_ID 수: {len(sheet_cache_ids)}")
                    self.log_message(f"      - 시트명 캐시의 첫 5개 ID: {sheet_cache_ids[:5]}")
                else:
                    self.log_message(f"      - 시트명 '{norm_sheet_name}' 캐시에 없음")

                # 각 행 처리
                processed_rows = 0
                
                for row in range(header_row + 1, worksheet.max_row + 1):
                    string_id = worksheet.cell(row=row, column=string_id_col).value
                    if not string_id:
                        continue
                    
                    # 🔧 STRING_ID를 반드시 문자열로 변환 (핵심 수정!)
                    string_id = str(string_id).strip()
                    if not string_id:
                        continue
                    
                    processed_rows += 1
                    
                    # 🔧 진행 상황 로그 (100행마다)
                    if processed_rows % 100 == 0:
                        self.log_message(f"    📈 진행 상황: {processed_rows}/{data_rows}행 처리됨")
                    
                    # 중복 STRING_ID 확인
                    if string_id in self.duplicate_ids and len(self.duplicate_ids[string_id]) > 1:
                        duplicate_ids_in_sheet.add(string_id)
                    
                    # 번역 데이터 가져오기 (3단계 우선순위)
                    trans_data = None
                    match_type = "없음"
                    
                    # 🔧 상세 디버깅 (첫 5개 ID만)
                    debug_detail = processed_rows <= 5
                    
                    # 🔧 특정 테스트 ID에 대해서는 항상 상세 디버깅
                    test_ids = ['8004001', '4000001', '4000201']
                    is_test_id = string_id in test_ids
                    
                    if debug_detail or is_test_id:
                        self.log_message(f"      🔍 STRING_ID '{string_id}' 매칭 시도:")
                        self.log_message(f"        - current_file_name: '{current_file_name}'")
                        self.log_message(f"        - norm_sheet_name: '{norm_sheet_name}'")
                        
                        # 파일명 캐시 상세 확인
                        file_cache_exists = current_file_name in self.translation_file_cache
                        self.log_message(f"        - 파일명 캐시 존재: {file_cache_exists}")
                        
                        if file_cache_exists:
                            file_cache = self.translation_file_cache[current_file_name]
                            id_in_file_cache = string_id in file_cache
                            self.log_message(f"        - 파일명 캐시 내 ID 존재: {id_in_file_cache}")
                            self.log_message(f"        - 파일명 캐시 크기: {len(file_cache)}")
                            
                            # 파일 캐시의 첫 10개 ID 확인
                            cache_ids = list(file_cache.keys())[:10]
                            self.log_message(f"        - 파일명 캐시의 첫 10개 ID: {cache_ids}")
                        
                        # 시트명 캐시 상세 확인  
                        sheet_cache_exists = norm_sheet_name in self.translation_sheet_cache
                        self.log_message(f"        - 시트명 캐시 존재: {sheet_cache_exists}")
                        
                        if sheet_cache_exists:
                            sheet_cache = self.translation_sheet_cache[norm_sheet_name]
                            id_in_sheet_cache = string_id in sheet_cache
                            self.log_message(f"        - 시트명 캐시 내 ID 존재: {id_in_sheet_cache}")
                            self.log_message(f"        - 시트명 캐시 크기: {len(sheet_cache)}")
                        
                        # 전체 캐시 확인
                        id_in_global_cache = string_id in self.translation_cache
                        self.log_message(f"        - 전체 캐시 내 ID 존재: {id_in_global_cache}")
                        
                        # STRING_ID 타입 확인
                        self.log_message(f"        - STRING_ID 타입: {type(string_id)}")
                        self.log_message(f"        - STRING_ID 값: '{string_id}'")

                    # 1. 파일명 + STRING_ID 매칭 (최우선)
                    if (current_file_name in self.translation_file_cache and 
                        string_id in self.translation_file_cache[current_file_name]):
                        trans_data = self.translation_file_cache[current_file_name][string_id]
                        match_type = "파일명"
                        sheet_file_match += 1
                        if debug_detail or is_test_id:
                            self.log_message(f"        ✅ 파일명 매칭 성공: {string_id}")
                    
                    # 2. 시트명 + STRING_ID 매칭 (중간 우선순위)
                    elif (norm_sheet_name in self.translation_sheet_cache and 
                        string_id in self.translation_sheet_cache[norm_sheet_name]):
                        trans_data = self.translation_sheet_cache[norm_sheet_name][string_id]
                        match_type = "시트명"
                        sheet_sheet_match += 1
                        if debug_detail or is_test_id:
                            self.log_message(f"        ✅ 시트명 매칭 성공: {string_id}")
                    
                    # 3. STRING_ID만으로 매칭 (마지막 우선순위)
                    elif string_id in self.translation_cache:
                        trans_data = self.translation_cache[string_id]
                        match_type = "ID만"
                        sheet_id_match += 1
                        if debug_detail or is_test_id:
                            self.log_message(f"        ✅ ID만 매칭 성공: {string_id}")
                    
                    if not trans_data:
                        # 🔧 번역 데이터가 없는 경우 상세 로그
                        if debug_detail or is_test_id:
                            self.log_message(f"        ❌ 모든 매칭 실패: {string_id}")
                        elif processed_rows <= 10:
                            self.log_message(f"      ⚠️ 번역 데이터 없음: {string_id}")
                        continue
                    
                    row_updated = False
                    updated_langs = []
                    kr_changed = False  # 이 행에서 KR 원문이 변경되었는지 플래그
                    smart_applied = False  # 스마트 번역이 적용되었는지 플래그
                    
                    # 🔧 번역 데이터 상세 로그 (첫 3개만)
                    if debug_detail:
                        self.log_message(f"      📝 번역 데이터 내용: {string_id}")
                        for lang_key, lang_value in trans_data.items():
                            if lang_key in ['kr', 'en', 'cn', 'tw', 'th']:
                                self.log_message(f"        - {lang_key}: '{lang_value}' (타입: {type(lang_value)})")
                    
                    # 🔧 KR 원문 변경 여부 확인 및 스마트 번역 시도
                    current_kr_value = None
                    if 'kr' in lang_cols:
                        current_kr_value = worksheet.cell(row=row, column=lang_cols['kr']).value
                        cache_kr_value = trans_data.get('kr')
                        
                        # KR 값이 다른 경우 (원문 변경된 케이스)
                        if current_kr_value and cache_kr_value and str(current_kr_value).strip() != str(cache_kr_value).strip():
                            kr_changed = True
                            kr_changed_count += 1
                            kr_changed_items.append({
                                'string_id': string_id,
                                'sheet_name': sheet_name,
                                'current_kr': str(current_kr_value).strip(),
                                'cache_kr': str(cache_kr_value).strip(),
                                'match_type': match_type
                            })
                            
                            if debug_detail or is_test_id:
                                self.log_message(f"        🔄 KR 원문 변경 감지:")
                                self.log_message(f"          - 현재 KR: '{current_kr_value}'")
                                self.log_message(f"          - 캐시 KR: '{cache_kr_value}'")
                            
                            # 🧠 스마트 번역 시도 (KR이 다른 경우에만)
                            if smart_translation and current_kr_value:
                                current_kr_text = str(current_kr_value).strip()
                                if current_kr_text in self.kr_reverse_cache:
                                    # 현재 KR과 일치하는 다른 번역 데이터 발견!
                                    smart_trans_data = self.kr_reverse_cache[current_kr_text]
                                    trans_data = smart_trans_data  # 번역 데이터를 스마트 번역 데이터로 교체
                                    smart_applied = True
                                    smart_applied_count += 1
                                    smart_applied_items.append({
                                        'string_id': string_id,
                                        'sheet_name': sheet_name,
                                        'current_kr': current_kr_text,
                                        'original_match_type': match_type,
                                        'smart_source_string_id': smart_trans_data.get('string_id', 'Unknown')
                                    })
                                    
                                    if debug_detail or is_test_id:
                                        self.log_message(f"        🧠 스마트 번역 적용:")
                                        self.log_message(f"          - 일치 KR: '{current_kr_text}'")
                                        self.log_message(f"          - 소스 ID: {smart_trans_data.get('string_id', 'Unknown')}")
                    
                    # 각 언어별로 적용
                    for lang in selected_langs:
                        lang_lower = lang.lower()
                        
                        # 🔧 언어별 상세 로그 (첫 3개만)
                        if debug_detail:
                            self.log_message(f"      🌐 언어 처리: {lang} (소문자: {lang_lower})")
                            self.log_message(f"        - 언어 컬럼 존재: {lang in lang_cols}")
                            if lang in lang_cols:
                                self.log_message(f"        - 언어 컬럼 번호: {lang_cols[lang]}")
                            self.log_message(f"        - 번역 데이터 존재: {lang_lower in trans_data}")
                            if lang_lower in trans_data:
                                trans_value = trans_data[lang_lower]
                                self.log_message(f"        - 번역 값: '{trans_value}' (타입: {type(trans_value)}, 빈값여부: {not trans_value})")
                        
                        if lang in lang_cols and trans_data.get(lang_lower):
                            # 현재 값과 번역 값이 다른 경우에만 업데이트
                            current_value = worksheet.cell(row=row, column=lang_cols[lang]).value
                            trans_value = trans_data[lang_lower]
                            
                            # 🔧 값 비교 상세 로그 (첫 3개만)
                            if debug_detail:
                                self.log_message(f"        - 현재 값: '{current_value}' (타입: {type(current_value)})")
                                self.log_message(f"        - 번역 값: '{trans_value}' (타입: {type(trans_value)})")
                                self.log_message(f"        - 값이 다름: {current_value != trans_value}")
                                self.log_message(f"        - 번역 값 유효: {bool(trans_value)}")
                            
                            if trans_value and current_value != trans_value:
                                worksheet.cell(row=row, column=lang_cols[lang]).value = trans_value
                                worksheet.cell(row=row, column=lang_cols[lang]).fill = fill_green
                                row_updated = True
                                updated_langs.append(lang)
                                
                                if debug_detail:
                                    self.log_message(f"        ✅ 번역 적용됨: '{current_value}' → '{trans_value}'")
                            elif debug_detail:
                                if not trans_value:
                                    self.log_message(f"        ⚠️ 번역 값이 비어있음")
                                else:
                                    self.log_message(f"        ℹ️ 값이 동일해서 건너뜀")
                    
                    # 🔧 번역 적용일 기록 (찾은 모든 컬럼에 표시 - 수정된 부분)
                    if row_updated and record_date and apply_cols:
                        for target_name, col_num in apply_cols.items():
                            current_apply_val = worksheet.cell(row=row, column=col_num).value
                            if current_apply_val != "적용":
                                worksheet.cell(row=row, column=col_num).value = "적용"
                    
                    if row_updated:
                        sheet_updated += 1
                        # 🔧 첫 10개 업데이트에 대해서만 상세 로그
                        if sheet_updated <= 10:
                            kr_status = " (KR변경)" if kr_changed else ""
                            smart_status = " (스마트)" if smart_applied else ""
                            self.log_message(f"      🔄 번역 적용: {string_id} ({match_type} 매칭) - {', '.join(updated_langs)}{kr_status}{smart_status}")
                
                # 시트별 결과 통계 누적
                total_updated += sheet_updated
                file_match_count += sheet_file_match
                sheet_match_count += sheet_sheet_match
                id_only_match_count += sheet_id_match
                
                # 🔧 시트별 결과 로그
                self.log_message(f"    ✅ 시트 '{sheet_name}' 완료:")
                self.log_message(f"      - 처리된 행: {processed_rows}개")
                self.log_message(f"      - 업데이트된 행: {sheet_updated}개")
                self.log_message(f"      - 매칭 유형별: 파일명({sheet_file_match}) + 시트명({sheet_sheet_match}) + ID만({sheet_id_match})")
                if duplicate_ids_in_sheet:
                    self.log_message(f"      - 중복 ID: {len(duplicate_ids_in_sheet)}개")
            
            # 🔧 파일 전체 결과 로그
            self.log_message(f"  📊 파일 전체 결과:")
            self.log_message(f"    - 총 업데이트: {total_updated}개")
            self.log_message(f"    - 매칭 통계: 파일명({file_match_count}) + 시트명({sheet_match_count}) + ID만({id_only_match_count})")
            self.log_message(f"    - KR 원문 변경: {kr_changed_count}개")
            if smart_translation:
                self.log_message(f"    - 스마트 번역 적용: {smart_applied_count}개")
            
            # KR 변경 케이스가 있으면 상세 로그
            if kr_changed_items:
                self.log_message(f"  ⚠️ KR 원문 변경된 항목들 (새 번역 필요):")
                for item in kr_changed_items[:10]:  # 최대 10개만 표시
                    self.log_message(f"    - {item['string_id']} ({item['match_type']}): '{item['current_kr']}' ← '{item['cache_kr']}'")
                if len(kr_changed_items) > 10:
                    self.log_message(f"    ... 외 {len(kr_changed_items) - 10}개")
            
            # 스마트 번역 적용 케이스가 있으면 상세 로그
            if smart_applied_items:
                self.log_message(f"  🧠 스마트 번역 적용된 항목들:")
                for item in smart_applied_items[:10]:  # 최대 10개만 표시
                    self.log_message(f"    - {item['string_id']} ← {item['smart_source_string_id']}: '{item['current_kr']}'")
                if len(smart_applied_items) > 10:
                    self.log_message(f"    ... 외 {len(smart_applied_items) - 10}개")
            
            # 변경사항이 있으면 저장
            if total_updated > 0:
                self.log_message(f"  💾 파일 저장 중...")
                try:
                    workbook.save(file_path)
                    workbook.close()
                    self.log_message(f"  ✅ 파일 저장 완료")
                    self.log_message(f"🎉 파일 처리 완료: {file_name} (총 {total_updated}개 업데이트)")
                    return {
                        "status": "success",
                        "total_updated": total_updated,
                        "file_match_count": file_match_count,
                        "sheet_match_count": sheet_match_count,
                        "id_only_match_count": id_only_match_count,
                        "kr_changed_count": kr_changed_count,
                        "kr_changed_items": kr_changed_items,
                        "smart_applied_count": smart_applied_count,
                        "smart_applied_items": smart_applied_items
                    }
                except PermissionError:
                    self.log_message(f"  ❌ 파일 저장 권한 없음")
                    workbook.close()
                    return {
                        "status": "error",
                        "message": "파일 저장 권한이 없습니다. 파일이 다른 프로그램에서 열려있는지 확인하세요",
                        "error_type": "save_permission_denied"
                    }
                except Exception as save_error:
                    self.log_message(f"  ❌ 파일 저장 오류: {save_error}")
                    workbook.close()
                    return {
                        "status": "error",
                        "message": f"파일 저장 실패: {save_error}",
                        "error_type": "save_error"
                    }
            else:
                workbook.close()
                self.log_message(f"  ℹ️ 변경사항 없음")
                self.log_message(f"📝 파일 처리 완료: {file_name} (변경사항 없음)")
                return {
                    "status": "info",
                    "message": "변경사항 없음",
                    "kr_changed_count": kr_changed_count,
                    "kr_changed_items": kr_changed_items,
                    "smart_applied_count": smart_applied_count,
                    "smart_applied_items": smart_applied_items
                }
            
        except Exception as e:
            # 열려 있는 워크북 닫기 시도
            try:
                workbook.close()
            except:
                pass
            
            self.log_message(f"❌ 파일 처리 중 오류 발생: {file_name} - {str(e)}")
            return {
                "status": "error",
                "message": str(e),
                "error_type": "processing_error"
            }


    def check_external_links(self, workbook):
        """워크북에서 외부 링크 검사 (번역 도구용) - 검증된 최종 버전"""
        import re
        
        external_links = []
        
        # 외부 참조 패턴들 (검증된 버전)
        external_patterns = [
            r"'[^']*\.xl[sx]?[xm]?'!",  # '파일명.xlsx'! 또는 '경로\파일명.xlsx'!
            r'\[.*\.xl[sx]?[xm]?\]',    # [파일명.xlsx] 패턴
            r"'[A-Z]:[^']*\.xl[sx]?[xm]?'!", # 'C:\경로\파일명.xlsx'! 패턴  
            r'\\[^\\]*\.xl[sx]?[xm]?!', # \파일명.xlsx! 패턴
            r"=[^=]*'[A-Z]:[^']*'",     # =으로 시작하는 드라이브 경로
            r'\[\d+\]!',                # [숫자]! 패턴 (시트 참조)
        ]
        
        # #REF! 오류 패턴들 (검증된 버전)
        ref_error_patterns = [
            r'#REF!',                   # #REF! 오류
            r'OFFSET\(#REF!',          # OFFSET 함수에서 #REF! 오류
        ]
        
        try:
            # 방법 1: 워크북의 external_links 속성 확인
            if hasattr(workbook, 'external_links') and workbook.external_links:
                for link in workbook.external_links:
                    external_links.append(f"워크북_외부링크: {str(link)}")
            
            # 방법 2: 명명된 범위 검사 (가장 중요!) - 검증된 로직
            if hasattr(workbook, 'defined_names') and workbook.defined_names:
                # 딕셔너리 키로 접근 (검증된 방법)
                for name_key in workbook.defined_names.keys():
                    try:
                        defined_name = workbook.defined_names[name_key]
                        if hasattr(defined_name, 'value') and defined_name.value:
                            name_formula = str(defined_name.value)
                            
                            # #REF! 오류 우선 검사
                            ref_error_found = False
                            for ref_pattern in ref_error_patterns:
                                if re.search(ref_pattern, name_formula):
                                    external_links.append(f"명명된_범위_REF오류:{name_key} - {name_formula[:50]}")
                                    ref_error_found = True
                                    break
                            
                            # #REF! 오류가 없는 경우에만 외부 참조 패턴 검사
                            if not ref_error_found:
                                for pattern in external_patterns:
                                    if re.search(pattern, name_formula):
                                        external_links.append(f"명명된_범위_외부링크:{name_key} - {name_formula[:50]}")
                                        break
                    except Exception as e:
                        # 개별 명명된 범위 처리 중 오류가 발생해도 계속 진행
                        pass
            
            # 방법 3: 셀별 외부 참조 검사 (제한적으로)
            cell_count = 0
            for sheet_name in workbook.sheetnames:
                if cell_count >= 100:  # 번역 도구에서는 성능을 위해 더 제한적으로
                    break
                    
                worksheet = workbook[sheet_name]
                
                for row in worksheet.iter_rows():
                    for cell in row:
                        cell_count += 1
                        if cell_count > 100:
                            break
                            
                        # 공식이 있는 셀 검사
                        if cell.data_type == 'f' and cell.value:
                            formula = str(cell.value)
                            
                            # #REF! 오류 검사
                            for ref_pattern in ref_error_patterns:
                                if re.search(ref_pattern, formula):
                                    external_links.append(f"셀_REF오류:{sheet_name}!{cell.coordinate} - {formula[:50]}")
                                    break
                            else:
                                # 외부 참조 패턴 검사
                                for pattern in external_patterns:
                                    if re.search(pattern, formula):
                                        external_links.append(f"셀_외부링크:{sheet_name}!{cell.coordinate} - {formula[:50]}")
                                        break
                        
                        # #REF! 값 검사
                        elif cell.value and str(cell.value).startswith('#REF!'):
                            external_links.append(f"셀_REF값:{sheet_name}!{cell.coordinate} - {cell.value}")
                    
                    if cell_count > 100:
                        break
                        
        except Exception as e:
            # 외부 링크 검사 중 오류가 발생하면 무시하고 계속 진행
            pass
            
        return external_links[:10]  # 최대 10개만 반환

    def load_translation_cache_from_db(self, db_path):
        """데이터베이스에서 번역 데이터를 읽어 캐시를 생성합니다."""
        try:
            self.log_message(f"⚙️ DB 로딩 시작: {db_path}")
            conn = sqlite3.connect(db_path)
            cursor = conn.cursor()
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='translation_data'")
            if cursor.fetchone() is None:
                message = "'translation_data' 테이블이 DB에 없습니다."
                self.log_message(f"❌ {message}")
                return {"status": "error", "message": message}

            query = "SELECT * FROM translation_data"
            df = pd.read_sql_query(query, conn)
            conn.close()
            
            # 공통 캐시 구축 함수 호출
            self.build_cache_from_dataframe(df)
            
            # UI가 필요로 하는 모든 캐시 정보를 포함하여 반환
            return {
                "status": "success",
                "source_type": "DB",
                "translation_cache": self.translation_cache,
                "translation_file_cache": self.translation_file_cache,
                "translation_sheet_cache": self.translation_sheet_cache,
                "duplicate_ids": self.duplicate_ids,
                "kr_reverse_cache": self.kr_reverse_cache,
                "id_count": len(self.translation_cache)
            }
        except Exception as e:
            self.log_message(f"❌ 번역 DB 캐시 로딩 오류: {str(e)}")
            return {"status": "error", "message": str(e)}