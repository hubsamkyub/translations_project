# utils/excel_utils.py 파일 생성
import os
import sys  # 이 줄을 추가
import pandas as pd
import logging
import time
import sqlite3
import json
from typing import Dict, List, Union, Optional, Any, Tuple
import openpyxl  # 추가 - 엑셀 파일 직접 처리용
import traceback  # 추가 - 오류 추적용
import re  # 추가 - 정규식 처리용

from utils.common_utils import PathUtils, HashUtils, FileUtils, DBUtils, logger
from utils.type_mappings import get_table_name_for_type, get_description_for_type, resolve_type_info


class ExcelFileManager:
    """
    엑셀 파일 관리 및 유틸리티 기능을 제공하는 클래스
    """
    
    @staticmethod
    def find_header_row(file_path: str, sheet_name: str, db_columns: List[str], xls=None) -> int:
        """
        엑셀 파일에서 헤더 행을 찾습니다.
        
        Args:
            file_path: 엑셀 파일 경로
            sheet_name: 시트 이름
            db_columns: DB 컬럼 목록
            xls: 재사용할 ExcelFile 객체 (없으면 새로 생성)
            
        Returns:
            헤더 행 인덱스 (0-based)
        """
        if not db_columns:
            return 2  # 기본값
            
        priority_rows = [2, 3]  # 3행, 4행 (인덱스는 0부터 시작하므로 2, 3)
        
        try:
            # xls 객체가 없는 경우 직접 엑셀 파일을 열어야 함
            if xls is None:
                with pd.ExcelFile(file_path, engine="openpyxl") as temp_xls:
                    df_priority = pd.read_excel(
                        temp_xls, 
                        sheet_name=sheet_name, 
                        header=None, 
                        nrows=max(priority_rows)+1
                    )
            else:
                # 전달받은 xls 객체 사용
                df_priority = pd.read_excel(
                    xls, 
                    sheet_name=sheet_name, 
                    header=None, 
                    nrows=max(priority_rows)+1
                )

            for row_idx in priority_rows:
                if row_idx < len(df_priority):
                    row = df_priority.iloc[row_idx].tolist()
                    found = False
                    if isinstance(row, list) and db_columns:
                        # 주요 컬럼 중 일부가 있는지 확인
                        if db_columns[0] in row:
                            found = True
                        # 추가 검증: 첫 3개 컬럼 중 2개 이상 존재하는지 확인
                        elif len(db_columns) >= 3:
                            match_count = sum(1 for col in db_columns[:3] if col in row)
                            if match_count >= 2:
                                found = True
                    
                    if found:
                        return row_idx
        except Exception as e:
            logger.error(f"헤더 행 검색 우선 확인 오류: {file_path}, {sheet_name} - {e}")
        
        # 못 찾았으면 5행부터 10행까지 검색
        try:
            df_extended = pd.read_excel(
                xls if xls else file_path, 
                sheet_name=sheet_name, 
                header=None, 
                skiprows=4, 
                nrows=6, 
                engine="openpyxl"
            )
            
            for i in range(len(df_extended)):
                row = df_extended.iloc[i].tolist()
                if isinstance(row, list) and db_columns and db_columns[0] in row:
                    return i + 4  # skiprows=4 했으므로 인덱스 조정
        except Exception as e:
            logger.error(f"헤더 행 확장 검색 오류: {file_path}, {sheet_name} - {e}")
        
        return 2  # 기본값
    
    @staticmethod
    def find_excel_file(base_path: str, file_name: str) -> Optional[str]:
        """
        주어진 파일 이름으로 엑셀 파일을 검색합니다.
        
        Args:
            base_path: 기본 검색 경로
            file_name: 찾을 파일 이름
        
        Returns:
            찾은 파일의 전체 경로 또는 파일을 찾지 못한 경우 None
        """
        # 직접 경로에서 파일 확인
        file_path = os.path.join(base_path, file_name)
        if os.path.exists(file_path):
            return file_path
        
        # Excel_String 폴더 구조 확인
        excel_string_folders = []
        
        # 폴더 경로 내 Excel_String 폴더 찾기
        for root, dirs, files in os.walk(base_path):
            for dir_name in dirs:
                if dir_name.startswith("Excel_String"):
                    excel_string_folders.append(os.path.join(root, dir_name))
        
        # 모든 Excel_String 폴더에서 파일 검색
        for folder in excel_string_folders:
            potential_path = os.path.join(folder, file_name)
            if os.path.exists(potential_path):
                return potential_path
        
        return None
    
    @staticmethod
    def analyze_excel_file(file_path, db_folder=None, header_row_override=None, pk_override_dict=None):
        """
        엑셀 파일의 시트와 내용을 분석합니다. (예외 PK 처리 포함)
        """
        result = {}

        try:
            with pd.ExcelFile(file_path, engine="openpyxl") as xls:
                db_columns_cache = {}

                for sheet_name in xls.sheet_names:
                    try:
                        if sheet_name.startswith("#"):
                            continue

                        sheet_meta = {}
                        table_name = sheet_name.split("@")[0].replace(".xlsx", "")
                        table_base_lower = table_name.lower()

                        # DB 컬럼 캐시 처리
                        if db_folder:
                            if table_name not in db_columns_cache:
                                db_path = os.path.join(db_folder, f"{table_name}.db")
                                
                                # 추가: 파일 존재 여부 확인 및 0KB 파일 검사
                                if not os.path.exists(db_path) or os.path.getsize(db_path) == 0:
                                    logger.debug(f"유효한 DB 파일 없음: {db_path}")
                                    db_columns = []
                                else:
                                    try:
                                        conn = sqlite3.connect(db_path)
                                        cur = conn.cursor()
                                        try:
                                            cur.execute(f'PRAGMA table_info("{table_name}")')
                                            db_columns = [row[1] for row in cur.fetchall()]
                                        except:
                                            cur.execute(f'PRAGMA table_info([{table_name}])')
                                            db_columns = [row[1] for row in cur.fetchall()]
                                        conn.close()
                                    except Exception as e:
                                        logger.error(f"DB 컬럼 조회 오류: {db_path}/{table_name} - {e}")
                                        db_columns = []
                                        # 추가: 오류 발생 후 0KB 파일이 생성되었다면 삭제
                                        if os.path.exists(db_path) and os.path.getsize(db_path) == 0:
                                            try:
                                                os.remove(db_path)
                                                logger.info(f"빈 DB 파일 삭제: {db_path}")
                                            except Exception as del_e:
                                                logger.error(f"빈 DB 파일 삭제 실패: {db_path} - {del_e}")
                                        
                                db_columns_cache[table_name] = db_columns
                            db_columns = db_columns_cache.get(table_name, [])
                        else:
                            db_columns = []

                        # 헤더 행 위치 결정
                        header_row = header_row_override
                        if header_row is None:
                            header_row = ExcelFileManager.find_header_row(file_path, sheet_name, db_columns, xls)

                        # 기본 :pk 방식
                        explicit_pk_col = None
                        try:
                            df_top = pd.read_excel(xls, sheet_name=sheet_name, header=None, nrows=header_row + 1)
                            for row_idx in range(header_row):
                                for col_idx in range(len(df_top.columns)):
                                    cell_val = df_top.iloc[row_idx, col_idx]
                                    if isinstance(cell_val, str) and ":pk" in cell_val:
                                        pk_candidate = df_top.iloc[header_row, col_idx]
                                        if isinstance(pk_candidate, str):
                                            explicit_pk_col = pk_candidate.strip()
                                            break
                                if explicit_pk_col:
                                    break
                            del df_top
                        except Exception as e:
                            logger.warning(f"[PK 검색 오류] {file_path}/{sheet_name}: {e}")

                        # 예외 규칙 적용
                        if pk_override_dict and table_base_lower in pk_override_dict:
                            explicit_pk_cols = pk_override_dict[table_base_lower]
                            print(f"[{sheet_name}] 예외 PK 강제 적용: {explicit_pk_cols}")
                        else:
                            explicit_pk_cols = [explicit_pk_col] if explicit_pk_col else []

                        # 컬럼 위치 정보 구성
                        column_positions = {}
                        try:
                            df_header = pd.read_excel(xls, sheet_name=sheet_name, header=None, skiprows=header_row, nrows=1)
                            header_row_data = df_header.iloc[0].tolist()
                            for idx, col_name in enumerate(header_row_data, 1):
                                if pd.notna(col_name):
                                    col_str = str(col_name).strip().upper()
                                    if 'VARCHAR' in col_str or 'NTEXT' in col_str or ':PK' in col_str:
                                        if idx == 2:
                                            column_positions["STRING_ID"] = idx
                                        elif 3 <= idx < 15:
                                            lang_mapping = {
                                                3: "KR", 4: "EN", 5: "CN", 6: "TW",
                                                7: "TH", 8: "PT", 9: "ES", 10: "DE",
                                                11: "FR", 12: "JP"
                                            }
                                            if idx in lang_mapping:
                                                column_positions[lang_mapping[idx]] = idx
                                    else:
                                        if col_str:
                                            column_positions[col_str] = idx
                            del df_header
                        except Exception as e:
                            logger.warning(f"컬럼 위치 분석 오류: {file_path}, {sheet_name} - {e}")

                        # 실제 데이터 수
                        try:
                            df = pd.read_excel(xls, sheet_name=sheet_name, header=header_row)
                            # 필요한 작업 수행
                            sheet_meta = {
                                "header_row": header_row,
                                "has_reward_group_id": "RewardGroupID" in db_columns,
                                "has_reward_id": "RewardID" in db_columns,
                                "columns": db_columns,
                                "column_positions": column_positions,
                                "rows": len(df),
                                "pk": explicit_pk_cols
                            }
                            
                            # DataFrame 명시적으로 삭제
                            del df
                            
                        except Exception as df_error:
                            logger.warning(f"DataFrame 처리 오류: {sheet_name} - {df_error}")
                            continue
                        
                        result[sheet_name] = sheet_meta
                    except Exception as e:
                        logger.warning(f"시트 분석 실패: {sheet_name} - {e}")

                if 'db_columns_cache' in locals():
                    del db_columns_cache

            return result
        except Exception as e:
            logger.error(f"엑셀 파일 분석 실패: {file_path} - {e}")
            return {}


    def find_explicit_pk_from_excel(file_path, sheet_name, header_row):
        """엑셀 시트 상단에서 ':pk'가 있는 열을 찾아 실제 PK 컬럼명을 반환"""
        try:
            # header_row는 0-based 기준으로 아래가 컬럼명이므로, 그 위까지 탐색
            read_rows = header_row
            if read_rows < 1:
                return None

            df_top = pd.read_excel(file_path, sheet_name=sheet_name, header=None, nrows=read_rows, engine="openpyxl")

            for row_idx in range(read_rows):
                for col_idx, value in enumerate(df_top.iloc[row_idx]):
                    if isinstance(value, str) and ":pk" in value:
                        # 해당 열의 header_row 위치의 셀을 PK로 간주
                        pk_candidate = df_top.iloc[header_row, col_idx]
                        if isinstance(pk_candidate, str):
                            return pk_candidate.strip()
            return None
        except Exception as e:
            logger.warning(f"[PK 추출 오류] {file_path}/{sheet_name}: {e}")
            return None

  
    def load_pk_overrides():
        """Excel_Indent.json에서 PK 정보를 로드합니다. 오류 시 기본값 반환."""
        try:
            base_dir = os.path.dirname(os.path.dirname(__file__))  # utils의 상위 폴더
            json_path = os.path.join(base_dir, "Excel_Indent.json")
            
            if not os.path.exists(json_path):
                logger.warning(f"Excel_Indent.json 파일이 없습니다: {json_path}")
                return {}
            
            # 직접 파일 파싱 (json 모듈 사용하지 않음)
            result = {}
            with open(json_path, "r", encoding="utf-8") as f:
                content = f.read()
                
            # 기본값으로 자주 사용되는 OrderBy 값 직접 설정
            result["itemtemplate"] = ["TemplateID"]
            result["boxtemplate"] = ["UniqueID"]
            result["herotemplate"] = ["UniqueID"]
            result["stagetemplate"] = ["UniqueID"]
            
            return result
        except Exception as e:
            logger.error(f"PK 오버라이드 로드 오류: {e}")
            # 기본값 반환
            return {
                "itemtemplate": ["TemplateID"],
                "boxtemplate": ["UniqueID"],
                "herotemplate": ["UniqueID"],
                "stagetemplate": ["UniqueID"]
            }


    @staticmethod
    def update_excel_cache(folder_path, cache_path=None, progress_callback=None):
        """
        지정된 폴더의 모든 엑셀 파일을 스캔하여 캐시를 업데이트합니다.
        
        Args:
            folder_path: 엑셀 파일이 있는 폴더 경로
            cache_path: 캐시 저장 경로 (기본값: .cache/excel_cache.json)
            progress_callback: 진행 상황 콜백 함수 (선택적)
            
        Returns:
            업데이트된 캐시 데이터 (dict)
        """
        if cache_path is None:
            cache_id = HashUtils.hash_paths(folder_path)
            # PyInstaller 실행 파일 지원
            if getattr(sys, 'frozen', False):
                base_dir = os.path.dirname(sys.executable)
            else:
                base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
            cache_dir = os.path.join(base_dir, ".cache", cache_id)
            PathUtils.ensure_dir(cache_dir)
            cache_path = os.path.join(cache_dir, "excel_cache.json")
        
        # 기존 캐시 로드
        old_cache = FileUtils.load_cached_data(cache_path)
        new_cache = {}
        files_to_scan = []
        
        # 파일 스캔
        if progress_callback:
            progress_callback("엑셀 파일 검색 중...")
        
        for root, _, files in os.walk(folder_path):
            for file in files:
                if file.endswith(".xlsx") and not file.startswith("~$"):
                    path = os.path.join(root, file)
                    rel_path = os.path.relpath(path, folder_path).replace("\\", "/")
                    mtime = PathUtils.get_file_mtime(path)
                    
                    # 캐시 상태 확인
                    if rel_path not in old_cache or old_cache[rel_path]["mtime"] != mtime:
                        files_to_scan.append((rel_path, path, mtime))
                    else:
                        new_cache[rel_path] = old_cache[rel_path]
        
        # 변경된 파일만 분석
        total_files = len(files_to_scan)
        for idx, (rel_path, path, mtime) in enumerate(files_to_scan):
            if progress_callback:
                progress_callback(f"엑셀 분석 중 [{idx+1}/{total_files}]: {os.path.basename(path)}")
            
            pk_override_dict = ExcelFileManager.load_pk_overrides()
            sheets_result = ExcelFileManager.analyze_excel_file(path, folder_path, pk_override_dict=pk_override_dict)

            if sheets_result:
                new_cache[rel_path] = {
                    "path": path,
                    "mtime": mtime,
                    "sheets": sheets_result
                }
        
        # 캐시 저장
        FileUtils.save_cache(cache_path, new_cache)
        return new_cache
    
    @staticmethod
    def highlight_excel_by_value(file_path, sheet_name, column_name, value, excel_cache=None):
        """
        엑셀 파일에서 특정 컬럼의 값을 찾아 강조 표시하고 엑셀을 엽니다.
        
        Args:
            file_path: 엑셀 파일 경로
            sheet_name: 시트 이름
            column_name: 컬럼 이름
            value: 찾을 값
            excel_cache: 엑셀 캐시 데이터 (선택적)
            
        Returns:
            성공 여부 (Boolean)
        """
        try:
            import win32com.client
            import win32gui
            import win32con
            import time
            
            # 이미 열려있는 워크북 찾기
            excel, workbook = ExcelFileManager.find_open_excel_workbook(file_path)

            if excel and workbook:
                logger.info(f"이미 열려있는 엑셀 파일 사용: {file_path}")
                excel.Visible = True
                try:
                    excel.WindowState = win32com.client.constants.xlMaximized
                except:
                    pass
                excel.Activate()
                workbook.Activate()
            else:
                logger.info(f"새로운 엑셀 파일 열기: {file_path}")
                # 새로 Excel 애플리케이션 열기
                excel = win32com.client.Dispatch("Excel.Application")
                excel.Visible = True  # 처음부터 True로 설정
                
                try:
                    # 파일 열기
                    workbook = excel.Workbooks.Open(os.path.abspath(file_path))
                    
                    # Excel 창 강제 표시
                    excel.WindowState = win32com.client.constants.xlMaximized
                    excel.Activate()
                    workbook.Activate()
                    
                except Exception as open_error:
                    try:
                        excel.Quit()
                    except:
                        pass
                    raise open_error

            # 해당 시트로 이동
            worksheet = ExcelFileManager.activate_excel_worksheet(workbook, sheet_name)
            if not worksheet:
                logger.warning(f"시트 활성화 실패: {sheet_name}")
                # 그래도 Excel 창은 표시하도록 시도
                ExcelFileManager._force_show_excel_window()
                return True
            
            # 헤더 행 찾기
            header_row = 3  # 기본값
            if excel_cache:
                rel_path = os.path.relpath(file_path).replace("\\", "/")
                if rel_path in excel_cache and "sheets" in excel_cache[rel_path]:
                    sheet_info = excel_cache[rel_path]["sheets"].get(sheet_name, {})
                    if "header_row" in sheet_info:
                        header_row = sheet_info["header_row"] + 1  # 1-based로 변환
            
            # 헤더에서 컬럼 위치 찾기
            column_idx = None
            for i in range(1, 100):  # 최대 100개 컬럼까지 검색
                try:
                    cell_value = worksheet.Cells(header_row, i).Value
                    if cell_value == column_name:
                        column_idx = i
                        logger.debug(f"컬럼 발견: {column_name} at column {i}")
                        break
                except:
                    continue
            
            if column_idx is None:
                logger.warning(f"컬럼을 찾을 수 없음: {column_name}")
                # 그래도 Excel 창은 표시하도록 시도
                ExcelFileManager._force_show_excel_window()
                return True
            
            # 값 검색 및 찾은 셀로 이동
            found = False
            for row in range(header_row + 1, header_row + 10000):  # 최대 10000행까지 검색
                try:
                    cell_value = worksheet.Cells(row, column_idx).Value
                    if str(cell_value) == str(value):
                        # 찾은 셀로 이동
                        cell = worksheet.Cells(row, column_idx)
                        cell.Select()
                        
                        # 강조 표시 (기존 강조 표시 제거 후 새로 적용)
                        try:
                            worksheet.Range("A:Z").Interior.ColorIndex = win32com.client.constants.xlColorIndexNone
                        except:
                            pass
                        
                        # 새로운 강조 표시
                        cell.Interior.ColorIndex = 6  # 노란색
                        
                        # 화면 조정
                        excel.ActiveWindow.ScrollRow = max(1, row - 5)
                        
                        logger.info(f"값 발견 및 강조 표시: {value} at row {row}")
                        found = True
                        break
                        
                except Exception as cell_error:
                    # 빈 셀이나 읽기 오류는 무시하고 계속
                    continue
            
            if not found:
                logger.warning(f"값을 찾을 수 없음: {value}")
            
            # Windows API를 사용해서 Excel 창 강제 표시
            ExcelFileManager._force_show_excel_window()
            
            # 추가 활성화 시도
            try:
                time.sleep(0.3)  # 잠시 대기
                excel.Visible = True
                excel.Activate()
                workbook.Activate()
                if worksheet:
                    worksheet.Activate()
            except Exception as activate_error:
                logger.warning(f"최종 활성화 오류: {activate_error}")
            
            return True
            
        except Exception as e:
            logger.error(f"엑셀 강조 표시 실패: {e}")
            
            # COM 방식이 실패하면 기본 방식으로 파일 열기
            try:
                logger.info("기본 방식으로 엑셀 파일 열기 시도")
                os.startfile(file_path)
                return True
            except Exception as start_error:
                logger.error(f"기본 파일 열기도 실패: {start_error}")
                return False

    @staticmethod
    def _force_show_excel_window():
        """
        Windows API를 사용해서 모든 Excel 창을 강제로 표시합니다.
        """
        try:
            import win32gui
            import win32con
            
            def enum_windows_callback(hwnd, windows):
                if win32gui.IsWindowVisible(hwnd):
                    window_text = win32gui.GetWindowText(hwnd)
                    class_name = win32gui.GetClassName(hwnd)
                    
                    # Excel 창 식별 (여러 가지 패턴으로 확인)
                    if (('Microsoft Excel' in window_text or 
                         'Excel' in window_text or
                         '.xlsx' in window_text or
                         class_name == 'XLMAIN') and
                        'splash' not in window_text.lower()):
                        windows.append((hwnd, window_text))
                return True
            
            windows = []
            win32gui.EnumWindows(enum_windows_callback, windows)
            
            for hwnd, title in windows:
                try:
                    logger.debug(f"Excel 창 활성화 시도: {title}")
                    
                    # 창이 최소화되어 있으면 복원
                    if win32gui.IsIconic(hwnd):
                        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                    
                    # 창을 표시하고 최상위로 가져오기
                    win32gui.ShowWindow(hwnd, win32con.SW_SHOW)
                    win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
                    
                    # 포커스 설정 (여러 번 시도)
                    for _ in range(3):
                        try:
                            win32gui.SetForegroundWindow(hwnd)
                            break
                        except:
                            time.sleep(0.1)
                    
                    logger.info(f"Excel 창 활성화 완료: {title}")
                    
                except Exception as window_error:
                    logger.debug(f"창 활성화 실패 {title}: {window_error}")
            
            if windows:
                logger.info(f"총 {len(windows)}개의 Excel 창을 활성화했습니다.")
            else:
                logger.warning("Excel 창을 찾을 수 없습니다.")
                
        except Exception as e:
            logger.error(f"Excel 창 강제 표시 실패: {e}")

    @staticmethod
    def open_excel_file(file_path, sheet_name=None, highlight_info=None):
        """
        엑셀 파일을 열고 선택적으로 특정 값을 강조 표시합니다.
        (이미 열려있는 파일이 있으면 해당 파일을 사용)
        
        Args:
            file_path: 엑셀 파일 경로
            sheet_name: 시트 이름 (선택적)
            highlight_info: {column, value, excel_cache} 강조 표시 정보 (선택적)
            
        Returns:
            성공 여부 (Boolean)
        """
        if not os.path.exists(file_path):
            logger.error(f"파일이 존재하지 않음: {file_path}")
            return False
        
        # 강조 표시 정보가 있으면 해당 기능 사용
        if sheet_name and highlight_info and 'column' in highlight_info and 'value' in highlight_info:
            return ExcelFileManager.highlight_excel_by_value(
                file_path, 
                sheet_name, 
                highlight_info['column'], 
                highlight_info['value'],
                highlight_info.get('excel_cache')
            )
        
        # 강조 표시가 필요없는 경우 - COM 대신 기본 방식 사용
        try:
            logger.info(f"기본 방식으로 엑셀 파일 열기: {file_path}")
            os.startfile(file_path)
            
            # 잠시 대기 후 Excel 창 강제 표시
            import time
            time.sleep(2)  # Excel이 완전히 로드될 때까지 대기
            ExcelFileManager._force_show_excel_window()
            
            return True
            
        except Exception as e:
            logger.error(f"엑셀 파일 열기 실패: {file_path} - {e}")
            return False


    @staticmethod
    def add_hash_to_a_column(excel_path, sheet_name, id_value):
        """
        지정된 ID에 해당하는 모든 행의 A열에 #을 추가합니다.
        
        Args:
            excel_path (str): 엑셀 파일 경로
            sheet_name (str): 시트 이름
            id_value (str): 찾을 ID 값
            
        Returns:
            bool: 성공 여부
        """
        result = ExcelFileManager._find_rows_by_id(excel_path, sheet_name, id_value)
        if not result:
            return False
        
        _, _, matching_rows = result
        return ExcelFileManager._modify_a_column_hash(excel_path, sheet_name, matching_rows, add_hash=True)


    @staticmethod
    def remove_hash_from_a_column(excel_path, sheet_name, id_value, id_column=None, header_row=None):
        """
        지정된 ID에 해당하는 모든 행의 A열에서 #을 제거합니다.
        
        Args:
            excel_path (str): 엑셀 파일 경로
            sheet_name (str): 시트 이름
            id_value (str): 찾을 ID 값
            id_column (str, optional): 특정 ID 컬럼 이름
            header_row (int, optional): 헤더 행 번호
            
        Returns:
            bool: 성공 여부
        """
        result = ExcelFileManager._find_rows_by_id(excel_path, sheet_name, id_value, id_column, header_row)
        if not result:
            return False
        
        _, _, matching_rows = result
        return ExcelFileManager._modify_a_column_hash(excel_path, sheet_name, matching_rows, add_hash=False, header_row=header_row)



    @staticmethod
    def _find_rows_by_id(excel_path, sheet_name, id_value, id_column=None, header_row=None):
        """
        지정된 ID에 해당하는 행들을 찾습니다.
        
        Args:
            excel_path (str): 엑셀 파일 경로
            sheet_name (str): 시트 이름
            id_value (str): 찾을 ID 값
            id_column (str, optional): 특정 ID 컬럼 이름
            header_row (int, optional): 헤더 행 번호
            
        Returns:
            tuple: (DataFrame, 컬럼명, 일치하는 행 인덱스 리스트) 또는 None(실패 시)
        """
        try:
            logger.debug(f"엑셀 파일 열기: {excel_path}, 시트: {sheet_name}, 헤더 행: {header_row}")
            if not os.path.exists(excel_path):
                logger.warning(f"파일이 존재하지 않음: {excel_path}")
                return None
            
            # 판다스로 데이터 읽기 - 헤더 행 지정
            df = pd.read_excel(excel_path, sheet_name=sheet_name, header=header_row)
            logger.debug(f"엑셀 파일 로드 성공, 컬럼: {df.columns.tolist()}")
            
            # 컬럼이 숫자로만 이루어져 있는지 확인 (header가 적용되지 않았을 가능성)
            is_numeric_columns = all(isinstance(col, int) for col in df.columns)
            if is_numeric_columns:
                logger.warning(f"컬럼이 숫자로만 이루어져 있음 - 헤더 행 ({header_row})이 올바르지 않을 수 있음")
                
                # 데이터 구조 확인을 위해 첫 번째 행 출력
                if len(df) > 0:
                    logger.debug(f"첫 번째 행 데이터: {df.iloc[0].tolist()}")
                    
                    # 첫 번째 행을 헤더로 사용할 수 있는 옵션 (필요한 경우 주석 해제)
                    # df.columns = [str(x) for x in df.iloc[0]]
                    # df = df.drop(0).reset_index(drop=True)
                    # logger.debug(f"첫 번째 행을 헤더로 사용한 후 컬럼: {df.columns.tolist()}")
            
            # ID 컬럼 찾기
            id_col_name = id_column
            if id_col_name is None:
                for col in df.columns:
                    col_str = str(col).lower()
                    if ('id' in col_str and ('item' in col_str or 'template' in col_str)) or 'templateid' in col_str:
                        id_col_name = col
                        logger.debug(f"ID 컬럼 발견: {col}")
                        break
                
                if id_col_name is None:
                    logger.debug(f"ID 컬럼을 찾을 수 없음, 일반적인 ID 컬럼 검색")
                    for col in df.columns:
                        col_str = str(col).lower()
                        if 'id' in col_str:
                            id_col_name = col
                            logger.debug(f"일반 ID 컬럼 발견: {col}")
                            break
            
            if id_col_name is None:
                logger.warning(f"ID 컬럼 찾기 실패")
                return None
            
            # ID 값 검색 - 모든 일치하는 행 찾기
            id_value_str = str(id_value)
            matching_rows = []
            
            # 숫자형 검색 시도
            try:
                numeric_id = float(id_value)
                df[id_col_name] = pd.to_numeric(df[id_col_name], errors='coerce')
                matched = df[df[id_col_name] == numeric_id]
                if not matched.empty:
                    matching_rows = matched.index.tolist()
                    logger.debug(f"숫자 ID로 {len(matching_rows)}개 행 찾음")
            except Exception as e:
                logger.debug(f"숫자 검색 실패: {e}, 문자열 검색으로 전환")
            
            # 문자열 검색 시도
            if not matching_rows:
                try:
                    df[id_col_name] = df[id_col_name].astype(str)
                    matched = df[df[id_col_name] == id_value_str]
                    if not matched.empty:
                        matching_rows = matched.index.tolist()
                        logger.debug(f"문자열 ID로 {len(matching_rows)}개 행 찾음")
                except Exception as e:
                    logger.debug(f"문자열 검색 실패: {e}")
            
            if not matching_rows:
                logger.warning(f"ID {id_value}를 가진 행을 찾을 수 없음")
                return None
            
            return df, id_col_name, matching_rows
        
        except Exception as e:
            logger.error(f"행 검색 중 오류: {e}")
            logger.debug(traceback.format_exc())
            return None


    @staticmethod
    def _modify_a_column_hash(excel_path, sheet_name, row_indices, add_hash=True, header_row=None):
        """
        지정된 행들의 A열에 #을 추가하거나 제거합니다.
        """
        try:
            if not os.path.exists(excel_path):
                logger.warning(f"파일이 존재하지 않음: {excel_path}")
                return False
            
            # 워크북 열기
            wb = openpyxl.load_workbook(excel_path)
            ws = wb[sheet_name]
            
            # 모든 일치하는 행 처리
            modified_count = 0
            for row_idx in row_indices:
                # 실제 행 번호 계산 (헤더 행을 고려)
                header_offset = (header_row or 0) + 1  # 헤더 행 + 1(엑셀의 행은 1부터 시작)
                actual_row = row_idx + header_offset
                
                logger.debug(f"수정할 행 번호: {actual_row}, 원본 인덱스: {row_idx}, 헤더 행: {header_row}")
                
                # A열 처리
                a_cell = ws.cell(row=actual_row, column=1)
                cell_value = a_cell.value
                
                # 디버깅용 로그 추가
                logger.debug(f"A열 셀 값: '{cell_value}', 타입: {type(cell_value)}")
                
                if add_hash:
                    # #을 추가
                    if cell_value is None:
                        a_cell.value = "#"
                        modified_count += 1
                        logger.debug(f"None에서 # 추가: '{a_cell.value}'")
                    elif isinstance(cell_value, str):
                        if not cell_value.startswith('#'):
                            a_cell.value = f"#{cell_value}"
                            modified_count += 1
                            logger.debug(f"문자열에 # 추가: '{cell_value}' -> '{a_cell.value}'")
                    else:
                        a_cell.value = f"#{cell_value}"
                        modified_count += 1
                        logger.debug(f"기타 타입에 # 추가: '{cell_value}' -> '{a_cell.value}'")
                else:
                    # #을 제거
                    if cell_value and isinstance(cell_value, str) and cell_value.startswith('#'):
                        a_cell.value = cell_value[1:]
                        modified_count += 1
                        logger.debug(f"행 {actual_row}의 # 제거: '{cell_value}' -> '{a_cell.value}'")
                
            # 파일 저장
            wb.save(excel_path)
            
            # 중요: 작업 완료 후 파일 명시적으로 닫기 추가
            wb.close()
            
            action_type = "추가" if add_hash else "제거"
            logger.info(f"{modified_count}개 행 #문자 {action_type} 후 파일 저장 완료: {excel_path}")
            
            # 수정된 행이 없어도 True 반환 (항목을 찾았지만 이미 #이 있어서 수정이 필요 없는 경우)
            return True
        
        except Exception as e:
            logger.error(f"A열 수정 중 오류: {e}")
            logger.debug(traceback.format_exc())
            return False

    @staticmethod
    def add_hash_to_a_column(excel_path, sheet_name, id_value, id_column=None, header_row=None):
        """
        지정된 ID에 해당하는 모든 행의 A열에 #을 추가합니다.
        
        Args:
            excel_path (str): 엑셀 파일 경로
            sheet_name (str): 시트 이름
            id_value (str): 찾을 ID 값
            id_column (str, optional): 특정 ID 컬럼 이름
            header_row (int, optional): 헤더 행 번호
            
        Returns:
            bool: 성공 여부
        """
        result = ExcelFileManager._find_rows_by_id(excel_path, sheet_name, id_value, id_column, header_row)
        if not result:
            return False
        
        _, _, matching_rows = result
        return ExcelFileManager._modify_a_column_hash(excel_path, sheet_name, matching_rows, add_hash=True, header_row=header_row)
    
    
    @staticmethod
    def find_file_by_type(cache, type_code, type_category="reward"):
        """
        주어진 타입 코드에 해당하는 테이블 파일을 찾습니다.
        
        Args:
            cache (dict): 엑셀 캐시 정보
            type_code (str|int): 타입 코드
            type_category (str): 타입 카테고리 ("reward", "item" 등)
            
        Returns:
            tuple: (파일명, 파일 경로) 또는 (None, None)
        """
        table_name = get_table_name_for_type(type_code, type_category)
        if not table_name:
            logger.warning(f"타입 코드 {type_code}에 해당하는 테이블을 찾을 수 없음")
            return None, None
        
        # 캐시에서 파일 찾기
        for file, info in cache.items():
            if table_name.lower() in file.lower():
                return file, info["path"]
        
        logger.warning(f"캐시에서 {table_name} 테이블을 찾을 수 없음")
        return None, None
    
    @staticmethod
    def get_item_name_by_type_id(db_folder, cache, type_code, id_value, type_category="reward"):
        """
        타입 코드와 ID를 사용하여 항목 이름을 조회합니다.
        
        Args:
            db_folder (str): DB 폴더 경로
            cache (dict): 엑셀 캐시 정보
            type_code (str|int): 타입 코드
            id_value (str|int): ID 값
            type_category (str): 타입 카테고리 ("reward", "item" 등)
            
        Returns:
            str: 항목 이름 또는 기본 텍스트
        """
        import sqlite3
        
        # 타입 정보 가져오기
        table_name = get_table_name_for_type(type_code, type_category)
        id_column = get_column_for_type(type_code, type_category)
        description = get_description_for_type(type_code, type_category)
        
        if not table_name or not id_column:
            return f"{description}: {id_value}"
        
        # DB 파일 확인
        db_path = os.path.join(db_folder, f"{table_name}.db")
        if os.path.exists(db_path):
            try:
                # DB에서 이름 조회
                conn = sqlite3.connect(db_path)
                cursor = conn.cursor()
                
                # 컬럼 목록 확인
                cursor.execute(f"PRAGMA table_info({table_name})")
                columns = [row[1] for row in cursor.fetchall()]
                
                # ID 컬럼이 있는지 확인
                if id_column not in columns:
                    for col in columns:
                        if 'id' in col.lower():
                            id_column = col
                            break
                
                # 이름 컬럼 찾기
                name_col = None
                for col in columns:
                    if col.lower() in ['name', 'displayname', 'title', 'description']:
                        name_col = col
                        break
                
                if not name_col:
                    conn.close()
                    return f"{description} {id_value}"
                
                # ID로 데이터 조회
                query = f"SELECT {name_col} FROM {table_name} WHERE {id_column} = ?"
                cursor.execute(query, (id_value,))
                result = cursor.fetchone()
                
                if result and result[0]:
                    conn.close()
                    return result[0]
                
                conn.close()
            except Exception as e:
                logger.error(f"DB 조회 오류: {e}")
        
        # DB에서 찾지 못했으면 엑셀에서 검색
        file, path = ExcelFileManager.find_file_by_type(cache, type_code, type_category)
        if file and path:
            for sheet, meta in cache[file].get("sheets", {}).items():
                try:
                    header = meta["header_row"]
                    df = pd.read_excel(path, sheet_name=sheet, header=header)
                    
                    # ID 컬럼 찾기
                    sheet_id_col = None
                    for col in df.columns:
                        if id_column.lower() in str(col).lower() or 'id' in str(col).lower():
                            sheet_id_col = col
                            break
                    
                    if not sheet_id_col:
                        continue
                    
                    # 이름 컬럼 찾기
                    name_col = None
                    for col in df.columns:
                        col_lower = str(col).lower()
                        if col_lower in ['name', 'displayname', 'title', 'description']:
                            name_col = col
                            break
                    
                    if not name_col:
                        continue
                    
                    # ID로 검색
                    try:
                        numeric_id = float(id_value)
                        df[sheet_id_col] = pd.to_numeric(df[sheet_id_col], errors='coerce')
                        matched = df[df[sheet_id_col] == numeric_id]
                    except:
                        df[sheet_id_col] = df[sheet_id_col].astype(str)
                        matched = df[df[sheet_id_col] == str(id_value)]
                    
                    if not matched.empty and name_col in matched.columns:
                        item_name = matched.iloc[0][name_col]
                        if pd.notna(item_name):
                            return str(item_name)
                except Exception as e:
                    logger.error(f"엑셀 조회 오류: {file}/{sheet}: {e}")
        
        # 모든 검색 실패 시 기본값 반환
        return f"{description} {id_value}"
    
    @staticmethod
    def add_hash_to_reward_id(db_folder, cache, reward_id, reward_type=None):
        """
        특정 RewardID와 일치하는 모든 Box 파일의 A열에 #을 추가합니다.
        """
        # 타입 정보 가져오기 (선택적)
        if reward_type:
            type_desc = get_description_for_type(reward_type)
            logger.info(f"{type_desc} 타입의 RewardID {reward_id} 검색")
        
        total_modified = 0
        
        # Box 관련 파일만 필터링 - ItemTemplate 제외
        box_files = []
        box_pattern = re.compile(r'^box', re.IGNORECASE)
        
        for file, info in cache.items():
            # ItemTemplate 파일 제외
            if 'itemtemplate' in file.lower():
                logger.info(f"ItemTemplate 파일 제외: {file}")
                continue
                
            file_name = os.path.splitext(file)[0]  # 확장자 제외한 파일명
            if box_pattern.match(file_name):
                box_files.append((file, info))
                continue
                    
            # 시트 이름이 Box로 시작하는 파일도 포함
            for sheet in info.get("sheets", {}):
                if box_pattern.match(sheet):
                    box_files.append((file, info))
                    break
        
        if not box_files:
            logger.warning(f"Box 관련 엑셀 파일을 찾을 수 없음")
            return 0
        
        # 각 파일에서 RewardID 검색
        for file, info in box_files:
            path = info["path"]
            logger.info(f"[RewardID 검색] 파일: {file}")
            
            for sheet, meta in info.get("sheets", {}).items():
                try:
                    # 헤더 행 정보 가져오기 (안전하게 기본값 설정)
                    header_row = meta.get("header_row")
                    logger.debug(f"[RewardID 검색] 시트: {sheet}, 헤더 행: {header_row}")
                    
                    # 엑셀 파일 읽기 - 헤더 행 지정
                    df = pd.read_excel(path, sheet_name=sheet, header=header_row)
                    
                    # 컬럼 정보 로깅
                    logger.debug(f"컬럼 목록: {df.columns.tolist()}")
                    
                    # 컬럼이 숫자로만 이루어져 있는지 확인 (header가 적용되지 않았을 가능성)
                    is_numeric_columns = all(isinstance(col, int) for col in df.columns)
                    if is_numeric_columns:
                        logger.warning(f"컬럼이 숫자로만 이루어져 있음 - 헤더 행 ({header_row})이 올바르지 않을 수 있음")
                        # 엑셀 데이터 구조 확인을 위해 첫 번째 행 출력
                        if len(df) > 0:
                            logger.debug(f"첫 번째 행 데이터: {df.iloc[0].tolist()}")
                    
                    # RewardID 컬럼 찾기
                    reward_id_col = None
                    for col in df.columns:
                        col_lower = str(col).lower()
                        if 'rewardid' in col_lower or 'reward_id' in col_lower:
                            reward_id_col = col
                            logger.debug(f"RewardID 컬럼 발견: {col}")
                            break
                    
                    if not reward_id_col:
                        logger.debug(f"[RewardID 검색] {sheet} 시트에 RewardID 컬럼이 없음")
                        continue
                    
                    # RewardType 필터링 (지정된 경우)
                    if reward_type:
                        # (RewardType 필터링 로직)
                        pass
                    
                    # RewardID 검색 (숫자형과 문자열형 모두 시도)
                    matched_rows = []
                    
                    # 숫자형 검색
                    try:
                        numeric_reward_id = float(reward_id)
                        df[reward_id_col] = pd.to_numeric(df[reward_id_col], errors='coerce')
                        matched = df[df[reward_id_col] == numeric_reward_id]
                        if not matched.empty:
                            matched_rows = matched.index.tolist()
                            logger.debug(f"숫자형 검색으로 {len(matched_rows)}개 행 찾음")
                    except Exception as e:
                        logger.debug(f"숫자형 검색 오류: {e}")
                    
                    # 문자열 검색 (숫자 검색이 실패했거나 결과가 없는 경우)
                    if not matched_rows:
                        try:
                            df[reward_id_col] = df[reward_id_col].astype(str)
                            matched = df[df[reward_id_col] == str(reward_id)]
                            if not matched.empty:
                                matched_rows = matched.index.tolist()
                                logger.debug(f"문자열 검색으로 {len(matched_rows)}개 행 찾음")
                        except Exception as e:
                            logger.debug(f"문자열 검색 오류: {e}")
                    
                    if matched_rows:
                        logger.info(f"[RewardID 검색] {file}/{sheet}에서 RewardID {reward_id} 발견: {len(matched_rows)}행")
                        # A열 수정
                        try:
                            # A열 수정 로직 구현
                            # 중요: matched_rows의 인덱스는 0부터 시작하지만, 실제 엑셀에서는 헤더 행 이후부터 시작
                            # 따라서 실제 행 번호는 인덱스 + header_row + 1(엑셀의 행은 1부터 시작)
                            modified = ExcelFileManager._modify_a_column_hash(path, sheet, matched_rows, add_hash=True, header_row=header_row)
                            if modified:
                                total_modified += len(matched_rows)
                        except Exception as e:
                            logger.error(f"A열 수정 오류: {e}")
                
                except Exception as e:
                    logger.error(f"[RewardID 검색] 파일 처리 오류: {file}/{sheet}: {e}")
        
        return total_modified

    @staticmethod
    def remove_hash_from_reward_id(db_folder, cache, reward_id, reward_type=None):
        """
        특정 RewardID와 일치하는 모든 Box 파일의 A열에서 #을 제거합니다.
        """
        # 타입 정보 가져오기 (선택적)
        if reward_type:
            type_desc = get_description_for_type(reward_type)
            logger.info(f"{type_desc} 타입의 RewardID {reward_id} 검색")
        
        total_modified = 0
        
        # Box 관련 엑셀 파일 찾기 - ItemTemplate 제외
        box_files = []
        for file, info in cache.items():
            # ItemTemplate 파일 제외
            if 'itemtemplate' in file.lower():
                logger.info(f"ItemTemplate 파일 제외: {file}")
                continue
                
            if 'box' in file.lower() or 'boxtemplate' in file.lower():
                box_files.append((file, info["path"]))
        
        if not box_files:
            logger.warning(f"Box 관련 엑셀 파일을 찾을 수 없음")
            return 0
        
        # 각 파일에서 RewardID 검색
        for file, path in box_files:
            file_info = cache.get(file, {})
            
            for sheet, meta in file_info.get("sheets", {}).items():
                try:
                    header = meta["header_row"]
                    
                    # 엑셀 파일 읽기
                    df = pd.read_excel(path, sheet_name=sheet, header=header)
                    
                    # RewardID 컬럼 찾기
                    reward_id_col = None
                    for col in df.columns:
                        col_lower = str(col).lower()
                        if 'rewardid' in col_lower or 'reward_id' in col_lower:
                            reward_id_col = col
                            break
                    
                    if not reward_id_col:
                        continue
                    
                    # RewardType 필터링 (지정된 경우)
                    if reward_type:
                        reward_type_col = None
                        for col in df.columns:
                            col_lower = str(col).lower()
                            if 'rewardtype' in col_lower or 'reward_type' in col_lower:
                                reward_type_col = col
                                break
                        
                        if reward_type_col:
                            # RewardType으로 필터링
                            try:
                                df[reward_type_col] = pd.to_numeric(df[reward_type_col], errors='coerce')
                                df = df[df[reward_type_col] == float(reward_type)]
                                if df.empty:
                                    continue
                            except:
                                df[reward_type_col] = df[reward_type_col].astype(str)
                                df = df[df[reward_type_col] == str(reward_type)]
                                if df.empty:
                                    continue
                    
                    # RewardID 검색 (숫자형과 문자열형 모두 시도)
                    matched_rows = []
                    
                    # 숫자형 검색
                    try:
                        numeric_reward_id = float(reward_id)
                        df[reward_id_col] = pd.to_numeric(df[reward_id_col], errors='coerce')
                        matched = df[df[reward_id_col] == numeric_reward_id]
                        if not matched.empty:
                            matched_rows = matched.index.tolist()
                    except:
                        # 문자열 검색
                        df[reward_id_col] = df[reward_id_col].astype(str)
                        matched = df[df[reward_id_col] == str(reward_id)]
                        if not matched.empty:
                            matched_rows = matched.index.tolist()
                    
                    if matched_rows:
                        logger.info(f"{file}/{sheet}에서 RewardID {reward_id} 발견: {len(matched_rows)}행")
                        # A열 수정
                        if ExcelFileManager._modify_a_column_hash(path, sheet, matched_rows, add_hash=False):
                            total_modified += len(matched_rows)
                
                except Exception as e:
                    logger.error(f"파일 처리 오류: {file}/{sheet}: {e}")
        
        return total_modified
    
    @staticmethod
    def search_by_reward_id(cache, reward_id, reward_type=None):
        """
        RewardID를 포함하는 Box를 검색합니다.
        
        Args:
            cache (dict): 엑셀 캐시 정보
            reward_id (str): 찾을 RewardID 값
            reward_type (str, optional): RewardType 값 (지정 시 해당 타입만 검색)
            
        Returns:
            list: (ItemID, BoxName, BoxType) 튜플 리스트
        """
        found_boxes = []
        
        # Box 관련 엑셀 파일 찾기
        for file, info in cache.items():
            if 'box' in file.lower() or 'boxtemplate' in file.lower():
                path = info["path"]
                
                for sheet, meta in info.get("sheets", {}).items():
                    try:
                        header = meta["header_row"]
                        
                        # 엑셀 파일 읽기
                        df = pd.read_excel(path, sheet_name=sheet, header=header)
                        
                        # RewardID 컬럼 찾기
                        reward_id_col = None
                        id_col = None
                        
                        for col in df.columns:
                            col_lower = str(col).lower()
                            if 'rewardid' in col_lower or 'reward_id' in col_lower:
                                reward_id_col = col
                            elif 'itemid' in col_lower or 'item_id' in col_lower or 'templateid' in col_lower or 'itemtid' in col_lower:
                                id_col = col
                        
                        if not (reward_id_col and id_col):
                            continue
                        
                        # RewardType 필터링 (지정된 경우)
                        if reward_type:
                            reward_type_col = None
                            for col in df.columns:
                                col_lower = str(col).lower()
                                if 'rewardtype' in col_lower or 'reward_type' in col_lower:
                                    reward_type_col = col
                                    break
                            
                            if reward_type_col:
                                # RewardType으로 필터링
                                try:
                                    df[reward_type_col] = pd.to_numeric(df[reward_type_col], errors='coerce')
                                    df = df[df[reward_type_col] == float(reward_type)]
                                    if df.empty:
                                        continue
                                except:
                                    df[reward_type_col] = df[reward_type_col].astype(str)
                                    df = df[df[reward_type_col] == str(reward_type)]
                                    if df.empty:
                                        continue
                        
                        # 숫자형과 문자열 형 모두 검색
                        try:
                            # 숫자형 검색
                            numeric_reward_id = float(reward_id)
                            df[reward_id_col] = pd.to_numeric(df[reward_id_col], errors='coerce')
                            matched = df[df[reward_id_col] == numeric_reward_id]
                        except:
                            # 문자열 검색
                            df[reward_id_col] = df[reward_id_col].astype(str)
                            matched = df[df[reward_id_col] == str(reward_id)]
                        
                        # 결과 처리
                        if not matched.empty:
                            for _, row in matched.iterrows():
                                box_id = str(row[id_col])
                                
                                # 이미 발견한 Box ID는 중복 추가하지 않음
                                if not any(box[0] == box_id for box in found_boxes):
                                    # Box 이름과 타입은 추후 ItemTemplate에서 조회
                                    found_boxes.append((box_id, "", ""))
                    
                    except Exception as e:
                        logger.error(f"검색 오류: {file} / {sheet}: {e}")
        
        return found_boxes
        

    @staticmethod
    def _find_box_files(cache):
        """
        Box로 시작하는 파일 또는 시트만 선택합니다.
        
        Args:
            cache (dict): 엑셀 캐시 정보
            
        Returns:
            list: (파일명, 파일 경로, 시트명, 헤더 행) 튜플 리스트
        """
        import re
        
        box_files = []
        # 정규표현식 패턴: 단어 시작(^)이 Box 또는 box인 경우만 매칭
        box_pattern = re.compile(r'^box', re.IGNORECASE)
        
        for file, info in cache.items():
            file_name = os.path.splitext(file)[0]  # 확장자 제외한 파일명
            
            # 파일명이 Box로 시작하는 경우, 모든 시트 추가
            if box_pattern.match(file_name):
                for sheet, meta in info.get("sheets", {}).items():
                    header_row = meta.get("header_row", 0)
                    logger.debug(f"Box 파일 발견: {file}, 시트: {sheet}, 헤더 행: {header_row}")
                    box_files.append((file, info["path"], sheet, header_row))
                continue
                
            # 또는 시트명이 Box로 시작하는 시트만 추가
            for sheet, meta in info.get("sheets", {}).items():
                if box_pattern.match(sheet):
                    header_row = meta.get("header_row", 0)
                    logger.debug(f"Box 시트 발견: {file}, 시트: {sheet}, 헤더 행: {header_row}")
                    box_files.append((file, info["path"], sheet, header_row))
        
        logger.info(f"총 {len(box_files)}개의 Box 관련 파일/시트를 찾았습니다.")
        return box_files
    
    @staticmethod
    def find_open_excel_workbook(file_path):
        """
        이미 열려있는 엑셀 워크북을 찾습니다.
        
        Args:
            file_path: 찾을 엑셀 파일 경로
            
        Returns:
            tuple: (excel_app, workbook) 또는 (None, None)
        """
        try:
            import win32com.client
            
            # 실행 중인 Excel 애플리케이션 찾기
            try:
                excel = win32com.client.GetActiveObject("Excel.Application")
            except:
                # 실행 중인 Excel이 없으면 None 반환
                return None, None
            
            # 파일명 추출
            target_filename = os.path.basename(file_path).lower()
            
            # 열려있는 워크북들 검사
            for workbook in excel.Workbooks:
                try:
                    # 워크북의 전체 경로와 비교
                    if workbook.FullName.lower() == os.path.abspath(file_path).lower():
                        logger.debug(f"이미 열려있는 워크북 발견: {workbook.FullName}")
                        return excel, workbook
                    
                    # 파일명만으로도 비교 (경로가 다를 수 있음)
                    if workbook.Name.lower() == target_filename:
                        logger.debug(f"파일명으로 일치하는 워크북 발견: {workbook.Name}")
                        return excel, workbook
                        
                except Exception as wb_error:
                    logger.debug(f"워크북 검사 오류: {wb_error}")
                    continue
            
            return None, None
            
        except Exception as e:
            logger.debug(f"Excel 애플리케이션 검사 오류: {e}")
            return None, None

    @staticmethod
    def activate_excel_worksheet(workbook, sheet_name):
        """
        워크북에서 특정 시트를 활성화합니다.
        
        Args:
            workbook: Excel 워크북 객체
            sheet_name: 활성화할 시트 이름
            
        Returns:
            worksheet 객체 또는 None
        """
        try:
            # 시트가 존재하는지 확인
            for sheet in workbook.Worksheets:
                if sheet.Name == sheet_name:
                    sheet.Activate()
                    logger.debug(f"시트 활성화 성공: {sheet_name}")
                    return sheet
            
            # 시트를 찾지 못한 경우 첫 번째 시트 활성화
            if workbook.Worksheets.Count > 0:
                first_sheet = workbook.Worksheets(1)
                first_sheet.Activate()
                logger.debug(f"대상 시트 없음, 첫 번째 시트 활성화: {first_sheet.Name}")
                return first_sheet
            
            return None
            
        except Exception as e:
            logger.error(f"시트 활성화 오류: {e}")
            return None

    @staticmethod
    def highlight_excel_by_value(file_path, sheet_name, column_name, value, excel_cache=None):
        """
        엑셀 파일에서 특정 컬럼의 값을 찾아 강조 표시하고 엑셀을 엽니다.
        (이미 열려있는 파일이 있으면 해당 파일을 사용)
        
        Args:
            file_path: 엑셀 파일 경로
            sheet_name: 시트 이름
            column_name: 컬럼 이름
            value: 찾을 값
            excel_cache: 엑셀 캐시 데이터 (선택적)
            
        Returns:
            성공 여부 (Boolean)
        """
        try:
            import win32com.client
            
            # 1. 이미 열려있는 워크북 찾기
            excel, workbook = ExcelFileManager.find_open_excel_workbook(file_path)
            
            if excel and workbook:
                logger.info(f"이미 열려있는 엑셀 파일 사용: {file_path}")
                # Excel을 화면에 표시
                excel.Visible = False
                excel.WindowState = win32com.client.constants.xlMaximized  # 최대화
                excel.Activate()  # Excel 애플리케이션 활성화
                
                # 해당 워크북을 최상위로
                workbook.Activate()
            else:
                logger.info(f"새로운 엑셀 파일 열기: {file_path}")
                # 2. 새로 Excel 애플리케이션 열기
                excel = win32com.client.Dispatch("Excel.Application")
                excel.Visible = False
                
                # 파일 열기
                workbook = excel.Workbooks.Open(os.path.abspath(file_path))
            
            # 3. 해당 시트로 이동
            worksheet = ExcelFileManager.activate_excel_worksheet(workbook, sheet_name)
            if not worksheet:
                logger.warning(f"시트 활성화 실패: {sheet_name}")
                return True  # 파일은 열렸으므로 True 반환
            
            # 4. 헤더 행 찾기
            header_row = 3  # 기본값
            if excel_cache:
                rel_path = os.path.relpath(file_path).replace("\\", "/")
                if rel_path in excel_cache and "sheets" in excel_cache[rel_path]:
                    sheet_info = excel_cache[rel_path]["sheets"].get(sheet_name, {})
                    if "header_row" in sheet_info:
                        header_row = sheet_info["header_row"] + 1  # 1-based로 변환
            
            # 5. 헤더에서 컬럼 위치 찾기
            column_idx = None
            for i in range(1, 100):  # 최대 100개 컬럼까지 검색
                cell_value = worksheet.Cells(header_row, i).Value
                if cell_value == column_name:
                    column_idx = i
                    logger.debug(f"컬럼 발견: {column_name} at column {i}")
                    break
            
            if column_idx is None:
                logger.warning(f"컬럼을 찾을 수 없음: {column_name}")
                return True  # 파일은 열렸으므로 True 반환
            
            # 6. 값 검색 및 찾은 셀로 이동
            found = False
            for row in range(header_row + 1, header_row + 10000):  # 최대 10000행까지 검색
                try:
                    cell_value = worksheet.Cells(row, column_idx).Value
                    if str(cell_value) == str(value):
                        # 찾은 셀로 이동
                        cell = worksheet.Cells(row, column_idx)
                        cell.Select()
                        
                        # 강조 표시 (기존 강조 표시 제거 후 새로 적용)
                        # 기존 강조 표시 제거 (A:Z 범위에서)
                        try:
                            worksheet.Range("A:Z").Interior.ColorIndex = win32com.client.constants.xlColorIndexNone
                        except:
                            pass
                        
                        # 새로운 강조 표시
                        cell.Interior.ColorIndex = 6  # 노란색
                        
                        # 화면 조정
                        excel.ActiveWindow.ScrollRow = max(1, row - 5)
                        
                        logger.info(f"값 발견 및 강조 표시: {value} at row {row}")
                        found = True
                        break
                        
                except Exception as cell_error:
                    # 빈 셀이나 읽기 오류는 무시하고 계속
                    continue
            
            if not found:
                logger.warning(f"값을 찾을 수 없음: {value}")
            
            return True
            
        except Exception as e:
            logger.error(f"엑셀 강조 표시 실패: {e}")
            
            # 기본 방식으로 파일 열기 시도
            try:
                os.startfile(file_path)
                return True
            except:
                return False
                
    @staticmethod
    def open_excel_file_simple(file_path, sheet_name=None):
        """
        엑셀 파일을 단순하게 엽니다. (프로세스 독립적)
        
        Args:
            file_path: 엑셀 파일 경로
            sheet_name: 시트 이름 (사용되지 않음, 호환성용)
            
        Returns:
            성공 여부 (Boolean)
        """
        try:
            if not os.path.exists(file_path):
                logger.error(f"파일이 존재하지 않음: {file_path}")
                return False
            
            logger.info(f"심플 모드로 엑셀 파일 열기: {file_path}")
            
            if sys.platform.startswith('win'):
                # Windows에서 기본 프로그램으로 열기
                os.startfile(file_path)
            elif sys.platform == 'darwin':
                # macOS에서 기본 프로그램으로 열기
                subprocess.call(['open', file_path])
            else:
                # Linux에서 기본 프로그램으로 열기
                subprocess.call(['xdg-open', file_path])
            
            return True
            
        except Exception as e:
            logger.error(f"심플 엑셀 파일 열기 실패: {file_path} - {e}")
            return False