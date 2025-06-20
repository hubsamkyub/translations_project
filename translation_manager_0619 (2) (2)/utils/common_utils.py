import os
import json
import hashlib
import logging
import sqlite3
import pandas as pd
import time
from pathlib import Path
from functools import lru_cache
from typing import Dict, List, Union, Optional, Any, Tuple

# 로깅 설정
def setup_logging(log_file: str = 'app.log', console_level: int = logging.INFO, file_level: int = logging.DEBUG) -> logging.Logger:
    """
    애플리케이션 로깅 설정
    
    Args:
        log_file: 로그 파일 경로
        console_level: 콘솔 출력 로그 레벨
        file_level: 파일 출력 로그 레벨
    
    Returns:
        설정된 로거 객체
    """
    logger = logging.getLogger('app')
    logger.setLevel(logging.DEBUG)
    
    # 모든 핸들러 제거 (중복 방지)
    for handler in logger.handlers[:]:
        logger.removeHandler(handler)
    
    # 콘솔 핸들러
    console_handler = logging.StreamHandler()
    console_handler.setLevel(console_level)
    console_formatter = logging.Formatter('%(levelname)s - %(message)s')
    console_handler.setFormatter(console_formatter)
    logger.addHandler(console_handler)
    
    # 파일 핸들러
    file_handler = logging.FileHandler(log_file, encoding='utf-8', mode='w')
    file_handler.setLevel(file_level)
    file_formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
    file_handler.setFormatter(file_formatter)
    logger.addHandler(file_handler)
    
    return logger

# 기본 로거 생성
logger = setup_logging()

# 파일 및 경로 유틸리티
class PathUtils:
    @staticmethod
    def normalize_path(file_path: Optional[str]) -> str:
        """
        파일 경로를 정규화합니다.
        
        Args:
            file_path: 정규화할 파일 경로
            
        Returns:
            정규화된 경로 문자열
        """
        if not file_path:
            return ""
        
        # Path 객체 사용하여 정규화
        return str(Path(file_path)).replace("\\", "/")
    
    @staticmethod
    def get_file_mtime(path: str) -> int:
        """
        파일의 수정 시간을 가져옵니다.
        
        Args:
            path: 파일 경로
            
        Returns:
            파일 수정 시간 (정수, 초 단위)
        """
        try:
            return int(os.path.getmtime(path))
        except Exception as e:
            logger.warning(f"파일 수정 시간 얻기 실패: {path} - {e}")
            return -1
    
    @staticmethod
    def get_display_name(file_path: str) -> str:
        """
        파일 경로에서 표시용 파일명을 추출합니다.
        
        Args:
            file_path: 파일 경로
            
        Returns:
            파일명
        """
        return os.path.basename(file_path) if file_path else ""
    
    @staticmethod
    def ensure_dir(dir_path: str) -> None:
        """
        디렉토리가 존재하는지 확인하고, 없으면 생성합니다.
        
        Args:
            dir_path: 디렉토리 경로
        """
        if dir_path:
            os.makedirs(dir_path, exist_ok=True)
    
    @staticmethod
    def find_files(base_path: str, pattern: str, recursive: bool = True) -> List[str]:
        """
        특정 패턴의 파일을 찾습니다.
        
        Args:
            base_path: 기본 검색 경로
            pattern: 검색할 파일 패턴 (예: *.xlsx)
            recursive: 하위 디렉토리까지 검색 여부
            
        Returns:
            발견된 파일 경로 목록
        """
        found_files = []
        
        if not os.path.exists(base_path):
            logger.warning(f"기본 경로가 존재하지 않음: {base_path}")
            return found_files
        
        if recursive:
            for root, _, files in os.walk(base_path):
                for filename in files:
                    if filename.endswith(pattern) and not filename.startswith("~$"):
                        found_files.append(os.path.join(root, filename))
        else:
            for filename in os.listdir(base_path):
                if filename.endswith(pattern) and not filename.startswith("~$"):
                    found_files.append(os.path.join(base_path, filename))
        
        return found_files
    
    @staticmethod
    def get_file_identifier(rel_path: str) -> str:
        """
        상대 경로에서 파일을 고유하게 식별할 수 있는 ID 생성
        
        Args:
            rel_path: 상대 경로
            
        Returns:
            파일 식별자
        """
        norm_path = PathUtils.normalize_path(rel_path)
        filename = os.path.basename(norm_path)
        
        # Excel_String 폴더 내부 구조 추출
        parts = norm_path.split('/')
        for i, part in enumerate(parts):
            if part.startswith("Excel_String"):
                if i < len(parts) - 1:
                    return norm_path.replace("/", "_")
        
        # Excel_String 폴더가 없는 경우 파일명만 반환
        return filename

# 해시 관련 유틸리티
class HashUtils:
    @staticmethod
    def hash_paths(*paths: str) -> str:
        """
        여러 경로를 해시하여 고유 ID 생성
        
        Args:
            *paths: 해시할 경로들
            
        Returns:
            MD5 해시 문자열
        """
        normalized_paths = []
        
        for path in paths:
            if path is None:
                raise ValueError("경로가 None입니다")
            
            path_str = str(path).strip() or "."
            if not path_str:
                raise ValueError("경로가 비어 있습니다")
            
            normalized_paths.append(os.path.abspath(os.path.normpath(path_str)))
        
        key = "|".join(normalized_paths)
        return hashlib.md5(key.encode("utf-8")).hexdigest()
    
    @staticmethod
    def get_cache_dir(base_cache_dir: str, *paths: str) -> str:
        """
        캐시 디렉토리 경로 생성 및 보장
        
        Args:
            base_cache_dir: 기본 캐시 디렉토리
            *paths: 해시할 경로들
            
        Returns:
            캐시 디렉토리 경로
        """
        cache_id = HashUtils.hash_paths(*paths)
        cache_dir = os.path.join(base_cache_dir, cache_id)
        PathUtils.ensure_dir(cache_dir)
        return cache_dir

# 파일 입출력 유틸리티
class FileUtils:
    @staticmethod
    def load_json(file_path: str, default_value: Any = None) -> Any:
        """
        JSON 파일을 로드합니다.
        
        Args:
            file_path: JSON 파일 경로
            default_value: 파일이 없거나 오류 시 반환할 기본값
            
        Returns:
            로드된
            JSON 데이터 또는 기본값
        """
        if not os.path.exists(file_path):
            logger.debug(f"파일이 존재하지 않음: {file_path}, 기본값 반환")
            return {} if default_value is None else default_value
        
        try:
            with open(file_path, "r", encoding="utf-8") as f:
                data = json.load(f)
            logger.debug(f"파일 로드 성공: {file_path}")
            return data
        except Exception as e:
            logger.error(f"파일 로드 실패: {file_path} - {e}")
            return {} if default_value is None else default_value
    
    @staticmethod
    def save_json(file_path: str, data: Any, indent: int = 2, backup: bool = True) -> bool:
        """
        데이터를 JSON 파일로 저장합니다.
        
        Args:
            file_path: 저장할 파일 경로
            data: 저장할 데이터
            indent: JSON 들여쓰기 간격
            backup: 기존 파일 백업 여부
            
        Returns:
            성공 여부 (True/False)
        """
        try:
            # 디렉토리 생성
            dir_path = os.path.dirname(file_path)
            if dir_path:
                PathUtils.ensure_dir(dir_path)
            
            # 백업 생성
            if backup and os.path.exists(file_path):
                backup_file = f"{file_path}.bak"
                try:
                    import shutil
                    shutil.copy2(file_path, backup_file)
                    logger.debug(f"백업 파일 생성: {backup_file}")
                except Exception as backup_err:
                    logger.warning(f"백업 생성 실패: {backup_err}")
            
            # 파일 저장
            with open(file_path, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=indent, ensure_ascii=False)
            
            logger.debug(f"파일 저장 완료: {file_path}")
            return True
        except Exception as e:
            logger.error(f"파일 저장 실패: {file_path} - {e}")
            return False
    
    @staticmethod
    def load_cached_data(cache_path: str) -> Dict:
        """
        캐시 데이터를 로드합니다.
        
        Args:
            cache_path: 캐시 파일 경로
            
        Returns:
            캐시 데이터 (없을 경우 빈 딕셔너리)
        """
        return FileUtils.load_json(cache_path, {})

    @staticmethod
    def truncate(value: Any, max_len: int = 100) -> str:
        """
        너무 긴 문자열을 자르고 ... 표시
        
        Args:
            value: 잘라낼 값
            max_len: 최대 길이
            
        Returns:
            잘린 문자열
        """
        text = str(value)
        return text if len(text) <= max_len else text[:max_len] + "..."

    @staticmethod
    def save_cache(path: str, data: Dict) -> bool:
        """
        캐시 데이터를 저장합니다.
        
        Args:
            path: 캐시 파일 경로
            data: 저장할 캐시 데이터
            
        Returns:
            성공 여부 (True/False)
        """
        return FileUtils.save_json(path, data)

# 데이터베이스 유틸리티
class DBUtils:
    @staticmethod
    @lru_cache(maxsize=100)
    def get_columns_from_db(db_path: str, table_name: str) -> List[str]:
        """
        DB에서 컬럼 목록을 가져옵니다. (캐싱됨)
        
        Args:
            db_path: DB 파일 경로
            table_name: 테이블 이름
            
        Returns:
            컬럼 목록
        """
        if not os.path.exists(db_path):
            return []
            
        try:
            conn = sqlite3.connect(db_path)
            cur = conn.cursor()
            cur.execute(f"PRAGMA table_info({table_name})")
            columns = [row[1] for row in cur.fetchall()]
            conn.close()
            return columns
        except Exception as e:
            logger.error(f"DB 컬럼 조회 오류: {db_path}/{table_name} - {e}")
            return []
    
    @staticmethod
    def execute_query(db_path: str, query: str, params: tuple = (), fetch_all: bool = True) -> Union[List[tuple], None]:
        """
        SQL 쿼리를 실행하고 결과를 반환합니다.
        
        Args:
            db_path: DB 파일 경로
            query: SQL 쿼리
            params: 쿼리 파라미터
            fetch_all: 모든 결과 반환 여부 (False면 첫 번째 결과만)
            
        Returns:
            쿼리 결과 또는 None (오류 시)
        """
        if not os.path.exists(db_path):
            logger.warning(f"DB 파일이 존재하지 않음: {db_path}")
            return None
        
        try:
            conn = sqlite3.connect(db_path)
            cursor = conn.cursor()
            cursor.execute(query, params)
            
            if fetch_all:
                result = cursor.fetchall()
            else:
                result = cursor.fetchone()
                
            conn.close()
            return result
        except Exception as e:
            logger.error(f"쿼리 실행 오류: {db_path} - {query} - {e}")
            return None
    
    @staticmethod
    def create_table(db_path: str, table_name: str, columns_def: str) -> bool:
        """
        테이블을 생성합니다.
        
        Args:
            db_path: DB 파일 경로
            table_name: 테이블 이름
            columns_def: 컬럼 정의 문자열
            
        Returns:
            성공 여부 (True/False)
        """
        # 디렉토리 생성
        dir_path = os.path.dirname(db_path)
        PathUtils.ensure_dir(dir_path)
        
        try:
            conn = sqlite3.connect(db_path)
            cursor = conn.cursor()
            cursor.execute(f"CREATE TABLE IF NOT EXISTS {table_name} ({columns_def})")
            conn.commit()
            conn.close()
            return True
        except Exception as e:
            logger.error(f"테이블 생성 오류: {db_path}/{table_name} - {e}")
            return False
    
    @staticmethod
    def create_fts_table(db_path: str, table_name: str, columns: List[str]) -> bool:
        """
        FTS (전문 검색) 테이블을 생성합니다.
        
        Args:
            db_path: DB 파일 경로
            table_name: 테이블 이름
            columns: 컬럼 목록
            
        Returns:
            성공 여부 (True/False)
        """
        # 디렉토리 생성
        dir_path = os.path.dirname(db_path)
        PathUtils.ensure_dir(dir_path)
        
        try:
            conn = sqlite3.connect(db_path)
            cursor = conn.cursor()
            
            # 기존 테이블 확인
            cursor.execute(f"SELECT name FROM sqlite_master WHERE type='table' AND name='{table_name}'")
            if cursor.fetchone():
                conn.close()
                return True  # 이미 존재하면 성공으로 간주
                
            # FTS5 테이블 생성
            columns_str = ", ".join(columns)
            cursor.execute(f"CREATE VIRTUAL TABLE {table_name} USING fts5 ({columns_str})")
            
            conn.commit()
            conn.close()
            return True
        except Exception as e:
            logger.error(f"FTS 테이블 생성 오류: {db_path}/{table_name} - {e}")
            return False
    
    @staticmethod
    def insert_many(db_path: str, table_name: str, columns: List[str], values: List[tuple]) -> int:
        """
        여러 레코드를 테이블에 삽입합니다.
        
        Args:
            db_path: DB 파일 경로
            table_name: 테이블 이름
            columns: 컬럼 목록
            values: 삽입할 값 목록 (튜플의 리스트)
            
        Returns:
            삽입된 레코드 수
        """
        if not values:
            return 0
            
        try:
            conn = sqlite3.connect(db_path)
            cursor = conn.cursor()
            
            placeholders = ", ".join(["?"] * len(columns))
            columns_str = ", ".join(columns)
            query = f"INSERT INTO {table_name} ({columns_str}) VALUES ({placeholders})"
            
            cursor.executemany(query, values)
            inserted_count = cursor.rowcount
            
            conn.commit()
            conn.close()
            
            return inserted_count
        except Exception as e:
            logger.error(f"대량 삽입 오류: {db_path}/{table_name} - {e}")
            return 0

# Excel 관련 유틸리티
class ExcelUtils:
    """호환성을 위한 레이어: excel_utils.py의 ExcelFileManager를 사용"""
    
    @staticmethod
    def find_header_row(file_path, sheet_name, db_columns, xls=None):
        from utils.excel_utils import ExcelFileManager
        return ExcelFileManager.find_header_row(file_path, sheet_name, db_columns, xls)
    
    @staticmethod
    def find_excel_file(base_path, file_name):
        from utils.excel_utils import ExcelFileManager
        return ExcelFileManager.find_excel_file(base_path, file_name)

# 성능 측정 및 로깅 유틸리티
class PerformanceUtils:
    @staticmethod
    def timed_function(func):
        """
        함수 실행 시간을 측정하는 데코레이터
        
        사용 예:
        @PerformanceUtils.timed_function
        def my_function():
            # 코드
        """
        def wrapper(*args, **kwargs):
            start_time = time.time()
            logger.debug(f"{func.__name__} 실행 시작")
            
            try:
                result = func(*args, **kwargs)
                elapsed = time.time() - start_time
                logger.debug(f"{func.__name__} 실행 완료 ({elapsed:.4f}초)")
                return result
            except Exception as e:
                elapsed = time.time() - start_time
                logger.error(f"{func.__name__} 실행 중 오류 ({elapsed:.4f}초): {e}")
                raise
                
        return wrapper
    
    @staticmethod
    def log_performance(message: str, start_time: float) -> float:
        """
        경과 시간을 로깅하고 현재 시간 반환
        
        Args:
            message: 로그 메시지
            start_time: 시작 시간
            
        Returns:
            현재 시간 (다음 측정의 시작점으로 사용 가능)
        """
        current_time = time.time()
        elapsed = current_time - start_time
        logger.debug(f"{message}: {elapsed:.4f}초")
        return current_time
    
    @staticmethod
    def create_timer():
        """
        타이머 객체 생성
        
        Returns:
            타이머 객체(start, elapsed, log 메서드 포함)
        """
        class Timer:
            def __init__(self):
                self.start_time = time.time()
                self.last_time = self.start_time
                self.checkpoints = {}
            
            def elapsed(self):
                """총 경과 시간 반환"""
                return time.time() - self.start_time
            
            def checkpoint(self, name):
                """체크포인트 설정"""
                now = time.time()
                elapsed = now - self.last_time
                self.checkpoints[name] = elapsed
                self.last_time = now
                logger.debug(f"[타이머] {name}: {elapsed:.4f}초")
                return elapsed
            
            def summary(self):
                """모든 체크포인트 요약"""
                total = self.elapsed()
                result = {
                    "total": total,
                    "checkpoints": self.checkpoints
                }
                logger.debug(f"[타이머 요약] 총 {total:.4f}초, {len(self.checkpoints)}개 체크포인트")
                return result
        
        return Timer()

# 메시지 박스 유틸리티 (tkinter 사용)
def show_message(parent, message_type, title, message, log_level=None, progress_update=None):
    """
    통합 메시지 박스 표시 함수
    
    Args:
        parent: 부모 윈도우
        message_type: 메시지 박스 타입 ('info', 'warning', 'error', 'yesno')
        title: 메시지 박스 제목
        message: 표시할 메시지
        log_level: 로깅 레벨 (None이면 로깅하지 않음)
        progress_update: 진행 창 업데이트 함수 (None이면 업데이트 없음)
    
    Returns:
        yesno 타입인 경우 사용자 선택 결과(True/False), 아니면 None
    """
    # tkinter 임포트(필요할 때만)
    import tkinter as tk
    from tkinter import messagebox
    
    # 로깅 처리
    if log_level:
        log_func = getattr(logging, log_level.lower(), logging.info)
        log_func(f"메시지 박스 ({message_type}): {title} - {message}")
    
    # 진행 창 업데이트
    if progress_update:
        success = message_type != 'error'
        progress_update(f"{title}: {message}", success=success)
    
    # 메시지 박스 타입에 따라 다른 함수 호출
    if message_type == 'info':
        messagebox.showinfo(title, message, parent=parent)
        return None
    elif message_type == 'warning':
        messagebox.showwarning(title, message, parent=parent)
        return None
    elif message_type == 'error':
        messagebox.showerror(title, message, parent=parent)
        return None
    elif message_type == 'yesno':
        return messagebox.askyesno(title, message, parent=parent)
    else:
        # 지원하지 않는 타입은 기본적으로 info로 처리
        logging.warning(f"지원하지 않는 메시지 타입: {message_type}, info로 대체합니다.")
        messagebox.showinfo(title, message, parent=parent)
        return None