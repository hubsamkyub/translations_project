import sqlite3
import pandas as pd
import os
import json
import time
from pathlib import Path
import hashlib
import logging

from utils.cache_utils import load_cached_data, save_cache
from utils.common_utils import PathUtils, FileUtils, DBUtils, logger, HashUtils
from utils.excel_utils import ExcelFileManager

# 기존 함수 유지하되 내부는 common_utils 호출로 대체
def normalize_path(path):
    return PathUtils.normalize_path(path)

def get_file_identifier(rel_path):
    return PathUtils.get_file_identifier(rel_path)

# 시작 로그
logging.debug("===== 모듈 로딩 시작 =====")

def get_per_file_db_path(rel_path: str, cache_root_dir: str) -> str:
    """
    프리셋 전환시에도 동일 파일은 동일 DB 사용하기 위한 DB 경로 계산
    
    Args:
        rel_path: 엑셀 파일의 상대 경로
        cache_root_dir: 캐시 루트 디렉토리
        
    Returns:
        DB 파일 경로
    """
    logger.debug(f"DB 경로 계산: {rel_path}, {cache_root_dir}")
    
    # cache_root_dir이 .cache로 끝나는지 확인
    if not os.path.basename(os.path.dirname(cache_root_dir)) == '.cache':
        logger.warning(f"경고: 캐시 루트 디렉토리가 .cache 내부가 아님: {cache_root_dir}")
        
        # 강제로 .cache 디렉토리 내부로 경로 변경
        script_dir = os.path.dirname(os.path.abspath(__file__))
        cache_id = HashUtils.hash_paths(cache_root_dir)  # HashUtils 사용
        cache_root_dir = os.path.join(script_dir, ".cache", cache_id)
        logger.debug(f"캐시 경로 수정: {cache_root_dir}")
    
    file_id = PathUtils.get_file_identifier(rel_path)  # PathUtils 사용
    db_filename = Path(file_id).stem + ".db"
    final_path = os.path.join(cache_root_dir, "string_dbs", db_filename)
    logger.debug(f"최종 DB 경로: {final_path}")
    return final_path


def build_single_string_db(rel_path, excel_cache, root_excel_folder, db_path):
    """
    단일 엑셀 파일에 대한 String DB 생성
    
    Args:
        rel_path: 엑셀 파일의 상대 경로
        excel_cache: 엑셀 캐시 데이터
        root_excel_folder: 엑셀 루트 폴더 경로
        db_path: 생성할 DB 파일 경로
        
    Returns:
        생성 성공 여부 (Boolean)
    """
    import sqlite3
    import pandas as pd
    from pathlib import Path
    
    # 파일 경로 정규화
    rel_path = PathUtils.normalize_path(rel_path)
    full_path = os.path.join(root_excel_folder, rel_path)
    file = Path(rel_path).name

    # 1. 파일 존재 확인
    if not os.path.exists(full_path):
        logger.error(f"파일이 존재하지 않음: {full_path}")
        return False

    # 2. 캐시에 파일이 있는지 확인
    if rel_path not in excel_cache:
        logger.error(f"캐시에 파일 정보 없음: {rel_path}")
        return False

    # 3. 시트 정보 확인
    sheets = excel_cache.get(rel_path, {}).get("sheets", {})
    if not sheets:
        logger.warning(f"유효한 시트 없음: {rel_path}")
        return False

    # 4. 데이터 유효성 확인
    valid_data = False
    valid_sheets_count = 0
    
    for sheet_name, sheet_meta in sheets.items():
        # 기본 유효성 검사 - 컬럼 정보가 있는지 확인
        if (sheet_meta and 
            isinstance(sheet_meta.get("column_positions"), dict) and 
            len(sheet_meta.get("column_positions", {})) > 0):
            valid_data = True
            valid_sheets_count += 1
            
    if not valid_data:
        logger.warning(f"유효한 데이터 없음: {rel_path}")
        return False
        
    logger.info(f"처리 시작: {file} (유효 시트 {valid_sheets_count}개)")
    
    # 5. DB 디렉토리 및 파일 준비
    try:
        # DB 경로 디렉터리 생성
        db_dir = os.path.dirname(db_path)
        PathUtils.ensure_dir(db_dir)
        
        # 기존 DB 파일이 있으면 삭제
        if os.path.exists(db_path):
            os.remove(db_path)
    except Exception as e:
        logger.error(f"DB 준비 오류: {db_path} - {e}")
        return False

    # 6. 새 DB 생성 및 테이블 설정
    try:
        # DB 생성에는 DBUtils 활용 가능
        columns_def = """
            string_id TEXT, file TEXT, sheet TEXT,
            help TEXT, origin TEXT, request TEXT,
            kr TEXT, en TEXT, cn TEXT, tw TEXT, th TEXT, pt TEXT, es TEXT, de TEXT, fr TEXT, jp TEXT,
            added TEXT, applied TEXT
        """
        
        # 가상 테이블(FTS5)은 DBUtils로 직접 만들기 어려우므로 원래 코드 유지
        conn = sqlite3.connect(db_path)
        cur = conn.cursor()

        cur.execute("""
            CREATE VIRTUAL TABLE string_data USING fts5 (
                string_id, file, sheet,
                help, origin, request,
                kr, en, cn, tw, th, pt, es, de, fr, jp,
                added, applied
            );
        """)

        insert_columns = [
            "string_id", "file", "sheet",
            "help", "origin", "request",
            "kr", "en", "cn", "tw", "th", "pt", "es", "de", "fr", "jp",
            "added", "applied"
        ]
        insert_sql = f"INSERT INTO string_data ({','.join(insert_columns)}) VALUES ({','.join(['?']*len(insert_columns))})"

        # 7. 각 시트별 데이터 처리
        records_inserted = 0
        
        for sheet_name, sheet_meta in sheets.items():
            try:
                header_row = sheet_meta.get("header_row", 0)
                column_positions = sheet_meta.get("column_positions", {})
                
                if not column_positions:
                    logger.warning(f"컬럼 정보 없음: {file} / {sheet_name}")
                    continue

                # ExcelFileManager를 활용하여 엑셀 파일 읽기
                # (기존 방식 유지, 필요시 ExcelFileManager에 적절한 함수 추가)
                df = pd.read_excel(full_path, sheet_name=sheet_name, header=header_row, dtype=str).fillna("")
                
                if df.empty:
                    logger.warning(f"빈 시트: {file} / {sheet_name}")
                    continue
                    
                records = []

                # 각 행 처리
                for _, row in df.iterrows():
                    record = []
                    for col in insert_columns:
                        if col == "file":
                            val = rel_path  # 상대 경로 전체를 저장
                        elif col == "sheet":
                            val = sheet_name
                        else:
                            pos = column_positions.get(col.upper())
                            if pos is not None and 0 <= pos - 1 < len(df.columns):
                                val = row.iloc[pos - 1]
                            else:
                                val = ""
                        record.append(val)
                    records.append(record)

                # DBUtils.insert_many 대신 직접 실행 (FTS5 테이블 특성상)
                if records:
                    cur.executemany(insert_sql, records)
                    records_inserted += len(records)
                    logger.info(f"{file} / {sheet_name}: {len(records)}행 삽입")

            except Exception as e:
                logger.error(f"시트 처리 오류: {file} / {sheet_name} - {e}")
                continue

        # 8. 트랜잭션 커밋 및 연결 종료
        conn.commit()
        conn.close()
        
        # 데이터가 하나도 없는 경우 DB 파일 삭제
        if records_inserted == 0:
            logger.warning(f"삽입된 레코드 없음: {db_path}")
            if os.path.exists(db_path):
                os.remove(db_path)
            return False
            
        logger.info(f"완료: {file} - 총 {records_inserted}행 삽입")
        return True
        
    except Exception as e:
        logger.error(f"DB 생성 오류: {db_path} - {e}")
        # 오류 발생 시 불완전한 DB 파일 삭제
        try:
            if os.path.exists(db_path):
                os.remove(db_path)
        except:
            pass
        return False


def build_or_update_all_string_dbs(excel_cache_path, root_excel_folder, cache_root_dir, changed_files=None, progress_callback=None):
    """
    모든 String 관련 파일의 DB를 구축하거나 업데이트합니다.
    
    Args:
        excel_cache_path: 엑셀 캐시 파일 경로
        root_excel_folder: 엑셀 루트 폴더 경로
        cache_root_dir: 캐시 루트 디렉토리
        changed_files: 변경된 파일 목록 (None이면 모든 파일 처리)
        progress_callback: 진행 상황 콜백 함수
    """
    logger.debug("===== build_or_update_all_string_dbs 시작 =====")
    logger.debug(f"엑셀 캐시 경로: {excel_cache_path}")
    logger.debug(f"루트 엑셀 폴더: {root_excel_folder}")
    logger.debug(f"캐시 루트 디렉터리: {cache_root_dir}")
    
    # string_dbs 디렉터리 생성 
    string_dbs_dir = os.path.join(cache_root_dir, "string_dbs")
    PathUtils.ensure_dir(string_dbs_dir)
    logger.debug(f"string_dbs 디렉터리 생성/확인: {string_dbs_dir}")
    
    db_mtime_path = os.path.join(cache_root_dir, "string_dbs.mtime.json")
    logger.debug(f"DB mtime 경로: {db_mtime_path}")
    
    # 캐시 로드 및 경로 정규화
    cache = FileUtils.load_cached_data(excel_cache_path)
    logger.debug(f"엑셀 캐시 항목 수: {len(cache)}")
    
    cache = {
        PathUtils.normalize_path(rel_path): meta
        for rel_path, meta in cache.items()
    }
    
    # mtime 정보 로드
    mtime_map = FileUtils.load_cached_data(db_mtime_path)
    
    # 처리할 파일 목록 준비
    string_files = get_string_files_to_process(cache, root_excel_folder, changed_files, progress_callback)
    
    # 각 파일 처리
    updated_count = process_string_files(string_files, cache, root_excel_folder, 
                                         cache_root_dir, mtime_map, progress_callback)
    
    # 결과 저장 및 마무리
    FileUtils.save_cache(db_mtime_path, mtime_map)
    
    elapsed_time = int(time.time() - time.time())
    result_msg = f"✅ 파일 단위 DB 갱신 완료: {updated_count}개, {elapsed_time}초 소요"
    logger.debug(result_msg)
    logger.debug("===== build_or_update_all_string_dbs 완료 =====")
    
    if progress_callback:
        progress_callback(result_msg)
    return updated_count


def get_string_files_to_process(cache, root_excel_folder, changed_files=None, progress_callback=None):
    """
    처리해야 할 String 관련 파일 목록을 생성합니다.
    
    Args:
        cache: 엑셀 캐시 데이터
        root_excel_folder: 엑셀 루트 폴더 경로
        changed_files: 변경된 파일 목록 (None이면 모든 String 파일 처리)
        progress_callback: 진행 상황 콜백 함수
        
    Returns:
        처리할 파일 경로 목록
    """
    # 변경된 파일 목록이 제공된 경우, 해당 파일만 처리
    if changed_files and len(changed_files) > 0:
        msg = f"변경된 {len(changed_files)}개 파일만 처리합니다."
        logger.debug(msg)
        if progress_callback:
            progress_callback(f"[String DB] {msg}")
        files_to_process = set(changed_files)
    else:
        msg = "모든 String 관련 파일을 처리합니다."
        logger.debug(msg)
        if progress_callback:
            progress_callback(f"[String DB] {msg}")
        files_to_process = None

    # 처리할 파일 목록 생성
    string_files = []
    for rel_path in cache.keys():
        rel_path = PathUtils.normalize_path(rel_path)
        parts = rel_path.split("/")
        filename = os.path.basename(rel_path)

        # String 관련 파일 필터링
        is_string_file = any(p.startswith("Excel_String") for p in parts) or filename.startswith("String")
        if not is_string_file:
            logger.debug(f"String 파일 아님 (건너뜀): {rel_path}")
            continue

        # 변경된 파일만 처리하도록 필터링
        if files_to_process is not None and rel_path not in files_to_process:
            logger.debug(f"변경되지 않은 파일 (건너뜀): {rel_path}")
            continue
            
        # 파일 존재 확인
        full_path = os.path.join(root_excel_folder, rel_path)
        if not os.path.exists(full_path):
            logger.warning(f"파일 존재하지 않음 (건너뜀): {full_path}")
            continue
            
        string_files.append(rel_path)
    
    # 총 처리할 파일 수 표시
    total_files = len(string_files)
    logger.debug(f"총 처리할 파일 수: {total_files}")
    
    if progress_callback:
        progress_callback(f"[String DB] 총 {total_files}개 파일 처리 예정")
        
    return string_files

def process_string_files(string_files, cache, root_excel_folder, cache_root_dir, 
                        mtime_map, progress_callback=None):
    """
    String 파일 목록을 처리하여 DB를 생성합니다.
    
    Args:
        string_files: 처리할 파일 경로 목록
        cache: 엑셀 캐시 데이터
        root_excel_folder: 엑셀 루트 폴더 경로
        cache_root_dir: 캐시 루트 디렉토리
        mtime_map: 파일 수정 시간 정보
        progress_callback: 진행 상황 콜백 함수
        
    Returns:
        업데이트된 파일 수
    """
    updated_count = 0
    start_time = time.time()
    
    # 각 파일 처리
    for idx, rel_path in enumerate(string_files, 1):
        logger.debug(f"파일 처리 시작 ({idx}/{len(string_files)}): {rel_path}")
        
        if progress_callback:
            progress_callback(f"[String DB] ({idx}/{len(string_files)}) 처리 중: {os.path.basename(rel_path)}")
            
        full_path = os.path.join(root_excel_folder, rel_path)
        
        # 현재 mtime 가져오기
        current_mtime = PathUtils.get_file_mtime(full_path)
        
        # 파일 고유 식별자 사용
        file_id = PathUtils.get_file_identifier(rel_path)
        
        # 마지막 mtime 확인
        last_mtime = mtime_map.get(file_id)
        
        # 변경 감지 및 처리
        if last_mtime == current_mtime:
            logger.debug(f"파일 변경 없음 (건너뜀): {rel_path}")
            continue
            
        # DB 경로 계산 및 생성
        db_path = get_per_file_db_path(rel_path, cache_root_dir)
        logger.debug(f"DB 경로: {db_path}")
        
        # DB 생성
        ok = build_single_string_db(rel_path, cache, root_excel_folder, db_path)
        
        if ok:
            mtime_map[file_id] = current_mtime
            updated_count += 1
    
    elapsed_time = time.time() - start_time
    logger.info(f"처리 완료: {updated_count}개 파일, {elapsed_time:.2f}초 소요")
    return updated_count

def should_rebuild_string_db(cache_path, db_path, mtime_cache_path=None, root_excel_folder=None):
    """
    String DB의 재구축 필요 여부를 확인합니다.
    
    Args:
        cache_path: 엑셀 캐시 파일 경로
        db_path: DB 파일 경로
        mtime_cache_path: mtime 캐시 파일 경로 (기본값: db_path와 동일 디렉토리의 string_dbs.mtime.json)
        root_excel_folder: 엑셀 루트 폴더 경로 (선택적)
        
    Returns:
        재구축 필요 여부와 변경된 파일 목록을 포함한 dict 또는 단순 bool
    """
    # 타입 검사 추가
    if not isinstance(cache_path, str) or not isinstance(db_path, str):
        logger.error(f"경로 타입 오류: cache_path={type(cache_path)}, db_path={type(db_path)}")
        cache_path = str(cache_path) if cache_path is not None else ""
        db_path = str(db_path) if db_path is not None else ""

    # mtime_cache_path가 없으면 기본값 생성
    if mtime_cache_path is None:
        mtime_cache_path = os.path.join(os.path.dirname(db_path), "string_dbs.mtime.json")
        
    # 기본 검사: 필수 파일 및 디렉토리 확인
    if not _check_required_files(cache_path, db_path, mtime_cache_path):
        return True
        
    # 캐시 데이터 로드
    excel_cache = FileUtils.load_cached_data(cache_path)
    old_mtimes = FileUtils.load_cached_data(mtime_cache_path)
    
    # 엑셀 기본 경로 확인
    excel_base = _get_excel_base_path(root_excel_folder, cache_path)
    
    # 파일 변경 확인
    changed_files = _check_file_changes(excel_cache, excel_base, old_mtimes)
    
    # 변경사항에 따른 결과 반환
    return _prepare_rebuild_result(changed_files, mtime_cache_path, old_mtimes)

def _check_required_files(cache_path, db_path, mtime_cache_path):
    """필수 파일 및 디렉토리 존재 여부 확인"""
    string_dbs_dir = os.path.dirname(db_path)
    
    if not os.path.exists(string_dbs_dir):
        logger.debug(f"string_dbs 디렉터리가 존재하지 않음: {string_dbs_dir}")
        PathUtils.ensure_dir(string_dbs_dir)
        return False
        
    if not os.path.exists(cache_path):
        logger.debug(f"엑셀 캐시 파일이 존재하지 않음: {cache_path}")
        return False
        
    if not os.path.exists(mtime_cache_path):
        logger.debug(f"mtime 캐시 파일이 존재하지 않음: {mtime_cache_path}")
        return False
        
    return True

def _get_excel_base_path(root_excel_folder, cache_path):
    """엑셀 기본 경로 결정"""
    if root_excel_folder and os.path.exists(root_excel_folder):
        return root_excel_folder
    else:
        excel_dir = os.path.dirname(cache_path)
        return os.path.dirname(excel_dir)

def _check_file_changes(cache, excel_base, old_mtimes):
    """파일 변경 사항 확인"""
    changed_files = []
    
    for rel_path, meta in cache.items():
        rel_path = PathUtils.normalize_path(rel_path)
        parts = rel_path.split("/")
        filename = os.path.basename(rel_path)

        # String 관련 파일만 처리
        if not (any(p.startswith("Excel_String") for p in parts) or filename.startswith("String")):
            continue

        # 파일 경로 확인
        full_path = os.path.join(excel_base, rel_path)
        if not os.path.exists(full_path):
            # 상대 경로로 다시 시도
            alt_path = os.path.join(os.path.dirname(os.path.dirname(cache_path)), rel_path)
            if os.path.exists(alt_path):
                full_path = alt_path
                logger.debug(f"대체 경로 사용: {full_path}")
            else:
                logger.debug(f"파일을 찾을 수 없음: {rel_path}")
                changed_files.append(f"{rel_path} (파일 없음)")
                continue

        # mtime 비교
        try:
            current_mtime = PathUtils.get_file_mtime(full_path)
            file_id = PathUtils.get_file_identifier(rel_path)
            cached_mtime = old_mtimes.get(file_id)
            
            if cached_mtime is None:
                logger.debug(f"새 파일 발견: {rel_path}")
                changed_files.append(f"{rel_path} (새 파일)")
            elif cached_mtime != current_mtime:
                logger.debug(f"파일 변경 감지: {rel_path}")
                changed_files.append(f"{rel_path} (mtime 변경)")
        except Exception as e:
            logger.error(f"mtime 확인 중 오류: {rel_path} - {e}")
            changed_files.append(f"{rel_path} (오류: {str(e)})")
    
    return changed_files

def _prepare_rebuild_result(changed_files, mtime_cache_path, old_mtimes):
    """재구축 결과 준비"""
    if changed_files:
        logger.debug(f"재구축 필요: {len(changed_files)}개 파일 변경됨")
        
        # 로그에 처음 5개 변경 파일만 표시
        for i, file in enumerate(changed_files[:5], 1):
            logger.debug(f"  {i}. {file}")
        
        if len(changed_files) > 5:
            logger.debug(f"  ... 외 {len(changed_files) - 5}개 파일")
        
        # mtime 업데이트 (last_check만)
        try:
            updated_mtimes = old_mtimes.copy()
            updated_mtimes["last_check"] = int(time.time())
            FileUtils.save_cache(mtime_cache_path, updated_mtimes)
        except Exception as e:
            logger.error(f"mtime 캐시 저장 실패: {e}")
        
        # 변경된 파일 정보가 있는 딕셔너리 반환
        return {
            "rebuild_needed": True,
            "changed_files": [f.split(" (")[0] for f in changed_files]  # 괄호 내용 제거
        }
    else:
        logger.debug("모든 파일이 최신 상태, 재구축 불필요")
        return {"rebuild_needed": False}
    

def load_or_build_string_db(excel_cache_path, db_path, root_excel_folder, progress_callback=None):
    """
    캐시 기반 DB 생성/갱신 및 mtime 기록
    개선: 프리셋 전환 시나리오 지원
    
    Args:
        excel_cache_path: 엑셀 캐시 파일 경로
        db_path: DB 파일 경로
        root_excel_folder: 엑셀 파일 루트 폴더
        progress_callback: 진행 상황 콜백 함수 (선택적)
    """
    # 타입 검사 및 경로 정규화
    if not isinstance(db_path, str):
        err_msg = f"db_path 타입 오류: {type(db_path)}"
        logger.error(err_msg)
        print(f"[오류] {err_msg}")
        db_path = str(db_path)  # 문자열로 변환 시도
    
    # 콜백 함수가 제공된 경우 진행 상황 업데이트
    if progress_callback:
        progress_callback("FTS5 DB 초기화 중...")
    
    # 경로 및 디렉토리 구성
    string_dbs_dir = os.path.join(os.path.dirname(db_path), "string_dbs")
    PathUtils.ensure_dir(string_dbs_dir)
    
    mtime_cache_path = os.path.join(os.path.dirname(db_path), "string_dbs.mtime.json")
    logger.debug(f"mtime 캐시 경로: {mtime_cache_path}")

    # 재구축 필요 여부 확인
    change_info = should_rebuild_string_db(
        excel_cache_path,
        db_path,
        mtime_cache_path,
        root_excel_folder=root_excel_folder
    )
    
    # 이전 버전 호환성 처리
    if isinstance(change_info, bool):
        rebuild_needed = change_info
        changed_files = []
    else:
        rebuild_needed = change_info.get("rebuild_needed", False)
        changed_files = change_info.get("changed_files", [])
    
    if not rebuild_needed:
        logger.info("이미 최신 상태, DB 로딩 완료")
        if progress_callback:
            progress_callback("[String DB] 이미 최신 상태, DB 로딩 완료")
        return
    
    logger.info(f"재구축 필요: {len(changed_files)}개 파일 변경됨")
    
    if progress_callback:
        progress_callback(f"[String DB] 재구축 필요: {len(changed_files)}개 파일 변경됨")
    
    logger.info("DB 재구축 시작")
    
    if progress_callback:
        progress_callback("[String DB] DB 재구축 시작...")
    
    # 캐시 로드
    excel_cache = FileUtils.load_cached_data(excel_cache_path)
    
    # 기존 mtime 정보 로드
    mtime_map = FileUtils.load_cached_data(mtime_cache_path)
    
    # 변경된 파일만 처리 또는 전체 처리
    if changed_files:
        files_to_process = changed_files
        msg = f"변경된 {len(files_to_process)}개 파일만 처리합니다."
    else:
        files_to_process = []
        for rel_path in excel_cache.keys():
            rel_path = PathUtils.normalize_path(rel_path)  # 경로 정규화
            parts = rel_path.split("/")
            filename = os.path.basename(rel_path)

            # String 관련 파일만 처리
            is_string_file = any(p.startswith("Excel_String") for p in parts) or filename.startswith("String")
            if is_string_file:
                files_to_process.append(rel_path)
        
        msg = f"총 {len(files_to_process)}개 String 관련 파일을 처리합니다."
    
    logger.info(msg)
    if progress_callback:
        progress_callback(f"[String DB] {msg}")
    
    # 각 파일 처리
    updated_count = 0
    start_time = time.time()
    
    for idx, rel_path in enumerate(files_to_process, 1):
        logger.debug(f"파일 처리 시작 ({idx}/{len(files_to_process)}): {rel_path}")
        
        if progress_callback:
            progress_callback(f"[String DB] ({idx}/{len(files_to_process)}) 처리 중: {os.path.basename(rel_path)}")
            
        full_path = os.path.join(root_excel_folder, rel_path)
        
        # 파일 존재 확인
        if not os.path.exists(full_path):
            logger.warning(f"파일이 존재하지 않음: {full_path}")
            continue
            
        current_mtime = PathUtils.get_file_mtime(full_path)
        
        # 파일 고유 식별자 사용
        file_id = PathUtils.get_file_identifier(rel_path)
        
        # DB 경로 구성
        db_file = f"{os.path.splitext(file_id)[0]}.db"
        db_file_path = os.path.join(string_dbs_dir, db_file)
        
        # DB 생성 시작
        ok = build_single_string_db(rel_path, excel_cache, root_excel_folder, db_file_path)
        
        if ok:
            mtime_map[file_id] = current_mtime  # 파일명 기반으로 mtime 저장
            updated_count += 1

    # mtime 정보 저장
    FileUtils.save_cache(mtime_cache_path, mtime_map)
    
    # 결과 로깅
    elapsed_time = int(time.time() - start_time)
    result_msg = f"✅ 파일 단위 DB 갱신 완료: {updated_count}개, {elapsed_time}초 소요"
    logger.info(result_msg)
    
    if progress_callback:
        progress_callback(result_msg)


def search_string_db(
    db_path: str,
    keyword: str,
    columns: list[str],
    match_exact: bool = False,
    match_case: bool = False,
    match_word: bool = False,
    use_regex: bool = False
) -> list[dict]:
    """
    FTS5 DB를 이용하여 문자열 검색
    - columns: KR, EN, CN, TW 등 검색 대상 컬럼
    - 반환: file, sheet, string_id, 매칭된 컬럼들의 dict 리스트
    """
    import os
    
    # string_dbs 디렉터리 체크
    string_dbs_dir = os.path.join(os.path.dirname(db_path), "string_dbs")
    if os.path.exists(string_dbs_dir):
        # 메인 DB 파일이 있는지 확인
        if not os.path.exists(db_path):
            # string_dbs 디렉터리 내의 모든 DB 검색
            return search_all_string_dbs(
                keyword=keyword,
                columns=columns,
                db_dir=os.path.dirname(db_path),
                match_exact=match_exact,
                match_case=match_case,
                match_word=match_word,
                use_regex=use_regex
            )

    # 검색 컬럼 필터 생성
    conditions = []
    keyword_query = keyword

    if not match_case:
        keyword_query = keyword_query.lower()

    for col in columns:
        if match_exact:
            # 완전 일치
            conditions.append(f"{col} = ?")
        elif use_regex:
            # SQLite는 regex 지원 X → fallback X
            # 실제 사용 시 pandas fallback 필요
            pass
        elif match_word:
            # 단어 매칭
            import re
            pattern = rf'\b{re.escape(keyword_query)}\b'
            conditions.append(f"{col} MATCH ?")
            keyword_query = pattern
        else:
            # 부분일치
            conditions.append(f"{col} MATCH ?")

    where_clause = " OR ".join(conditions)
    query = f"""
        SELECT file, sheet, string_id, kr, en, cn, tw, th, pt, es, de, fr, jp
        FROM string_data
        WHERE {where_clause}
    """

    values = [keyword_query] * len(conditions)

    # DBUtils.execute_query 사용
    rows = DBUtils.execute_query(db_path, query, values)
    if rows is None:
        return []

    result = []
    for row in rows:
        row_dict = {
            "file": row[0],
            "sheet": row[1],
            "STRING_ID": row[2],
            "KR": row[3], "EN": row[4], "CN": row[5], "TW": row[6],
            "TH": row[7], "PT": row[8], "ES": row[9], "DE": row[10],
            "FR": row[11], "JP": row[12],
        }

        # 매칭 컬럼 체크
        matched_columns = []
        for col in columns:
            cell = row_dict.get(col.upper(), "")
            if keyword.lower() in str(cell).lower():
                matched_columns.append(col.upper())

        row_dict["matched"] = matched_columns
        result.append(row_dict)

    return result

def search_all_string_dbs(
    keyword: str,
    columns: list[str],
    db_dir: str,
    match_exact: bool = False,
    match_case: bool = False,
    match_word: bool = False,
    use_regex: bool = False
) -> list[dict]:
    """
    지정된 디렉토리의 모든 DB 파일에서 문자열 검색
    
    Args:
        keyword: 검색할 키워드
        columns: 검색할 컬럼 목록
        db_dir: DB 파일이 있는 디렉토리
        match_exact: 정확히 일치하는 항목만 검색
        match_case: 대소문자 구분
        match_word: 단어 단위로 검색
        use_regex: 정규식 사용
        
    Returns:
        검색 결과 딕셔너리 목록
    """
    import os
    
    # string_dbs 디렉터리 확인
    string_dbs_dir = os.path.join(db_dir, "string_dbs")
    if os.path.exists(string_dbs_dir):
        db_dir = string_dbs_dir  # string_dbs 디렉터리가 있으면 사용
    
    results = []
    db_files = [f for f in os.listdir(db_dir) if f.endswith(".db")]
    keyword_processed = keyword if match_case else keyword.lower()

    for db_file in db_files:
        db_path = os.path.join(db_dir, db_file)
        
        # 각 DB 파일에서 검색
        try:
            # 검색 조건 생성
            where_conditions = []
            values = []

            # fallback to `LIKE` for partial match
            for col in columns:
                if match_exact:
                    where_conditions.append(f"{col} = ?")
                    values.append(keyword_processed)
                elif match_word:
                    pattern = rf"% {keyword_processed} %"  # weak workaround
                    where_conditions.append(f"{col} LIKE ?")
                    values.append(pattern)
                elif use_regex:
                    # SQLite doesn't support regex natively
                    # fallback handled in result filter
                    where_conditions.append(f"{col} LIKE ?")
                    values.append(f"%{keyword_processed}%")
                else:
                    where_conditions.append(f"{col} LIKE ?")
                    values.append(f"%{keyword_processed}%")

            sql = f"""
                SELECT file, sheet, string_id, kr, en, cn, tw, th, pt, es, de, fr, jp
                FROM string_data
                WHERE {" OR ".join(where_conditions)}
            """
            
            # DBUtils.execute_query 사용
            rows = DBUtils.execute_query(db_path, sql, values)
            if rows is None:
                continue

            # 결과 처리
            for row in rows:
                result = {
                    "file": row[0],
                    "sheet": row[1],
                    "STRING_ID": row[2],
                    "KR": row[3], "EN": row[4], "CN": row[5], "TW": row[6],
                    "TH": row[7], "PT": row[8], "ES": row[9], "DE": row[10],
                    "FR": row[11], "JP": row[12]
                }
                
                # 매칭된 컬럼 확인
                matched = []
                for col in columns:
                    cell = result.get(col.upper(), "")
                    comp = cell if match_case else str(cell).lower()
                    if keyword_processed in comp:
                        matched.append(col.upper())
                
                result["matched"] = matched
                results.append(result)

        except Exception as e:
            logger.error(f"검색 실패: {db_file} - {e}")

    return results