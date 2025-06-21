import json
import os
from utils.common_utils import FileUtils
from utils.excel_utils import ExcelFileManager

# 프로그램 실행 디렉토리 기준 설정 파일 위치 지정
CONFIG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")
HISTORY_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "search_history.json")

def load_search_history(kind="string"):
    """검색 히스토리 로드"""
    path = f"search_history_{kind}.json"
    return FileUtils.load_json(path, [])

def save_search_history(history, kind="string"):
    """검색 히스토리 저장"""
    path = f"search_history_{kind}.json"
    return FileUtils.save_json(path, history)

def load_config(filename=CONFIG_FILE):
    """설정 파일 로드 (오류 시 기본값 반환)"""
    try:
        print(f"설정 파일 로드 시도: {filename}")
        if os.path.exists(filename):
            print(f"설정 파일 존재함: {filename}")
            with open(filename, "r", encoding="utf-8") as f:
                config = json.load(f)
                
                # 경로 정규화 처리
                if "data_path" in config and config["data_path"]:
                    config["data_path"] = os.path.normpath(config["data_path"])
                if "db_path" in config and config["db_path"]:
                    config["db_path"] = os.path.normpath(config["db_path"])
                
                # 프리셋 경로도 정규화
                if "presets" in config:
                    for preset_id, preset in config["presets"].items():
                        if "data" in preset and preset["data"]:
                            preset["data"] = os.path.normpath(preset["data"])
                
                print(f"설정 파일 로드 성공 (정규화 후): {config}")
                return config
        else:
            print(f"설정 파일이 존재하지 않음: {filename}")
    except Exception as e:
        print(f"설정 파일 로드 오류: {e}")
        import traceback
        traceback.print_exc()
    
    # 기본 설정 - 프리셋 구조 포함
    default_config = {"data_path": "", "presets": {}}
    print(f"기본 설정 반환: {default_config}")
    return default_config

def save_config(filename, config):
    """설정 파일 저장 (백업 생성)"""
    try:
        # 파일 경로 처리 - 상대 경로인 경우 현재 디렉토리 기준으로 처리
        if not os.path.dirname(filename):
            # 파일명만 있는 경우 (예: 'config.json')
            full_path = os.path.join(os.getcwd(), filename)
        else:
            # 경로가 포함된 경우
            full_path = os.path.abspath(filename)
        
        # 백업 생성
        if os.path.exists(full_path):
            backup_file = f"{full_path}.bak"
            try:
                import shutil
                shutil.copy2(full_path, backup_file)
            except Exception as backup_err:
                print(f"설정 백업 실패: {backup_err}")
        
        # 디렉토리가 존재하는지 확인하고 생성 (경로가 있는 경우만)
        dir_path = os.path.dirname(full_path)
        if dir_path:
            os.makedirs(dir_path, exist_ok=True)
        
        # 설정 저장 - 전체 경로 사용
        with open(full_path, "w", encoding="utf-8") as f:
            json.dump(config, f, indent=2, ensure_ascii=False)
        
        print(f"설정 저장 완료: {full_path}")
        return True
    except Exception as e:
        print(f"설정 파일 저장 오류: {e}")
        import traceback
        traceback.print_exc()
        return False

def get_preset_paths(config):
    """프리셋 경로 정보 반환"""
    return config.get("presets", {})

def open_excel_from_result(tree, item):
    """트리뷰에서 선택한 항목의 엑셀 파일 열기"""
    if not item:
        return
    tags = tree.item(item, "tags")
    if not tags:
        return

    path = tags[0]
    values = tree.item(item, "values")
    if len(values) < 3:
        return

    sheet, string_id = values[1], values[2]
    ExcelFileManager.highlight_excel_by_value(path, sheet, "STRING_ID", string_id)

def open_excel_and_highlight(path, sheet, column_name, target_value, excel_cache=None):
    """
    엑셀 파일을 열고 특정 값을 강조 표시합니다.
    
    Args:
        path: 엑셀 파일 경로
        sheet: 시트 이름
        column_name: 컬럼 이름
        target_value: 찾을 값
        excel_cache: 엑셀 캐시 (선택적)
        
    Returns:
        성공 여부 (Boolean)
    """
    return ExcelFileManager.highlight_excel_by_value(path, sheet, column_name, target_value, excel_cache)