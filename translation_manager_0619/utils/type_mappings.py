# utils/type_mappings.py

import os
import json
import logging

# 로거 설정
logger = logging.getLogger(__name__)

# 타입 매핑 정보
REWARD_TYPE_MAPPINGS = {
    # HeroTemplate 관련
    "10": {"table": "HeroTemplate", "column": "BaseHeroID", "description": "영웅"},
    "11": {"table": "HeroTemplate", "column": "BaseHeroID", "description": "영웅 조각"},
    
    # ItemTemplate 관련
    "20": {"table": "ItemTemplate", "column": "TemplateID", "description": "아이템"},
    "21": {"table": "ItemTemplate", "column": "TemplateID", "description": "장비"},
    
    # 기타 리소스
    "30": {"table": "GoodsMaxValue", "column": "GoodsType", "description": "재화"},
    "40": {"table": "BoxTemplate", "column": "ItemTID", "description": "상자"},
    "41": {"table": "BoxTemplate", "column": "ItemTID", "description": "선택 상자"},
    "50": {"table": "TicketMaxValue", "column": "TicketType", "description": "티켓"},
    "70": {"table": "WisdomBookTemplate", "column": "TemplateID", "description": "도감"},
    "80": {"table": "CostumeTemplate", "column": "TemplateID", "description": "코스튬"}
}

# 다른 타입 매핑들 (필요에 따라 추가)
ITEM_TYPE_MAPPINGS = {
    "6000": {"description": "랜덤박스", "category": "box"},
    "6001": {"description": "패키지", "category": "package"},
    "6002": {"description": "선택박스", "category": "select_box"}
    # 추가 아이템 타입...
}

# 다른 enum 매핑들
# HERO_TYPE_MAPPINGS = {...}
# COSTUME_TYPE_MAPPINGS = {...}

def get_mapping_file_path(file_name, search_folders=None):
    """
    매핑 파일의 경로를 찾습니다.
    
    Args:
        file_name (str): 매핑 파일 이름 (예: "typecode_mapping.json")
        search_folders (list, optional): 검색할 폴더 경로 리스트
        
    Returns:
        str: 찾은 파일 경로 또는 None (파일을 찾지 못한 경우)
    """
    if search_folders is None:
        # 기본 검색 경로
        search_folders = [
            os.getcwd(),  # 현재 디렉토리
            os.path.dirname(os.path.abspath(__file__)),  # 현재 모듈 디렉토리
            os.path.dirname(os.path.dirname(os.path.abspath(__file__))),  # 상위 디렉토리
            os.path.join(os.path.expanduser("~"), "Documents")  # 사용자 Documents 폴더
        ]
    
    for folder in search_folders:
        path = os.path.join(folder, file_name)
        if os.path.exists(path):
            logger.debug(f"매핑 파일 발견: {path}")
            return path
    
    logger.warning(f"매핑 파일을 찾을 수 없음: {file_name}")
    return None

def load_mappings_from_file(file_name, default_mappings=None):
    """
    JSON 파일에서 매핑 정보를 로드합니다.
    
    Args:
        file_name (str): 매핑 파일 이름
        default_mappings (dict, optional): 파일이 없을 경우 사용할 기본 매핑
        
    Returns:
        dict: 로드된 매핑 정보
    """
    if default_mappings is None:
        default_mappings = {}
    
    file_path = get_mapping_file_path(file_name)
    if file_path:
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                mappings = json.load(f)
                logger.debug(f"매핑 파일 로드 성공: {len(mappings)} 항목")
                return mappings
        except Exception as e:
            logger.error(f"매핑 파일 로드 오류: {e}")
    
    logger.info(f"기본 매핑 사용: {len(default_mappings)} 항목")
    return default_mappings

def get_table_for_type(type_code, type_category="reward"):
    """
    주어진 타입 코드에 해당하는 테이블 정보를 반환합니다.
    
    Args:
        type_code (str|int): 타입 코드
        type_category (str): 타입 카테고리 ("reward", "item" 등)
        
    Returns:
        dict: 테이블 정보 (table, column, description) 또는 None
    """
    # 타입 코드 문자열로 변환
    type_code_str = str(type_code)
    
    # 타입 카테고리에 따라 적절한 매핑 사용
    if type_category == "reward":
        # RewardType 매핑 로드 시도
        custom_mappings = load_mappings_from_file("reward_type_mappings.json", REWARD_TYPE_MAPPINGS)
        return custom_mappings.get(type_code_str)
    
    elif type_category == "item":
        # ItemType 매핑 로드 시도
        custom_mappings = load_mappings_from_file("item_type_mappings.json", ITEM_TYPE_MAPPINGS)
        return custom_mappings.get(type_code_str)
    
    # 다른 카테고리들...
    
    logger.warning(f"알 수 없는 타입 카테고리: {type_category}, 코드: {type_code}")
    return None

def get_table_name_for_type(type_code, type_category="reward"):
    """
    주어진 타입 코드에 해당하는 테이블 이름만 반환합니다.
    
    Args:
        type_code (str|int): 타입 코드
        type_category (str): 타입 카테고리 ("reward", "item" 등)
        
    Returns:
        str: 테이블 이름 또는 None
    """
    mapping = get_table_for_type(type_code, type_category)
    if mapping and "table" in mapping:
        return mapping["table"]
    return None

def get_column_for_type(type_code, type_category="reward"):
    """
    주어진 타입 코드에 해당하는 ID 컬럼 이름을 반환합니다.
    
    Args:
        type_code (str|int): 타입 코드
        type_category (str): 타입 카테고리 ("reward", "item" 등)
        
    Returns:
        str: 컬럼 이름 또는 None
    """
    mapping = get_table_for_type(type_code, type_category)
    if mapping and "column" in mapping:
        return mapping["column"]
    return None

def get_description_for_type(type_code, type_category="reward"):
    """
    주어진 타입 코드에 해당하는 설명을 반환합니다.
    
    Args:
        type_code (str|int): 타입 코드
        type_category (str): 타입 카테고리 ("reward", "item" 등)
        
    Returns:
        str: 설명 또는 None
    """
    mapping = get_table_for_type(type_code, type_category)
    if mapping and "description" in mapping:
        return mapping["description"]
    return f"타입-{type_code}"

def resolve_type_info(type_code, id_value, type_category="reward"):
    """
    타입 코드와 ID를 사용하여 항목 정보를 조회합니다.
    
    Args:
        type_code (str|int): 타입 코드
        id_value (str|int): ID 값
        type_category (str): 타입 카테고리 ("reward", "item" 등)
        
    Returns:
        dict: 항목 정보 (table, name, id, description 등)
    """
    # 기본 정보
    result = {
        "type": str(type_code),
        "id": str(id_value),
        "description": get_description_for_type(type_code, type_category),
        "table": get_table_name_for_type(type_code, type_category)
    }
    
    # 테이블 정보가 없으면 여기서 종료
    if not result["table"]:
        result["name"] = f"{result['description']}: {id_value}"
        return result
    
    # 여기서 DB 또는 Excel에서 실제 이름 조회 가능 (선택적)
    # DB 조회 로직은 호출자 쪽에서 구현하는 것이 더 적합할 수 있음
    
    return result