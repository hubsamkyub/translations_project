import logging
from utils.common_utils import PathUtils, ExcelUtils, show_message

def normalize_file_path(file_path):
    """파일 경로 정규화"""
    return PathUtils.normalize_path(file_path)

def get_display_name(file_path):
    """파일 경로에서 표시용 파일명 추출"""
    return PathUtils.get_display_name(file_path)

def find_excel_file(base_path, file_name):
    """엑셀 파일 검색"""
    return ExcelUtils.find_excel_file(base_path, file_name)