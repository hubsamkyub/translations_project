# build.spec

# -*- mode: python ; coding: utf-8 -*-

import sys

# PyInstaller가 실행되는 환경의 인코딩을 UTF-8로 설정 (한글 경로 문제 방지)
sys.stdout.reconfigure(encoding='utf-8')

# --- 분석 단계 ---
# 이 단계에서 스크립트의 의존성을 분석하고 필요한 파일들을 찾아냅니다.
a = Analysis(
    ['translation_main.py'],  #<-- 1. 메인 실행 파일
    pathex=['translation_manager_0621'],  #<-- 2. 모듈 검색 경로 추가 (가장 중요)
    binaries=[],
    datas=[
        # (원본 파일 경로, 빌드 후 폴더 내 위치)
        # --- 3. 프로그램 실행에 필요한 데이터 파일들을 여기에 추가합니다. ---
        ('tools/config.json', 'tools'),
        ('utils/config.json', 'utils')
        # 아래 두 파일은 실제 위치에 맞게 경로를 수정해야 할 수 있습니다.
        # ('translation_manager_0621/tools/string_exception_rules.json', 'tools'),
        # ('translation_manager_0621/Excel_Indent.json', '.')
    ],
    hiddenimports=[
        # --- 4. PyInstaller가 자동으로 찾지 못할 수 있는 숨겨진 의존성을 추가합니다. ---
        'win32com.client',
        'openpyxl.cell._writer',
        'pandas.io.excel._openpyxl'
    ],
    hookspath=[],
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=None,
    noarchive=False
)

# --- 실행 파일 생성 단계 ---
pyz = PYZ(a.pure, a.zipped_data, cipher=None)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='TranslationAutomationTool',  #<-- 5. 생성될 .exe 파일의 이름
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,  # UPX가 설치된 경우 실행 파일 압축 (파일 크기 감소)
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  #<-- 6. GUI 애플리케이션이므로 콘솔 창을 띄우지 않음
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None #<-- 아이콘 파일(.ico) 경로를 지정할 수 있습니다.
)

# --- 최종 폴더 구성 단계 ---
# 위에서 생성된 파일들을 하나의 폴더로 모읍니다.
coll = COLLECT(
    exe,
    a.datas,
    a.binaries,
    a.zipfiles,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='TranslationAutomationTool' #<-- 7. 결과물이 담길 폴더 이름
)