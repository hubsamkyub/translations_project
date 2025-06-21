import sqlite3
import os
import pandas as pd
import re
import os.path

class DBCompareManager:
    def __init__(self, parent_window=None):
        self.parent = parent_window
        self.compare_results = []  # 비교 결과 저장

    def extract_file_sheet_from_table(self, table_name):
        """테이블 이름에서 파일명과 시트명 추출"""
        parts = table_name.split('_', 1)
        if len(parts) > 1:
            file_name = parts[0] + ".xlsx"
            sheet_name = parts[1] if "_" in table_name else table_name
        else:
            file_name = table_name + ".xlsx"
            sheet_name = table_name
        
        return file_name, sheet_name

    def compare_all_databases(self, db_pairs, changed_kr=True, new_items=True, deleted_items=True, progress_callback=None):
        """선택된 모든 DB 파일을 비교"""
        if not db_pairs:
            return {"status": "error", "message": "비교할 DB 파일 목록이 없습니다."}
        
        # 결과 초기화
        self.compare_results = []
        
        # 비교 옵션 가져오기
        total_changes = 0
        success_count = 0
        error_list = []
        
        # 각 DB 파일 쌍에 대해 비교 실행
        for idx, db_pair in enumerate(db_pairs):
            if progress_callback:
                progress_callback(f"DB 비교 중: {db_pair['file_name']} ({idx+1}/{len(db_pairs)})",
                                 idx+1, len(db_pairs))
            
            try:
                # 단일 DB 쌍 비교
                changes = self.compare_db_pair(
                    db_pair['original_path'], 
                    db_pair['compare_path'], 
                    changed_kr, 
                    new_items, 
                    deleted_items
                )
                total_changes += changes
                success_count += 1
                
                if progress_callback:
                    progress_callback(f"{db_pair['file_name']} 비교 완료: {changes}개 변경사항 발견",
                                    idx+1, len(db_pairs))
                
            except Exception as e:
                error_msg = f"DB 비교 실패 ({db_pair['file_name']}): {e}"
                error_list.append(error_msg)
                
                if progress_callback:
                    progress_callback(error_msg, idx+1, len(db_pairs))
            
        return {
            "status": "success",
            "success_count": success_count,
            "total_count": len(db_pairs),
            "total_changes": total_changes,
            "errors": error_list,
            "compare_results": self.compare_results
        }

    def compare_db_pair(self, original_db_path, compare_db_path, changed_kr=True, new_items=True, deleted_items=True):
        """단일 DB 쌍 비교"""
        changes_count = 0
        
        try:
            # DB 연결
            conn = sqlite3.connect(original_db_path)
            cursor = conn.cursor()
            
            # 비교 DB를 연결
            cursor.execute(f"ATTACH DATABASE '{compare_db_path}' AS compare_db;")
            
            # 테이블 목록 가져오기
            cursor.execute("SELECT name FROM main.sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%'")
            original_tables = [table[0] for table in cursor.fetchall()]
            
            cursor.execute("SELECT name FROM compare_db.sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%'")
            compare_tables = [table[0] for table in cursor.fetchall()]
            
            # 공통 테이블 찾기
            common_tables = set(original_tables).intersection(set(compare_tables))
            
            # DB 파일명 추출
            db_file_name = os.path.basename(original_db_path)
            
            # 각 테이블별로 비교
            for table in common_tables:
                if table.startswith("String") and not table.startswith("sqlite"):
                    # 테이블 비교 및 결과 추가
                    table_changes = self.compare_table_for_pair(conn, table, db_file_name, changed_kr, new_items, deleted_items)
                    changes_count += table_changes
            
            # 원본에만 있는 테이블 처리 (신규 테이블)
            if new_items:
                new_tables = set(original_tables) - set(compare_tables)
                for table in new_tables:
                    if table.startswith("String") and not table.startswith("sqlite"):
                        table_changes = self.process_unique_table_for_pair(conn, table, 'main', '신규 테이블', db_file_name)
                        changes_count += table_changes
            
            # 비교본에만 있는 테이블 처리 (삭제된 테이블)
            if deleted_items:
                deleted_tables = set(compare_tables) - set(original_tables)
                for table in deleted_tables:
                    if table.startswith("String") and not table.startswith("sqlite"):
                        table_changes = self.process_unique_table_for_pair(conn, table, 'compare_db', '삭제된 테이블', db_file_name)
                        changes_count += table_changes
            
            # DB 분리 및 연결 종료
            cursor.execute("DETACH DATABASE compare_db;")
            conn.close()
            
            return changes_count
        
        except Exception as e:
            # 연결이 열려있으면 닫기 시도
            try:
                cursor.execute("DETACH DATABASE IF EXISTS compare_db;")
                conn.close()
            except:
                pass
            raise e


    def compare_table_for_pair(self, conn, table, db_file_name, changed_kr=True, new_items=True, deleted_items=True):
        """DB 쌍의 테이블 비교 (신규/삭제 표시 방식 수정)"""
        cursor = conn.cursor()
        changes_count = 0
        
        try:
            # 테이블 구조 확인 - 원본 테이블
            cursor.execute(f"PRAGMA main.table_info('{table}')")
            original_columns = {row[1]: row[0] for row in cursor.fetchall()}
            
            # 테이블 구조 확인 - 비교 테이블
            cursor.execute(f"PRAGMA compare_db.table_info('{table}')")
            compare_columns = {row[1]: row[0] for row in cursor.fetchall()}
            
            # 필요한 컬럼이 양쪽 다 있는지 확인
            has_string_id = 'STRING_ID' in original_columns and 'STRING_ID' in compare_columns
            has_kr = 'KR' in original_columns and 'KR' in compare_columns
            
            if not has_string_id or not has_kr:
                return 0
            
            # 변경된 KR 값 찾기 (비교본에 있지만 원본과 값이 다른 항목)
            if changed_kr:
                query = f"""
                SELECT b.STRING_ID, a.KR, b.KR
                FROM main."{table}" a
                INNER JOIN compare_db."{table}" b ON a.STRING_ID = b.STRING_ID
                WHERE a.KR != b.KR AND a.KR IS NOT NULL AND b.KR IS NOT NULL
                """
                
                cursor.execute(query)
                results = cursor.fetchall()
                
                # 테이블 이름에서 파일명과 시트명 추출
                file_name, sheet_name = self.extract_file_sheet_from_table(table)
                
                # 변경 후
                for string_id, original_kr, compare_kr in results:
                    self.compare_results.append({
                        "file_name": db_file_name,
                        "sheet_name": sheet_name,
                        "type": "변경됨",
                        "string_id": string_id,
                        "kr": compare_kr,            # 비교본(새 버전)의 KR 값
                        "original_kr": original_kr  # 원본(이전 버전)의 KR 값
                    })
                    changes_count += 1
                            
            # 원본에만 있는 항목 찾기 (신규) - 🔧 수정
            if new_items:
                query = f"""
                SELECT a.STRING_ID, a.KR
                FROM main."{table}" a
                LEFT JOIN compare_db."{table}" b ON a.STRING_ID = b.STRING_ID
                WHERE b.STRING_ID IS NULL AND a.KR IS NOT NULL
                """
                
                cursor.execute(query)
                results = cursor.fetchall()
                
                # 테이블 이름에서 파일명과 시트명 추출
                file_name, sheet_name = self.extract_file_sheet_from_table(table)
                
                for string_id, kr in results:
                    self.compare_results.append({
                        "file_name": db_file_name,
                        "sheet_name": sheet_name,
                        "string_id": string_id,
                        "kr": "",               # 🔧 수정: 신규는 빈 값
                        "original_kr": kr,      # 🔧 수정: 원본 값은 original_kr에
                        "type": "신규"
                    })
                    changes_count += 1
            
            # 비교본에만 있는 항목 찾기 (삭제) - 기존 유지
            if deleted_items:
                query = f"""
                SELECT b.STRING_ID, b.KR
                FROM compare_db."{table}" b
                LEFT JOIN main."{table}" a ON b.STRING_ID = a.STRING_ID
                WHERE a.STRING_ID IS NULL AND b.KR IS NOT NULL
                """
                
                cursor.execute(query)
                results = cursor.fetchall()
                
                # 테이블 이름에서 파일명과 시트명 추출
                file_name, sheet_name = self.extract_file_sheet_from_table(table)
                
                for string_id, kr in results:
                    self.compare_results.append({
                        "file_name": db_file_name,
                        "sheet_name": sheet_name,
                        "string_id": string_id,
                        "kr": kr,              # 삭제는 kr에 (기존 유지)
                        "original_kr": "",     # 삭제는 original_kr 빈 값
                        "type": "삭제됨"
                    })
                    changes_count += 1
        
        except sqlite3.Error as e:
            print(f"테이블 '{table}' 비교 중 오류 발생: {e}")
        
        return changes_count


    def process_unique_table_for_pair(self, conn, table, db_prefix, type_label, db_file_name):
        """DB 쌍의 고유 테이블 처리 (신규/삭제 표시 방식 수정)"""
        cursor = conn.cursor()
        changes_count = 0
        
        try:
            # 테이블 구조 확인
            cursor.execute(f"PRAGMA {db_prefix}.table_info('{table}')")
            columns = {row[1]: row[0] for row in cursor.fetchall()}
            
            # 필요한 컬럼이 있는지 확인
            if 'STRING_ID' not in columns or 'KR' not in columns:
                return 0
            
            query = f"""
            SELECT STRING_ID, KR
            FROM {db_prefix}."{table}"
            WHERE KR IS NOT NULL
            """
            
            cursor.execute(query)
            results = cursor.fetchall()
            
            # 테이블 이름에서 파일명과 시트명 추출
            file_name, sheet_name = self.extract_file_sheet_from_table(table)
            
            for string_id, kr in results:
                if type_label == '신규 테이블':
                    # 🔧 수정: 신규 테이블은 original_kr에
                    self.compare_results.append({
                        "file_name": db_file_name,
                        "sheet_name": sheet_name,
                        "string_id": string_id,
                        "kr": "",               # 신규는 빈 값
                        "original_kr": kr,      # 원본 값은 original_kr에
                        "type": type_label
                    })
                else:  # 삭제된 테이블
                    # 기존 유지: 삭제는 kr에
                    self.compare_results.append({
                        "file_name": db_file_name,
                        "sheet_name": sheet_name,
                        "string_id": string_id,
                        "kr": kr,               # 삭제는 kr에
                        "original_kr": "",      # 삭제는 original_kr 빈 값
                        "type": type_label
                    })
                changes_count += 1
                
        except sqlite3.Error as e:
            print(f"테이블 '{table}' 처리 중 오류 발생: {e}")
        
        return changes_count



    
    def export_results_to_excel(self, output_path):
        """비교 결과를 엑셀 파일로 내보내기"""
        try:
            # 데이터프레임으로 변환
            df = pd.DataFrame(self.compare_results)
            
            # 엑셀로 저장
            df.to_excel(output_path, index=False)
            
            return {
                "status": "success",
                "path": output_path,
                "count": len(self.compare_results)
            }
        except Exception as e:
            return {
                "status": "error",
                "message": str(e)
            }
            
    def compare_translation_databases(self, db1_path, db2_path, languages=['kr', 'en', 'cn', 'tw', 'th'], progress_callback=None):
        """translation_data 테이블을 가진 두 DB 파일 비교 (TRANSLATION DB 비교용)"""
        if not db1_path or not db2_path:
            return {"status": "error", "message": "두 개의 DB 파일 경로가 모두 필요합니다."}
        
        if not os.path.exists(db1_path) or not os.path.exists(db2_path):
            return {"status": "error", "message": "DB 파일이 존재하지 않습니다."}
        
        # 결과 초기화
        self.compare_results = []
        changes_count = 0
        
        try:
            if progress_callback:
                progress_callback("DB 연결 및 테이블 확인 중...", 1, 5)
            
            # 첫 번째 DB 연결
            conn = sqlite3.connect(db1_path)
            cursor = conn.cursor()
            
            # 두 번째 DB 연결
            cursor.execute(f"ATTACH DATABASE '{db2_path}' AS db2;")
            
            # translation_data 테이블이 양쪽에 모두 있는지 확인
            cursor.execute("SELECT name FROM main.sqlite_master WHERE type='table' AND name='translation_data'")
            db1_has_table = cursor.fetchone() is not None
            
            cursor.execute("SELECT name FROM db2.sqlite_master WHERE type='table' AND name='translation_data'")
            db2_has_table = cursor.fetchone() is not None
            
            if not db1_has_table or not db2_has_table:
                return {"status": "error", "message": "translation_data 테이블이 한쪽 또는 양쪽 DB에 없습니다."}
            
            # DB 파일명 추출
            db1_name = os.path.basename(db1_path)
            db2_name = os.path.basename(db2_path)
            
            if progress_callback:
                progress_callback("첫 번째 DB에만 있는 항목 찾는 중...", 2, 5)
            
            # 1. 첫 번째 DB에만 있는 항목 (신규)
            query = """
            SELECT string_id, kr, en, cn, tw, th, file_name, sheet_name
            FROM main.translation_data 
            WHERE status = 'active' 
              AND string_id NOT IN (
                SELECT string_id 
                FROM db2.translation_data 
                WHERE status = 'active'
              )
            """
            
            cursor.execute(query)
            new_items = cursor.fetchall()
            
            # 1. 첫 번째 DB에만 있는 항목 (신규) - 🔧 수정
            for item in new_items:
                string_id, kr, en, cn, tw, th, file_name, sheet_name = item
                self.compare_results.append({
                    "db1_name": db1_name,
                    "db2_name": db2_name,
                    "file_name": file_name or "Unknown",
                    "sheet_name": sheet_name or "translation_data",
                    "type": "신규 (DB1에만 있음)",
                    "string_id": string_id,
                    "kr": "",               # 🔧 수정: 신규는 빈 값
                    "en": "",
                    "cn": "",
                    "tw": "",
                    "th": "",
                    "original_kr": kr,      # 🔧 수정: 원본 값들을 original_*에
                    "original_en": en,
                    "original_cn": cn,
                    "original_tw": tw,
                    "original_th": th
                })
                changes_count += 1
            
            if progress_callback:
                progress_callback("두 번째 DB에만 있는 항목 찾는 중...", 3, 5)
            
            # 2. 두 번째 DB에만 있는 항목 (삭제됨)
            query = """
            SELECT string_id, kr, en, cn, tw, th, file_name, sheet_name
            FROM db2.translation_data 
            WHERE status = 'active' 
              AND string_id NOT IN (
                SELECT string_id 
                FROM main.translation_data 
                WHERE status = 'active'
              )
            """
            
            cursor.execute(query)
            deleted_items = cursor.fetchall()
            
            # 2. 두 번째 DB에만 있는 항목 (삭제됨) - 기존 유지
            for item in deleted_items:
                string_id, kr, en, cn, tw, th, file_name, sheet_name = item
                self.compare_results.append({
                    "db1_name": db1_name,
                    "db2_name": db2_name,
                    "file_name": file_name or "Unknown",
                    "sheet_name": sheet_name or "translation_data",
                    "type": "삭제됨 (DB2에만 있음)",
                    "string_id": string_id,
                    "kr": kr,               # 삭제는 kr에 (기존 유지)
                    "en": en,
                    "cn": cn,
                    "tw": tw,
                    "th": th,
                    "original_kr": "",      # 삭제는 original_* 빈 값
                    "original_en": "",
                    "original_cn": "",
                    "original_tw": "",
                    "original_th": ""
                })
                changes_count += 1
            
            if progress_callback:
                progress_callback("내용이 다른 항목 찾는 중...", 4, 5)
            
            # 3. 양쪽에 모두 있지만 내용이 다른 항목 (변경됨)
            # 선택된 언어들에 대해서만 비교
            lang_conditions = []
            for lang in languages:
                lang_conditions.append(f"""
                    (db1.{lang} != db2.{lang} OR 
                     (db1.{lang} IS NULL) != (db2.{lang} IS NULL))
                """)
            
            lang_condition = " OR ".join(lang_conditions)
            
            query = f"""
            SELECT 
                db1.string_id,
                db1.kr, db2.kr,
                db1.en, db2.en,
                db1.cn, db2.cn,
                db1.tw, db2.tw,
                db1.th, db2.th,
                db1.file_name, db1.sheet_name
            FROM main.translation_data db1
            INNER JOIN db2.translation_data db2 ON db1.string_id = db2.string_id
            WHERE db1.status = 'active' 
              AND db2.status = 'active'
              AND ({lang_condition})
            ORDER BY db1.string_id
            """
            
            cursor.execute(query)
            changed_items = cursor.fetchall()
            
            for item in changed_items:
                (string_id, 
                 db1_kr, db2_kr, db1_en, db2_en, db1_cn, db2_cn, 
                 db1_tw, db2_tw, db1_th, db2_th, 
                 file_name, sheet_name) = item
                
                # 어떤 언어가 변경되었는지 확인
                changed_languages = []
                if db1_kr != db2_kr or (db1_kr is None) != (db2_kr is None):
                    changed_languages.append("KR")
                if db1_en != db2_en or (db1_en is None) != (db2_en is None):
                    changed_languages.append("EN")
                if db1_cn != db2_cn or (db1_cn is None) != (db2_cn is None):
                    changed_languages.append("CN")
                if db1_tw != db2_tw or (db1_tw is None) != (db2_tw is None):
                    changed_languages.append("TW")
                if db1_th != db2_th or (db1_th is None) != (db2_th is None):
                    changed_languages.append("TH")
                
                self.compare_results.append({
                    "db1_name": db1_name,
                    "db2_name": db2_name,
                    "file_name": file_name or "Unknown",
                    "sheet_name": sheet_name or "translation_data",
                    "type": f"변경됨 ({', '.join(changed_languages)})",
                    "string_id": string_id,
                    "kr": db2_kr,  # DB2 값 (새로운 값)
                    "en": db2_en,
                    "cn": db2_cn,
                    "tw": db2_tw,
                    "th": db2_th,
                    "original_kr": db1_kr,  # DB1 값 (원래 값)
                    "original_en": db1_en,
                    "original_cn": db1_cn,
                    "original_tw": db1_tw,
                    "original_th": db1_th
                })
                changes_count += 1
            
            if progress_callback:
                progress_callback("비교 완료!", 5, 5)
            
            # DB 분리 및 연결 종료
            cursor.execute("DETACH DATABASE db2;")
            conn.close()
            
            return {
                "status": "success",
                "db1_name": db1_name,
                "db2_name": db2_name,
                "total_changes": changes_count,
                "new_items": len(new_items),
                "deleted_items": len(deleted_items),
                "changed_items": len(changed_items),
                "compare_results": self.compare_results
            }
            
        except Exception as e:
            # 연결이 열려있으면 닫기 시도
            try:
                cursor.execute("DETACH DATABASE IF EXISTS db2;")
                conn.close()
            except:
                pass
            
            return {
                "status": "error",
                "message": str(e)
            }
        
    # db_compare_manager.py에 추가할 메서드들
    def detect_db_type(self, db_path):
        """DB 타입 자동 감지 (TRANSLATION DB vs STRING DB)"""
        try:
            conn = sqlite3.connect(db_path)
            cursor = conn.cursor()
            
            # translation_data 테이블이 있는지 확인
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='translation_data'")
            has_translation_table = cursor.fetchone() is not None
            
            if has_translation_table:
                # translation_data 테이블의 구조 확인
                cursor.execute("PRAGMA table_info(translation_data)")
                columns = [row[1] for row in cursor.fetchall()]
                
                # TRANSLATION DB의 필수 컬럼들이 있는지 확인
                required_columns = ['string_id', 'kr', 'en', 'cn', 'tw', 'th']
                if all(col in columns for col in required_columns):
                    conn.close()
                    return "TRANSLATION"
            
            # String으로 시작하는 테이블이 있는지 확인
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name LIKE 'String%'")
            string_tables = cursor.fetchall()
            
            conn.close()
            
            if string_tables:
                return "STRING"
            
            return "UNKNOWN"
            
        except Exception as e:
            print(f"DB 타입 감지 중 오류: {e}")
            return "UNKNOWN"

    def auto_compare_databases(self, db1_path, db2_path, options):
        """DB 타입을 자동으로 판단하여 적절한 비교 수행"""
        # DB 타입 감지
        db1_type = self.detect_db_type(db1_path)
        db2_type = self.detect_db_type(db2_path)
        
        if db1_type != db2_type:
            return {
                "status": "error",
                "message": f"DB 타입이 다릅니다. DB1: {db1_type}, DB2: {db2_type}"
            }
        
        if db1_type == "TRANSLATION":
            # TRANSLATION DB 비교
            result = self.compare_translation_databases(
                db1_path, 
                db2_path, 
                options.get("languages", ["kr", "en", "cn", "tw", "th"])
            )
            
            if result["status"] == "success":
                # 결과 형태를 통합 형태로 변환
                unified_results = []
                for item in result["compare_results"]:
                    unified_results.append({
                        "db_name": item.get("db1_name", ""),
                        "file_name": item.get("file_name", ""),
                        "sheet_name": item.get("sheet_name", ""),
                        "string_id": item.get("string_id", ""),
                        "type": item.get("type", ""),
                        "kr": item.get("kr", ""),
                        "original_kr": item.get("original_kr", "")
                    })
                
                result["compare_results"] = unified_results
                result["db_type"] = "TRANSLATION DB"
            
            return result
            
        elif db1_type == "STRING":
            # STRING DB 비교
            db_pair = {
                'file_name': os.path.basename(db1_path),
                'original_path': db1_path,
                'compare_path': db2_path
            }
            
            try:
                changes = self.compare_db_pair(
                    db1_path,
                    db2_path,
                    options.get("changed_kr", True),
                    options.get("new_items", True),
                    options.get("deleted_items", True)
                )
                
                # 결과 형태를 통합 형태로 변환
                unified_results = []
                for item in self.compare_results:
                    unified_results.append({
                        "db_name": item.get("file_name", ""),
                        "file_name": item.get("file_name", ""),
                        "sheet_name": item.get("sheet_name", ""),
                        "string_id": item.get("string_id", ""),
                        "type": item.get("type", ""),
                        "kr": item.get("kr", ""),
                        "original_kr": item.get("original_kr", "")
                    })
                
                return {
                    "status": "success",
                    "total_changes": changes,
                    "compare_results": unified_results,
                    "db_type": "STRING DB"
                }
                
            except Exception as e:
                return {
                    "status": "error",
                    "message": str(e)
                }
        
        else:
            return {
                "status": "error",
                "message": f"지원되지 않는 DB 타입입니다: {db1_type}"
            }

    def auto_compare_folder_databases(self, db_pairs, options, progress_callback=None):
        """폴더 내 DB들을 자동으로 타입 판단하여 비교"""
        if not db_pairs:
            return {"status": "error", "message": "비교할 DB 파일 목록이 없습니다."}
        
        # 첫 번째 DB 쌍으로 타입 확인
        first_pair = db_pairs[0]
        db_type = self.detect_db_type(first_pair['original_path'])
        
        if db_type == "TRANSLATION":
            # TRANSLATION DB들을 폴더별로 비교
            return self.compare_translation_folder_databases(db_pairs, options, progress_callback)
        
        elif db_type == "STRING":
            # 기존 STRING DB 폴더 비교 사용
            result = self.compare_all_databases(
                db_pairs,
                options.get("changed_kr", True),
                options.get("new_items", True),
                options.get("deleted_items", True),
                progress_callback
            )
            
            if result["status"] == "success":
                # 결과 형태를 통합 형태로 변환
                unified_results = []
                for item in result["compare_results"]:
                    unified_results.append({
                        "db_name": item.get("file_name", ""),
                        "file_name": item.get("file_name", ""),
                        "sheet_name": item.get("sheet_name", ""),
                        "string_id": item.get("string_id", ""),
                        "type": item.get("type", ""),
                        "kr": item.get("kr", ""),
                        "original_kr": item.get("original_kr", "")
                    })
                
                result["compare_results"] = unified_results
                result["db_type"] = "STRING DB"
            
            return result
        
        else:
            return {
                "status": "error",
                "message": f"지원되지 않는 DB 타입입니다: {db_type}"
            }

    def compare_translation_folder_databases(self, db_pairs, options, progress_callback=None):
        """TRANSLATION DB들을 폴더별로 비교"""
        self.compare_results = []
        total_changes = 0
        success_count = 0
        error_list = []
        
        for idx, db_pair in enumerate(db_pairs):
            if progress_callback:
                progress_callback(f"TRANSLATION DB 비교 중: {db_pair['file_name']} ({idx+1}/{len(db_pairs)})",
                                idx+1, len(db_pairs))
            
            try:
                # 각 TRANSLATION DB 쌍 비교
                result = self.compare_translation_databases(
                    db_pair['original_path'],
                    db_pair['compare_path'],
                    options.get("languages", ["kr", "en", "cn", "tw", "th"])
                )
                
                if result["status"] == "success":
                    # 결과를 전체 결과에 추가
                    for item in result["compare_results"]:
                        self.compare_results.append({
                            "db_name": db_pair['file_name'],
                            "file_name": item.get("file_name", ""),
                            "sheet_name": item.get("sheet_name", ""),
                            "string_id": item.get("string_id", ""),
                            "type": item.get("type", ""),
                            "kr": item.get("kr", ""),
                            "original_kr": item.get("original_kr", "")
                        })
                    
                    total_changes += result["total_changes"]
                    success_count += 1
                else:
                    error_list.append(f"{db_pair['file_name']}: {result['message']}")
                    
            except Exception as e:
                error_msg = f"TRANSLATION DB 비교 실패 ({db_pair['file_name']}): {e}"
                error_list.append(error_msg)
                
                if progress_callback:
                    progress_callback(error_msg, idx+1, len(db_pairs))
        
        return {
            "status": "success",
            "success_count": success_count,
            "total_count": len(db_pairs),
            "total_changes": total_changes,
            "errors": error_list,
            "compare_results": self.compare_results,
            "db_type": "TRANSLATION DB"
        }