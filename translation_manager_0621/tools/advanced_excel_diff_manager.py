#advanced_excel_diff_manager
# tools/advanced_excel_diff_manager.py
import os
import pandas as pd
import numpy as np

class AdvancedExcelDiffManager:
    """
    고급 엑셀 비교 기능의 데이터 처리 로직을 담당합니다.
    UI와 독립적으로 작동하며, 모든 연산은 pandas를 기반으로 합니다.
    """
    # --- [신규] 폴더 검색 기능 추가 ---
    def find_string_excels_in_folder(self, folder_path):
        """지정된 폴더에서 'string'으로 시작하는 엑셀 파일을 검색합니다."""
        if not folder_path or not os.path.isdir(folder_path):
            raise ValueError("유효한 폴더 경로가 아닙니다.")

        found_files = []
        for filename in os.listdir(folder_path):
            if filename.lower().startswith('string') and (filename.endswith('.xlsx') or filename.endswith('.xls')):
                if not filename.startswith('~$'): # 임시 파일 제외
                    found_files.append(filename)
        return found_files
    
    # --- [신규] 여러 파일로부터 데이터프레임 생성 기능 추가 ---
    def create_dataframe_from_files(self, folder_path, file_list, header_row):
        """지정된 파일 목록에서 'string' 시작 시트를 모두 읽어 하나의 데이터프레임으로 합칩니다."""
        all_dfs = []
        header_index = int(header_row) - 1

        for filename in file_list:
            file_path = os.path.join(folder_path, filename)
            try:
                xls = pd.ExcelFile(file_path)
                for sheet_name in xls.sheet_names:
                    if sheet_name.lower().startswith('string'):
                        df = pd.read_excel(xls, sheet_name=sheet_name, header=header_index, dtype=str).fillna('')
                        all_dfs.append(df)
            except Exception as e:
                raise IOError(f"'{filename}' 파일 처리 중 오류: {e}")

        if not all_dfs:
            raise ValueError("선택된 파일에서 'string'으로 시작하는 시트를 찾을 수 없습니다.")

        combined_df = pd.concat(all_dfs, ignore_index=True)
        return combined_df


    def get_sheet_names(self, file_path):
        """엑셀 파일의 모든 시트 이름을 리스트로 반환합니다."""
        try:
            if not file_path or not isinstance(file_path, str):
                raise ValueError("유효한 파일 경로가 아닙니다.")
            return pd.ExcelFile(file_path).sheet_names
        except FileNotFoundError:
            raise FileNotFoundError(f"파일을 찾을 수 없습니다: {file_path}")
        except Exception as e:
            raise IOError(f"파일을 읽는 중 오류 발생: {e}")

    # --- [수정] 미리 생성된 데이터프레임을 인자로 받을 수 있도록 변경 ---
    def run_comparison(self, key_columns, 
                     df_old=None, file_path_old=None, sheets_old=None, header_old=None, 
                     df_new=None, file_path_new=None, sheets_new=None, header_new=None, 
                     filter_options=None):
        try:
            # 원본 데이터 로드: 데이터프레임이 직접 제공되면 그대로 사용, 아니면 파일에서 로드
            if df_old is None:
                df_old = self._load_and_combine_sheets(file_path_old, sheets_old, header_old)

            # 대상 데이터 로드: 데이터프레임이 직접 제공되면 그대로 사용, 아니면 파일에서 로드
            if df_new is None:
                df_new = self._load_and_combine_sheets(file_path_new, sheets_new, header_new)

            # (이하 로직은 이전과 거의 동일)
            if filter_options and filter_options.get('column') and filter_options.get('text'):
                df_old = self._apply_filter(df_old, filter_options)
                df_new = self._apply_filter(df_new, filter_options)

            self._validate_inputs(df_old, "원본", key_columns)
            self._validate_inputs(df_new, "대상", key_columns)

            merged_df = pd.merge(df_old, df_new, on=key_columns, how='outer', suffixes=('_old', '_new'), indicator=True)
            report_df, stats = self._format_report(merged_df, key_columns)

            message = f"비교 완료. 추가: {stats['added']}, 삭제: {stats['deleted']}, 변경: {stats['modified']} 건"
            return {"status": "success", "message": message, "report_df": report_df}

        except Exception as e:
            return {"status": "error", "message": str(e)}

    def _load_and_combine_sheets(self, file_path, sheet_names, header_row):
        """지정된 시트들을 읽어 하나의 데이터프레임으로 합쳐 반환합니다."""
        all_sheets_df = []
        header_index = int(header_row) - 1
        for sheet in sheet_names:
            try:
                df = pd.read_excel(file_path, sheet_name=sheet, header=header_index, dtype=str).fillna('')
                all_sheets_df.append(df)
            except Exception as e:
                raise IOError(f"'{file_path}' 파일의 '{sheet}' 시트를 읽는 데 실패했습니다: {e}")

        if not all_sheets_df:
            raise ValueError(f"'{file_path}' 파일에서 유효한 데이터를 읽어오지 못했습니다. 시트 이름과 헤더 행을 확인하세요.")

        return pd.concat(all_sheets_df, ignore_index=True)

    def _apply_filter(self, df, filter_options):
        """데이터프레임에 필터링 조건을 적용합니다."""
        filter_col = filter_options['column']
        filter_text = filter_options['text']

        if filter_col not in df.columns:
            raise ValueError(f"필터링 컬럼 '{filter_col}'을 찾을 수 없습니다.")

        return df[df[filter_col].str.contains(filter_text, na=False)].copy()

    def _validate_inputs(self, df, source_name, key_columns):
        """데이터와 Key 컬럼의 유효성을 검사합니다."""
        for col in key_columns:
            if col not in df.columns:
                raise ValueError(f"{source_name} 파일의 컬럼에 '{col}'가 존재하지 않습니다.")

    def _format_report(self, merged_df, key_columns):
        """병합된 데이터프레임을 최종 보고서 형태로 가공합니다."""
        report_data = []

        # Key를 제외한 순수 데이터 컬럼 (접미사 없이)
        value_columns = [col for col in merged_df.columns if col.endswith('_old') and col.replace('_old', '') not in key_columns]
        value_columns = [col.replace('_old', '') for col in value_columns]

        stats = {'added': 0, 'deleted': 0, 'modified': 0}

        for _, row in merged_df.iterrows():
            details = {}
            is_modified = False

            if row['_merge'] == 'left_only':
                status = '삭제'
                stats['deleted'] += 1
            elif row['_merge'] == 'right_only':
                status = '추가'
                stats['added'] += 1
            elif row['_merge'] == 'both':
                for col in value_columns:
                    old_val = row.get(f'{col}_old', '')
                    new_val = row.get(f'{col}_new', '')
                    if old_val != new_val:
                        is_modified = True
                        details[col] = f'"{old_val}" → "{new_val}"'

                if is_modified:
                    status = '변경'
                    stats['modified'] += 1
                else:
                    continue # 변경 없으면 보고서에 미포함

            report_row = {'상태': status}
            report_row.update({k: row[k] for k in key_columns}) # 키 컬럼 추가
            report_row.update(details)
            report_data.append(report_row)

        final_df = pd.DataFrame(report_data)
        return final_df, stats