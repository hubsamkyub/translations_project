import pandas as pd
import numpy as np

class ExcelDiffManager:
    """
    두 엑셀 파일의 데이터를 비교하고 차이점을 보고서로 생성하는 로직을 담당하는 클래스.
    UI와 독립적으로 동작하도록 설계되었습니다.
    """
    def run_comparison(self, file_path1, file_path2, key_column, header1, header2, sheet_name1, sheet_name2):
        """
        비교 프로세스 전체를 실행하고 결과를 반환합니다.

        :return: 성공 또는 실패 상태와 메시지를 담은 딕셔너리.
                 성공 시: {'status': 'success', 'message': '...', 'stats': {...}}
                 실패 시: {'status': 'error', 'message': '...'}
        """
        try:
            # 1. 엑셀 파일 로드 및 유효성 검사
            df1 = self._load_dataframe(file_path1, header1, sheet_name1)
            df2 = self._load_dataframe(file_path2, header2, sheet_name2)
            self._validate_inputs(df1, df2, key_column)

            # 2. 데이터 비교 수행
            merged_df = self._perform_diff(df1, df2, key_column)

            # 3. 결과 보고서 포맷팅
            report_df, stats = self._format_report(merged_df, key_column, list(df1.columns))

            message = f"비교 완료. 추가: {stats['added']}, 삭제: {stats['deleted']}, 변경: {stats['modified']} 건"
            return {
                "status": "success",
                "message": message,
                "stats": stats,
                "report_df": report_df
            }

        except Exception as e:
            return {"status": "error", "message": f"오류 발생: {e}"}

    def save_report_to_excel(self, report_df, output_path):
        """
        생성된 보고서 데이터프레임을 스타일을 적용하여 엑셀 파일로 저장합니다.
        """
        try:
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                report_df.to_excel(writer, index=False, sheet_name='diff_report')

                # 스타일링
                workbook = writer.book
                worksheet = writer.sheets['diff_report']

                # 변경사항에 따라 색상 적용
                green_fill = pd.io.formats.style.Styler._background_gradient('mediumseagreen', 'mediumseagreen')
                red_fill = pd.io.formats.style.Styler._background_gradient('lightcoral', 'lightcoral')
                yellow_fill = pd.io.formats.style.Styler._background_gradient('gold', 'gold')

                for idx, status in enumerate(report_df['변경사항'], 1):
                    if status == '추가':
                        worksheet[f'A{idx+1}'].fill = green_fill
                    elif status == '삭제':
                        worksheet[f'A{idx+1}'].fill = red_fill
                    elif status == '변경':
                        worksheet[f'A{idx+1}'].fill = yellow_fill

            return {"status": "success", "message": f"보고서가 성공적으로 저장되었습니다: {output_path}"}
        except Exception as e:
            return {"status": "error", "message": f"보고서 저장 중 오류 발생: {e}"}

    def _load_dataframe(self, file_path, header_row, sheet_name):
        """엑셀 파일을 읽어 데이터프레임으로 반환합니다."""
        if not file_path or not isinstance(file_path, str):
            raise ValueError("파일 경로가 유효하지 않습니다.")
        try:
            # header는 0부터 시작하므로 사용자가 입력한 값(1부터 시작)에서 1을 빼줍니다.
            return pd.read_excel(file_path, header=int(header_row) - 1, sheet_name=sheet_name, dtype=str).fillna('')
        except FileNotFoundError:
            raise FileNotFoundError(f"파일을 찾을 수 없습니다: {file_path}")
        except Exception as e:
            raise IOError(f"'{file_path}' 파일의 '{sheet_name}' 시트를 읽는 중 오류가 발생했습니다: {e}")

    def _validate_inputs(self, df1, df2, key_column):
        """입력된 데이터프레임과 Key 컬럼의 유효성을 검사합니다."""
        if key_column not in df1.columns:
            raise ValueError(f"파일 1의 컬럼에 '{key_column}'가 존재하지 않습니다.")
        if key_column not in df2.columns:
            raise ValueError(f"파일 2의 컬럼에 '{key_column}'가 존재하지 않습니다.")

    def _perform_diff(self, df1, df2, key_column):
        """두 데이터프레임을 비교하여 병합된 결과물을 반환합니다."""
        # 각 DataFrame에 접미사 추가하여 병합
        return pd.merge(df1, df2, on=key_column, how='outer', suffixes=('_old', '_new'), indicator=True)

    def _format_report(self, merged_df, key_column, original_columns):
        """병합된 데이터프레임을 최종 보고서 형태로 가공합니다."""
        report_data = []

        # Key와 변경 상태를 제외한 순수 데이터 컬럼
        value_columns = [col for col in original_columns if col != key_column]

        for _, row in merged_df.iterrows():
            key_value = row[key_column]
            status = ""
            details = {}

            if row['_merge'] == 'both': # 변경 또는 동일
                is_modified = False
                for col in value_columns:
                    old_val = row[f'{col}_old']
                    new_val = row[f'{col}_new']
                    if old_val != new_val:
                        is_modified = True
                        details[col] = f'"{old_val}" → "{new_val}"'
                if is_modified:
                    status = '변경'
                else:
                    continue # 동일한 행은 보고서에 포함하지 않음

            elif row['_merge'] == 'left_only': # 삭제
                status = '삭제'
                for col in value_columns:
                    details[col] = f'"{row[f"{col}_old"]}"'

            elif row['_merge'] == 'right_only': # 추가
                status = '추가'
                for col in value_columns:
                    details[col] = f'"{row[f"{col}_new"]}"'

            report_row = {'변경사항': status, key_column: key_value}
            report_row.update(details)
            report_data.append(report_row)

        if not report_data:
            return pd.DataFrame(), {'added': 0, 'deleted': 0, 'modified': 0}

        final_report_df = pd.DataFrame(report_data)

        # 통계 계산
        stats = final_report_df['변경사항'].value_counts().to_dict()
        stats = {
            'added': stats.get('추가', 0),
            'deleted': stats.get('삭제', 0),
            'modified': stats.get('변경', 0)
        }

        return final_report_df, stats