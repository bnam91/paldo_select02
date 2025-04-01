import pandas as pd
import re
from PyQt5.QtWidgets import QMessageBox

class ExcelHandler:
    @staticmethod
    def format_phone_number(number):
        """전화번호 형식 정리"""
        if pd.isna(number):
            return ""
            
        # 숫자만 추출
        number_str = str(number)
        digits_only = re.sub(r'\D', '', number_str)
        
        # 11자리 전화번호 포맷팅
        if len(digits_only) == 11:
            return f"{digits_only[:3]}-{digits_only[3:7]}-{digits_only[7:]}"
        # 10자리 전화번호 포맷팅
        elif len(digits_only) == 10:
            return f"{digits_only[:3]}-{digits_only[3:6]}-{digits_only[6:]}"
        else:
            return number_str
    
    @staticmethod
    def load_excel_file(file_path, parent, header_mapping):
        """엑셀 파일 로드 및 전처리"""
        try:
            # 엑셀 파일 로드
            original_df = pd.read_excel(file_path)
            
            # 연락처/이름/희망상품 칼럼 찾기 및 데이터 전처리
            contact_column_idx = -1
            name_column_idx = -1
            product_column_idx = -1
            url_column_idx = -1
            
            # 매핑된 헤더 정보로 칼럼 매핑
            for original_header, mapped_header in header_mapping.items():
                for i, col in enumerate(original_df.columns):
                    if original_header in str(col):
                        original_df.rename(columns={col: original_header}, inplace=True)
                        break
            
            # 칼럼 인덱스 찾기 및 데이터 전처리
            for col in original_df.columns:
                if ("연락처" in col or "전화" in col) and not ("카톡" in col or "아이디" in col):
                    # 연락처 칼럼 인덱스 저장
                    contact_column_idx = original_df.columns.get_loc(col)
                    original_df[col] = original_df[col].apply(ExcelHandler.format_phone_number)
                # 이름 칼럼 인덱스 저장
                if "성함" in col or "이름" in col or "닉네임" in col:
                    name_column_idx = original_df.columns.get_loc(col)
                # 희망상품 칼럼 인덱스 저장
                if "희망상품" in col or "희망 상품" in col:
                    product_column_idx = original_df.columns.get_loc(col)
                # URL 칼럼 인덱스 저장
                col_str = str(col).lower()
                if "url" in col_str or "계정 링크" in col or "블로그" in col:
                    url_column_idx = original_df.columns.get_loc(col)
            
            # C열부터 N열 선택 (K열, M열 제외)
            columns_to_show = []
            for col in original_df.columns[2:14]:  # C(인덱스 2)부터 N(인덱스 13)까지
                col_idx = original_df.columns.get_loc(col)
                # K열(인덱스 10)과 M열(인덱스 12)는 제외
                if col_idx != 10 and col_idx != 12:
                    columns_to_show.append(col)
            
            # 필터링된 데이터프레임 생성
            filtered_df = original_df[columns_to_show]
            
            return {
                'original_df': original_df, 
                'filtered_df': filtered_df,
                'contact_column_idx': contact_column_idx,
                'name_column_idx': name_column_idx,
                'product_column_idx': product_column_idx,
                'url_column_idx': url_column_idx
            }
            
        except Exception as e:
            QMessageBox.critical(parent, "오류", f"엑셀 파일 로드 중 오류 발생: {str(e)}")
            return None 