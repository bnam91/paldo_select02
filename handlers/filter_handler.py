import pandas as pd

class FilterHandler:
    @staticmethod
    def clean_url(url):
        """URL에서 http://, https:// 접두사 제거"""
        url = str(url).strip().lower()
        # http://, https:// 접두사 제거
        if url.startswith('http://'):
            url = url[7:]
        elif url.startswith('https://'):
            url = url[8:]
        # www. 접두사 제거 (선택적)
        if url.startswith('www.'):
            url = url[4:]
        return url
    
    @staticmethod
    def apply_product_filter(df, search_text):
        """상품명 검색 필터 적용"""
        if not search_text:
            return df
            
        # C열(인덱스 2) 기준으로 필터링
        c_column = df.columns[2]
        mask = df[c_column].astype(str).str.contains(search_text, case=False, na=False)
        return df[mask]
    
    @staticmethod
    def apply_single_product_filter(df):
        """단일 상품 필터 적용"""
        c_column = df.columns[2]
        # 쉼표가 포함되지 않은 항목만 선택 (단일 상품)
        mask = ~df[c_column].astype(str).str.contains(',', na=False)
        return df[mask]
    
    @staticmethod
    def apply_contact_search_filter(df, search_text, name_column_idx, contact_column_idx, url_column_idx):
        """이름/연락처/URL 검색 필터 적용"""
        if not search_text:
            return df
            
        name_mask = pd.Series(False, index=df.index)
        contact_mask = pd.Series(False, index=df.index)
        url_mask = pd.Series(False, index=df.index)
        
        # 이름 칼럼 검색
        if name_column_idx != -1:
            name_mask = df.iloc[:, name_column_idx].astype(str).str.contains(
                search_text, case=False, na=False, regex=False
            )
        
        # 연락처 칼럼 검색
        if contact_column_idx != -1:
            contact_mask = df.iloc[:, contact_column_idx].astype(str).str.contains(
                search_text, case=False, na=False, regex=False
            )
        
        # URL 칼럼 검색 (접두사 제외 버전)
        if url_column_idx != -1:
            # 검색어에서 접두사 제거
            clean_search_text = FilterHandler.clean_url(search_text)
            
            # 데이터에서도 접두사 제거하여 비교
            url_mask = df.iloc[:, url_column_idx].apply(
                lambda x: clean_search_text in FilterHandler.clean_url(x)
            )
        
        # 이름, 연락처 또는 URL에 검색어 포함된 항목 선택
        combined_mask = name_mask | contact_mask | url_mask
        return df[combined_mask]
    
    @staticmethod
    def apply_completed_filter(df, row_status):
        """완료 상태 항목 제외 필터 적용"""
        # 완료 상태(4)가 아닌 행만 선택
        completed_indices = [idx for idx, status in row_status.items() if status == 4]
        return df[~df.index.isin(completed_indices)]
    
    @staticmethod
    def apply_channel_filter(df, selected_channels):
        """채널 필터 적용"""
        if not selected_channels:
            return df
            
        # D열(인덱스 3) 기준으로 필터링
        d_column = df.columns[3]
        
        # 선택된 채널 포함 여부 체크를 위한 마스크
        channel_mask = df[d_column].apply(
            lambda x: any(channel in str(x) for channel in selected_channels)
        )
        
        return df[channel_mask]

    def filter_by_name_contact(self, df, name_column_idx, contact_column_idx, url_column_idx, search_text):
        """이름 또는 연락처로 필터링"""
        if not search_text or search_text.strip() == "":
            return df
        
        search_text = search_text.lower().strip()
        
        # 이름 또는 연락처에 검색어가 포함된 행만 선택
        combined_mask = False
        
        # 이름 칼럼이 있는 경우
        if name_column_idx >= 0:
            # regex=False로 설정하여 정규식으로 해석되지 않도록 함
            name_mask = df.iloc[:, name_column_idx].astype(str).str.contains(
                search_text, case=False, na=False, regex=False)
            combined_mask = combined_mask | name_mask
        
        # 연락처 칼럼이 있는 경우
        if contact_column_idx >= 0:
            # regex=False로 설정하여 정규식으로 해석되지 않도록 함
            contact_mask = df.iloc[:, contact_column_idx].astype(str).str.contains(
                search_text, case=False, na=False, regex=False)
            combined_mask = combined_mask | contact_mask
        
        # URL 칼럼이 있는 경우
        if url_column_idx >= 0:
            # regex=False로 설정하여 정규식으로 해석되지 않도록 함
            url_mask = df.iloc[:, url_column_idx].astype(str).str.contains(
                search_text, case=False, na=False, regex=False)
            combined_mask = combined_mask | url_mask
        
        # 필터링된 데이터프레임 반환
        return df[combined_mask] 