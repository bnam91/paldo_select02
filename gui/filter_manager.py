import os
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QLabel, QCheckBox, 
                            QComboBox, QHBoxLayout, QGroupBox)
from handlers import FilterHandler
import pandas as pd

class FilterManager:
    """필터 관련 기능을 관리하는 클래스"""
    
    def __init__(self, parent):
        """
        초기화
        
        Args:
            parent: ExcelViewer 클래스의 인스턴스
        """
        self.parent = parent
    
    def apply_filters(self):
        """현재 필터 설정에 따라 데이터 필터링"""
        # 원본 데이터가 없으면 리턴
        if self.parent.original_df is None:
            return
        
        # 필터링을 위한 마스크 생성 (모든 행 선택)
        mask = pd.Series([True] * len(self.parent.original_df))
        
        # 선택된 상품에 따른 필터링
        selected_product = self.parent.product_combo.currentText()
        # '전체'가 아닌 경우에만 필터링 적용
        if selected_product and selected_product != "전체":
            # 희망 상품 컬럼이 있는 경우에만 필터링 적용
            if self.parent.product_column_idx >= 0:
                product_col_name = self.parent.original_df.columns[self.parent.product_column_idx]
                # 상품명이 포함된 행만 선택
                product_mask = self.parent.original_df[product_col_name].str.contains(
                    selected_product, case=False, na=False, regex=False)
                mask = mask & product_mask
        
        # 2. 단일 상품 필터 적용
        if self.parent.single_product_checkbox.isChecked() and selected_product:
            try:
                filtered_original = FilterHandler.apply_single_product_filter(self.parent.original_df)
                
                # '상품명' 열이 있는지 확인
                if '상품명' in filtered_original.columns:
                    mask = mask & (filtered_original['상품명'] == selected_product)
                elif self.parent.product_column_idx >= 0:
                    # '상품명' 열이 없으면 희망상품 컬럼을 대신 사용
                    product_col_name = self.parent.original_df.columns[self.parent.product_column_idx]
                    # 정확히 일치하는 상품만 필터링
                    mask = mask & (self.parent.original_df[product_col_name] == selected_product)
                else:
                    self.parent.status_label.setText("단일 상품 필터링을 위한 '상품명' 열을 찾을 수 없습니다.")
            except Exception as e:
                self.parent.status_label.setText(f"단일 상품 필터 적용 중 오류: {str(e)}")
        
        # 3. 이름/연락처/URL 검색 필터 적용
        contact_search_text = self.parent.contact_search_input.text().strip()
        if contact_search_text:
            filtered_original = FilterHandler.apply_contact_search_filter(
                self.parent.original_df, contact_search_text, 
                self.parent.name_column_idx, self.parent.contact_column_idx, self.parent.url_column_idx
            )
            
            # 인덱스 직접 비교 대신 불리언 마스크 생성
            contact_mask = pd.Series(False, index=self.parent.original_df.index)
            contact_mask[filtered_original.index] = True
            mask = mask & contact_mask
        
        # 4. 상태별 필터 적용
        selected_statuses = [status for status, checkbox in self.parent.status_checkboxes.items() 
                          if checkbox.isChecked()]
        
        if len(selected_statuses) < 5:  # 5개 상태가 모두 선택되지 않은 경우
            # 선택된 상태 값과 일치하는 행만 남김
            status_mask = self.parent.original_df.index.to_series().apply(
                lambda idx: self.parent.row_status.get(idx, 0) in selected_statuses
            )
            mask = mask & status_mask
        
        # 5. 채널 필터 적용
        selected_channels = [channel for channel, checkbox in self.parent.channel_checkboxes.items() 
                            if checkbox.isChecked()]
        
        if selected_channels and len(selected_channels) < len(self.parent.channel_list):
            filtered_original = FilterHandler.apply_channel_filter(self.parent.original_df, selected_channels)
            
            # 인덱스 직접 비교 대신 불리언 마스크 생성
            channel_mask = pd.Series(False, index=self.parent.original_df.index)
            channel_mask[filtered_original.index] = True
            mask = mask & channel_mask
        
        # 아무것도 선택되지 않았으면 경고
        if len(selected_channels) == 0:
            self.parent.status_label.setText("최소 하나의 채널을 선택해주세요.")
            return
        
        if len(self.parent.original_df[mask]) == 0:
            self.parent.status_label.setText("필터 조건에 맞는 데이터가 없습니다.")
            self.parent.update_status_statistics()  # 빈 결과도 통계 업데이트
            return
        
        # C열부터 N열 선택 (K열, M열 제외)
        columns_to_show = []
        for col in self.parent.original_df.columns[2:14]:  # C(인덱스 2)부터 N(인덱스 13)까지
            col_idx = self.parent.original_df.columns.get_loc(col)
            # K열(인덱스 10)과 M열(인덱스 12)는 제외
            if col_idx != 10 and col_idx != 12:
                columns_to_show.append(col)
        
        # 필터링된 데이터프레임 적용
        self.parent.filtered_df = self.parent.original_df[mask][columns_to_show]
        
        # 테이블 업데이트
        self.parent.table_manager.update_table(self.parent.filtered_df)
        
        # 필터 상태 메시지 업데이트
        filter_msg = []
        if selected_product:
            filter_msg.append(f"상품검색: '{selected_product}'")
        if contact_search_text:
            filter_msg.append(f"이름/연락처/URL: '{contact_search_text}'")
        if self.parent.single_product_checkbox.isChecked():
            filter_msg.append("단일 상품만")

        # 상태 필터 메시지 추가
        if len(selected_statuses) < 5:
            status_names = {0: "미정", 1: "선정", 2: "대기", 3: "제외", 4: "완료"}
            selected_status_names = [status_names[s] for s in selected_statuses]
            filter_msg.append(f"상태: {', '.join(selected_status_names)}")

        if len(selected_channels) < len(self.parent.channel_list):
            filter_msg.append(f"채널: {', '.join(selected_channels)}")
        
        if filter_msg:
            self.parent.status_label.setText(f"적용된 필터: {', '.join(filter_msg)}")
        else:
            self.parent.status_label.setText("모든 데이터가 표시됩니다.")
        
        # 상태 메시지 업데이트를 위한 타이머 시작
        self.parent.status_timer.start(3000)  # 3초 후 업데이트
    
    def reset_filter(self):
        """모든 필터 초기화"""
        if self.parent.original_df is None:
            return
        
        # 검색어 필드 초기화
        self.parent.product_combo.setCurrentIndex(0)  # '전체'로 설정
        self.parent.contact_search_input.clear()  # 이름/연락처 검색 필드도 초기화
        
        # 단일 상품 체크박스 초기화
        self.parent.single_product_checkbox.setChecked(False)
        
        # 모든 상태 체크박스 선택
        for checkbox in self.parent.status_checkboxes.values():
            checkbox.setChecked(True)
        
        # 모든 채널 체크박스 선택
        for checkbox in self.parent.channel_checkboxes.values():
            checkbox.setChecked(True)
        
        # C열부터 N열 선택 (K열, M열 제외)
        columns_to_show = []
        for col in self.parent.original_df.columns[2:14]:  # C(인덱스 2)부터 N(인덱스 13)까지
            col_idx = self.parent.original_df.columns.get_loc(col)
            # K열(인덱스 10)과 M열(인덱스 12)는 제외
            if col_idx != 10 and col_idx != 12:
                columns_to_show.append(col)
        
        # 선택된 열만 포함하는 데이터프레임 생성
        self.parent.filtered_df = self.parent.original_df[columns_to_show]
        
        # 테이블 업데이트
        self.parent.table_manager.update_table(self.parent.filtered_df)
        
        self.parent.status_label.setText("필터가 초기화되었습니다.") 