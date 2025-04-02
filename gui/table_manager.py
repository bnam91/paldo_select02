import pandas as pd
import webbrowser
from PyQt5.QtWidgets import QTableWidgetItem, QApplication
from PyQt5.QtGui import QColor
from widgets import URLTableWidgetItem, StatusButton

class TableManager:
    """테이블 관련 기능을 관리하는 클래스"""
    
    def __init__(self, parent):
        """
        초기화
        
        Args:
            parent: ExcelViewer 클래스의 인스턴스
        """
        self.parent = parent
        self.table = parent.table
        self.row_colors = parent.row_colors
        self.header_mapping = parent.header_mapping
        
        # 테이블 이벤트 연결
        self.table.cellClicked.connect(self.on_cell_clicked)
    
    def update_table(self, df):
        """테이블 위젯 데이터 업데이트"""
        if df is None or len(df) == 0:
            self.table.setRowCount(0)
            self.table.setColumnCount(0)
            return
        
        # 테이블 위젯 설정
        self.table.setRowCount(len(df))
        self.table.setColumnCount(len(df.columns) + 3)  # 상태 버튼 칼럼 + 지정상품 + 지정채널
        
        # 헤더 레이블 설정 - 매핑된 이름 사용
        header_labels = ["상태", "지정상품", "지정채널"]  # 첫 번째 열은 상태 버튼, 이후 두 개는 새 칼럼
        for col in df.columns:
            # 매핑 정보가 있으면 매핑된 이름 사용, 없으면 원래 이름 사용
            mapped_name = self.header_mapping.get(col, col)
            header_labels.append(mapped_name)
        
        self.table.setHorizontalHeaderLabels(header_labels)
        
        # 열 인덱스 찾기 (테이블 내에서의 인덱스)
        product_column_idx = -1  # 희망상품 
        channel_column_idx = -1  # 신청채널
        url_column_idx = -1      # URL
        name_column_idx = -1     # 이름 및 닉네임
        
        for i, col in enumerate(df.columns):
            col_str = str(col).lower()
            if "희망상품" in col or "희망 상품" in col:
                product_column_idx = i + 3  # +3은 상태 버튼과 지정상품, 지정채널 칼럼 때문
            elif "신청 채널" in col or "신청채널" in col:
                channel_column_idx = i + 3
            elif "url" in col_str or "계정 링크" in col or "블로그" in col:
                url_column_idx = i + 3
            elif "성함" in col or "이름" in col or "닉네임" in col:
                name_column_idx = i + 3
        
        # 데이터 채우기
        for row in range(len(df)):
            # 원본 데이터프레임에서의 인덱스(행 ID)
            row_id = df.index[row]
            
            # 상태 버튼 추가
            status_btn = StatusButton(row_id)
            status = 0  # 기본 상태
            
            # 저장된 상태가 있으면 적용
            if row_id in self.parent.row_status:
                status = self.parent.row_status[row_id]
                status_btn.set_status(status)
            
            self.table.setCellWidget(row, 0, status_btn)
            
            # 새 칼럼(지정상품, 지정채널)을 위한 빈 아이템 추가
            self.table.setItem(row, 1, QTableWidgetItem(""))
            self.table.setItem(row, 2, QTableWidgetItem(""))
            
            # 저장된 지정상품 정보가 있으면 표시
            if row_id in self.parent.assigned_products and status == 1:
                self.table.item(row, 1).setText(self.parent.assigned_products[row_id])
            
            # 저장된 지정채널 정보가 있으면 표시
            if row_id in self.parent.assigned_channels and status == 1:
                self.table.item(row, 2).setText(self.parent.assigned_channels[row_id])
            
            # 버튼 클릭 시 상태 저장 및 행 색상 변경을 위한 연결
            status_btn.clicked.connect(lambda checked, r=row_id, btn=status_btn, row_idx=row: 
                                     self.update_row_status(r, btn.get_status(), row_idx))
            
            # 데이터 행 채우기
            for col in range(len(df.columns)):
                # 실제 열 인덱스 (상태 버튼 칼럼과 2개의 추가 칼럼 때문에 +3)
                table_col_idx = col + 3
                
                # URL 필드인 경우 URLTableWidgetItem 사용
                if table_col_idx == url_column_idx:
                    url_text = str(df.iloc[row, col])
                    item = URLTableWidgetItem(url_text)
                else:
                    item = QTableWidgetItem(str(df.iloc[row, col]))
                
                self.table.setItem(row, table_col_idx, item)
            
            # 행 색상 항상 적용 (필터 후에도 색상 유지)
            self.color_row(row, status)
        
        # 칼럼 너비 설정
        self.table.setColumnWidth(0, 80)  # 상태 버튼 칼럼 너비 고정
        self.table.setColumnWidth(1, 150)  # 지정상품 칼럼 너비 고정
        self.table.setColumnWidth(2, 100)  # 지정채널 칼럼 너비 고정
        
        # 특정 칼럼 너비 고정
        if product_column_idx != -1:
            self.table.setColumnWidth(product_column_idx, 300)  # 희망상품 칼럼 너비
        
        if channel_column_idx != -1:
            self.table.setColumnWidth(channel_column_idx, 150)  # 신청채널 칼럼 너비
        
        if url_column_idx != -1:
            self.table.setColumnWidth(url_column_idx, 250)  # URL 칼럼 너비
        
        if name_column_idx != -1:
            self.table.setColumnWidth(name_column_idx, 200)  # 이름 및 닉네임 칼럼 너비
        
        # 나머지 칼럼 너비 자동 조정
        for i in range(3, self.table.columnCount()):
            if i not in [product_column_idx, channel_column_idx, url_column_idx, name_column_idx]:
                self.table.resizeColumnToContents(i)
        
        # 상태 업데이트를 위한 타이머 재시작
        self.parent.status_timer.start(3000)  # 3초 후 상태 메시지 업데이트
        
        # 상태 통계 업데이트
        self.parent.update_status_statistics()
    
    def update_row_status(self, row_id, status, row_idx):
        """행 상태 업데이트"""
        old_status = self.parent.row_status.get(row_id, 0)
        
        # 상태 저장
        self.parent.row_status[row_id] = status
        
        # 상태 변경 플래그 설정
        self.parent.is_state_modified = True
        
        # 행 색상 변경
        self.color_row(row_idx, status)
        
        # 지정상품 칼럼에 상태 표시 업데이트
        item_product = self.table.item(row_idx, 1)  # 1은 지정상품 칼럼 인덱스
        item_channel = self.table.item(row_idx, 2)  # 2는 지정채널 칼럼 인덱스
        
        if item_product and item_channel:
            if status == 1:  # 선정 상태
                # 콤보박스에서 선택된 상품명 가져오기
                selected_product = self.parent.product_combo.currentText()
                
                # '전체'가 선택되었거나 선택된 항목이 없는 경우 '선정완료'로 표시
                if not selected_product or selected_product == "전체":
                    display_text = "선정완료"
                else:
                    display_text = selected_product
                    
                item_product.setText(display_text)
                
                # 지정상품 정보 저장
                self.parent.assigned_products[row_id] = display_text
                
                # 지정채널 업데이트
                selected_channel = self.parent.get_selected_channel()
                if selected_channel:
                    item_channel.setText(selected_channel)
                    # 지정채널 정보 저장
                    self.parent.assigned_channels[row_id] = selected_channel
            else:
                # 다른 상태일 때는 비움
                item_product.setText("")
                item_channel.setText("")
                
                # 지정상품 및 채널 정보 삭제
                if row_id in self.parent.assigned_products:
                    del self.parent.assigned_products[row_id]
                if row_id in self.parent.assigned_channels:
                    del self.parent.assigned_channels[row_id]
        
        # 선정(1) -> 다른 상태로 변경된 경우, 관련 완료 상태 해제
        if old_status == 1 and status != 1 and self.parent.contact_column_idx != -1:
            self.clear_completed_status_for_contact(row_id)
        
        # 다른 상태 -> 선정(1) 상태로 변경된 경우, 동일 연락처 행들을 완료로 변경
        elif status == 1 and self.parent.contact_column_idx != -1:
            self.mark_duplicate_contacts_as_completed(row_id)
        
        # 상태 통계 업데이트
        self.parent.update_status_statistics()
        
        # 테이블 리프레시
        self.update_table(self.parent.filtered_df)
        
        # UI 강제 업데이트
        QApplication.processEvents()
    
    def color_row(self, row, status):
        """행 배경색 설정"""
        # 상태에 따라 행 배경색 설정
        bgColor = self.row_colors[status]
        
        # 모든 셀에 배경색 적용
        for col in range(1, self.table.columnCount()):  # 상태 버튼 제외하고 적용
            item = self.table.item(row, col)
            if item:
                if bgColor:
                    item.setBackground(QColor(bgColor))
                else:
                    # 기본 색상으로 되돌리기
                    item.setBackground(QColor(255, 255, 255))
    
    def on_cell_clicked(self, row, column):
        """테이블 셀 클릭 이벤트 핸들러"""
        # 클릭된 테이블 식별 (sender() 메서드 사용)
        sender_table = self.parent.sender()
        
        # URL 열 인덱스 확인 (테이블 내에서의 인덱스)
        url_table_idx = -1
        for i, col in enumerate(self.parent.filtered_df.columns):
            col_str = str(col).lower()
            if "url" in col_str or "계정 링크" in col or "블로그" in col:
                url_table_idx = i + 3  # +3은 상태, 지정상품, 지정채널 칼럼 때문
                break
        
        if column == url_table_idx:
            # 클릭된 셀의 텍스트 가져오기 (실제 클릭된 테이블에서)
            url_text = sender_table.item(row, column).text()
            if url_text and url_text.strip():
                # URL 형식 확인 및 수정
                url = url_text.strip()
                if not url.startswith(('http://', 'https://')):
                    url = 'https://' + url
                    
                try:
                    # 기본 웹 브라우저로 URL 열기
                    webbrowser.open(url)
                except Exception as e:
                    self.parent.status_label.setText(f"URL을 열 수 없습니다: {str(e)}")
    
    def clear_completed_status_for_contact(self, row_id):
        """연락처 관련 완료 상태 해제"""
        # 해당 행의 연락처 확인
        if self.parent.contact_column_idx == -1 or row_id not in self.parent.original_df.index:
            return
        
        # 변경하고자 하는 행의 연락처 가져오기
        contact = self.parent.original_df.iloc[row_id, self.parent.contact_column_idx]
        if pd.isna(contact):
            return
        
        contact = str(contact)
        
        # 동일 연락처를 가진 행 중 완료 상태인 항목 찾기
        if contact in self.parent.contact_rows:
            for related_row_id in self.parent.contact_rows[contact]:
                if related_row_id in self.parent.row_status and self.parent.row_status[related_row_id] == 4:
                    # 완료 상태 해제하고 원래 상태로 되돌림
                    if related_row_id in self.parent.original_status:
                        self.parent.row_status[related_row_id] = self.parent.original_status[related_row_id]
                        del self.parent.original_status[related_row_id]
                    else:
                        # 원래 상태 정보가 없으면 미정(0)으로 설정
                        self.parent.row_status[related_row_id] = 0
    
    def mark_duplicate_contacts_as_completed(self, row_id):
        """동일 연락처 행들 완료 상태로 변경"""
        # 해당 행의 연락처 확인
        if self.parent.contact_column_idx == -1 or row_id not in self.parent.original_df.index:
            return
        
        # 변경하고자 하는 행의 연락처 가져오기
        contact = self.parent.original_df.iloc[row_id, self.parent.contact_column_idx]
        if pd.isna(contact):
            return
        
        contact = str(contact)
        print(f"Processing contact: {contact}")  # 로그 추가
        
        # 동일 연락처를 가진 다른 행들 찾기
        if contact in self.parent.contact_rows:
            for related_row_id in self.parent.contact_rows[contact]:
                # 현재 행은 건너뜀
                if related_row_id == row_id:
                    continue
                
                # 완료 상태가 아닌 행만 처리
                current_status = self.parent.row_status.get(related_row_id, 0)
                if current_status != 4:
                    # 기존 상태 저장 후 완료 상태로 변경
                    self.parent.original_status[related_row_id] = current_status
                    self.parent.row_status[related_row_id] = 4
                    print(f"Row {related_row_id} marked as completed")  # 로그 추가

    def update_table_widget(self, table_widget, df):
        """특정 테이블 위젯 데이터 업데이트"""
        if df is None or len(df) == 0:
            table_widget.setRowCount(0)
            table_widget.setColumnCount(0)
            return
        
        # 테이블 위젯 설정
        table_widget.setRowCount(len(df))
        table_widget.setColumnCount(len(df.columns) + 3)  # 상태 버튼 칼럼 + 지정상품 + 지정채널
        
        # 헤더 레이블 설정 - 매핑된 이름 사용
        header_labels = ["상태", "지정상품", "지정채널"]  # 첫 번째 열은 상태 버튼, 이후 두 개는 새 칼럼
        for col in df.columns:
            # 매핑 정보가 있으면 매핑된 이름 사용, 없으면 원래 이름 사용
            mapped_name = self.header_mapping.get(col, col)
            header_labels.append(mapped_name)
        
        table_widget.setHorizontalHeaderLabels(header_labels)
        
        # 열 인덱스 찾기 (테이블 내에서의 인덱스)
        product_column_idx = -1  # 희망상품 
        channel_column_idx = -1  # 신청채널
        url_column_idx = -1      # URL
        name_column_idx = -1     # 이름 및 닉네임
        
        for i, col in enumerate(df.columns):
            col_str = str(col).lower()
            if "희망상품" in col or "희망 상품" in col:
                product_column_idx = i + 3  # +3은 상태 버튼과 지정상품, 지정채널 칼럼 때문
            elif "신청 채널" in col or "신청채널" in col:
                channel_column_idx = i + 3
            elif "url" in col_str or "계정 링크" in col or "블로그" in col:
                url_column_idx = i + 3
            elif "성함" in col or "이름" in col or "닉네임" in col:
                name_column_idx = i + 3
        
        # 데이터 채우기
        for row in range(len(df)):
            # 원본 데이터프레임에서의 인덱스(행 ID)
            row_id = df.index[row]
            
            # 상태 버튼 추가
            status_btn = StatusButton(row_id)
            status = 0  # 기본 상태
            
            # 저장된 상태가 있으면 적용
            if row_id in self.parent.row_status:
                status = self.parent.row_status[row_id]
                status_btn.set_status(status)
            
            table_widget.setCellWidget(row, 0, status_btn)
            
            # 새 칼럼(지정상품, 지정채널)을 위한 빈 아이템 추가
            table_widget.setItem(row, 1, QTableWidgetItem(""))
            table_widget.setItem(row, 2, QTableWidgetItem(""))
            
            # 저장된 지정상품 정보가 있으면 표시
            if row_id in self.parent.assigned_products and status == 1:
                table_widget.item(row, 1).setText(self.parent.assigned_products[row_id])
            
            # 저장된 지정채널 정보가 있으면 표시
            if row_id in self.parent.assigned_channels and status == 1:
                table_widget.item(row, 2).setText(self.parent.assigned_channels[row_id])
            
            # 버튼 클릭 시 상태 저장 및 행 색상 변경을 위한 연결
            status_btn.clicked.connect(lambda checked, r=row_id, btn=status_btn, row_idx=row, table=table_widget: 
                                     self.update_row_status_for_table(r, btn.get_status(), row_idx, table))
            
            # 데이터 행 채우기
            for col in range(len(df.columns)):
                # 실제 열 인덱스 (상태 버튼 칼럼과 2개의 추가 칼럼 때문에 +3)
                table_col_idx = col + 3
                
                # URL 필드인 경우 URLTableWidgetItem 사용
                if table_col_idx == url_column_idx:
                    url_text = str(df.iloc[row, col])
                    item = URLTableWidgetItem(url_text)
                else:
                    item = QTableWidgetItem(str(df.iloc[row, col]))
                
                table_widget.setItem(row, table_col_idx, item)
            
            # 행 색상 항상 적용 (필터 후에도 색상 유지)
            self.color_row_for_table(table_widget, row, status)
        
        # 칼럼 너비 설정
        table_widget.setColumnWidth(0, 80)  # 상태 버튼 칼럼 너비 고정
        table_widget.setColumnWidth(1, 150)  # 지정상품 칼럼 너비 고정
        table_widget.setColumnWidth(2, 100)  # 지정채널 칼럼 너비 고정
        
        # 특정 칼럼 너비 고정
        if product_column_idx != -1:
            table_widget.setColumnWidth(product_column_idx, 300)  # 희망상품 칼럼 너비
        
        if channel_column_idx != -1:
            table_widget.setColumnWidth(channel_column_idx, 150)  # 신청채널 칼럼 너비
        
        if url_column_idx != -1:
            table_widget.setColumnWidth(url_column_idx, 250)  # URL 칼럼 너비
        
        if name_column_idx != -1:
            table_widget.setColumnWidth(name_column_idx, 200)  # 이름 및 닉네임 칼럼 너비
        
        # 나머지 칼럼 너비 자동 조정
        for i in range(3, table_widget.columnCount()):
            if i not in [product_column_idx, channel_column_idx, url_column_idx, name_column_idx]:
                table_widget.resizeColumnToContents(i)

    def update_row_status_for_table(self, row_id, status, row_idx, table_widget):
        """특정 테이블의 행 상태 업데이트"""
        old_status = self.parent.row_status.get(row_id, 0)
        
        # 상태 저장
        self.parent.row_status[row_id] = status
        
        # 상태 변경 플래그 설정
        self.parent.is_state_modified = True
        
        # 행 색상 변경
        self.color_row_for_table(table_widget, row_idx, status)
        
        # 지정상품 칼럼에 상태 표시 업데이트
        item_product = table_widget.item(row_idx, 1)  # 1은 지정상품 칼럼 인덱스
        item_channel = table_widget.item(row_idx, 2)  # 2는 지정채널 칼럼 인덱스
        
        if item_product and item_channel:
            if status == 1:  # 선정 상태
                # 콤보박스에서 선택된 상품명 가져오기
                selected_product = self.parent.product_combo.currentText()
                
                # '전체'가 선택되었거나 선택된 항목이 없는 경우 '선정완료'로 표시
                if not selected_product or selected_product == "전체":
                    display_text = "선정완료"
                else:
                    display_text = selected_product
                    
                item_product.setText(display_text)
                
                # 지정상품 정보 저장
                self.parent.assigned_products[row_id] = display_text
                
                # 지정채널 업데이트
                selected_channel = self.parent.get_selected_channel()
                if selected_channel:
                    item_channel.setText(selected_channel)
                    # 지정채널 정보 저장
                    self.parent.assigned_channels[row_id] = selected_channel
            else:
                # 다른 상태일 때는 비움
                item_product.setText("")
                item_channel.setText("")
                
                # 지정상품 및 채널 정보 삭제
                if row_id in self.parent.assigned_products:
                    del self.parent.assigned_products[row_id]
                if row_id in self.parent.assigned_channels:
                    del self.parent.assigned_channels[row_id]
        
        # 선정(1) -> 다른 상태로 변경된 경우, 관련 완료 상태 해제
        if old_status == 1 and status != 1 and self.parent.contact_column_idx != -1:
            self.clear_completed_status_for_contact(row_id)
        
        # 다른 상태 -> 선정(1) 상태로 변경된 경우, 동일 연락처 행들을 완료로 변경
        elif status == 1 and self.parent.contact_column_idx != -1:
            self.mark_duplicate_contacts_as_completed(row_id)
        
        # 상태 통계 업데이트
        self.parent.update_status_statistics()
        
        # 테이블 리프레시 - 각 탭의 테이블 업데이트
        current_tab_index = self.parent.tab_widget.currentIndex()
        if current_tab_index == 0:  # 데이터 탭
            self.update_table(self.parent.filtered_df)
        else:
            # 현재 탭의 테이블에 데이터 업데이트
            self.update_table_widget(table_widget, self.parent.filtered_df)
        
        # UI 강제 업데이트
        QApplication.processEvents()

    def color_row_for_table(self, table_widget, row, status):
        """특정 테이블의 행 배경색 설정"""
        # 상태에 따라 행 배경색 설정
        bgColor = self.row_colors[status]
        
        # 모든 셀에 배경색 적용
        for col in range(1, table_widget.columnCount()):  # 상태 버튼 제외하고 적용
            item = table_widget.item(row, col)
            if item:
                if bgColor:
                    item.setBackground(QColor(bgColor))
                else:
                    # 기본 색상으로 되돌리기
                    item.setBackground(QColor(255, 255, 255)) 