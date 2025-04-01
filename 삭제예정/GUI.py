import sys
import pandas as pd
import re
import json
import os
import webbrowser  # URL 열기를 위한 모듈 추가
from PyQt5.QtWidgets import (QApplication, QMainWindow, QTableWidget, QTableWidgetItem, 
                            QVBoxLayout, QWidget, QPushButton, QFileDialog, QLabel, 
                            QHBoxLayout, QLineEdit, QGroupBox, QCheckBox, QMessageBox)
from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtGui import QColor, QFont

class URLTableWidgetItem(QTableWidgetItem):
    """URL을 포함하는 테이블 아이템 클래스"""
    def __init__(self, url_text):
        super().__init__(url_text)
        self.url = url_text.strip()
        # 링크 스타일 적용
        font = QFont()
        font.setUnderline(True)
        self.setFont(font)
        self.setForeground(QColor("blue"))
        # 툴팁 설정
        self.setToolTip(f"클릭하여 열기: {self.url}")

class StatusButton(QPushButton):
    def __init__(self, row_id, parent=None):
        super().__init__(parent)
        self.row_id = row_id
        self.status = 0  # 0: 기본, 1: 초록(선정), 2: 노랑(대기), 3: 빨강(제외), 4: 회색(완료)
        self.setFixedSize(80, 30)  # 버튼 크기 증가
        self.setText("")
        self.clicked.connect(self.change_status)
        
    def change_status(self):
        # 완료 상태(4)는 수동으로 변경할 수 없음
        if self.status == 4:
            return
            
        self.status = (self.status + 1) % 4  # 0, 1, 2, 3만 순환
        self.update_color()
        
    def update_color(self):
        if self.status == 0:
            self.setStyleSheet("")
            self.setText("")
        elif self.status == 1:
            self.setStyleSheet("background-color: #CCFFCC; color: #006600;")  # 파스텔 초록
            self.setText("선정")
        elif self.status == 2:
            self.setStyleSheet("background-color: #FFFACD; color: #8B8000;")  # 파스텔 노랑
            self.setText("대기")
        elif self.status == 3:
            self.setStyleSheet("background-color: #FFCCCC; color: #CC0000;")  # 파스텔 빨강
            self.setText("제외")
        elif self.status == 4:
            self.setStyleSheet("background-color: #999999; color: #FFFFFF;")  # 진한 회색, 흰색 텍스트
            self.setText("완료")
    
    def get_status(self):
        return self.status
    
    def set_status(self, status):
        self.status = status
        self.update_color()

class ExcelViewer(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("엑셀 데이터 뷰어")
        self.setGeometry(100, 100, 1200, 600)
        
        # 원본 데이터프레임 저장 변수
        self.original_df = None
        self.filtered_df = None
        self.excel_file_path = ""  # 현재 로드된 엑셀 파일 경로
        
        # 상태 버튼 데이터 저장
        self.row_status = {}  # {row_id: status}
        
        # 연락처 인덱스 저장
        self.contact_column_idx = -1  # 연락처 컬럼 인덱스
        self.name_column_idx = -1     # 이름 컬럼 인덱스
        self.product_column_idx = -1  # 희망상품 칼럼 인덱스
        
        # 연락처별 선정된 행 ID 저장
        self.contact_selection = {}  # {연락처: 선정된_행_ID}
        
        # 완료 상태로 변경된 행들의 원래 상태 저장
        self.original_status = {}  # {row_id: 원래_상태}
        
        # 연락처별 관련 행 ID 저장
        self.contact_rows = {}  # {연락처: [row_id1, row_id2, ...]}
        
        # 상태에 따른 행 배경색
        self.row_colors = {
            0: "",  # 기본 - 색상 없음
            1: "#EAFFEA",  # 선정 - 파스텔 초록
            2: "#FFFEF0",  # 대기 - 파스텔 노랑
            3: "#FFF0F0",  # 제외 - 파스텔 빨강
            4: "#999999"   # 완료 - 진한 회색
        }
        
        # 헤더 매핑 정보 설정
        self.header_mapping = {
            "● 희망상품(복수 신청가능)": "희망상품",
            "● 신청 채널을 선택해주세요.": "신청채널",
            "● 계정 링크 입력해주세요 (블로그 및 인스타 주소)": "URL",
            "● 팔로워수 혹은 평균 일 방문자수 선택": "일방문 및 팔로워수",
            "● 이웃활동을 열심히 하시는 편이신가요?": "이웃활동",
            "● 성함 (닉네임) --- ex) 홍길동 (해운대럭키가이)": "이름 및 닉네임",
            "● 연락처 ( 예- 01021456993 )": "연락처",
            "● 카톡아이디(연락처 오입력 시 연락)": "카톡아이디"
        }
        
        # 채널 목록 정의
        self.channel_list = ["블로그", "인스타 - 피드", "인스타 - 릴스"]
        self.channel_checkboxes = {}
        
        # 중앙 위젯 설정
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # 레이아웃 설정
        layout = QVBoxLayout(central_widget)
        
        # 파일 선택 버튼
        self.file_btn = QPushButton("엑셀 파일 선택")
        self.file_btn.clicked.connect(self.load_excel_file)
        layout.addWidget(self.file_btn)
        
        # 상태 라벨
        self.status_label = QLabel("파일을 선택해주세요.")
        layout.addWidget(self.status_label)
        
        # 필터 영역 설정 - 가로 레이아웃으로 변경
        filter_layout = QHBoxLayout()
        
        # 텍스트 검색 영역 - 세로 레이아웃으로 변경
        text_search_layout = QVBoxLayout()
        
        # 검색 섹션 그룹박스
        search_group = QGroupBox("상품 검색")
        search_inner_layout = QVBoxLayout()
        
        # 검색어 입력 필드
        self.filter_label = QLabel("검색어:")
        search_inner_layout.addWidget(self.filter_label)
        
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("상품명 검색 (예: 뽀로로)")
        search_inner_layout.addWidget(self.search_input)
        
        # 단일 상품 필터 체크박스 추가
        self.single_product_checkbox = QCheckBox("단일 상품만 보기")
        # 즉시 반영을 위해 stateChanged 대신 clicked 시그널 사용
        self.single_product_checkbox.clicked.connect(self.apply_filters)
        search_inner_layout.addWidget(self.single_product_checkbox)
        
        # 검색 버튼
        self.search_btn = QPushButton("검색")
        self.search_btn.clicked.connect(self.apply_filters)
        search_inner_layout.addWidget(self.search_btn)
        
        search_group.setLayout(search_inner_layout)
        text_search_layout.addWidget(search_group)
        
        # 이름/연락처/URL 검색 그룹박스 추가
        contact_search_group = QGroupBox("이름/연락처/URL 검색")
        contact_search_layout = QHBoxLayout()  # 가로 레이아웃으로 변경
        
        # 이름/연락처 검색 입력 필드
        self.contact_search_label = QLabel("검색어:")
        contact_search_layout.addWidget(self.contact_search_label, 1)  # 비율 1
        
        self.contact_search_input = QLineEdit()
        self.contact_search_input.setPlaceholderText("이름, 연락처, URL 검색")
        contact_search_layout.addWidget(self.contact_search_input, 3)  # 비율 3 (가장 넓게)
        
        # 검색 버튼
        self.contact_search_btn = QPushButton("검색")
        self.contact_search_btn.clicked.connect(self.apply_filters)
        contact_search_layout.addWidget(self.contact_search_btn, 1)  # 비율 1
        
        contact_search_group.setLayout(contact_search_layout)
        text_search_layout.addWidget(contact_search_group)
        
        # 초기화 버튼
        self.reset_btn = QPushButton("필터 초기화")
        self.reset_btn.clicked.connect(self.reset_filter)
        text_search_layout.addWidget(self.reset_btn)
        
        # 상태 필터 그룹박스 추가
        status_filter_group = QGroupBox("상태 필터")
        status_filter_layout = QVBoxLayout()
        
        # 완료 상태 제외 체크박스
        self.hide_completed_checkbox = QCheckBox("완료 상태 항목 제외")
        self.hide_completed_checkbox.clicked.connect(self.apply_filters)  # 체크 상태 변경 시 필터 적용
        status_filter_layout.addWidget(self.hide_completed_checkbox)
        
        status_filter_group.setLayout(status_filter_layout)
        text_search_layout.addWidget(status_filter_group)  # text_search_layout에 추가
        
        filter_layout.addLayout(text_search_layout)
        
        # 채널 필터 그룹박스
        channel_group = QGroupBox("채널 필터")
        channel_layout = QVBoxLayout()  # 세로 레이아웃으로 변경
        
        # 채널 체크박스 생성
        for channel in self.channel_list:
            checkbox = QCheckBox(channel)
            checkbox.setChecked(True)  # 기본적으로 모든 채널 선택
            # 즉시 반영을 위해 clicked 신호 사용
            checkbox.clicked.connect(self.apply_filters)
            self.channel_checkboxes[channel] = checkbox
            channel_layout.addWidget(checkbox)
        
        channel_group.setLayout(channel_layout)
        filter_layout.addWidget(channel_group)
        
        layout.addLayout(filter_layout)
        
        # 통계 및 저장 버튼 영역
        stats_save_layout = QHBoxLayout()
        
        # 통계 정보 표시 레이블
        self.stats_label = QLabel("상태 통계: 데이터를 불러오세요")
        stats_save_layout.addWidget(self.stats_label, 3)  # 3:1 비율로 공간 할당
        
        # 저장 버튼 영역 (가로 배치)
        save_buttons_layout = QHBoxLayout()
        
        # 현재 화면 저장 버튼
        self.save_btn = QPushButton("현재 화면 저장")
        self.save_btn.clicked.connect(self.save_current_view)
        save_buttons_layout.addWidget(self.save_btn)
        
        # 작업 상태 저장 버튼
        self.save_state_btn = QPushButton("작업 상태 저장")
        self.save_state_btn.clicked.connect(self.save_work_state)
        save_buttons_layout.addWidget(self.save_state_btn)
        
        # 작업 상태 불러오기 버튼
        self.load_state_btn = QPushButton("상태 불러오기")
        self.load_state_btn.clicked.connect(self.load_work_state)
        save_buttons_layout.addWidget(self.load_state_btn)
        
        # 버튼 레이아웃을 통계 영역에 추가
        stats_save_layout.addLayout(save_buttons_layout, 2)
        
        layout.addLayout(stats_save_layout)
        
        # 테이블 위젯 생성
        self.table = QTableWidget()
        # 셀 클릭 이벤트 연결
        self.table.cellClicked.connect(self.on_cell_clicked)
        layout.addWidget(self.table)
        
        # 상태 메시지 업데이트를 위한 타이머
        self.status_timer = QTimer()
        self.status_timer.setSingleShot(True)
        self.status_timer.timeout.connect(self.clear_status_after_delay)
        
        # URL 컬럼 인덱스 저장
        self.url_column_idx = -1  # URL 칼럼 인덱스
    
    def update_status_statistics(self):
        """현재 테이블에 표시된 행의 상태 통계 업데이트"""
        if self.filtered_df is None:
            self.stats_label.setText("상태 통계: 데이터를 불러오세요")
            return
            
        total_rows = len(self.filtered_df)
        
        # 현재 필터링된 데이터의 인덱스 목록
        current_indices = self.filtered_df.index.tolist()
        
        # 상태별 개수 계산
        status_counts = {0: 0, 1: 0, 2: 0, 3: 0, 4: 0}
        
        for idx in current_indices:
            status = self.row_status.get(idx, 0)  # 상태가 없으면 기본값 0
            status_counts[status] += 1
            
        # 통계 텍스트 생성
        stats_text = f"총 {total_rows}개 행 | "
        stats_text += f"선정: {status_counts[1]}개 | "
        stats_text += f"대기: {status_counts[2]}개 | "
        stats_text += f"제외: {status_counts[3]}개 | "
        stats_text += f"완료: {status_counts[4]}개"
        
        self.stats_label.setText(stats_text)
    
    def save_current_view(self):
        """현재 화면에 표시된 데이터를 엑셀 파일로 저장"""
        if self.filtered_df is None or len(self.filtered_df) == 0:
            self.status_label.setText("저장할 데이터가 없습니다.")
            return
            
        # 파일 저장 다이얼로그 열기
        file_path, _ = QFileDialog.getSaveFileName(self, "파일 저장", "", "Excel Files (*.xlsx)")
        
        if not file_path:
            return  # 사용자가 취소함
            
        try:
            # 상태 정보 추가
            export_df = self.filtered_df.copy()
            
            # 상태 정보 칼럼 추가
            export_df['상태'] = ''
            
            # 행 인덱스와 상태 정보 매핑
            for idx in export_df.index:
                if idx in self.row_status:
                    status_code = self.row_status[idx]
                    if status_code == 1:
                        export_df.at[idx, '상태'] = '선정'
                    elif status_code == 2:
                        export_df.at[idx, '상태'] = '대기'
                    elif status_code == 3:
                        export_df.at[idx, '상태'] = '제외'
                    elif status_code == 4:
                        export_df.at[idx, '상태'] = '완료'
            
            # 필요하면 헤더 매핑을 적용한 컬럼명으로 변경
            mapped_columns = {}
            for col in export_df.columns:
                if col in self.header_mapping:
                    mapped_columns[col] = self.header_mapping[col]
            
            if mapped_columns:
                export_df = export_df.rename(columns=mapped_columns)
            
            # 엑셀 파일로 저장
            export_df.to_excel(file_path, index=False)
            
            self.status_label.setText(f"현재 화면이 '{file_path}'에 저장되었습니다.")
            
        except Exception as e:
            self.status_label.setText(f"저장 중 오류 발생: {str(e)}")
    
    def format_phone_number(self, phone):
        """전화번호 형식 검증 및 변환"""
        try:
            # 숫자만 추출
            numbers_only = re.sub(r'[^\d]', '', str(phone))
            
            # 010으로 시작하고 11자리인지 확인
            if len(numbers_only) == 11 and numbers_only.startswith('010'):
                # 하이픈 추가하여 형식화
                return f'{numbers_only[:3]}-{numbers_only[3:7]}-{numbers_only[7:]}'
            else:
                return '오류'
        except:
            return '오류'
    
    def load_excel_file(self):
        # 파일 다이얼로그 열기
        file_path, _ = QFileDialog.getOpenFileName(self, "엑셀 파일 선택", "", "Excel Files (*.xlsx *.xls)")
        
        if file_path:
            try:
                # 엑셀 파일 읽기
                self.original_df = pd.read_excel(file_path)
                
                # 상태 버튼 데이터 초기화
                self.row_status = {}
                
                # 연락처 컬럼 찾기
                self.contact_column_idx = -1
                for i, col in enumerate(self.original_df.columns):
                    col_str = str(col).lower()
                    if "연락처" in col_str:
                        self.contact_column_idx = i
                        break
                
                # 연락처 데이터 전처리
                for col in self.original_df.columns:
                    if ("연락처" in col or "전화" in col) and not ("카톡" in col or "아이디" in col):
                        # 연락처 칼럼 인덱스 저장
                        self.contact_column_idx = self.original_df.columns.get_loc(col)
                        self.original_df[col] = self.original_df[col].apply(self.format_phone_number)
                    # 이름 칼럼 인덱스 저장
                    if "성함" in col or "이름" in col or "닉네임" in col:
                        self.name_column_idx = self.original_df.columns.get_loc(col)
                    # 희망상품 칼럼 인덱스 저장
                    if "희망상품" in col or "희망 상품" in col:
                        self.product_column_idx = self.original_df.columns.get_loc(col)
                    # URL 칼럼 인덱스 저장
                    col_str = str(col).lower()
                    if "url" in col_str or "계정 링크" in col or "블로그" in col:
                        self.url_column_idx = self.original_df.columns.get_loc(col)
                
                # C열부터 N열 선택 (K열, M열 제외)
                columns_to_show = []
                for col in self.original_df.columns[2:14]:  # C(인덱스 2)부터 N(인덱스 13)까지
                    col_idx = self.original_df.columns.get_loc(col)
                    # K열(인덱스 10)과 M열(인덱스 12)는 제외
                    if col_idx != 10 and col_idx != 12:
                        columns_to_show.append(col)
                
                # 선택된 열만 포함하는 데이터프레임 생성
                self.filtered_df = self.original_df[columns_to_show]
                
                # 테이블 업데이트
                self.update_table(self.filtered_df)
                
                self.excel_file_path = file_path  # 파일 경로 저장
                
                self.status_label.setText(f"파일 '{file_path}'이(가) 로드되었습니다.")
                
                # 상태 통계 업데이트
                self.update_status_statistics()
                
            except Exception as e:
                self.status_label.setText(f"오류 발생: {str(e)}")
        else:
            self.status_label.setText("파일이 선택되지 않았습니다.")
    
    def update_table(self, df):
        """테이블 위젯 데이터 업데이트"""
        if df is None or len(df) == 0:
            self.table.setRowCount(0)
            self.table.setColumnCount(0)
            return
            
        # 테이블 위젯 설정
        self.table.setRowCount(len(df))
        self.table.setColumnCount(len(df.columns) + 1)  # 상태 버튼 칼럼 추가
        
        # 헤더 레이블 설정 - 매핑된 이름 사용
        header_labels = ["상태"]  # 첫 번째 열은 상태 버튼
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
                product_column_idx = i + 1  # +1은 상태 버튼 칼럼 때문
            elif "신청 채널" in col or "신청채널" in col:
                channel_column_idx = i + 1
            elif "url" in col_str or "계정 링크" in col or "블로그" in col:
                url_column_idx = i + 1
            elif "성함" in col or "이름" in col or "닉네임" in col:
                name_column_idx = i + 1
        
        # 데이터 채우기
        for row in range(len(df)):
            # 원본 데이터프레임에서의 인덱스(행 ID)
            row_id = df.index[row]
            
            # 상태 버튼 추가
            status_btn = StatusButton(row_id)
            status = 0  # 기본 상태
            
            # 저장된 상태가 있으면 적용
            if row_id in self.row_status:
                status = self.row_status[row_id]
                status_btn.set_status(status)
            
            self.table.setCellWidget(row, 0, status_btn)
            
            # 버튼 클릭 시 상태 저장 및 행 색상 변경을 위한 연결
            status_btn.clicked.connect(lambda checked, r=row_id, btn=status_btn, row_idx=row: 
                                       self.update_row_status(r, btn.get_status(), row_idx))
            
            # 데이터 행 채우기
            for col in range(len(df.columns)):
                # 실제 열 인덱스 (상태 버튼 칼럼 때문에 +1)
                table_col_idx = col + 1
                
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
        
        # 특정 칼럼 너비 고정
        if product_column_idx != -1:
            self.table.setColumnWidth(product_column_idx, 300)  # 희망상품 칼럼 너비를 300픽셀로 고정
        
        if channel_column_idx != -1:
            self.table.setColumnWidth(channel_column_idx, 150)  # 신청채널 칼럼 너비를 150픽셀로 고정
        
        if url_column_idx != -1:
            self.table.setColumnWidth(url_column_idx, 250)  # URL 칼럼 너비를 250픽셀로 고정
        
        if name_column_idx != -1:
            self.table.setColumnWidth(name_column_idx, 200)  # 이름 및 닉네임 칼럼 너비를 200픽셀로 고정
        
        # 나머지 칼럼 너비 자동 조정
        for i in range(1, self.table.columnCount()):
            if i not in [product_column_idx, channel_column_idx, url_column_idx, name_column_idx]:
                self.table.resizeColumnToContents(i)
        
        # 상태 업데이트를 위한 타이머 재시작
        self.status_timer.start(3000)  # 3초 후 상태 메시지 업데이트
        
        # 상태 통계 업데이트
        self.update_status_statistics()
    
    def update_row_status(self, row_id, status, row_idx):
        old_status = self.row_status.get(row_id, 0)
        
        # 상태 저장
        self.row_status[row_id] = status
        
        # 행 색상 변경
        self.color_row(row_idx, status)
        
        # 선정(1) -> 다른 상태로 변경된 경우, 관련 완료 상태 해제
        if old_status == 1 and status != 1 and self.contact_column_idx != -1:
            self.clear_completed_status_for_contact(row_id)
            
        # 다른 상태 -> 선정(1) 상태로 변경된 경우, 동일 연락처 행들을 완료로 변경
        elif status == 1 and self.contact_column_idx != -1:
            self.mark_duplicate_contacts_as_completed(row_id)
        
        # 상태 통계 업데이트
        self.update_status_statistics()
    
    def clear_completed_status_for_contact(self, unselected_row_id):
        """선정 취소된 행과 동일 연락처를 가진 완료 상태 행들을 원래 상태로 복원"""
        if self.original_df is None or self.contact_column_idx == -1:
            return
            
        # 선정 취소된 행의 연락처 값 가져오기
        contact = self.original_df.iloc[unselected_row_id, self.contact_column_idx]
        
        if pd.isna(contact) or contact == '' or contact == '오류':
            return  # 유효하지 않은 연락처면 처리 안 함
            
        # 연락처별 선정된 행 ID 목록에서 제거
        if contact in self.contact_selection and self.contact_selection[contact] == unselected_row_id:
            del self.contact_selection[contact]
            
        # 동일 연락처를 가진 완료 상태 행들 찾기
        completed_rows = []
        for idx, row in self.original_df.iterrows():
            if (idx != unselected_row_id and 
                row[self.contact_column_idx] == contact and 
                self.row_status.get(idx, 0) == 4 and
                idx in self.original_status):
                completed_rows.append(idx)
                
        # 완료 상태 행들을 원래 상태로 복원
        for idx in completed_rows:
            original_status = self.original_status.get(idx, 0)
            self.row_status[idx] = original_status
            
            # 원래 상태 기록에서 제거
            if idx in self.original_status:
                del self.original_status[idx]
                
        # 테이블 업데이트 - 현재 표시된 테이블에서 해당 행을 찾아 색상 변경
        for row in range(self.table.rowCount()):
            cell_widget = self.table.cellWidget(row, 0)
            if isinstance(cell_widget, StatusButton):
                row_id = cell_widget.row_id
                if row_id in completed_rows:
                    original_status = self.row_status[row_id]
                    cell_widget.set_status(original_status)
                    self.color_row(row, original_status)
    
    def mark_duplicate_contacts_as_completed(self, selected_row_id):
        """선정된 행과 동일한 연락처를 가진 다른 행들을 완료(상태=4)로 변경"""
        if self.original_df is None or self.contact_column_idx == -1:
            return
            
        # 선정된 행의 연락처 값 가져오기
        selected_contact = self.original_df.iloc[selected_row_id, self.contact_column_idx]
        
        if pd.isna(selected_contact) or selected_contact == '' or selected_contact == '오류':
            return  # 유효하지 않은 연락처면 처리 안 함
            
        # 연락처별 선정된 행 ID 저장
        self.contact_selection[selected_contact] = selected_row_id
        
        # 동일한 연락처를 가진 다른 행 찾기
        matching_indices = []
        for idx, row in self.original_df.iterrows():
            if idx != selected_row_id and row[self.contact_column_idx] == selected_contact:
                # 원래 상태 저장
                if idx not in self.original_status:
                    self.original_status[idx] = self.row_status.get(idx, 0)
                    
                matching_indices.append(idx)
        
        # 해당 행들의 상태를 완료(4)로 변경
        for idx in matching_indices:
            self.row_status[idx] = 4
        
        # 테이블 업데이트 - 현재 표시된 테이블에서 해당 행을 찾아 색상 변경
        for row in range(self.table.rowCount()):
            cell_widget = self.table.cellWidget(row, 0)
            if isinstance(cell_widget, StatusButton):
                row_id = cell_widget.row_id
                if row_id in matching_indices:
                    cell_widget.set_status(4)
                    self.color_row(row, 4)
    
    def color_row(self, row, status):
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
    
    def save_work_state(self):
        """작업 상태 중간 저장"""
        if self.original_df is None:
            QMessageBox.warning(self, "경고", "저장할 작업이 없습니다. 먼저 엑셀 파일을 로드해주세요.")
            return
            
        if not self.excel_file_path:
            QMessageBox.warning(self, "경고", "원본 엑셀 파일 경로를 찾을 수 없습니다.")
            return
            
        # 저장할 파일 경로 선택
        file_path, _ = QFileDialog.getSaveFileName(
            self, "작업 상태 저장", "", "작업 상태 파일 (*.workstate)"
        )
        
        if not file_path:
            return  # 사용자가 취소함
            
        try:
            # 저장할 상태 정보
            state_data = {
                "excel_file_path": self.excel_file_path,  # 원본 엑셀 파일 경로
                "row_status": {str(k): v for k, v in self.row_status.items()},  # 행 상태 정보 (인덱스는 문자열로 변환)
                "original_status": {str(k): v for k, v in self.original_status.items()},  # 원래 상태 정보
                "contact_selection": {k: str(v) for k, v in self.contact_selection.items() if k},  # 연락처별 선정 정보
                "timestamp": pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S")  # 저장 시간
            }
            
            # JSON 파일로 저장
            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(state_data, f, ensure_ascii=False, indent=2)
                
            self.status_label.setText(f"작업 상태가 '{file_path}'에 저장되었습니다.")
            
        except Exception as e:
            QMessageBox.critical(self, "오류", f"작업 상태 저장 중 오류 발생: {str(e)}")
    
    def load_work_state(self):
        """저장된 작업 상태 불러오기"""
        # 불러올 파일 경로 선택
        file_path, _ = QFileDialog.getOpenFileName(
            self, "작업 상태 불러오기", "", "작업 상태 파일 (*.workstate)"
        )
        
        if not file_path:
            return  # 사용자가 취소함
            
        try:
            # JSON 파일에서 상태 정보 불러오기
            with open(file_path, 'r', encoding='utf-8') as f:
                state_data = json.load(f)
                
            excel_path = state_data.get("excel_file_path", "")
            
            # 엑셀 파일이 존재하는지 확인
            if not os.path.exists(excel_path):
                answer = QMessageBox.question(
                    self, 
                    "엑셀 파일 없음", 
                    f"원본 엑셀 파일을 찾을 수 없습니다:\n{excel_path}\n\n새 엑셀 파일을 선택하시겠습니까?",
                    QMessageBox.Yes | QMessageBox.No, 
                    QMessageBox.Yes
                )
                
                if answer == QMessageBox.Yes:
                    new_excel_path, _ = QFileDialog.getOpenFileName(
                        self, "엑셀 파일 선택", "", "Excel 파일 (*.xlsx *.xls)"
                    )
                    if new_excel_path:
                        excel_path = new_excel_path
                    else:
                        return  # 파일 선택 취소
                else:
                    return  # 작업 취소
            
            # 엑셀 파일 로드
            self._load_excel_file(excel_path)
            
            # 상태 정보 복원
            row_status = {int(k): v for k, v in state_data.get("row_status", {}).items()}
            original_status = {int(k): v for k, v in state_data.get("original_status", {}).items()}
            contact_selection = {k: int(v) for k, v in state_data.get("contact_selection", {}).items() if k}
            
            # 상태 정보 적용
            self.row_status = row_status
            self.original_status = original_status
            self.contact_selection = contact_selection
            
            # 테이블 업데이트
            self.update_table(self.filtered_df)
            
            self.status_label.setText(f"작업 상태가 '{file_path}'에서 불러와졌습니다.")
            
        except Exception as e:
            QMessageBox.critical(self, "오류", f"작업 상태 불러오기 중 오류 발생: {str(e)}")
    
    def _load_excel_file(self, file_path):
        """엑셀 파일 로드 내부 메서드"""
        try:
            # 엑셀 파일 로드
            self.original_df = pd.read_excel(file_path)
            self.excel_file_path = file_path  # 파일 경로 저장
            
            # 연락처/이름/희망상품 칼럼 찾기 및 데이터 전처리
            self.contact_column_idx = -1
            self.name_column_idx = -1
            self.product_column_idx = -1
            
            # 매핑된 헤더 정보로 칼럼 매핑
            for original_header, mapped_header in self.header_mapping.items():
                for i, col in enumerate(self.original_df.columns):
                    if original_header in str(col):
                        self.original_df.rename(columns={col: original_header}, inplace=True)
                        break
            
            # 칼럼 인덱스 찾기 및 데이터 전처리
            for col in self.original_df.columns:
                if ("연락처" in col or "전화" in col) and not ("카톡" in col or "아이디" in col):
                    # 연락처 칼럼 인덱스 저장
                    self.contact_column_idx = self.original_df.columns.get_loc(col)
                    self.original_df[col] = self.original_df[col].apply(self.format_phone_number)
                # 이름 칼럼 인덱스 저장
                if "성함" in col or "이름" in col or "닉네임" in col:
                    self.name_column_idx = self.original_df.columns.get_loc(col)
                # 희망상품 칼럼 인덱스 저장
                if "희망상품" in col or "희망 상품" in col:
                    self.product_column_idx = self.original_df.columns.get_loc(col)
                # URL 칼럼 인덱스 저장
                col_str = str(col).lower()
                if "url" in col_str or "계정 링크" in col or "블로그" in col:
                    self.url_column_idx = self.original_df.columns.get_loc(col)
            
            # C열부터 N열 선택 (K열, M열 제외)
            columns_to_show = []
            for col in self.original_df.columns[2:14]:  # C(인덱스 2)부터 N(인덱스 13)까지
                col_idx = self.original_df.columns.get_loc(col)
                # K열(인덱스 10)과 M열(인덱스 12)는 제외
                if col_idx != 10 and col_idx != 12:
                    columns_to_show.append(col)
            
            # 필터링된 데이터프레임 생성
            self.filtered_df = self.original_df[columns_to_show]
            
            # 테이블 업데이트
            self.update_table(self.filtered_df)
            
            return True
            
        except Exception as e:
            QMessageBox.critical(self, "오류", f"엑셀 파일 로드 중 오류 발생: {str(e)}")
            return False
    
    def clear_status_after_delay(self):
        # 필터 상태 메시지 업데이트
        filter_msg = []
        
        search_text = self.search_input.text().strip()
        if search_text:
            filter_msg.append(f"검색어: '{search_text}'")
            
        if self.single_product_checkbox.isChecked():
            filter_msg.append("단일 상품만")
            
        selected_channels = [channel for channel, checkbox in self.channel_checkboxes.items() 
                            if checkbox.isChecked()]
        if len(selected_channels) < len(self.channel_list):
            filter_msg.append(f"선택된 채널: {', '.join(selected_channels)}")
            
        if filter_msg:
            self.status_label.setText(f"현재 적용된 필터: {', '.join(filter_msg)}")
        else:
            self.status_label.setText("모든 데이터가 표시됩니다.")
    
    def clean_url(self, url):
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

    def apply_filters(self):
        if self.original_df is None:
            self.status_label.setText("먼저 엑셀 파일을 로드해주세요.")
            return
        
        try:
            # 필터링 시작할 원본 데이터
            filtered_original = self.original_df.copy()
            
            # 1. 텍스트 검색 필터 적용
            search_text = self.search_input.text().strip()
            if search_text:
                # C열(인덱스 2) 기준으로 필터링
                c_column = filtered_original.columns[2]
                mask = filtered_original[c_column].astype(str).str.contains(search_text, case=False, na=False)
                filtered_original = filtered_original[mask]
            
            # 2. 단일 상품 필터 적용
            if self.single_product_checkbox.isChecked():
                c_column = filtered_original.columns[2]
                # 쉼표가 포함되지 않은 항목만 선택 (단일 상품)
                mask = ~filtered_original[c_column].astype(str).str.contains(',', na=False)
                filtered_original = filtered_original[mask]
            
            # 3. 이름/연락처/URL 검색 필터 적용
            contact_search_text = self.contact_search_input.text().strip()
            if contact_search_text:
                name_mask = pd.Series(False, index=filtered_original.index)
                contact_mask = pd.Series(False, index=filtered_original.index)
                url_mask = pd.Series(False, index=filtered_original.index)
                
                # 이름 칼럼 검색
                if self.name_column_idx != -1:
                    name_mask = filtered_original.iloc[:, self.name_column_idx].astype(str).str.contains(
                        contact_search_text, case=False, na=False
                    )
                
                # 연락처 칼럼 검색
                if self.contact_column_idx != -1:
                    contact_mask = filtered_original.iloc[:, self.contact_column_idx].astype(str).str.contains(
                        contact_search_text, case=False, na=False
                    )
                
                # URL 칼럼 검색 (접두사 제외 버전)
                if self.url_column_idx != -1:
                    # 검색어에서 접두사 제거
                    clean_search_text = self.clean_url(contact_search_text)
                    
                    # 데이터에서도 접두사 제거하여 비교
                    url_mask = filtered_original.iloc[:, self.url_column_idx].apply(
                        lambda x: clean_search_text in self.clean_url(x)
                    )
                
                # 이름, 연락처 또는 URL에 검색어 포함된 항목 선택
                combined_mask = name_mask | contact_mask | url_mask
                filtered_original = filtered_original[combined_mask]
            
            # 4. 완료 상태 항목 제외 필터 적용
            if self.hide_completed_checkbox.isChecked():
                # 완료 상태(4)가 아닌 행만 선택
                completed_indices = [idx for idx, status in self.row_status.items() if status == 4]
                filtered_original = filtered_original[~filtered_original.index.isin(completed_indices)]
            
            # 5. 채널 필터 적용
            selected_channels = [channel for channel, checkbox in self.channel_checkboxes.items() 
                                if checkbox.isChecked()]
            
            if selected_channels and len(selected_channels) < len(self.channel_list):
                # D열(인덱스 3) 기준으로 필터링
                d_column = filtered_original.columns[3]
                
                # 선택된 채널 포함 여부 체크를 위한 마스크
                channel_mask = filtered_original[d_column].apply(
                    lambda x: any(channel in str(x) for channel in selected_channels)
                )
                
                filtered_original = filtered_original[channel_mask]
            
            # 아무것도 선택되지 않았으면 경고
            if len(selected_channels) == 0:
                self.status_label.setText("최소 하나의 채널을 선택해주세요.")
                return
                
            if len(filtered_original) == 0:
                self.status_label.setText("필터 조건에 맞는 데이터가 없습니다.")
                self.update_status_statistics()  # 빈 결과도 통계 업데이트
                return
            
            # C열부터 N열 선택 (K열, M열 제외)
            columns_to_show = []
            for col in filtered_original.columns[2:14]:  # C(인덱스 2)부터 N(인덱스 13)까지
                col_idx = filtered_original.columns.get_loc(col)
                # K열(인덱스 10)과 M열(인덱스 12)는 제외
                if col_idx != 10 and col_idx != 12:
                    columns_to_show.append(col)
            
            # 필터링된 데이터프레임 생성
            self.filtered_df = filtered_original[columns_to_show]
            
            # 테이블 업데이트
            self.update_table(self.filtered_df)
            
            # 상태 메시지 업데이트
            filter_msg = []
            if search_text:
                filter_msg.append(f"상품검색: '{search_text}'")
            if contact_search_text:
                filter_msg.append(f"이름/연락처: '{contact_search_text}'")
            if self.single_product_checkbox.isChecked():
                filter_msg.append("단일 상품만")
            if self.hide_completed_checkbox.isChecked():
                filter_msg.append("완료 상태 제외")
            if len(selected_channels) < len(self.channel_list):
                filter_msg.append(f"선택된 채널: {', '.join(selected_channels)}")
                
            if filter_msg:
                self.status_label.setText(f"필터 적용 ({', '.join(filter_msg)}) - {len(self.filtered_df)}개의 행이 표시됩니다.")
            else:
                self.status_label.setText(f"모든 데이터가 표시됩니다. (총 {len(self.filtered_df)}개 행)")
            
        except Exception as e:
            self.status_label.setText(f"필터 적용 중 오류 발생: {str(e)}")
    
    def reset_filter(self):
        if self.original_df is None:
            return
            
        # 검색어 필드 초기화
        self.search_input.clear()
        self.contact_search_input.clear()  # 이름/연락처 검색 필드도 초기화
        
        # 단일 상품 체크박스 초기화
        self.single_product_checkbox.setChecked(False)
        
        # 완료 상태 제외 체크박스 초기화
        self.hide_completed_checkbox.setChecked(False)
        
        # 모든 채널 체크박스 선택
        for checkbox in self.channel_checkboxes.values():
            checkbox.setChecked(True)
            
        # C열부터 N열 선택 (K열, M열 제외)
        columns_to_show = []
        for col in self.original_df.columns[2:14]:  # C(인덱스 2)부터 N(인덱스 13)까지
            col_idx = self.original_df.columns.get_loc(col)
            # K열(인덱스 10)과 M열(인덱스 12)는 제외
            if col_idx != 10 and col_idx != 12:
                columns_to_show.append(col)
        
        # 선택된 열만 포함하는 데이터프레임 생성
        self.filtered_df = self.original_df[columns_to_show]
        
        # 테이블 업데이트
        self.update_table(self.filtered_df)
        
        self.status_label.setText("필터가 초기화되었습니다.")
    
    def on_cell_clicked(self, row, column):
        """테이블 셀 클릭 이벤트 핸들러"""
        # URL 열이 클릭되었는지 확인
        if self.filtered_df is None:
            return
        
        # URL 열 인덱스 확인 (테이블 내에서의 인덱스)
        url_table_idx = -1
        for i, col in enumerate(self.filtered_df.columns):
            col_str = str(col).lower()
            if "url" in col_str or "계정 링크" in col or "블로그" in col:
                url_table_idx = i + 1
                break
        
        if column == url_table_idx:
            # 클릭된 셀의 텍스트 가져오기
            url_text = self.table.item(row, column).text()
            if url_text and url_text.strip():
                # URL 형식 확인 및 수정
                url = url_text.strip()
                if not url.startswith(('http://', 'https://')):
                    url = 'https://' + url
                    
                try:
                    # 기본 웹 브라우저로 URL 열기
                    webbrowser.open(url)
                except Exception as e:
                    self.status_label.setText(f"URL을 열 수 없습니다: {str(e)}")
    
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ExcelViewer()
    window.show()
    sys.exit(app.exec_())
