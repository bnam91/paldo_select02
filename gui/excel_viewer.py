import sys
import pandas as pd
import re
import json
import os
import webbrowser
from PyQt5.QtWidgets import (QMainWindow, QTableWidget, QTableWidgetItem, 
                            QVBoxLayout, QWidget, QPushButton, QFileDialog, QLabel, 
                            QHBoxLayout, QMessageBox, QGridLayout, QTabWidget, QInputDialog, 
                            QComboBox, QCheckBox)
from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtGui import QColor
import datetime

from widgets import StatusButton, URLTableWidgetItem
from handlers import ExcelHandler, FilterHandler
from gui.ui_components import UIComponents
from gui.tab_manager import TabManager
from gui.table_manager import TableManager
from gui.filter_manager import FilterManager

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
        self.url_column_idx = -1      # URL 칼럼 인덱스
        
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
        
        # 채널 목록
        self.channel_list = ['블로그', '인스타 - 피드', '인스타 - 릴스', '유튜브', '유튜브 - 쇼츠']
        
        # 채널 체크박스 저장
        self.channel_checkboxes = {}
        
        # 상태 체크박스 저장
        self.status_checkboxes = {}
        
        # 상품 목록 (콤보박스에 표시할 항목들)
        self.product_list = []  # '전체' 항목 제거
        
        # 테이블 위젯 생성
        self.table = QTableWidget()
        
        # 상태 메시지 업데이트를 위한 타이머
        self.status_timer = QTimer()
        self.status_timer.setSingleShot(True)
        self.status_timer.timeout.connect(self.clear_status_after_delay)
        
        # 행별 지정상품 정보 저장 
        self.assigned_products = {}  # {row_id: 지정상품명}
        
        # 행별 지정채널 정보 저장
        self.assigned_channels = {}  # {row_id: 지정채널명}
        
        # 상태 저장 관련 변수
        self.last_save_path = ""
        self.auto_save_interval = 5  # 분 단위
        self.auto_save_timer = QTimer(self)
        self.auto_save_timer.timeout.connect(self.auto_save)
        self.auto_save_timer.start(30 * 1000)  # 30초를 밀리초로 변환
        self.is_state_modified = False  # 상태가 수정되었는지 여부
        
        # UI 초기화 먼저 수행 (tab_widget 생성)
        self.init_ui()
        
        # 매니저 클래스 초기화 (UI 초기화 후에)
        self.table_manager = TableManager(self)
        self.filter_manager = FilterManager(self)
        self.tab_manager = TabManager(self)
        
        # 채널 체크박스 이벤트 연결
        if hasattr(self, 'connect_channel_checkbox_events'):
            self.connect_channel_checkbox_events()
        
        # 탭 설정
        self.tab_manager.setup_tabs()
        
        # 탭 변경 이벤트 연결
        self.tab_widget.currentChanged.connect(self.on_tab_changed)
        
        self.organize_contacts_by_row()  # 연락처별 행 ID 저장
    
    def init_ui(self):
        """UI 초기화"""
        # 메인 위젯 및 레이아웃 설정
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)
        
        # 상단 필터 및 버튼 영역 - 3x2 레이아웃
        grid_layout = QGridLayout()
        
        # 1. 엑셀 파일 로드 버튼 (위쪽 왼쪽)
        self.load_btn = QPushButton("엑셀 파일 로드")
        self.load_btn.clicked.connect(self.load_excel)
        # 스타일 변경
        self.load_btn.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;  /* 녹색 배경 */
                color: white;  /* 흰색 글자 */
                font-weight: bold;  /* 굵은 글자 */
                border-radius: 5px;  /* 둥근 모서리 */
                padding: 10px;  /* 여백 */
            }
            QPushButton:hover {
                background-color: #45a049;  /* 호버 시 더 진한 녹색 */
            }
        """)
        grid_layout.addWidget(self.load_btn, 0, 0)
        
        # 2. 상태 저장 버튼 (위쪽 중앙)
        self.save_state_btn = QPushButton("상태 저장")
        self.save_state_btn.clicked.connect(self.save_work_state)
        grid_layout.addWidget(self.save_state_btn, 0, 1)
        
        # 3. 상태 불러오기 버튼 (위쪽 오른쪽)
        self.load_state_btn = QPushButton("상태 불러오기")
        self.load_state_btn.clicked.connect(self.load_work_state)
        grid_layout.addWidget(self.load_state_btn, 0, 2)
        
        # 4. 상품 검색 (아래쪽 왼쪽)
        search_group = QWidget()
        search_layout = QVBoxLayout(search_group)
        
        # 상품 검색 레이블
        search_label = QLabel("선정할 체험단을 선택하세요 :")
        search_layout.addWidget(search_label)
        
        # 상품 검색 콤보박스
        self.product_combo = QComboBox()
        self.product_combo.addItems(self.product_list)
        self.product_combo.currentIndexChanged.connect(self.apply_filters)
        search_layout.addWidget(self.product_combo)
        
        # 단일 상품 체크박스 (기존과 동일)
        self.single_product_checkbox = QCheckBox("단일 상품만 보기")
        self.single_product_checkbox.stateChanged.connect(self.apply_filters)
        search_layout.addWidget(self.single_product_checkbox)
        
        grid_layout.addWidget(search_group, 1, 0)
        
        # 5. 채널 필터 (아래쪽 중앙)
        channel_group = UIComponents.create_channel_filter_group(self, self.channel_list)
        grid_layout.addWidget(channel_group, 1, 1)
        
        # 6. 상태 필터 (아래쪽 오른쪽)
        status_filter_group = UIComponents.create_status_filter_group(self)
        grid_layout.addWidget(status_filter_group, 1, 2)
        
        # 상단 필터 및 버튼 레이아웃을 메인 레이아웃에 추가
        layout.addLayout(grid_layout)
        
        # 이름/연락처/URL 검색 (3x2 아래)
        contact_search_group = UIComponents.create_contact_search_group(self)
        layout.addWidget(contact_search_group)
        
        # 버튼 영역 (초기화 버튼)
        buttons_layout = QGridLayout()
        
        # 필터 초기화 버튼
        self.reset_btn = QPushButton("필터 초기화")
        self.reset_btn.clicked.connect(self.reset_filter)
        buttons_layout.addWidget(self.reset_btn, 0, 0)
        
        # 빈 칼럼 추가
        buttons_layout.addWidget(QWidget(), 0, 1)
        
        # URL로 보기 콤보박스 추가
        self.url_view_combo = QComboBox()
        self.url_view_combo.addItems(["전체", "미정", "대기", "선정", "제외"])
        buttons_layout.addWidget(self.url_view_combo, 0, 2)
        
        # URL로 보기 버튼 추가
        self.url_view_btn = QPushButton("URL띄우기")
        self.url_view_btn.clicked.connect(self.open_urls_in_table)
        buttons_layout.addWidget(self.url_view_btn, 0, 3)
        
        # 현재 화면 저장 버튼
        self.save_btn = QPushButton("현재화면 엑셀저장")
        self.save_btn.clicked.connect(self.save_current_view)
        self.save_btn.setStyleSheet("""
            QPushButton {
                background-color: #FF9800;  /* 주황색 배경 */
                color: white;  /* 흰색 글자 */
                font-weight: bold;  /* 굵은 글자 */
                border-radius: 5px;  /* 둥근 모서리 */
                padding: 10px;  /* 여백 */
            }
            QPushButton:hover {
                background-color: #FB8C00;  /* 호버 시 더 진한 주황색 */
            }
        """)
        buttons_layout.addWidget(self.save_btn, 0, 4)
        
        # 현재 화면 저장2 버튼 추가
        self.save_btn_2 = QPushButton("준비중")
        self.save_btn_2.setEnabled(False)  # 클릭 불가 설정
        self.save_btn_2.setStyleSheet("""
            QPushButton {
                background-color: #808080;  /* 회색 배경 */
                color: white;  /* 흰색 글자 */
                font-weight: bold;  /* 굵은 글자 */
                border-radius: 5px;  /* 둥근 모서리 */
                padding: 10px;  /* 여백 */
            }
            QPushButton:hover {
                background-color: #A9A9A9;  /* 호버 시 더 밝은 회색 */
            }
        """)
        buttons_layout.addWidget(self.save_btn_2, 0, 5)
        
        layout.addLayout(buttons_layout)
        
        # 상태 표시 영역
        status_layout = QHBoxLayout()
        
        # 현재 상태 메시지
        self.status_label = QLabel("프로그램이 준비되었습니다. 엑셀 파일을 로드해주세요.")
        status_layout.addWidget(self.status_label, 7)  # 비율 7 (더 넓게)
        
        # 마지막 자동 저장 시간 라벨
        self.last_auto_save_label = QLabel("마지막 자동 저장: 없음")
        status_layout.addWidget(self.last_auto_save_label, 3)  # 비율 3
        
        layout.addLayout(status_layout)
        
        # 통계 및 저장 영역
        stats_save_layout = QHBoxLayout()
        
        # 상태 통계 영역
        self.stats_label = QLabel("상태 통계 ▶ ")
        stats_save_layout.addWidget(self.stats_label, 5)  # 비율 5
        
        layout.addLayout(stats_save_layout)
        
        # 탭 위젯 생성
        self.tab_widget = QTabWidget()
        
        # 탭 스타일 설정
        self.tab_widget.setStyleSheet("""
            QTabWidget::pane { 
                border: 1px solid #C2C7CB;
                background: white;
            }
            
            QTabWidget::tab-bar {
                left: 5px;
            }
            
            QTabBar::tab {
                background: #F0F0F0;
                border: 1px solid #C4C4C3;
                border-bottom-color: #C2C7CB;
                border-top-left-radius: 4px;
                border-top-right-radius: 4px;
                min-width: 8ex;
                padding: 8px 12px;
                margin-right: 2px;
                color: #444444;
            }
            
            QTabBar::tab:selected {
                background: #4A86E8;
                color: white;
                font-weight: bold;
                border: 1px solid #3A76D8;
            }
            
            QTabBar::tab:!selected {
                margin-top: 2px;
            }
            
            QTabBar::tab:hover {
                background: #E0E0E0;
            }
            
            QTabBar::tab:selected:hover {
                background: #5A96F8;
            }
        """)
        
        layout.addWidget(self.tab_widget)
    
    # 필터 관련 메서드들 (FilterManager로 위임)
    def apply_filters(self):
        """모든 필터를 적용하여 테이블 업데이트"""
        self.filter_manager.apply_filters()
    
    def reset_filter(self):
        """모든 필터 초기화"""
        self.filter_manager.reset_filter()
    
    # 엑셀 로드 관련 메서드
    def load_excel(self):
        """엑셀 파일 로드 및 테이블에 표시"""
        file_path, _ = QFileDialog.getOpenFileName(self, "엑셀 파일 선택", "", "Excel Files (*.xlsx *.xls)")
        if not file_path:
            return
        
        try:
        # 엑셀 파일 로드
            self.excel_file_path = file_path
            
            # ExcelHandler 클래스를 통해 엑셀 로드
            result = ExcelHandler.load_excel_file(
                file_path=file_path, 
                parent=self, 
                header_mapping=self.header_mapping
            )
            
            # 반환 값 타입에 따른 처리
            if isinstance(result, pd.DataFrame):
                self.original_df = result
            elif isinstance(result, dict):
                # 딕셔너리 키 확인 및 처리
                keys = list(result.keys())
                if 'dataframe' in keys:
                    self.original_df = result['dataframe']
                elif 'df' in keys:
                    self.original_df = result['df']
                elif len(keys) > 0 and isinstance(result[keys[0]], pd.DataFrame):
                    # 첫 번째 키에 DataFrame이 있는 경우
                    self.original_df = result[keys[0]]
                else:
                    # 디버깅을 위해 키 목록 출력
                    key_list = ', '.join(keys)
                    self.status_label.setText(f"엑셀 로드 오류: 딕셔너리에 DataFrame 없음. 키 목록: {key_list}")
                    return
            else:
                self.status_label.setText(f"엑셀 로드 오류: 반환된 데이터 형식 ({type(result).__name__})")
                return
            
            if self.original_df is None or len(self.original_df) == 0:
                self.status_label.setText("엑셀 파일이 비어있거나 로드할 수 없습니다.")
                return
            
            # 필요한 인덱스 찾기 (연락처, 이름, 상품, URL 등)
            self.find_important_indices()
            
            # 연락처별 행 ID 저장
            self.organize_contacts_by_row()
            
            # 선택된 열만 보여주기
            columns_to_show = []
            for col in self.original_df.columns[2:14]:  # C(인덱스 2)부터 N(인덱스 13)까지
                col_idx = self.original_df.columns.get_loc(col)
                # K열(인덱스 10)과 M열(인덱스 12)는 제외
                if col_idx != 10 and col_idx != 12:
                    columns_to_show.append(col)
            
            self.filtered_df = self.original_df[columns_to_show]
            
            # 테이블 업데이트
            self.table_manager.update_table(self.filtered_df)
            
            # 상품 목록 추출 및 콤보박스 업데이트 부분
            if self.product_column_idx >= 0:
                # 모든 상품 추출
                product_column = self.original_df.columns[self.product_column_idx]
                products = []
                
                # 각 행에서 상품 추출
                for item in self.original_df[product_column].dropna():
                    # 여러 상품이 포함된 경우 분리
                    if ',' in item:
                        products.extend([p.strip() for p in item.split(',')])
                    else:
                        products.append(item.strip())
                
                # 중복 제거 및 정렬
                self.product_list = sorted(list(set(products)))
                
                # 콤보박스 업데이트
                self.product_combo.clear()
                for product in self.product_list:
                    self.product_combo.addItem(product)
                
                # 상품 목록을 기반으로 탭 업데이트
                self.update_tabs_from_products()
            
            self.status_label.setText(f"'{os.path.basename(file_path)}' 파일을 불러왔습니다.")
            
        except Exception as e:
            self.status_label.setText(f"엑셀 로드 중 오류: {str(e)}")
    
    def find_important_indices(self):
        """중요 컬럼 인덱스 찾기"""
        for i, col in enumerate(self.original_df.columns):
            col_lower = str(col).lower()
            
            # 상품 컬럼 찾기
            if "희망상품" in col or "희망 상품" in col:
                self.product_column_idx = i
            
            # 연락처 컬럼 찾기
            elif "연락처" in col:
                self.contact_column_idx = i
            
            # 이름 컬럼 찾기
            elif "성함" in col or "이름" in col or "닉네임" in col:
                self.name_column_idx = i
            
            # URL 컬럼 찾기
            elif "url" in col_lower or "계정 링크" in col or "블로그" in col:
                self.url_column_idx = i
    
    def organize_contacts_by_row(self):
        """연락처별 행 ID 저장"""
        if self.contact_column_idx == -1:
            return
        
        self.contact_rows = {}
        
        for row_id, row in self.original_df.iterrows():
            contact = row.iloc[self.contact_column_idx]
            if pd.notna(contact):
                contact = str(contact).strip()
                if contact not in self.contact_rows:
                    self.contact_rows[contact] = []
                self.contact_rows[contact].append(row_id)
    
    # 저장 및 불러오기 관련 메서드
    def save_current_view(self):
        """현재 선택된 탭에 표시된 데이터를 엑셀 파일로 저장"""
        # 현재 선택된 탭 인덱스 가져오기
        current_tab_index = self.tab_widget.currentIndex()
        
        # 현재 탭이 선택되지 않았거나 데이터가 없는 경우
        if current_tab_index < 0 or self.filtered_df is None or len(self.filtered_df) == 0:
            QMessageBox.warning(self, "저장 오류", "저장할 데이터가 없습니다.")
            return
        
        # 현재 선택된 탭의 이름 가져오기
        tab_name = self.tab_widget.tabText(current_tab_index)
        
        # '+' 탭인 경우 저장 중단
        if tab_name == "+":
            QMessageBox.warning(self, "저장 오류", "해당 탭에는 저장할 데이터가 없습니다.")
            return
        
        # 현재 탭의 위젯과 테이블 찾기
        current_tab = self.tab_widget.widget(current_tab_index)
        current_table = None
        
        for child in current_tab.children():
            if isinstance(child, QTableWidget):
                current_table = child
                break
        
        if current_table is None:
            QMessageBox.warning(self, "저장 오류", "현재 탭에서 테이블을 찾을 수 없습니다.")
            return
        
        # 파일 저장 대화상자 표시 (기본 파일명에 탭 이름 포함)
        file_path, _ = QFileDialog.getSaveFileName(
            self, "현재 화면 저장", f"{tab_name}_데이터.xlsx", "Excel Files (*.xlsx)")
        
        if not file_path:
            return  # 사용자가 취소함
        
        # 확장자 확인 및 추가
        if not file_path.endswith('.xlsx'):
            file_path += '.xlsx'
        
        try:
            # 현재 탭의 테이블 데이터를 데이터프레임으로 변환
            df_to_save = self.table_to_dataframe(current_table)
            
            # 데이터프레임이 비어있는 경우
            if df_to_save is None or df_to_save.empty:
                QMessageBox.warning(self, "저장 오류", "저장할 데이터가 없습니다.")
                return
            
            # 엑셀 파일로 저장
            df_to_save.to_excel(file_path, index=False)
            
            self.status_label.setText(f"현재 화면이 '{file_path}'에 저장되었습니다.")
        except Exception as e:
            QMessageBox.critical(self, "저장 오류", f"파일 저장 중 오류가 발생했습니다: {str(e)}")

    def table_to_dataframe(self, table):
        """테이블 위젯의 데이터를 데이터프레임으로 변환"""
        # 행과 열 수 가져오기
        rows = table.rowCount()
        cols = table.columnCount()
        
        if rows == 0 or cols == 0:
            return None
        
        # 헤더 가져오기
        headers = []
        for col in range(cols):
            header_item = table.horizontalHeaderItem(col)
            if header_item:
                headers.append(header_item.text())
            else:
                headers.append(f"Column {col}")
        
        # 데이터 수집
        data = []
        for row in range(rows):
            row_data = []
            for col in range(cols):
                # 상태 버튼 열은 건너뛰기 (첫 번째 열)
                if col == 0:
                    # 상태 버튼에서 상태값 가져오기
                    widget = table.cellWidget(row, col)
                    if hasattr(widget, 'get_status'):
                        status_value = widget.get_status()
                        status_text = ["미정", "선정", "대기", "제외", "완료"][status_value]
                        row_data.append(status_text)
                    else:
                        row_data.append("")
                    continue
                
                # 일반 셀은 텍스트 가져오기
                item = table.item(row, col)
                if item:
                    row_data.append(item.text())
                else:
                    row_data.append("")
            
            data.append(row_data)
        
        # 데이터프레임 생성
        return pd.DataFrame(data, columns=headers)

    def save_work_state(self, auto_save=False):
        """현재 작업 상태를 JSON 파일로 저장"""
        # 수정된 내용이 없으면 저장 안함 (자동 저장인 경우)
        if auto_save and not self.is_state_modified:
            return
        
        # 이전에 저장된 경로가 있으면 그 경로를 기본 경로로 설정
        default_path = self.last_save_path if self.last_save_path else ""
        
        # 자동 저장이 아닌 경우에만 파일 대화상자 표시
        if not auto_save:
            file_path, _ = QFileDialog.getSaveFileName(
                self, "상태 저장", default_path, "JSON Files (*.json)")
            if not file_path:
                return
        else:
            # 자동 저장인 경우 마지막 저장 경로 사용
            if not self.last_save_path:
                # 자동 저장 경로가 없는 경우 현재 디렉토리에 자동 저장
                file_path = "autosave_state.json"
            else:
                file_path = self.last_save_path
        
        # 확장자 확인 및 추가
        if not file_path.endswith('.json'):
            file_path += '.json'
        
        # 백업 파일 생성 (이전 파일이 있는 경우)
        if os.path.exists(file_path):
            # 타임스탬프 추가
            timestamp = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
            
            # 백업 폴더 경로 설정
            backup_dir = os.path.join(os.path.dirname(file_path), "bak")
            
            # 백업 폴더가 없으면 생성
            if not os.path.exists(backup_dir):
                os.makedirs(backup_dir)
            
            # 백업 파일 경로 설정
            backup_path = os.path.join(backup_dir, f"{os.path.basename(file_path)}.{timestamp}.bak")
            
            try:
                import shutil
                shutil.copy2(file_path, backup_path)
            except Exception:
                pass  # 백업 실패해도 계속 진행
        
        # 저장할 데이터 구성
        state_data = {
            'row_status': self.row_status,
            'assigned_products': self.assigned_products,
            'assigned_channels': self.assigned_channels,  # 지정채널 정보 추가
            'version': '1.1'  # 버전 정보 추가
        }
        
        try:
            # 정수 키를 문자열로 변환 (JSON은 키로 문자열만 허용)
            row_status_str = {str(k): v for k, v in self.row_status.items()}
            assigned_products_str = {str(k): v for k, v in self.assigned_products.items()}
            assigned_channels_str = {str(k): v for k, v in self.assigned_channels.items()}
            
            state_data['row_status'] = row_status_str
            state_data['assigned_products'] = assigned_products_str
            state_data['assigned_channels'] = assigned_channels_str
            
            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(state_data, f, ensure_ascii=False, indent=2)
            
            # 저장 경로 기억
            self.last_save_path = file_path
            
            # 상태 변경 플래그 초기화
            self.is_state_modified = False
            
            if not auto_save:
                self.status_label.setText(f"작업 상태가 '{file_path}'에 저장되었습니다.")
            else:
                self.status_label.setText(f"작업 상태가 자동 저장되었습니다: {file_path}")
        except Exception as e:
            if not auto_save:
                QMessageBox.critical(self, "저장 오류", f"상태 저장 중 오류가 발생했습니다: {str(e)}")

    def load_work_state(self):
        """저장된 작업 상태를 JSON 파일에서 불러오기"""
        # 이전에 저장된 경로가 있으면 그 경로를 기본 경로로 설정
        default_path = self.last_save_path if self.last_save_path else ""
        
        file_path, _ = QFileDialog.getOpenFileName(
            self, "상태 불러오기", default_path, "JSON Files (*.json)")
        if not file_path:
            return
        
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                state_data = json.load(f)
            
            # 버전 확인
            version = state_data.get('version', '1.0')
            
            # 데이터 불러오기
            row_status_str = state_data.get('row_status', {})
            assigned_products_str = state_data.get('assigned_products', {})
            assigned_channels_str = state_data.get('assigned_channels', {})
            
            # 문자열 키를 정수로 변환
            self.row_status = {int(k): v for k, v in row_status_str.items()}
            self.assigned_products = {int(k): v for k, v in assigned_products_str.items()}
            self.assigned_channels = {int(k): v for k, v in assigned_channels_str.items()}
            
            # 테이블 업데이트
            if self.filtered_df is not None:
                self.table_manager.update_table(self.filtered_df)
                
                # 각 탭의 테이블도 업데이트
                self.update_all_tabs()
        
            # 상태 통계 업데이트
            self.update_status_statistics()

            # 저장 경로 기억
            self.last_save_path = file_path
            
            # 상태 변경 플래그 초기화
            self.is_state_modified = False
            
            self.status_label.setText(f"작업 상태가 '{file_path}'에서 불러와졌습니다.")
        except Exception as e:
            QMessageBox.critical(self, "불러오기 오류", f"상태 불러오기 중 오류가 발생했습니다: {str(e)}")

    def auto_save(self):
        """자동 저장 실행"""
        # 등록된 상태가 있는 경우에만 자동 저장
        if self.row_status and self.is_state_modified:
            # 불러온 엑셀 파일명에 "_중간저장" 추가
            if self.excel_file_path:
                base_name = os.path.splitext(os.path.basename(self.excel_file_path))[0]
                auto_save_path = os.path.join(os.path.dirname(self.excel_file_path), f"{base_name}_중간저장.json")
            else:
                auto_save_path = "autosave_state.json"
            
            # 자동 저장 경로 설정
            self.last_save_path = auto_save_path
            self.save_work_state(auto_save=True)
            
            # 자동 저장 시간 업데이트
            current_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            self.last_auto_save_label.setText(f"마지막 자동 저장: {current_time}")

    def update_all_tabs(self):
        """모든 탭의 테이블 업데이트"""
        for i in range(self.tab_widget.count()):
            tab_text = self.tab_widget.tabText(i)
            if tab_text != "+" and tab_text != "데이터":
                tab = self.tab_widget.widget(i)
                for child in tab.children():
                    if isinstance(child, QTableWidget):
                        self.update_tab_table(child, tab_text)
                        break

    def update_tabs_from_products(self):
        """상품 목록을 기반으로 탭 업데이트"""
        # 기존 데이터 탭 제외하고 추가 탭 제거
        while self.tab_widget.count() > 1:
            self.tab_widget.removeTab(1)
        
        # "+" 탭 임시 저장
        add_tab_button = self.add_tab_button
        
        # 상품별 탭 추가
        for i, product in enumerate(self.product_list[:5]):  # 상위 5개만 탭으로 추가
            new_tab = QWidget()
            new_tab_layout = QVBoxLayout(new_tab)
            
            # 테이블 위젯 생성 및 추가
            new_tab_table = self.tab_manager.create_table_widget()
            new_tab_layout.addWidget(new_tab_table)
            
            # 테이블 데이터 표시 (필터링된 데이터)
            if self.filtered_df is not None:
                self.update_tab_table(new_tab_table, product)
            
            self.tab_widget.addTab(new_tab, product)
            
            # 탭 객체 저장
            setattr(self, f"tab_{i+2}", new_tab)
        
        # "+" 탭 다시 추가
        self.tab_widget.addTab(add_tab_button, "+")
        
        # 탭 이름을 콤보박스에 추가
        self.tab_manager.update_tab_combo()

    def update_tab_table(self, table, product_name):
        """탭의 테이블 데이터 업데이트"""
        # '데이터' 탭이 아닌 경우에만 추가 필터링 적용
        if product_name != "데이터" and self.filtered_df is not None and self.product_column_idx >= 0:
            # 1. 원본 데이터로부터 시작
            df = self.original_df.copy()
            
            # 2. 검색 필터 적용 (이름/연락처/URL 검색)
            contact_search_text = self.contact_search_input.text().strip().lower()
            if contact_search_text:
                # 이름, 연락처, URL 칼럼에서 검색어 포함 여부 확인
                contact_mask = False
                
                # 연락처 칼럼이 있는 경우
                if self.contact_column_idx >= 0:
                    contact_mask = contact_mask | df.iloc[:, self.contact_column_idx].astype(str).str.contains(
                        contact_search_text, case=False, na=False)
                
                # 이름 칼럼이 있는 경우
                if self.name_column_idx >= 0:
                    contact_mask = contact_mask | df.iloc[:, self.name_column_idx].astype(str).str.contains(
                        contact_search_text, case=False, na=False)
                
                # URL 칼럼이 있는 경우
                if self.url_column_idx >= 0:
                    contact_mask = contact_mask | df.iloc[:, self.url_column_idx].astype(str).str.contains(
                        contact_search_text, case=False, na=False)
                
                df = df[contact_mask]
            
            # 3. 상품명 필터 적용 (특정 탭의 상품만 표시)
            product_col_name = self.original_df.columns[self.product_column_idx]
            product_mask = df[product_col_name].str.contains(
                product_name, case=False, na=False, regex=False)
            df = df[product_mask]
            
            # 4. 상태 필터 (선정 상태만)
            status_filtered_rows = []
            for idx in df.index:
                status = self.row_status.get(idx, 0)
                # 상태가 선정(1)이고 지정상품 셀이 탭의 이름과 같은 경우
                if status == 1 and self.assigned_products.get(idx) == product_name:
                    status_filtered_rows.append(idx)
            
            if status_filtered_rows:
                tab_filtered_df = self.original_df.loc[status_filtered_rows, self.filtered_df.columns]
            else:
                # 일치하는 데이터가 없으면 빈 데이터프레임 생성
                tab_filtered_df = self.filtered_df.head(0)  # 빈 데이터프레임이지만 열 구조 유지
            
            # 테이블 위젯 업데이트
            self.table_manager.update_table_widget(table, tab_filtered_df)
        else:
            # 데이터 탭은 기존 필터링된 데이터 표시 (모든 필터 적용)
            self.table_manager.update_table_widget(table, self.filtered_df)

    def on_tab_changed(self, index):
        """탭 변경 시 호출되는 메서드"""
        # 필터링된 데이터가 없으면 리턴
        if self.filtered_df is None:
            return
        
        # 현재 탭 이름 가져오기
        tab_text = self.tab_widget.tabText(index)
        
        # + 탭은 무시
        if tab_text == "+":
            return
        
        # 채널 필터 활성화/비활성화 (데이터 탭에서만 활성화)
        self.toggle_channel_filter_controls(tab_text == "데이터")
        
        # 현재 탭의 위젯 가져오기
        current_tab = self.tab_widget.widget(index)
        
        # 탭에서 테이블 위젯 찾기
        for child in current_tab.children():
            if isinstance(child, QTableWidget):
                # 해당 탭의 테이블 업데이트
                self.update_tab_table(child, tab_text)
                
                # 해당 탭의 상태 통계 업데이트
                if tab_text == "데이터":
                    # 데이터 탭은 전체 통계 표시
                    self.update_status_statistics()
                else:
                    # 다른 탭은 해당 탭의 필터링된 데이터에 대한 통계 표시
                    self.update_tab_statistics(tab_text)
                break

    def toggle_channel_filter_controls(self, enabled):
        """채널 필터 컨트롤 활성화/비활성화"""
        # 모든 채널 체크박스 활성화/비활성화
        for checkbox in self.channel_checkboxes.values():
            checkbox.setEnabled(enabled)

    def update_tab_statistics(self, tab_name):
        """특정 탭의 상태 통계 업데이트"""
        if not self.row_status or self.filtered_df is None or self.product_column_idx < 0:
            self.stats_label.setText(f"탭 '{tab_name}' 통계 ▶ 데이터 없음")
            return
        
        # 현재 선택된 탭의 테이블 가져오기
        current_tab_index = self.tab_widget.currentIndex()
        current_tab = self.tab_widget.widget(current_tab_index)
        current_table = None
        
        for child in current_tab.children():
            if isinstance(child, QTableWidget):
                current_table = child
                break
        
        if current_table is None or current_table.rowCount() == 0:
            self.stats_label.setText(f"탭 '{tab_name}' 통계 ▶ 데이터 없음")
            return
        
        # 상태별 카운트
        status_count = {0: 0, 1: 0, 2: 0, 3: 0, 4: 0}
        
        # 지정채널별 카운트 
        channel_count = {}
        
        # 현재 화면에 표시된 데이터에 대한 상태 및 채널 카운트
        for row in range(current_table.rowCount()):
            # 상태 버튼에서 상태 가져오기
            status_btn = current_table.cellWidget(row, 0)
            if isinstance(status_btn, StatusButton):
                current_status = status_btn.get_status()
                status_count[current_status] += 1
                
                # 지정채널 카운트
                channel_item = current_table.item(row, 2)  # 지정채널 열 인덱스
                if channel_item:
                    channel_name = channel_item.text()
                    channel_count[channel_name] = channel_count.get(channel_name, 0) + 1
        
        # 통계 텍스트 생성
        stats_text = f"탭 '{tab_name}' 통계 ▶ "
        status_texts = []
        status_names = {0: "미정", 1: "선정", 2: "대기", 3: "제외", 4: "완료"}
        
        for status, count in status_count.items():
            if count > 0:
                status_texts.append(f"{status_names[status]}: {count}명")
        
        # 현재 화면의 총 인원 추가
        visible_rows = current_table.rowCount()
        status_texts.append(f"현재 화면: {visible_rows}명")
        
        stats_text += ", ".join(status_texts)
        
        # 지정채널 통계 추가
        if channel_count:
            channel_texts = []
            for channel, count in channel_count.items():
                channel_texts.append(f"{channel}: {count}명")
            
            stats_text += " ▶ 지정채널: " + ", ".join(channel_texts)
        
        self.stats_label.setText(stats_text)

    def toggle_all_channel_checkboxes(self, state):
        """모든 채널 체크박스 상태 변경"""
        # Qt.Checked 대신 정수로 비교 (2 = 체크됨)
        checked = state == 2
        
        # 이벤트 연쇄 방지를 위해 시그널 차단
        for checkbox in self.channel_checkboxes.values():
            # 체크박스 시그널 차단
            checkbox.blockSignals(True)
            # 체크박스 상태 변경
            checkbox.setChecked(checked)
            # 시그널 차단 해제
            checkbox.blockSignals(False)
        
        # 필터 적용
        self.apply_filters()

    def update_select_all_checkbox_state(self):
        """개별 체크박스 상태에 따라 전체선택 체크박스 상태 업데이트"""
        # 모든 체크박스가 선택되었는지 확인
        all_checked = all(checkbox.isChecked() for checkbox in self.channel_checkboxes.values())
        
        # 전체선택 체크박스 시그널 차단
        self.select_all_checkbox.blockSignals(True)
        # 상태 설정
        self.select_all_checkbox.setChecked(all_checked)
        # 시그널 차단 해제
        self.select_all_checkbox.blockSignals(False)

    def get_selected_channel(self):
        """채널 필터에서 선택된 채널을 반환 (2개 이상 선택된 경우 메시지 표시)"""
        selected_channels = []
        
        for channel, checkbox in self.channel_checkboxes.items():
            if checkbox.isChecked():
                selected_channels.append(channel)
        
        # 선택된 채널이 2개 이상인 경우 메시지 표시
        if len(selected_channels) > 1:
            QMessageBox.warning(self, "채널 선택 경고", 
                               "지정채널에 입력하기 위해 채널 필터를 하나만 선택해주세요.")
            return None
        
        # 선택된 채널이 없는 경우
        if not selected_channels:
            return None
        
        # 선택된 채널이 하나인 경우
        return selected_channels[0]

    def on_channel_checkbox_clicked(self, clicked_channel, checked):
        """채널 체크박스 클릭 이벤트 핸들러"""
        # 체크된 경우에만 다른 체크박스 해제
        if checked:
            # 다른 모든 체크박스 해제
            for channel, checkbox in self.channel_checkboxes.items():
                if channel != clicked_channel:
                    checkbox.blockSignals(True)  # 시그널 차단
                    checkbox.setChecked(False)   # 체크 해제
                    checkbox.blockSignals(False) # 시그널 차단 해제
        
        # 필터 적용
        self.apply_filters() 

    def clear_status_after_delay(self):
        """일정 시간 후 상태 메시지 지우기"""
        # 자동 저장 메시지나 임시 상태 메시지를 삭제하거나 기본 메시지로 교체
        current_text = self.status_label.text()
        
        # 자동 저장 메시지인 경우
        if "자동 저장" in current_text:
            self.status_label.setText("작업 중입니다.")
        # 일반적인 작업 완료 메시지인 경우
        elif "저장되었습니다" in current_text or "불러와졌습니다" in current_text:
            self.status_label.setText("작업 중입니다.")

    def update_status_statistics(self):
        """상태 통계 업데이트"""
        if not self.row_status:
            self.stats_label.setText("상태 통계 ▶ 데이터 없음")
            return
        
        # 상태별 카운트
        status_count = {0: 0, 1: 0, 2: 0, 3: 0, 4: 0}
        for status in self.row_status.values():
            status_count[status] = status_count.get(status, 0) + 1
        
        # 지정채널별 카운트 (선정 상태인 것만)
        channel_count = {}
        for row_id, channel in self.assigned_channels.items():
            if self.row_status.get(row_id, 0) == 1:  # 선정된 것만 카운트
                channel_count[channel] = channel_count.get(channel, 0) + 1
        
        # 통계 텍스트 생성
        stats_text = "상태 통계 ▶ "
        status_texts = []
        status_names = {0: "미정", 1: "선정", 2: "대기", 3: "제외", 4: "완료"}
        
        for status, count in status_count.items():
            if count > 0:
                status_texts.append(f"{status_names[status]}: {count}명")
        
        # 현재 필터링된 데이터의 행 수 추가
        visible_rows = 0
        if self.filtered_df is not None:
            visible_rows = len(self.filtered_df)
        status_texts.append(f"현재 화면: {visible_rows}명")
        
        stats_text += ", ".join(status_texts)
        
        # 지정채널 통계 추가
        if channel_count:
            channel_texts = []
            for channel, count in channel_count.items():
                channel_texts.append(f"{channel}: {count}명")
            
            stats_text += " | 지정채널: " + ", ".join(channel_texts)
        
        self.stats_label.setText(stats_text)

    def open_urls_in_table(self):
        """현재 테이블에 표시된 URL을 선택된 상태에 따라 열기"""
        if self.filtered_df is None or self.filtered_df.empty:
            QMessageBox.warning(self, "URL 열기 오류", "표시된 데이터가 없습니다.")
            return
        
        # 현재 선택된 탭의 테이블 가져오기
        current_tab_index = self.tab_widget.currentIndex()
        current_tab = self.tab_widget.widget(current_tab_index)
        current_table = None
        
        for child in current_tab.children():
            if isinstance(child, QTableWidget):
                current_table = child
                break
        
        if current_table is None or current_table.rowCount() == 0:
            QMessageBox.warning(self, "URL 열기 오류", "선택한 탭에 데이터가 없습니다.")
            return
        
        # URL 열 인덱스 확인 (테이블 내에서의 인덱스)
        url_table_idx = -1
        for i, col in enumerate(self.filtered_df.columns):
            col_str = str(col).lower()
            if "url" in col_str or "계정 링크" in col or "블로그" in col:
                url_table_idx = i + 3  # +3은 상태, 지정상품, 지정채널 칼럼 때문
                break
        
        if url_table_idx == -1:
            QMessageBox.warning(self, "URL 열기 오류", "URL 열을 찾을 수 없습니다.")
            return
        
        # 선택된 상태 가져오기
        selected_status_text = self.url_view_combo.currentText()
        status_mapping = {"미정": 0, "대기": 2, "선정": 1, "제외": 3}
        selected_status = status_mapping.get(selected_status_text, None)
        
        # 테이블의 모든 URL 수집
        urls = []
        for row in range(current_table.rowCount()):
            item = current_table.item(row, url_table_idx)
            if item:
                url_text = item.text().strip()
                if url_text:
                    # URL 형식 확인 및 수정
                    url = url_text
                    if not url.startswith(('http://', 'https://')):
                        url = 'https://' + url
                    
                    # 상태에 따라 URL 추가
                    status_btn = current_table.cellWidget(row, 0)  # 상태 버튼은 cellWidget으로 가져옴
                    if isinstance(status_btn, StatusButton):
                        current_status = status_btn.get_status()
                        if selected_status_text == "전체" or current_status == selected_status:
                            urls.append(url)
        
        if not urls:
            QMessageBox.warning(self, "URL 열기 오류", "선택한 탭에 해당 상태의 데이터가 없습니다.")
            return
        
        # 사용자 확인 팝업
        reply = QMessageBox.question(self, 'URL 열기 확인', 
                                     f"{len(urls)}개의 URL을 열겠습니까?", 
                                     QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        
        if reply == QMessageBox.Yes:
            # URL을 20개씩 열기
            self.open_urls_in_batches(urls, batch_size=20, delay=10000)  # 10초 = 10000ms

    def open_urls_in_batches(self, urls, batch_size, delay):
        """URL을 배치 단위로 열기"""
        if not urls:
            self.status_label.setText("모든 URL이 열렸습니다.")
            return
        
        # 현재 배치의 URL 열기
        for url in urls[:batch_size]:
            try:
                webbrowser.open(url)
            except Exception as e:
                self.status_label.setText(f"URL을 열 수 없습니다: {str(e)}")
        
        # 진행 상황 업데이트
        opened_count = len(urls[:batch_size])
        total_count = len(urls)
        self.status_label.setText(f"{opened_count}/{total_count} URL 열림")
        
        # 남은 URL이 있으면 타이머로 다음 배치 예약
        if len(urls) > batch_size:
            QTimer.singleShot(delay, lambda: self.open_urls_in_batches(urls[batch_size:], batch_size, delay))
        else:
            self.status_label.setText("모든 URL이 열렸습니다.")

    def save_current_view_2(self):
        """현재 화면을 엑셀로 저장하는 기능"""
        # save_current_view와 동일한 로직을 사용
        self.save_current_view() 