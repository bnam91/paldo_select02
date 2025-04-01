from PyQt5.QtWidgets import (QVBoxLayout, QHBoxLayout, QGroupBox, QLabel, 
                           QLineEdit, QPushButton, QCheckBox, QFrame, QGridLayout)

class UIComponents:
    @staticmethod
    def create_search_filter_group(owner):
        """검색 필터 그룹 생성"""
        search_group = QGroupBox("상품 검색")
        search_inner_layout = QVBoxLayout()
        
        # 검색어 입력 필드
        search_inner_layout.addWidget(QLabel("검색어:"))
        
        owner.search_input = QLineEdit()
        owner.search_input.setPlaceholderText("상품명 검색 (예: 뽀로로)")
        search_inner_layout.addWidget(owner.search_input)
        
        # 단일 상품 체크박스
        owner.single_product_checkbox = QCheckBox("단일 상품만 표시")
        search_inner_layout.addWidget(owner.single_product_checkbox)
        
        # 검색 버튼
        owner.search_btn = QPushButton("검색")
        owner.search_btn.clicked.connect(owner.apply_filters)
        search_inner_layout.addWidget(owner.search_btn)
        
        search_group.setLayout(search_inner_layout)
        return search_group
    
    @staticmethod
    def create_contact_search_group(owner):
        """이름/연락처/URL 검색 그룹 생성"""
        contact_search_group = QGroupBox("이름/연락처/URL 검색")
        contact_search_layout = QHBoxLayout()  # 가로 레이아웃으로 변경
        
        # 이름/연락처 검색 입력 필드
        owner.contact_search_label = QLabel("검색어:")
        contact_search_layout.addWidget(owner.contact_search_label, 1)  # 비율 1
        
        owner.contact_search_input = QLineEdit()
        owner.contact_search_input.setPlaceholderText("이름, 연락처, URL 검색")
        contact_search_layout.addWidget(owner.contact_search_input, 3)  # 비율 3 (가장 넓게)
        
        # 검색 버튼
        owner.contact_search_btn = QPushButton("검색")
        owner.contact_search_btn.clicked.connect(owner.apply_filters)
        contact_search_layout.addWidget(owner.contact_search_btn, 1)  # 비율 1
        
        contact_search_group.setLayout(contact_search_layout)
        return contact_search_group
    
    @staticmethod
    def create_status_filter_group(owner):
        """상태 필터 그룹 생성"""
        status_filter_group = QGroupBox("상태 필터")
        status_filter_layout = QVBoxLayout()
        
        # 상태 필터 레이블
        status_label = QLabel("상태별 표시:")
        status_filter_layout.addWidget(status_label)
        
        # 가로 레이아웃으로 체크박스들 배치
        checkboxes_layout = QHBoxLayout()
        
        # 상태별 체크박스 추가
        owner.status_checkboxes = {}
        status_names = {
            0: "미정",
            1: "선정",
            2: "대기", 
            3: "제외",
            4: "완료"
        }
        
        # 각 상태별 체크박스 생성 및 추가
        for status_code, status_name in status_names.items():
            checkbox = QCheckBox(status_name)
            checkbox.setChecked(True)  # 기본적으로 모든 상태 선택
            checkbox.clicked.connect(owner.apply_filters)  # 체크 상태 변경 시 필터 적용
            owner.status_checkboxes[status_code] = checkbox
            checkboxes_layout.addWidget(checkbox)
        
        # 가로 레이아웃을 메인 레이아웃에 추가
        status_filter_layout.addLayout(checkboxes_layout)
        
        status_filter_group.setLayout(status_filter_layout)
        return status_filter_group
    
    @staticmethod
    def create_channel_filter_group(parent, channel_list):
        """채널 필터 그룹 생성"""
        channel_group = QGroupBox("채널 필터")
        
        # 메인 레이아웃은 세로 배치 유지
        main_layout = QVBoxLayout(channel_group)
        
        # 안내 레이블 추가
        guide_label = QLabel("채널을 선택하세요 (하나만 선택 가능):")
        main_layout.addWidget(guide_label)
        
        # 구분선 추가
        line = QFrame()
        line.setFrameShape(QFrame.HLine)
        line.setFrameShadow(QFrame.Sunken)
        main_layout.addWidget(line)
        
        # 채널 체크박스를 그리드로 배치할 레이아웃 생성
        channels_grid = QGridLayout()
        channels_grid.setSpacing(10)  # 간격 설정
        
        # '인스타' 항목이 없으면 추가
        if "인스타" not in channel_list:
            channel_list = list(channel_list) + ["인스타"]
        
        # 한 줄에 표시할 체크박스 수
        items_per_row = 3
        
        # 채널별 체크박스 생성 및 그리드에 배치
        for i, channel in enumerate(channel_list):
            checkbox = QCheckBox(channel)
            checkbox.setChecked(True)  # 기본값: 선택
            
            # 그리드에 추가 (행, 열 계산)
            row = i // items_per_row
            col = i % items_per_row
            channels_grid.addWidget(checkbox, row, col)
            
            # 체크박스 객체 저장
            parent.channel_checkboxes[channel] = checkbox
        
        # 그리드 레이아웃을 메인 레이아웃에 추가
        main_layout.addLayout(channels_grid)
        
        # UI 초기화 완료 후 이벤트 연결을 위한 메서드 추가
        parent.connect_channel_checkbox_events = lambda: UIComponents._connect_channel_events(parent)
        
        return channel_group

    @staticmethod
    def _connect_channel_events(parent):
        """채널 체크박스 이벤트 연결 (UI 초기화 후 호출)"""
        # 모든 채널 체크박스에 이벤트 연결
        for channel, checkbox in parent.channel_checkboxes.items():
            # 체크박스에 클릭 이벤트 연결
            checkbox.clicked.connect(
                lambda state, ch=channel: parent.on_channel_checkbox_clicked(ch, state)) 