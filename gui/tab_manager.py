import webbrowser
from PyQt5.QtWidgets import QWidget, QInputDialog, QVBoxLayout

class TabManager:
    """탭 관련 기능을 관리하는 클래스"""
    
    def __init__(self, parent):
        """
        초기화
        
        Args:
            parent: ExcelViewer 클래스의 인스턴스
        """
        self.parent = parent
        self.tab_widget = parent.tab_widget
    
    def setup_tabs(self):
        """탭 초기 설정"""
        # 탭 더블클릭 이벤트 연결
        self.tab_widget.tabBar().tabBarDoubleClicked.connect(self.rename_tab)
        
        # 탭 클릭 이벤트 연결
        self.tab_widget.tabBarClicked.connect(self.on_tab_clicked)
        
        # 데이터 탭 생성
        self.parent.data_tab = QWidget()
        data_tab_layout = QVBoxLayout(self.parent.data_tab)
        
        # 테이블 위젯 설정
        data_tab_layout.addWidget(self.parent.table)
        
        # 데이터 탭 추가
        self.tab_widget.addTab(self.parent.data_tab, "데이터")
        
        # 빈 탭 2개 추가 (테이블 포함)
        self.parent.tab2 = QWidget()
        tab2_layout = QVBoxLayout(self.parent.tab2)
        tab2_table = self.create_table_widget()
        tab2_layout.addWidget(tab2_table)
        self.tab_widget.addTab(self.parent.tab2, "탭 1")
        
        self.parent.tab3 = QWidget()
        tab3_layout = QVBoxLayout(self.parent.tab3)
        tab3_table = self.create_table_widget()
        tab3_layout.addWidget(tab3_table)
        self.tab_widget.addTab(self.parent.tab3, "탭 2")
        
        # 새 탭 추가 버튼(+) 추가
        self.parent.add_tab_button = QWidget()
        self.tab_widget.addTab(self.parent.add_tab_button, "+")
        
        # 데이터 탭 생성 및 설정 후
        # 기존 탭 이름들을 콤보박스에 추가
        for i in range(1, 3):  # 탭 1과 탭 2
            tab_name = f"탭 {i}"
            self.update_combo_with_tab_name(tab_name)
    
    def create_table_widget(self):
        """테이블 위젯 생성 및 설정"""
        from PyQt5.QtWidgets import QTableWidget
        table = QTableWidget()
        # 기본 테이블과 동일한 설정 적용
        table.setColumnCount(self.parent.table.columnCount())
        table.setRowCount(self.parent.table.rowCount())
        
        # 테이블 이벤트 연결
        table.cellClicked.connect(self.parent.table_manager.on_cell_clicked)
        
        return table

    def rename_tab(self, index):
        """탭 이름 변경 함수"""
        # 데이터 탭(인덱스 0)은 이름 변경 불가능
        if index == 0:
            return
            
        # 현재 탭 이름 가져오기
        current_name = self.tab_widget.tabText(index)
        
        # 새 이름 입력 대화상자 표시
        new_name, ok = QInputDialog.getText(self.parent, '탭 이름 변경', 
                                         '새 탭 이름을 입력하세요:', 
                                         text=current_name)
        
        # 사용자가 확인을 누르고 이름이 비어있지 않으면 변경
        if ok and new_name:
            # 콤보박스에서 기존 탭 이름 제거
            old_index = self.parent.product_combo.findText(current_name)
            if old_index >= 0:
                self.parent.product_combo.removeItem(old_index)
            
            # 탭 이름 변경
            self.tab_widget.setTabText(index, new_name)
            
            # 새 탭 이름을 콤보박스에 추가
            if new_name != "+": # + 탭은 콤보박스에 추가하지 않음
                self.update_combo_with_tab_name(new_name)

    def on_tab_clicked(self, index):
        """탭 클릭 이벤트 핸들러"""
        # "+" 탭이 클릭되었는지 확인
        if self.tab_widget.tabText(index) == "+":
            # 클릭된 "+" 탭을 다시 마지막 위치로 이동
            self.tab_widget.removeTab(index)
            
            # 새 탭 생성
            new_tab = QWidget()
            new_tab_layout = QVBoxLayout(new_tab)
            new_tab_table = self.create_table_widget()
            new_tab_layout.addWidget(new_tab_table)
            
            new_tab_index = self.tab_widget.count()  # 새 탭이 추가될 위치 (+ 탭 이전)
            new_tab_name = f"새 탭 {new_tab_index}"
            self.tab_widget.addTab(new_tab, new_tab_name)
            
            # "+" 탭 다시 추가
            self.tab_widget.addTab(self.parent.add_tab_button, "+")
            
            # 새로 생성된 탭으로 포커스 이동
            self.tab_widget.setCurrentIndex(new_tab_index)
            
            # 새 탭 이름을 콤보박스에 추가
            self.update_combo_with_tab_name(new_tab_name)

    def update_combo_with_tab_name(self, tab_name):
        """탭 이름을 콤보박스에 추가"""
        # 이미 콤보박스에 있는지 확인
        if self.parent.product_combo.findText(tab_name) < 0:
            # 없으면 추가
            self.parent.product_combo.addItem(tab_name)

    def update_tab_combo(self):
        """현재 탭 이름을 콤보박스에 업데이트"""
        # 기존 항목 제거
        self.parent.product_combo.clear()
        
        # '전체' 항목 추가
        self.parent.product_combo.addItem("전체")
        
        # 데이터 탭을 제외한 모든 탭 이름 추가
        for i in range(1, self.tab_widget.count() - 1):  # 데이터 탭과 + 탭 제외
            tab_name = self.tab_widget.tabText(i)
            if tab_name != "+":
                self.update_combo_with_tab_name(tab_name) 