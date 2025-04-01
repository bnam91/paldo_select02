from PyQt5.QtWidgets import QPushButton

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