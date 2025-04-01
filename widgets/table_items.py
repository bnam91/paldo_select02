from PyQt5.QtWidgets import QTableWidgetItem
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