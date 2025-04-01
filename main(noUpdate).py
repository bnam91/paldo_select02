import sys
from PyQt5.QtWidgets import QApplication
from gui import ExcelViewer

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ExcelViewer()
    window.show()
    sys.exit(app.exec_()) 