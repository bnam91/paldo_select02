import os
import sys
import PyQt5
from PyQt5.QtWidgets import QApplication
from gui import ExcelViewer
from release_updater import ReleaseUpdater

# Qt 플러그인 경로 설정
os.environ['QT_QPA_PLATFORM_PLUGIN_PATH'] = os.path.join(os.path.dirname(PyQt5.__file__), 'Qt5', 'plugins')

# 환경 변수나 설정 파일에서 값 읽기
owner = os.environ.get("GITHUB_OWNER", "bnam91")
repo = os.environ.get("GITHUB_REPO", "paldo_select02")

def run_program():
    # PyQt5 애플리케이션 실행
    app = QApplication(sys.argv)
    window = ExcelViewer()
    window.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    try:
        updater = ReleaseUpdater(owner=owner, repo=repo)
        update_success = updater.update_to_latest()
        
        if update_success:
            print("프로그램을 실행합니다...")
        else:
            print("업데이트 실패, 이전 버전으로 계속 진행합니다...")
        
        # 업데이트 결과와 상관없이 프로그램 실행
        run_program()
        
    except Exception as e:
        print(f"예상치 못한 오류 발생: {e}")
        # 오류가 발생해도 프로그램 실행
        run_program()