"""
eHojo BudgetChecker - 애플리케이션 진입점

이 모듈은 eHojo BudgetChecker 애플리케이션의 시작점입니다.
GUI 창을 생성하고 이벤트 루프를 실행합니다.

사용 방법:
    python main.py

작성자: eHojo BudgetChecker Team
버전: 0.1.0
"""

from frontend.app import App


def main():
    """
    애플리케이션 메인 진입점.

    GUI 애플리케이션 창을 생성하고 실행합니다.
    사용자가 창을 닫을 때까지 이벤트 루프가 계속 실행됩니다.

    Returns:
        None
    """
    # GUI 애플리케이션 인스턴스 생성 및 실행
    app = App()
    app.mainloop()


if __name__ == '__main__':
    # 이 파일이 직접 실행될 때만 main() 함수 호출
    # (다른 파일에서 import 될 때는 실행되지 않음)
    main()

# ========== PyInstaller로 단일 exe 빌드 ==========
# 아래 명령어를 터미널에서 실행하면 단일 exe 파일이 생성됩니다.
#
# pyinstaller --onefile --windowed --icon=assets/icon.ico --add-data "assets;assets" --exclude-module numpy --exclude-module PIL --exclude-module lxml --exclude-module charset_normalizer --exclude-module scipy --exclude-module pandas main.py
#
# --onefile         : 모든 파일을 하나의 exe로 묶음
# --windowed        : 콘솔 창 없이 GUI만 실행 (Windows용)
# --icon            : exe 파일 아이콘 설정 (ico 파일 필요)
# --add-data        : assets 폴더를 exe에 포함 (아이콘 이미지 등)
# --exclude-module  : 불필요 라이브러리 제외 (파일 크기 최적화)
#
# ※ ico 파일 변환: png → ico 변환이 필요합니다
#    pip install Pillow
#    python -c "from PIL import Image; Image.open('assets/icon.png').save('assets/icon.ico')"
#
# ※ macOS에서는 --add-data 구분자가 세미콜론(;) 대신 콜론(:)입니다
#    pyinstaller --onefile --windowed --icon=assets/icon.ico --add-data "assets:assets" --exclude-module numpy --exclude-module PIL --exclude-module lxml --exclude-module charset_normalizer --exclude-module scipy --exclude-module pandas main.py
