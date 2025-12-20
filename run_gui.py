import sys
import os

# src 폴더를 path에 추가하여 모듈 import 가능하게 함
sys.path.append(os.path.join(os.path.dirname(__file__), 'src'))

def close_splash():
    """PyInstaller 스플래시 닫기"""
    try:
        import pyi_splash
        pyi_splash.close()
    except ImportError:
        pass  # 개발 모드에서는 무시

if __name__ == "__main__":
    from hwpconv.gui import HwpConverterApp
    close_splash()  # GUI 로드 완료 후 스플래시 닫기
    app = HwpConverterApp()
    app.mainloop()
