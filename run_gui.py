
import sys
import os

# src 폴더를 path에 추가하여 모듈 import 가능하게 함
sys.path.append(os.path.join(os.path.dirname(__file__), 'src'))

if __name__ == "__main__":
    from hwpconv.gui import HwpConverterApp
    app = HwpConverterApp()
    app.mainloop()
