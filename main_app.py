# main_app.py
import sys
from PyQt5.QtWidgets import QApplication
from ui_elements import UniApp
import multiprocessing

if __name__ == "__main__":
    multiprocessing.freeze_support() # 关键：在Windows下用于PyInstaller打包多进程
    try:
        app = QApplication(sys.argv)
        window = UniApp()
        window.show()
        sys.exit(app.exec_())
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
        sys.exit(1)