"""
main_app.py

这是应用程序的主入口点，负责创建并运行 PyQt5 应用程序。
"""
import sys
from PyQt5.QtWidgets import QApplication
from ui_elements import UniApp

if __name__ == '__main__':
    app = QApplication(sys.argv)
    win = UniApp()
    win.show()
    sys.exit(app.exec_())