import sys
from PyQt5.QtWidgets import QApplication
from ui_elements import UniApp # 从 ui_elements 导入 UniApp

if __name__ == '__main__':
    app = QApplication(sys.argv)
    win = UniApp()
    win.show()
    sys.exit(app.exec_())