"""Personal Toolbox — điểm khởi động (giao diện PySide6).

Chạy:  python main.py
"""
import sys

from PySide6.QtWidgets import QApplication

from app_qt import theme
from app_qt.main_window import MainWindow


def main():
    app = QApplication(sys.argv)
    app.setStyleSheet(theme.build_stylesheet())   # nạp "CSS" (theme.qss) của app
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
