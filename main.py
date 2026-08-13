"""Personal Toolbox — điểm khởi động (giao diện PySide6).

Chạy:  python main.py
"""
import sys

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QApplication

from app.core import debuglog
from app_qt import theme
from app_qt.main_window import MainWindow


def main():
    # Gắn hook ghi log trước khi dựng giao diện: app chạy bằng pythonw.exe
    # (không console) nên đây là nơi duy nhất còn đọc lại được traceback.
    debuglog.install()
    app = QApplication(sys.argv)
    # Tắt hiệu ứng "unroll" khi mở dropdown: cửa sổ chính frameless +
    # trong suốt (WA_TranslucentBackground) nên khi combobox chạy animation
    # xổ xuống, Windows phải composite lại toàn bộ layered window mỗi khung
    # hình → thấy như UI nháy 1 cái trước khi dropdown hiện ra.
    app.setEffectEnabled(Qt.UI_AnimateCombo, False)
    app.setStyleSheet(theme.build_stylesheet())   # nạp "CSS" (theme.qss) của app
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
