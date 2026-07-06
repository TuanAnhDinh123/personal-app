"""Personal Toolbox - điểm khởi động ứng dụng.

Chạy:  python main.py
"""
from app.ui import widgets
from app.ui.main_window import MainWindow


def main():
    widgets.install_loading_buttons()   # gắn hiệu ứng loading cho mọi nút
    app = MainWindow()
    app.mainloop()


if __name__ == "__main__":
    main()
