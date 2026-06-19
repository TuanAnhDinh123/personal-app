"""Personal Toolbox - điểm khởi động ứng dụng.

Chạy:  python main.py
"""
from app.ui.main_window import MainWindow


def main():
    app = MainWindow()
    app.mainloop()


if __name__ == "__main__":
    main()
