"""Personal Toolbox — điểm khởi động bản GIAO DIỆN CŨ (Tkinter).

Giữ lại để đối chiếu / rollback. Bản chính thức giờ là giao diện PySide6
(chạy bằng `python main.py`). Muốn xem lại giao diện cũ: `python main_tk.py`.
"""
from app.ui import widgets
from app.ui.main_window import MainWindow


def main():
    widgets.install_loading_buttons()   # gắn hiệu ứng loading cho mọi nút
    app = MainWindow()
    widgets.install_copy_support(app)    # cho phép sao chép chữ ở mọi nơi
    app.mainloop()


if __name__ == "__main__":
    main()
