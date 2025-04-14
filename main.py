import tkinter as tk
from tkinter import ttk

import ttkbootstrap as tb
from ttkbootstrap.constants import *
from ttkbootstrap import Style

from services.bai_toan_1 import BaiToanA1Frame
from services.bai_toan_2 import BaiToanA2Frame
from services.send_mail import SendMailFrame

class App(tb.Window):
    def __init__(self):
        super().__init__(themename="flatly")  # hoặc darkly, morph, litera...
        self.title("Personal App")
        self.geometry("1000x500")

        # Sidebar
        self.sidebar = tb.Frame(self, width=200)
        self.sidebar.pack(side=LEFT, fill=Y)

        self.content = tb.Frame(self, bootstyle="light")
        self.content.pack(side=RIGHT, fill=BOTH, expand=YES)

        # Treeview trong sidebar
        self.tree = ttk.Treeview(self.sidebar, show="tree")
        self.tree.pack(padx=10, pady=10, fill=Y, expand=True)

        a = self.tree.insert("", "end", text="Bài toán A", open=True)
        self.tree.insert(a, "end", text="Bài toán A1")
        self.tree.insert(a, "end", text="Bài toán A2")

        send_mail = self.tree.insert("", "end", text="Send Mail")

        # Gắn sự kiện chọn
        self.tree.bind("<<TreeviewSelect>>", self.on_tree_select)

        # Hiện giao diện đầu tiên (nếu muốn)
        self.current_frame = None

    def on_tree_select(self, event):
        selected = self.tree.focus()
        ten = self.tree.item(selected)["text"]

        # Xóa frame cũ
        if self.current_frame:
            self.current_frame.destroy()

        # Tạo frame mới
        if ten == "Bài toán A1":
            self.current_frame = BaiToanA1Frame(self.content)
        elif ten == "Bài toán A2":
            self.current_frame = BaiToanA2Frame(self.content)
        elif ten == "Send Mail":
            self.current_frame = SendMailFrame(self.content)
        else:
            self.current_frame = None

        if self.current_frame:
            self.current_frame.pack(fill=BOTH, expand=YES, padx=20, pady=20)

# ------------------------
if __name__ == "__main__":
    app = App()
    app.mainloop()