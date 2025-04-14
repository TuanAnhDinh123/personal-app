import tkinter as tk
from tkinter import ttk, messagebox

import ttkbootstrap as tb
from ttkbootstrap.constants import *
from ttkbootstrap import Style

class BaiToanA2Frame(tk.Frame):
    def __init__(self, master):
        super().__init__(master)
        tk.Label(self, text="Giao diện Bài toán A2").pack(pady=10)
        tk.Button(self, text="Chạy xử lý A2", command=self.xu_ly).pack(pady=5)

    def xu_ly(self):
        messagebox.showinfo("A2", "Xử lý A2 hoàn tất")