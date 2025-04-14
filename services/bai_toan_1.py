import tkinter as tk
from tkinter import ttk, messagebox

import ttkbootstrap as tb
from ttkbootstrap.constants import *
from ttkbootstrap import Style

class BaiToanA1Frame(tk.Frame):
    def __init__(self, master):
        super().__init__(master)
        tk.Label(self, text="Giao diện Bài toán A1").pack(pady=10)
        tk.Button(self, text="Xử lý A1 - Bước 1", command=self.buoc1).pack(pady=5)
        tk.Button(self, text="Xử lý A1 - Bước 2", command=self.buoc2).pack(pady=5)

    def buoc1(self):
        messagebox.showinfo("A1", "Đã xử lý bước 1 của A1")

    def buoc2(self):
        messagebox.showinfo("A1", "Đã xử lý bước 2 của A1")