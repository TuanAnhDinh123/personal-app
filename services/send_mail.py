import tkinter as tk
from tkinter import ttk, messagebox

import ttkbootstrap as tb
from ttkbootstrap.constants import *
from ttkbootstrap import Style

class SendMailFrame(tk.Frame):
    def __init__(self, master):
        super().__init__(master)
        tk.Label(self, text="Giao diện Bài toán Send Mail").pack(pady=10)
        tk.Button(self, text="Chạy xử lý Send mail", command=self.xu_ly).pack(pady=5)

    def xu_ly(self):
        messagebox.showinfo("Send mail", "Xử lý Send mail hoàn tất")