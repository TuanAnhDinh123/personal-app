import tkinter as tk
from tkinter import ttk, messagebox
import tkinter.font as tkFont

import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from ttkbootstrap import Style

import pandas as pd
import os
import win32com.client as win32
from datetime import datetime
from tkinter import filedialog, messagebox

class SendMailFrame(ttk.Frame):
  def __init__(self, master):
    super().__init__(master, padding=20)
    self.file_path = None
    self.custom_font = tkFont.Font(family="Segoe UI", size=11)
    
    # Label
    self.label = ttk.Label(self, text="Chọn file Excel chứa thông tin gửi mail:")
    self.label.pack(pady=(0, 10))

    # # Entry để hiển thị tên file (readonly)
    # self.file_entry = ttk.Entry(
    #     self,
    #     textvariable=self.file_path,
    #     font=self.custom_font,
    #     state="readonly",
    #     width=40,
    #     bootstyle="secondary"
    # )
    # self.file_entry.pack(side="left", fill="x", expand=True, padx=(0, 10), ipady=5)

    # # Nút chọn file
    # self.select_button = ttk.Button(
    #     self,
    #     text="Chọn file Excel",
    #     command=self.choose_file,
    #     bootstyle="primary"
    # )
    # # self.select_button.configure(font=self.custom_font)
    # self.select_button.pack(side="left", ipady=5, ipadx=10)
    
    # Nút chọn file
    self.select_button = ttk.Button(self, text="Chọn file Excel", command=self.choose_file)
    self.select_button.pack(pady=5)

    # Hiển thị tên file đã chọn
    self.file_label = ttk.Label(self, text="", foreground="gray")
    self.file_label.pack(pady=(0, 10))

    # Nút gửi mail
    self.send_button = ttk.Button(self, text="Gửi mail theo lịch", bootstyle=SUCCESS, command=self.send_scheduled_emails)
    self.send_button.pack(pady=20)

  def choose_file(self):
    path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if path:
      self.file_path = path
      self.file_label.config(text=os.path.basename(path))

  def send_scheduled_emails(self):
      if not self.file_path:
        messagebox.showwarning("Thiếu file", "Vui lòng chọn file Excel trước.")
        return

      try:
          df = pd.read_excel(self.file_path)
      except Exception as e:
          messagebox.showerror("Lỗi đọc file", str(e))
          return

      # Check cột bắt buộc
      if not all(col in df.columns for col in ["email", "time", "Linked"]):
          messagebox.showerror("Thiếu cột", "File phải có đủ 3 cột: email, time, Linked")
          return

      # Kiểm tra Linked
      invalid_links = df[~df["Linked"].apply(lambda x: isinstance(x, str) and os.path.isfile(x))]
      if not invalid_links.empty:
          messagebox.showwarning("Path không hợp lệ", f"Có {len(invalid_links)} đường dẫn không tồn tại.")
          return

      outlook = win32.Dispatch("Outlook.Application")
      namespace = outlook.GetNamespace("MAPI")
      accounts = namespace.Accounts
      for acc in accounts:
        print(acc.SmtpAddress)  # debug thử
        if "tuananhhehe19957@gmail.com" in acc.SmtpAddress:
          account = acc
          break

      if not account:
          messagebox.showerror("Không tìm thấy email gửi", "Không tìm thấy secondary email trong Outlook.")
          return

      count = 0
      for index, row in df.iterrows():
          try:
              mail = outlook.CreateItem(0)
              mail._oleobj_.Invoke(*(64209, 0, 8, 0, account))  # chọn From là secondary
              mail.To = row["email"]
              mail.Subject = "Test subject"
              mail.Body = "test content"
              mail.Attachments.Add(Source=row["Linked"])
              # Đặt giờ gửi
              send_time = pd.to_datetime(row["time"])
              mail.DeferredDeliveryTime = send_time
              mail.Save()  # Lưu vào Outbox
              count += 1
          except Exception as e:
              print(f"Lỗi gửi cho {row['email']}: {e}")

      messagebox.showinfo("Hoàn tất", f"Đã tạo {count} mail vào Outbox.")
