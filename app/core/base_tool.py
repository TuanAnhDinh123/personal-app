"""Lớp cha cho mọi công cụ (tool).

Mỗi tác vụ trong app là một class kế thừa BaseTool. Chỉ cần khai báo
metadata (tên, mô tả, icon, nhóm) và phần thân giao diện trong build_body().
Khung chung (thẻ trắng + nút thực hiện + thông báo) đã được lo sẵn.
"""
from abc import ABC, abstractmethod
from tkinter import messagebox

import tkinter as tk
import ttkbootstrap as ttk

from app.ui import theme, widgets


class BaseTool(ABC):
    # --- Metadata: ghi đè trong class con ---
    name: str = "Công cụ"
    description: str = ""
    icon: str = "🔧"
    category: str = "Khác"
    order: int = 100              # thứ tự hiển thị trong cùng một nhóm
    action_label: str = "Thực hiện"
    action_style: str = "primary"  # variant màu của nút chính (xem widgets.BUTTON_VARIANTS)
    action_icon: str = "play"      # icon vector của nút chính (xem widgets._ICON_DRAWERS)
    auto_startup: bool = False     # nếu True, MainWindow gọi startup() khi mở app
    show_on_home: bool = True      # False → ẩn thẻ ở Trang chủ (vẫn hiện ở sidebar)
    fills_height: bool = False     # True → trang chiếm full chiều cao (không bọc scroll)

    def build(self, parent) -> tk.Frame:
        """Dựng giao diện của tool — khung chung, hiếm khi cần ghi đè."""
        outer = tk.Frame(parent, bg=theme.CONTENT_BG)

        card = tk.Frame(
            outer, bg=theme.CARD_BG,
            highlightbackground=theme.BORDER, highlightthickness=1,
        )
        card.pack(fill="x", anchor="n")

        inner = tk.Frame(card, bg=theme.CARD_BG)
        inner.pack(fill="both", expand=True, padx=28, pady=24)

        self.build_body(inner)

        actions = tk.Frame(inner, bg=theme.CARD_BG)
        actions.pack(fill="x", pady=(22, 0))
        widgets.button(
            actions, text=self.action_label, variant=self.action_style,
            icon=self.action_icon, command=self.run,
        ).pack(side="left")

        return outer

    @abstractmethod
    def build_body(self, parent) -> None:
        """Dựng các ô nhập liệu của tool. Class con bắt buộc cài đặt."""
        ...

    def run(self) -> None:
        """Hành động khi bấm nút chính. Hiện chỉ hiển thị thông báo mẫu."""
        messagebox.showinfo("Hoàn tất", f'Task "{self.name}" đã hoàn thành ✅')

    def startup(self, window) -> None:
        """Chạy tự động khi mở app (chỉ khi auto_startup=True).

        `window` là MainWindow — dùng để điều hướng / hiện hộp thoại nếu cần.
        Mặc định không làm gì; tool cần tính năng nền sẽ ghi đè.
        """
