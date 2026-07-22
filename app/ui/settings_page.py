"""Màn hình "Cài đặt" — nội dung các thiết lập dùng chung cho toàn app.

Hiện có nhóm **AI (Gemini)**: API key + model mặc định. Tương lai muốn thêm
nhóm thiết lập mới → viết thêm một hàm `_build_<tên>(parent, data, fields)` và
gọi nó trong `build()`; nhớ thêm khóa tương ứng ở `app/core/settings.py`.
"""
from tkinter import messagebox

import tkinter as tk
import ttkbootstrap as ttk

from app.core import settings
from app.ui import theme, widgets


def _group_card(parent):
    """Một thẻ trắng bao quanh một nhóm thiết lập (giống khung của tool)."""
    card = tk.Frame(
        parent, bg=theme.CARD_BG,
        highlightbackground=theme.BORDER, highlightthickness=1,
    )
    card.pack(fill="x", anchor="n", pady=(0, 16))
    inner = tk.Frame(card, bg=theme.CARD_BG)
    inner.pack(fill="both", expand=True, padx=28, pady=24)
    return inner


def _build_ai_group(parent, data, fields):
    """Nhóm thiết lập AI: API key Gemini + model mặc định."""
    inner = _group_card(parent)
    widgets.section_label(inner, "AI (Gemini)")

    fields["api_key"] = widgets.text_row(inner, "API key Gemini", data["api_key"])
    widgets.hint(
        inner,
        "Lấy key miễn phí tại https://aistudio.google.com/apikey — key dùng "
        "chung cho mọi tính năng AI (vd Quét CV bằng AI) và được lưu trong máy.",
    )

    fields["ai_model"] = widgets.text_row(inner, "Model mặc định", data["ai_model"])
    widgets.hint(
        inner,
        "Mặc định 'gemini-3.6-flash'. Free tier giới hạn theo TỪNG model "
        "(vd 5 request/phút, 20/ngày) — nếu một model báo lỗi 429/503 do chạm "
        "trần, đổi sang model khác (gemini-3.5-flash, gemini-2.5-flash).",
    )


def build(parent):
    """Dựng nội dung màn hình Cài đặt, trả về frame ngoài cùng."""
    outer = tk.Frame(parent, bg=theme.CONTENT_BG)

    data = settings.load()
    fields = {}   # khóa cấu hình -> biến/widget để đọc lại khi lưu

    _build_ai_group(outer, data, fields)
    # (Tương lai) thêm nhóm thiết lập khác ở đây, ví dụ:
    # _build_email_group(outer, data, fields)

    actions = tk.Frame(outer, bg=theme.CONTENT_BG)
    actions.pack(fill="x", pady=(4, 0))

    def save():
        values = {
            "api_key":  fields["api_key"].get().strip(),
            "ai_model": fields["ai_model"].get().strip() or settings.DEFAULTS["ai_model"],
        }
        settings.save(values)
        messagebox.showinfo("Đã lưu", "Đã lưu cài đặt ✅")

    widgets.button(
        actions, text="Lưu cài đặt", variant="primary", icon="save",
        command=save,
    ).pack(side="left")

    return outer
