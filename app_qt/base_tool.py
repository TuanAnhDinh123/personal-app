"""Lớp cha cho mọi tool — bản PySide6 (thay cho app/core/base_tool.py).

Giữ NGUYÊN hợp đồng: khai báo metadata (name/description/icon/category/order…)
và build_body(). Khung chung (thẻ + nút thực hiện) lo sẵn. Chỉ khác: build_body
nhận một QWidget đã có sẵn QVBoxLayout để thêm các ô nhập (thay cho tk.Frame).
"""
from abc import ABC, abstractmethod

from PySide6.QtWidgets import QFrame, QHBoxLayout, QVBoxLayout, QWidget

from app_qt import dialogs, widgets


class BaseTool(ABC):
    # --- Metadata: ghi đè ở class con ---
    name: str = "Tool"
    description: str = ""
    icon: str = "🔧"
    category: str = "Other"
    order: int = 100
    action_label: str = "Run"
    action_style: str = "primary"     # variant nút chính
    action_icon: str = "play"
    auto_startup: bool = False
    show_on_home: bool = True
    fills_height: bool = False

    def build(self, parent=None) -> QWidget:
        """Dựng giao diện tool: thẻ trắng + body + nút thực hiện."""
        outer = QWidget(parent)
        outer_lay = QVBoxLayout(outer)
        outer_lay.setContentsMargins(0, 0, 0, 0)
        outer_lay.setSpacing(0)

        card = widgets.Card(outer)
        card_lay = QVBoxLayout(card)
        card_lay.setContentsMargins(22, 20, 22, 18)   # padding thẻ→nội dung chuẩn
        card_lay.setSpacing(4)

        body = QWidget(card)
        body_lay = QVBoxLayout(body)
        body_lay.setContentsMargins(0, 0, 0, 0)
        body_lay.setSpacing(6)
        self.build_body(body)
        card_lay.addWidget(body)

        actions = QHBoxLayout()
        actions.setContentsMargins(0, 18, 0, 0)
        btn = widgets.button(
            card, self.action_label, variant=self.action_style,
            icon=self.action_icon, command=self.run,
        )
        actions.addWidget(btn)
        for extra in self.extra_actions(card):
            actions.addWidget(extra)
        actions.addStretch(1)
        card_lay.addLayout(actions)

        outer_lay.addWidget(card)
        outer_lay.addStretch(1)
        self._page = outer
        return outer

    def extra_actions(self, parent: QWidget) -> list:
        """Nút phụ đặt cạnh nút chính (trả về list widget nút; mặc định không có)."""
        return []

    @abstractmethod
    def build_body(self, parent: QWidget) -> None:
        """Dựng các ô nhập của tool. Class con bắt buộc cài đặt."""
        ...

    def run(self) -> None:
        dialogs.success(self._page, "Done", f'Task "{self.name}" completed ✅')

    def startup(self, window) -> None:
        """Chạy tự động khi mở app (chỉ khi auto_startup=True)."""

    # --- tiện ích hộp thoại cho class con ---
    def info(self, title, msg):
        dialogs.info(getattr(self, "_page", None), title, msg)

    def error(self, title, msg):
        dialogs.error(getattr(self, "_page", None), title, msg)

    def confirm(self, title, msg, ok_label="OK", cancel_label="Cancel"):
        return dialogs.confirm(getattr(self, "_page", None), title, msg,
                               ok_label, cancel_label)
