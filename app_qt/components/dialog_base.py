"""Khung hộp thoại tùy biến dùng chung (frameless, bo góc, kéo được).

Trả về (dialog, card, layout) để bên gọi tự thêm nội dung + nút. Dùng cho các
hộp thoại soạn mail / đặt lịch / xác nhận nhiều lựa chọn.
"""
from PySide6.QtCore import Qt
from PySide6.QtWidgets import QDialog, QFrame, QHBoxLayout, QLabel, QVBoxLayout

from app_qt import widgets


class ShellDialog(QDialog):
    def __init__(self, parent):
        super().__init__(parent)
        self.setWindowFlags(Qt.FramelessWindowHint | Qt.Dialog)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setModal(True)
        self._drag = None

    def mousePressEvent(self, e):
        if e.button() == Qt.LeftButton:
            self._drag = e.globalPosition().toPoint() - self.frameGeometry().topLeft()

    def mouseMoveEvent(self, e):
        if self._drag is not None and e.buttons() & Qt.LeftButton:
            self.move(e.globalPosition().toPoint() - self._drag)

    def mouseReleaseEvent(self, e):
        self._drag = None


def build_dialog_shell(parent, title, min_width=520):
    """Dựng khung dialog. Trả về (dlg, card, content_layout)."""
    dlg = ShellDialog(parent)
    shell = QVBoxLayout(dlg)
    shell.setContentsMargins(24, 24, 24, 24)

    card = QFrame(dlg)
    card.setObjectName("Dialog")
    card.setMinimumWidth(min_width)
    widgets.add_shadow(card, blur=48, dy=12, alpha=70)
    shell.addWidget(card)

    lay = QVBoxLayout(card)
    lay.setContentsMargins(22, 20, 22, 18)   # padding thẻ→nội dung chuẩn
    lay.setSpacing(10)

    head = QHBoxLayout()
    t = QLabel(title)
    t.setObjectName("DialogTitle")
    head.addWidget(t, 1)
    close = QLabel("✕")
    close.setObjectName("DialogClose")
    close.setFixedSize(24, 24)
    close.setAlignment(Qt.AlignCenter)
    close.setCursor(Qt.PointingHandCursor)
    close.mousePressEvent = lambda _e: dlg.reject()
    head.addWidget(close, 0, Qt.AlignTop)
    lay.addLayout(head)

    return dlg, card, lay
