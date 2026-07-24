"""Hộp thoại tùy biến — thay QMessageBox mặc định của hệ điều hành.

Bo góc, không dùng khung Windows, header có chip icon màu theo loại thông báo,
nút bấm cùng phong cách app. Dùng: dialogs.info(parent, title, msg) / .error /
.success / .warning / .confirm(...).
"""
from PySide6.QtCore import Qt
from PySide6.QtWidgets import QFrame, QHBoxLayout, QLabel, QVBoxLayout

from app_qt import theme, widgets
from app_qt.components.modal import ModalDialog

_KINDS = {
    "info":    ("ℹ", theme.PALETTE["--accent"]),
    "success": ("✓", theme.PALETTE["--success"]),
    "warning": ("!", theme.PALETTE["--warning"]),
    "error":   ("✕", theme.PALETTE["--danger"]),
    "question": ("?", theme.PALETTE["--info"]),
}


class AppDialog(ModalDialog):
    # Hộp báo/xác nhận: cỡ nội dung tự co theo chữ, nhưng KHÔNG rộng quá cỡ sm.
    def __init__(self, parent, title, message, kind="info", buttons=None):
        super().__init__(parent, "sm")
        self._result = 0

        emoji, color = _KINDS.get(kind, _KINDS["info"])
        buttons = buttons or [("OK", "primary", 1)]

        # lề ngoài để lộ shadow
        shell = QVBoxLayout(self)
        shell.setContentsMargins(24, 24, 24, 24)

        card = QFrame(self)
        card.setObjectName("Dialog")
        card.setMaximumWidth(self.modal_w)   # tham chiếu cỡ sm (trần bề rộng)
        widgets.add_shadow(card, blur=48, dy=12, alpha=70)
        shell.addWidget(card)

        lay = QVBoxLayout(card)
        lay.setContentsMargins(22, 20, 22, 18)   # padding thẻ→nội dung chuẩn
        lay.setSpacing(14)

        # header: chip icon + tiêu đề + nút đóng
        head = QHBoxLayout()
        head.setSpacing(12)
        chip = QLabel(emoji)
        chip.setFixedSize(38, 38)
        chip.setAlignment(Qt.AlignCenter)
        r, g, b = widgets._hex_to_rgb(color)
        chip.setStyleSheet(
            f"background: rgba({r},{g},{b},0.15); color: {color};"
            f"border-radius: 11px; font-size: 18px; font-weight: 700;")
        head.addWidget(chip)
        t = QLabel(title)
        t.setObjectName("DialogTitle")
        head.addWidget(t, 1)
        close = QLabel("✕")
        close.setObjectName("DialogClose")
        close.setFixedSize(24, 24)
        close.setAlignment(Qt.AlignCenter)
        close.setCursor(Qt.PointingHandCursor)
        close.mousePressEvent = lambda _e: self.reject()
        head.addWidget(close, 0, Qt.AlignTop)
        lay.addLayout(head)

        # nội dung
        msg = QLabel(message)
        msg.setObjectName("DialogMsg")
        msg.setWordWrap(True)
        msg.setTextInteractionFlags(Qt.TextSelectableByMouse)
        msg.setMinimumWidth(360)
        msg.setMaximumWidth(460)
        lay.addWidget(msg)

        # nút
        row = QHBoxLayout()
        row.setContentsMargins(0, 6, 0, 0)
        row.addStretch(1)
        for label, variant, value in buttons:
            b = widgets.button(card, label, variant=variant,
                               command=lambda v=value: self._done(v))
            row.addWidget(b)
        lay.addLayout(row)

    def _done(self, value):
        self._result = value
        self.accept()

    def run(self):
        self.exec()
        return self._result


# --------------------------------------------------------------- API tiện dùng
def info(parent, title, message):
    return AppDialog(parent, title, message, "info").run()


def success(parent, title, message):
    return AppDialog(parent, title, message, "success").run()


def warning(parent, title, message):
    return AppDialog(parent, title, message, "warning").run()


def error(parent, title, message):
    return AppDialog(parent, title, message, "error").run()


def confirm(parent, title, message, ok_label="Đồng ý", cancel_label="Hủy"):
    """Trả về True nếu người dùng bấm nút đồng ý."""
    d = AppDialog(parent, title, message, "question", buttons=[
        (cancel_label, "neutral", 0),
        (ok_label, "primary", 1),
    ])
    return d.run() == 1
