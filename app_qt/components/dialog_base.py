"""Khung hộp thoại tùy biến dùng chung (frameless, bo góc, kéo được).

Trả về (dialog, card, layout) để bên gọi tự thêm nội dung + nút. Dùng cho các
hộp thoại soạn mail / đặt lịch / xác nhận nhiều lựa chọn.

Kích thước lấy từ hệ 3 cỡ chung (app_qt/components/modal.py): truyền
`size="sm"|"md"|"lg"`.
"""
from app_qt.components.modal import ModalDialog


def build_dialog_shell(parent, title, size="sm"):
    """Dựng khung dialog theo cỡ chuẩn. Trả về (dlg, card, content_layout)."""
    dlg = ModalDialog(parent, size)
    card, lay = dlg.build_shell(title, spacing=10)
    return dlg, card, lay
