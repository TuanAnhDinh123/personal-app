"""Màn hình "Cài đặt" (PySide6). Dùng lại app.core.settings (backend không đổi)."""
from PySide6.QtWidgets import QFrame, QHBoxLayout, QVBoxLayout, QWidget

from app.core import settings
from app_qt import dialogs, theme, widgets


def _group_card(parent_layout):
    card = widgets.Card()
    inner = QWidget(card)
    lay = QVBoxLayout(card)
    lay.setContentsMargins(28, 24, 28, 24)
    lay.setSpacing(6)
    inner_lay = QVBoxLayout(inner)
    inner_lay.setContentsMargins(0, 0, 0, 0)
    inner_lay.setSpacing(6)
    lay.addWidget(inner)
    parent_layout.addWidget(card)
    return inner


def build():
    """Trả về QWidget nội dung màn hình Cài đặt."""
    outer = QWidget()
    outer_lay = QVBoxLayout(outer)
    outer_lay.setContentsMargins(0, 0, 0, 0)
    outer_lay.setSpacing(16)

    data = settings.load()
    fields = {}

    # ---- Nhóm AI (Gemini) ----
    inner = _group_card(outer_lay)
    widgets.section_label(inner, "AI (Gemini)")
    fields["api_key"] = widgets.text_row(inner, "API key Gemini")
    fields["api_key"].set(data["api_key"])
    widgets.hint(
        inner,
        "Lấy key miễn phí tại https://aistudio.google.com/apikey — key dùng chung "
        "cho mọi tính năng AI (vd Quét CV bằng AI) và được lưu trong máy.")
    fields["ai_model"] = widgets.text_row(inner, "Model mặc định")
    fields["ai_model"].set(data["ai_model"])
    widgets.hint(
        inner,
        "Mặc định 'gemini-3.6-flash'. Free tier giới hạn theo TỪNG model; nếu một "
        "model báo lỗi 429/503 do chạm trần, đổi sang model khác "
        "(gemini-3.5-flash, gemini-2.5-flash).")

    # ---- Nút lưu ----
    # Thẻ tự chừa CARD_PAD cho bóng → nút (không phải thẻ) thêm lề trái CARD_PAD
    # để thẳng hàng mép thẻ nhìn thấy.
    actions = QHBoxLayout()
    actions.setContentsMargins(widgets.CARD_PAD, 4, 0, 0)

    def save():
        values = {
            "api_key": fields["api_key"].get().strip(),
            "ai_model": fields["ai_model"].get().strip() or settings.DEFAULTS["ai_model"],
        }
        settings.save(values)
        dialogs.success(outer, "Đã lưu", "Đã lưu cài đặt ✅")

    btn = widgets.button(outer, "Lưu cài đặt", variant="primary", icon="save", command=save)
    actions.addWidget(btn)
    actions.addStretch(1)
    outer_lay.addLayout(actions)
    outer_lay.addStretch(1)
    return outer
