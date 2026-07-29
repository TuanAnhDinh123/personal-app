"""Màn hình "Cài đặt" (PySide6). Dùng lại app.core.settings (backend không đổi)."""
from PySide6.QtWidgets import QFrame, QHBoxLayout, QVBoxLayout, QWidget

from app.core import outlook, settings
from app_qt import dialogs, theme, widgets


def _group_card(parent_layout):
    card = widgets.Card()
    inner = QWidget(card)
    lay = QVBoxLayout(card)
    lay.setContentsMargins(22, 20, 22, 18)   # padding thẻ→nội dung chuẩn
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
    fields["api_key"] = widgets.text_row(inner, "Gemini API key")
    fields["api_key"].set(data["api_key"])
    fields["ai_model"] = widgets.model_select_row(
        inner, "AI Model",
        lambda: settings.list_models(fields["api_key"].get()))
    fields["ai_model"].set(data["ai_model"])
    # Đổi API key thì cho phép tải lại danh sách model ở lần mở kế tiếp.
    fields["api_key"].widget.textChanged.connect(
        lambda *_: fields["ai_model"].widget.reset())

    # ---- Nhóm Xuất Excel ----
    inner = _group_card(outer_lay)
    widgets.section_label(inner, "Course roster export")
    fields["course_template_path"] = widgets.file_row(
        inner, "Template Excel file (course roster)", mode="file")
    fields["course_template_path"].set(data["course_template_path"])

    # ---- Nhóm Mail chúc mừng sinh nhật ----
    inner = _group_card(outer_lay)
    widgets.section_label(inner, "Birthday email")
    fields["birthday_images_folder"] = widgets.file_row(
        inner, "Birthday cards folder (Canva Bulk Create, filename = employee code)",
        mode="folder")
    fields["birthday_images_folder"].set(data["birthday_images_folder"])
    fields["birthday_from_account"] = widgets.model_select_row(
        inner, "Send birthday emails from account", outlook.list_accounts)
    fields["birthday_from_account"].set(data["birthday_from_account"])
    widgets.hint(inner, "Pick the Outlook account to send birthday emails from, "
                        "if it should differ from the account used elsewhere in the app. "
                        "Leave empty to use Outlook's default account.")

    # ---- Nút lưu ----
    # Thẻ tự chừa CARD_PAD cho bóng → nút (không phải thẻ) thêm lề trái CARD_PAD
    # để thẳng hàng mép thẻ nhìn thấy.
    actions = QHBoxLayout()
    actions.setContentsMargins(widgets.CARD_PAD, 4, 0, 0)

    def save():
        values = {
            "api_key": fields["api_key"].get().strip(),
            "ai_model": fields["ai_model"].get().strip() or settings.DEFAULTS["ai_model"],
            "course_template_path": fields["course_template_path"].get().strip(),
            "birthday_images_folder": fields["birthday_images_folder"].get().strip(),
            "birthday_from_account": fields["birthday_from_account"].get().strip(),
        }
        settings.save(values)
        dialogs.success(outer, "Saved", "Settings saved ✅")

    btn = widgets.button(outer, "Save settings", variant="primary", icon="save", command=save)
    actions.addWidget(btn)
    actions.addStretch(1)
    outer_lay.addLayout(actions)
    outer_lay.addStretch(1)
    return outer
