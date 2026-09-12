"""Màn hình "Cài đặt" (PySide6). Dùng lại app.core.settings (backend không đổi)."""
from PySide6.QtWidgets import QFrame, QHBoxLayout, QVBoxLayout, QWidget

from app.core import outlook, settings
from app_qt import dialogs, theme, widgets


def _int_or_default(text, key):
    """Ô số để trống (hoặc gõ rác) thì lấy lại giá trị mặc định, không lưu chuỗi
    rỗng — phần đọc setting mong đợi một con số."""
    try:
        return int(str(text).strip())
    except (TypeError, ValueError):
        return settings.DEFAULTS[key]


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

    # ---- Nhóm Đồng bộ nhân viên ----
    inner = _group_card(outer_lay)
    widgets.section_label(inner, "Employee data")
    fields["hc_excel_path"] = widgets.file_row(
        inner, "Personnel Data Excel file (HR headcount file)", mode="file")
    fields["hc_excel_path"].set(data["hc_excel_path"])
    widgets.hint(inner, "Used by Sync with Excel on the Employees screen: the app "
                        "reads the Personnel Data sheet and lists every difference "
                        "for you to confirm before anything is written.")

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
    fields["birthday_subject"] = widgets.text_row(inner, "Email subject")
    fields["birthday_subject"].set(data["birthday_subject"])
    widgets.hint(inner, "Use {name} for the employee's first name. The card image "
                        "is the whole message — there is no email body.")
    fields["birthday_catchup_days"] = widgets.digit_entry(
        inner, "Catch-up window (days)")
    fields["birthday_catchup_days"].set(str(data["birthday_catchup_days"]))
    widgets.hint(inner, "Each queued email is sent on the employee's own birthday, "
                        "the first time you open Personal Toolbox that day. If the app "
                        "isn't opened that day, it still goes out within this many days "
                        "— after that it is marked Missed and never sent automatically.")

    # ---- Nút lưu ----
    # Thẻ tự chừa CARD_PAD cho bóng → nút (không phải thẻ) thêm lề trái CARD_PAD
    # để thẳng hàng mép thẻ nhìn thấy.
    actions = QHBoxLayout()
    actions.setContentsMargins(widgets.CARD_PAD, 4, 0, 0)

    def save():
        # update() thay vì save(): giữ lại các thiết lập chung KHÔNG có ô nhập ở
        # màn hình này.
        settings.update(
            api_key=fields["api_key"].get().strip(),
            ai_model=fields["ai_model"].get().strip() or settings.DEFAULTS["ai_model"],
            course_template_path=fields["course_template_path"].get().strip(),
            hc_excel_path=fields["hc_excel_path"].get().strip(),
            birthday_images_folder=fields["birthday_images_folder"].get().strip(),
            birthday_from_account=fields["birthday_from_account"].get().strip(),
            birthday_subject=(fields["birthday_subject"].get().strip()
                              or settings.DEFAULTS["birthday_subject"]),
            birthday_catchup_days=_int_or_default(
                fields["birthday_catchup_days"].get(), "birthday_catchup_days"),
        )
        dialogs.success(outer, "Saved", "Settings saved ✅")

    btn = widgets.button(outer, "Save settings", variant="primary", icon="save", command=save)
    actions.addWidget(btn)
    actions.addStretch(1)
    outer_lay.addLayout(actions)
    outer_lay.addStretch(1)
    return outer
