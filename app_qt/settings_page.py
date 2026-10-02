"""Màn hình "Cài đặt" (PySide6). Dùng lại app.core.settings (backend không đổi)."""
from PySide6.QtWidgets import QFrame, QHBoxLayout, QVBoxLayout, QWidget

from app.core import outlook, settings, shared_db
from app_qt import dialogs, theme, widgets
from app_qt.components.task import Task


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

    # Hồ sơ người dùng KHÔNG ở màn hình này: nó lưu xuống bảng `users` của file
    # .db chứ không phải config.json, và sửa bằng modal riêng mở từ ô hồ sơ ở
    # sidebar (`app_qt/profile.py`). Mọi ô dưới đây đều thuộc config.json và
    # dùng chung đúng một nút *Save settings*.

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

    # ---- Nhóm Quyết định thôi việc ----
    inner = _group_card(outer_lay)
    widgets.section_label(inner, "Resignation decision")
    fields["resignation_template_path"] = widgets.file_row(
        inner, "Word template (.docx with mail merge fields)", mode="file")
    fields["resignation_template_path"].set(data["resignation_template_path"])
    fields["resignation_output_folder"] = widgets.file_row(
        inner, "Output folder", mode="folder")
    fields["resignation_output_folder"].set(data["resignation_output_folder"])
    # Ô này chỉ nhận SỐ THỨ TỰ; phần năm do app ghép vào (và tự đánh số lại từ 1
    # khi sang năm mới) nên cho gõ cả "2026-20" thì lần sau năm lại lệch.
    dec_year, dec_no = settings.next_decision_number(data)
    fields["resignation_decision_no"] = widgets.digit_entry(
        inner, "Next decision number")
    fields["resignation_decision_no"].set(str(dec_no))
    widgets.hint(inner, f"The next decision will be numbered "
                        f"{dec_year}-{dec_no:02d}, and the number goes up by one "
                        f"for every file exported. The year comes from the "
                        f"system clock: on 1 January the count restarts at 1. "
                        f"One file per employee, named <employee code>_<full "
                        f"name>.docx — export it by right-clicking a row on the "
                        f"Employees screen.")

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
                        "— after that it is marked Missed and never sent automatically. "
                        "Set it to 0 to turn catch-up off: anyone whose birthday has "
                        "already passed is marked Missed instead.")

    # ---- Nhóm DB dùng chung ----
    inner = _group_card(outer_lay)
    widgets.section_label(inner, "Shared database")
    fields["shared_db_folder"] = widgets.file_row(
        inner, "Shared database folder (network drive)", mode="folder")
    fields["shared_db_folder"].set(data["shared_db_folder"])
    widgets.hint(inner, "When set, several computers share one database. Each works "
                        "on its own copy and the app keeps them in step: it downloads "
                        "the newest shared copy at start-up and every few minutes, "
                        "and uploads your changes about "
                        f"{shared_db.GRACE} seconds after you stop editing, or when "
                        "you close the app (keeping the last "
                        f"{shared_db.HISTORY_KEEP} copies in a history subfolder). "
                        "Only one computer saves at a time — if someone else is "
                        "saving, you'll be asked to try again in a few seconds. If "
                        "the folder can't be reached (VPN off), you can still view "
                        "the data but not save. Leave it empty to keep the database "
                        "on this computer only.")
    fields["shared_db_read_only"] = widgets.checkbox(
        inner, "Read-only on this computer", checked=bool(data["shared_db_read_only"]))
    widgets.hint(inner, "For a computer that only needs to look at the data. It still "
                        "downloads the latest shared database (at start-up, every few "
                        "minutes, and with Refresh in the sidebar), but it can't save "
                        "changes and never uploads.")
    backup_row = QHBoxLayout()
    backup_row.setContentsMargins(0, 6, 0, 0)
    backup_btn = widgets.button(inner, "Back up now", variant="neutral",
                                icon="download")
    backup_row.addWidget(backup_btn)
    backup_row.addStretch(1)
    inner.layout().addLayout(backup_row)

    def backup_now():
        folder = fields["shared_db_folder"].get().strip()
        if not folder:
            dialogs.warning(outer, "Back up now",
                            "Choose a shared database folder first.")
            return
        if folder != shared_db.shared_folder():
            dialogs.warning(outer, "Back up now",
                            "The shared database folder has changed. "
                            "Click Save settings first.")
            return
        backup_btn.setEnabled(False)
        backup_btn.setText("Backing up…")

        def done(result):
            backup_btn.setEnabled(True)
            backup_btn.setText("Back up now")
            if result.ok:
                dialogs.success(outer, "Back up now", result.message)
            elif result.skipped:
                dialogs.warning(outer, "Back up now", result.message)
            else:
                dialogs.error(outer, "Back up now",
                              f"{result.message}\n\nDetails are in debug.log.")

        # Ổ mạng có thể treo hàng chục giây → chạy nền, không đứng giao diện.
        task = Task(lambda _emit: shared_db.push("manual"), outer)
        task.signals.finished.connect(done)
        outer._backup_task = task      # giữ tham chiếu để QThread không bị thu hồi
        task.start()

    backup_btn.clicked.connect(lambda *_: backup_now())

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
            resignation_template_path=fields["resignation_template_path"].get().strip(),
            resignation_output_folder=fields["resignation_output_folder"].get().strip(),
            # Ghi kèm NĂM của số vừa nhập: người dùng sửa số ở đây là đang nói
            # về năm nay, nếu không lần xuất tới lại tưởng số cũ của năm ngoái
            # và đánh số lại từ 1.
            resignation_decision_year=dec_year,
            resignation_decision_no=max(1, _int_or_default(
                fields["resignation_decision_no"].get(), "resignation_decision_no")),
            birthday_images_folder=fields["birthday_images_folder"].get().strip(),
            birthday_from_account=fields["birthday_from_account"].get().strip(),
            birthday_subject=(fields["birthday_subject"].get().strip()
                              or settings.DEFAULTS["birthday_subject"]),
            birthday_catchup_days=_int_or_default(
                fields["birthday_catchup_days"].get(), "birthday_catchup_days"),
            shared_db_folder=fields["shared_db_folder"].get().strip(),
            shared_db_read_only=bool(fields["shared_db_read_only"].get()),
        )
        # Đổi thư mục / chế độ chỉ đọc → kiểm lại ngay (ở nền) để ô trạng thái
        # ở sidebar và bản local cập nhật luôn, không đợi nhịp định kỳ.
        shared_db.check_in_background("settings")
        dialogs.success(outer, "Saved", "Settings saved ✅")

    btn = widgets.button(outer, "Save settings", variant="primary", icon="save", command=save)
    actions.addWidget(btn)
    actions.addStretch(1)
    outer_lay.addLayout(actions)
    outer_lay.addStretch(1)
    return outer
