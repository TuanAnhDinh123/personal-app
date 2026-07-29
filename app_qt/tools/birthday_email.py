"""Gửi mail chúc mừng sinh nhật — đính kèm ảnh thiệp xuất từ Canva Bulk Create.

Bấm "Send": tự quét bảng nhân viên tìm người có NGÀY SINH TRONG THÁNG HIỆN
TẠI, đối chiếu mã NV với thư mục ảnh đã cấu hình ở Cài đặt (tên file = mã NV)
→ hiện modal xác nhận (ai THIẾU ảnh bị đánh dấu, sẽ KHÔNG được gửi) → xác nhận
thì gửi qua Outlook (có thể dùng tài khoản riêng, khác tài khoản mặc định —
cũng cấu hình ở Cài đặt) → báo lại số mail gửi thành công + tên những người
chưa được gửi.
"""
import csv
import datetime
import os
import unicodedata

from PySide6.QtWidgets import (
    QCheckBox, QFileDialog, QFrame, QHBoxLayout, QLabel, QVBoxLayout, QWidget,
)

from app.core import cv_repository as repo
from app.core import outlook
from app.core import settings
from app_qt import dialogs, theme, widgets
from app_qt.base_tool import BaseTool
from app_qt.components.modal import ModalDialog

_IMAGE_EXTS = {".jpg", ".jpeg", ".png"}

_DEFAULT_SUBJECT = "Happy Birthday, {name}!"
_DEFAULT_BODY = (
    "Dear {name},\n\n"
    "Wishing you a very happy birthday! May your day be filled with joy, and "
    "the year ahead bring you good health, happiness and success.\n\n"
    "With warmest wishes,\nDLVN HR Team"
)


def _fill(text, name):
    return text.replace("{name}", name)


def _day_month(dob):
    """'dd/mm/yyyy' -> (ngày, tháng) dạng int, hoặc None nếu không đọc được."""
    try:
        d, m, _y = (dob or "").strip().split("/")
        return int(d), int(m)
    except (ValueError, AttributeError):
        return None


def _birth_month(dob):
    dm = _day_month(dob)
    return dm[1] if dm else None


def _strip_vn_accents(text):
    """Bỏ dấu tiếng Việt (đ/Đ xử lý riêng vì NFD không tách được ký tự này)."""
    text = (text or "").replace("đ", "d").replace("Đ", "D")
    text = unicodedata.normalize("NFD", text)
    text = "".join(ch for ch in text if unicodedata.category(ch) != "Mn")
    return unicodedata.normalize("NFC", text)


def _scan_images(folder):
    """Map mã NV (chuẩn hóa hoa) -> tên file ảnh, trong `folder`."""
    images = {}
    if folder and os.path.isdir(folder):
        for fname in os.listdir(folder):
            stem, ext = os.path.splitext(fname)
            stem = stem.strip()
            if ext.lower() in _IMAGE_EXTS and stem:
                images[stem.upper()] = fname
    return images


def _chip(parent, text, color):
    lbl = QLabel(text, parent)
    r, g, b = widgets._hex_to_rgb(color)
    lbl.setStyleSheet(
        f"background: rgba({r},{g},{b},0.15); color:{color}; border-radius:10px;"
        " padding:3px 10px; font-size:12px; font-weight:600;")
    return lbl


class BirthdayEmailTool(BaseTool):
    name = "Birthday emails"
    description = "Send birthday cards (Canva-exported images) to employees by email."
    icon = "🎂"
    category = "Office"
    order = 7

    def build(self, parent=None):
        card = widgets.Card(parent)
        lay = QVBoxLayout(card)
        lay.setContentsMargins(22, 20, 22, 18)
        lay.setSpacing(10)
        self._root = card

        widgets.section_label(card, "Email content")
        self.subject_field = widgets.text_row(card, "Subject")
        self.subject_field.set(_DEFAULT_SUBJECT)
        self.body_field = widgets.text_area(card, "Body", value=_DEFAULT_BODY, height=8)

        send_bar = QHBoxLayout()
        send_bar.setContentsMargins(0, 8, 0, 0)
        send_bar.addWidget(widgets.button(card, "Send", variant="primary",
                                          icon="mail", command=self._on_send))
        send_bar.addWidget(widgets.button(
            card, "Export CSV (missing cards)", variant="neutral",
            icon="file-text", command=self._export_missing_csv))
        send_bar.addStretch(1)
        lay.addLayout(send_bar)
        return card

    def build_body(self, parent):
        pass

    # -------------------------------------------------------------- quét + xác nhận
    def _on_send(self):
        if not outlook.available():
            dialogs.warning(self._root, "Outlook required",
                            "Sending email needs Outlook on Windows (pywin32).")
            return

        month = datetime.date.today().month
        matches = [e for e in repo.list_employees()
                  if _birth_month(e["date_of_birth"]) == month]
        if not matches:
            dialogs.info(self._root, "No birthdays",
                        "No employee has a birthday this month.")
            return

        folder = settings.get("birthday_images_folder")
        images = _scan_images(folder)

        rows = []
        for emp in matches:
            code_norm = (emp["code"] or "").strip().upper()
            image_name = images.get(code_norm)
            rows.append({
                "code": emp["code"],
                "name": emp["name"] or emp["full_name"] or emp["code"],
                "full_name": emp["full_name"] or emp["code"],
                "email": (emp["email"] or "").strip(),
                "date_of_birth": emp["date_of_birth"],
                "image_path": os.path.join(folder, image_name) if image_name else None,
            })

        dlg = _BirthdayConfirmDialog(self._root, rows, folder)
        if dlg.exec() != dlg.Accepted:
            return
        selected = dlg.selected_rows()
        if not selected:
            dialogs.info(self._root, "Nothing selected",
                        "No employee is ticked — nothing to send.")
            return
        self._send_all(selected)

    def _send_all(self, rows):
        account = settings.get("birthday_from_account").strip()
        subject_tpl = self.subject_field.get().strip() or _DEFAULT_SUBJECT
        body_tpl = self.body_field.get().strip() or _DEFAULT_BODY

        sent, not_sent = [], []
        for row in rows:
            display_name = row["full_name"]
            if not row["image_path"]:
                not_sent.append(f"{display_name} — no matching card")
                continue
            if not row["email"]:
                not_sent.append(f"{display_name} — missing email")
                continue
            try:
                outlook.send_mail(
                    row["email"], _fill(subject_tpl, row["name"]), _fill(body_tpl, row["name"]),
                    account_smtp=account or None, attachments=[row["image_path"]])
                sent.append(display_name)
            except Exception as exc:
                not_sent.append(f"{display_name} — send failed: {exc}")

        msg = f"Sent: {len(sent)}/{len(rows)}"
        if not_sent:
            msg += "\n\nNot sent:\n" + "\n".join(not_sent)
            dialogs.warning(self._root, "Done with skipped/failed", msg)
        else:
            dialogs.success(self._root, "Done", f"Sent {len(sent)} birthday email(s) ✅")

    # -------------------------------------------------------- xuất CSV cho Canva
    def _export_missing_csv(self):
        """Xuất CSV TOÀN BỘ nhân viên (không lọc theo tháng) CHƯA có card, để
        nạp vào Canva Bulk Create tạo 1 lần cho hết, khỏi phải làm lắt nhắt
        từng tháng. Cột `name` bỏ dấu tiếng Việt cho khớp mẫu thiệp. Cột
        date_of_birth ghi ngày/tháng SINH kèm NĂM HIỆN TẠI (không phải năm
        sinh thật) vì đây là ngày hiển thị trên thiệp, không phải để lộ tuổi."""
        today = datetime.date.today()
        all_employees = [e for e in repo.list_employees()
                         if _day_month(e["date_of_birth"])]
        if not all_employees:
            dialogs.info(self._root, "No employees",
                        "No employee has a usable date of birth.")
            return

        folder = settings.get("birthday_images_folder")
        images = _scan_images(folder)

        missing = []
        for emp in all_employees:
            code_norm = (emp["code"] or "").strip().upper()
            if code_norm in images:
                continue   # đã có card rồi, không cần xuất lại
            dm = _day_month(emp["date_of_birth"])
            raw_name = emp["name"] or emp["full_name"] or emp["code"] or ""
            missing.append({
                "code": emp["code"] or "",
                "name": _strip_vn_accents(raw_name),
                "date_of_birth": f"{dm[0]}/{dm[1]}/{today.year}",
            })

        if not missing:
            dialogs.info(self._root, "Nothing to export",
                        "Every employee already has a card.")
            return

        path, _ = QFileDialog.getSaveFileName(
            self._root, "Export employees without a card",
            "birthday_missing_cards.csv", "CSV (*.csv)", "",
            QFileDialog.Option.DontConfirmOverwrite)
        if not path:
            return
        if not path.lower().endswith(".csv"):
            path += ".csv"

        try:
            with open(path, "w", newline="", encoding="utf-8-sig") as f:
                writer = csv.DictWriter(f, fieldnames=["code", "name", "date_of_birth"])
                writer.writeheader()
                writer.writerows(missing)
        except OSError as exc:
            dialogs.error(self._root, "Can't write file", str(exc))
            return

        dialogs.success(
            self._root, "Exported",
            f"Exported {len(missing)} employee(s) without a card to:\n{path}\n\n"
            "In Canva Bulk Create: connect this CSV, replace the name/date text "
            "boxes with the \"name\"/\"date_of_birth\" columns, then under Advanced "
            "settings set \"Name each page using\" → code, so the downloaded card "
            "filenames match employee codes.")


class _BirthdayConfirmDialog(ModalDialog):
    """Modal xác nhận trước khi gửi: liệt kê người có sinh nhật tháng này, đánh
    dấu ai THIẾU ảnh thiệp (người đó sẽ không được gửi mail)."""

    def __init__(self, parent, rows, folder):
        super().__init__(parent, "md")
        card, lay = self.build_shell(f"Birthdays this month · {len(rows)}")
        self._checks = []   # [(row, QCheckBox)] — bỏ tick để loại người đó (vd chỉ gửi test cho mình)

        if not folder:
            widgets.hint(
                card, "⚠ No cards folder configured — set it in Settings → "
                     "Birthday email. Nobody will be emailed.")

        has_any_card = any(r["image_path"] for r in rows)
        self.select_all_cb = QCheckBox("Select all", card)
        self.select_all_cb.setChecked(has_any_card)
        self.select_all_cb.setEnabled(has_any_card)
        self.select_all_cb.stateChanged.connect(self._toggle_all)
        lay.addWidget(self.select_all_cb)

        body = QWidget()
        col = QVBoxLayout(body)
        col.setContentsMargins(0, 0, 8, 0)
        col.setSpacing(8)
        if not rows:
            empty = QLabel("No employee has a birthday this month.")
            empty.setObjectName("DialogMsg")
            col.addWidget(empty)
        for row in rows:
            col.addWidget(self._row_card(body, row))
        col.addStretch(1)
        sa = widgets.scroll_area(body)
        lay.addWidget(sa, 1)
        self.set_grow_region(sa)

        missing = sum(1 for r in rows if not r["image_path"])
        if missing:
            widgets.hint(
                card, f"⚠ {missing} employee(s) have no matching card and will "
                     "NOT be emailed.")

        foot = QHBoxLayout()
        foot.addWidget(widgets.button(card, "Send", variant="primary", icon="mail",
                                      command=self.accept))
        foot.addWidget(widgets.button(card, "Cancel", variant="neutral", icon="x",
                                      command=self.reject))
        foot.addStretch(1)
        lay.addLayout(foot)

    def selected_rows(self):
        """Danh sách row còn được tick — bỏ tick thì loại khỏi đợt gửi."""
        return [row for row, cb in self._checks if cb.isChecked()]

    def _toggle_all(self, _state=None):
        """Tick/bỏ tick "Select all" -> áp cho mọi dòng CÓ card (danh sách dài
        thì tick/bỏ hết 1 phát nhanh hơn tự bỏ từng người)."""
        checked = self.select_all_cb.isChecked()
        for row, cb in self._checks:
            if cb.isEnabled():
                cb.setChecked(checked)

    def _row_card(self, parent, row):
        box = QFrame(parent)
        box.setObjectName("DetailCard")
        h = QHBoxLayout(box)
        h.setContentsMargins(14, 10, 14, 10)
        h.setSpacing(8)

        has_card = bool(row["image_path"])
        cb = QCheckBox(box)
        cb.setChecked(has_card)      # thiếu card thì không có gì để gửi -> mặc định bỏ tick
        cb.setEnabled(has_card)
        h.addWidget(cb)
        self._checks.append((row, cb))

        name = QLabel(f"{row['full_name']}  ({row['code']})", box)
        name.setObjectName("DetailName")
        h.addWidget(name, 1)
        dob = QLabel(row["date_of_birth"] or "—", box)
        dob.setObjectName("Hint")
        h.addWidget(dob)
        if has_card:
            h.addWidget(_chip(box, "Ready", theme.PALETTE["--success"]))
        else:
            h.addWidget(_chip(box, "No card", theme.PALETTE["--danger"]))
        return box
